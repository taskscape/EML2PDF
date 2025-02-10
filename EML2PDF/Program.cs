using System.Diagnostics;
using System.Text;
using Microsoft.Extensions.Configuration;
using MimeKit;
using PuppeteerSharp;
using Serilog;

namespace EML2PDF;

internal static class Program
{
    private static bool DeleteAfterwards { get; set; }
    private static bool GetPDFFromAttachments { get; set; }
    private static string SeqAppName { get; set; }
    private static string SeqAddress { get; set; }

    public static async Task<int> Main(string[] args)
    {
        string? exePath = Process.GetCurrentProcess().MainModule?.FileName;
        string realExeDirectory = Path.GetDirectoryName(exePath);

        IConfiguration config = new ConfigurationBuilder()
            .SetBasePath(realExeDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        SeqAddress = config["Seq:ServerAddress"] ?? string.Empty;
        SeqAppName = config["Seq:AppName"] ?? string.Empty;
        DeleteAfterwards = bool.Parse(config["DeleteFileAfterProcessing"] ?? "false");
        GetPDFFromAttachments = bool.Parse(config["GetPDFFromAttachments"] ?? "false");

        Log.Logger = new LoggerConfiguration()
            .Enrich.WithProperty("Application", SeqAppName)
            .MinimumLevel.Debug()
            .WriteTo.File(
                path: $"{realExeDirectory}/logs/log-.txt",
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 7,
                outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}"
            )
            .WriteTo.Seq(SeqAddress)
            .CreateLogger();

        Log.Information("Program start: EML2PDF conversion");
        Log.Debug("Executable directory: {realExeDirectory}", realExeDirectory);

        string? emlFilePath;

        if (args.Length > 0)
        {
            emlFilePath = args[0].Trim().Trim('"');
            Log.Debug("Received EML file path from arguments: {emlFilePath}", emlFilePath);
        }
        else
        {
            Console.Write("Please enter the path to the EML file (you may include quotes if it has spaces): ");
            string userInput = Console.ReadLine() ?? string.Empty;
            emlFilePath = userInput.Trim().Trim('"');
            Log.Debug("Received EML file path from console input: {emlFilePath}", emlFilePath);
        }

        if (string.IsNullOrWhiteSpace(emlFilePath) || !File.Exists(emlFilePath))
        {
            Log.Warning("Invalid file path: {emlFilePath}", emlFilePath);
            Console.WriteLine($"Invalid file path: \"{emlFilePath}\"");
            return 1;
        }
        
        if (GetPDFFromAttachments)
        {
            MimeMessage outerMessage = await MimeMessage.LoadAsync(emlFilePath);
            (MimePart? deepestPdf, int pdfDepth) = GetDeepestNestedPdf(outerMessage);
            if (deepestPdf != null)
            {
                string outputFilePath = Path.ChangeExtension(emlFilePath, ".eml.pdf");
                if (File.Exists(outputFilePath))
                {
                    Log.Warning("Output file already exists: {outputFilePath} for input file: {emlFilePath}", outputFilePath, emlFilePath);
                    Console.WriteLine($"File {outputFilePath} already exists.");
                    return 0;
                }
                Log.Information("Saving PDF attachment from nested attachment to {outputFilePath} for input file: {emlFilePath}", outputFilePath, emlFilePath);
                await using (FileStream stream = File.Create(outputFilePath))
                {
                    await deepestPdf.Content.DecodeToAsync(stream);
                }
                Log.Information("PDF attachment saved successfully: {outputFilePath} for input file: {emlFilePath}", outputFilePath, emlFilePath);
                Console.WriteLine("RET-OUTPUT: " + outputFilePath);
                Console.WriteLine("PDF attachment extraction completed.");
                if (DeleteAfterwards)
                {
                    Log.Debug("DeleteAfterwards set to true, deleting input file: {emlFilePath}", emlFilePath);
                    File.Delete(emlFilePath);
                }
                else
                {
                    string newPath = Path.Combine(
                        Path.GetDirectoryName(emlFilePath),
                        $"{DateTime.UtcNow:yyyyMMdd HHmm}_{Path.GetFileName(emlFilePath)}.bak");
                    Log.Debug("Moving original EML file to backup: {newPath} for input file: {emlFilePath}", newPath, emlFilePath);
                    File.Move(emlFilePath, newPath);
                }
                await Log.CloseAndFlushAsync();
                return 0;
            }
            Log.Information("GetPDFFromAttachments is true, but no PDF attachment was found for input file: {emlFilePath}. Proceeding with HTML conversion.", emlFilePath);
        }
        
        MimeMessage outerMsg = await MimeMessage.LoadAsync(emlFilePath);
        (MimeMessage deepestMessage, int depth) = GetDeepestNestedEml(outerMsg);
        MimeMessage messageToConvert = depth > 0 ? deepestMessage : outerMsg;
        if (depth > 0)
        {
            Log.Information("Found nested EML attachment(s) in input file: {emlFilePath}. Using the most deeply nested EML for conversion.", emlFilePath);
        }

        string htmlContent = ParseMimeMessageToHtml(messageToConvert);
        Log.Information("Parsed EML file to HTML successfully for input file: {emlFilePath}", emlFilePath);

        if (!string.IsNullOrEmpty(htmlContent))
        {
            try
            {
                string outputFilePath = Path.ChangeExtension(emlFilePath, ".eml.pdf");
                if (File.Exists(outputFilePath))
                {
                    Log.Warning("Output file already exists: {outputFilePath} for input file: {emlFilePath}", outputFilePath, emlFilePath);
                    Console.WriteLine($"File {outputFilePath} already exists.");
                    return 0;
                }
                Log.Information("Starting PDF generation for {outputFilePath} for input file: {emlFilePath}", outputFilePath, emlFilePath);
                await SaveHtmlToPdf(htmlContent, outputFilePath);
                Log.Information("PDF generated successfully: {outputFilePath} for input file: {emlFilePath}", outputFilePath, emlFilePath);
                Console.WriteLine("RET-OUTPUT: " + outputFilePath);
                Console.WriteLine("Rendering .eml file to PDF completed.");
                if (DeleteAfterwards)
                {
                    Log.Debug("DeleteAfterwards set to true, deleting input file: {emlFilePath}", emlFilePath);
                    File.Delete(emlFilePath);
                }
                else
                {
                    string newPath = Path.Combine(
                        Path.GetDirectoryName(emlFilePath),
                        $"{DateTime.UtcNow:yyyyMMdd HHmm}_{Path.GetFileName(emlFilePath)}.bak");
                    Log.Debug("Moving original EML file to backup: {newPath} for input file: {emlFilePath}", newPath, emlFilePath);
                    File.Move(emlFilePath, newPath);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error while saving HTML to PDF for input file: {emlFilePath}", emlFilePath);
                Console.WriteLine($"Error while saving HTML to PDF: {ex.Message}");
                await Log.CloseAndFlushAsync();
                return 1;
            }
        }
        else
        {
            Log.Error("Failed to parse .eml file: HTML content is empty for input file: {emlFilePath}", emlFilePath);
            Console.WriteLine("Failed to parse .eml file.");
            await Log.CloseAndFlushAsync();
            return 1;
        }

        Log.Information("Program completed successfully for input file: {emlFilePath}", emlFilePath);
        await Log.CloseAndFlushAsync();
        return 0;
    }

    /// <summary>
    /// Recursively traverses the attachments in the given MimeMessage and returns the deepest nested EML.
    /// Returns a tuple containing the deepest message and the depth (0 means no nested EML was found).
    /// </summary>
    private static (MimeMessage message, int depth) GetDeepestNestedEml(MimeMessage message)
    {
        MimeMessage bestMessage = message;
        int bestDepth = 0;

        foreach (MimeEntity? attachment in message.Attachments)
        {
            switch (attachment)
            {
                case MessagePart mp:
                {
                    (MimeMessage nestedMessage, int depth) = GetDeepestNestedEml(mp.Message);
                    int candidateDepth = depth + 1;
                    if (candidateDepth > bestDepth)
                    {
                        bestMessage = nestedMessage;
                        bestDepth = candidateDepth;
                    }
                    break;
                }
                case MimePart part when
                    !string.IsNullOrEmpty(part.FileName) &&
                    part.FileName.EndsWith(".eml", StringComparison.OrdinalIgnoreCase):
                {
                    using MemoryStream ms = new MemoryStream();
                    part.Content.DecodeTo(ms);
                    ms.Position = 0;
                    MimeMessage nestedMessage = MimeMessage.Load(ms);
                    (MimeMessage candidateMessage, int depth) = GetDeepestNestedEml(nestedMessage);
                    int candidateDepth = depth + 1;
                    if (candidateDepth > bestDepth)
                    {
                        bestMessage = candidateMessage;
                        bestDepth = candidateDepth;
                    }
                    break;
                }
            }
        }

        return (bestMessage, bestDepth);
    }

    /// <summary>
    /// Recursively traverses the attachments in the given MimeMessage and returns the deepest nested PDF attachment.
    /// Returns a tuple containing the PDF (or null if none found) and its depth (0 means no PDF was found).
    /// </summary>
    private static (MimePart? pdf, int depth) GetDeepestNestedPdf(MimeMessage message)
    {
        MimePart? bestPdf = null;
        int bestDepth = 0;

        foreach (MimeEntity? attachment in message.Attachments)
        {
            switch (attachment)
            {
                case MessagePart mp:
                {
                    (MimePart? nestedPdf, int depth) = GetDeepestNestedPdf(mp.Message);
                    int candidateDepth = depth + 1;
                    if (nestedPdf != null && candidateDepth > bestDepth)
                    {
                        bestPdf = nestedPdf;
                        bestDepth = candidateDepth;
                    }
                    break;
                }
                case MimePart part when
                    !string.IsNullOrEmpty(part.FileName) &&
                    part.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase):
                {
                    int candidateDepth = 1; // Direct attachment
                    if (candidateDepth > bestDepth)
                    {
                        bestPdf = part;
                        bestDepth = candidateDepth;
                    }
                    break;
                }
            }
        }

        return (bestPdf, bestDepth);
    }

    /// <summary>
    /// Converts a preloaded MimeMessage to HTML.
    /// </summary>
    private static string ParseMimeMessageToHtml(MimeMessage message)
    {
        Log.Debug("Parsing MimeMessage to HTML for input file.");
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        TextPart? htmlPart = message.BodyParts
            .OfType<TextPart>()
            .FirstOrDefault(bp => bp.ContentType.MediaSubtype.Equals("html", StringComparison.OrdinalIgnoreCase));

        if (htmlPart != null)
        {
            Log.Debug("Found HTML part in MimeMessage for input file.");
            string htmlBody = DecodeEmailBody(htmlPart);
            htmlBody = ReplaceInlineImagesWithBase64(htmlBody, message.BodyParts);
            return htmlBody;
        }

        Log.Debug("No HTML part found in MimeMessage for input file, returning text body wrapped in <pre>.");
        return $"<pre>{DecodeEmailBody(message.TextBody ?? string.Empty, "utf-8")}</pre>";
    }

    /// <summary>
    /// Saves the provided HTML content to a PDF file at the specified output path.
    /// Uses PuppeteerSharp in headless mode.
    /// </summary>
    private static async Task SaveHtmlToPdf(string htmlContent, string outputPath)
    {
        Log.Debug("SaveHtmlToPdf called with outputPath: {outputPath}", outputPath);
        if (string.IsNullOrWhiteSpace(htmlContent))
            throw new ArgumentException("HTML content cannot be null or empty.", nameof(htmlContent));

        if (string.IsNullOrWhiteSpace(outputPath))
            throw new ArgumentException("Output path cannot be null or empty.", nameof(outputPath));

        await new BrowserFetcher().DownloadAsync();
        Log.Debug("Downloaded browser for PDF conversion.");
        await using IBrowser browser = await Puppeteer.LaunchAsync(new LaunchOptions
        {
            Headless = true,
            Args = ["--no-sandbox", "--disable-setuid-sandbox", "--disable-gpu"]
        });
        await using IPage page = await browser.NewPageAsync();
        Log.Debug("Browser launched and new page created for PDF conversion.");
        await page.SetContentAsync(htmlContent);

        int bodyHeight = await page.EvaluateExpressionAsync<int>("document.body.scrollHeight");
        Log.Debug("Measured body height: {bodyHeight}", bodyHeight);
        await page.PdfAsync(outputPath, new PdfOptions
        {
            PrintBackground = true,
            Width = "8.5in",
            Height = $"{bodyHeight}px"
        });

        Log.Information("PDF file created at {outputPath}", outputPath);
        await browser.CloseAsync();
        Log.Debug("Browser closed after PDF conversion.");
    }

    /// <summary>
    /// Decodes a MimeKit TextPart into a string using its declared charset or falls back to ISO-8859-1.
    /// </summary>
    private static string DecodeEmailBody(TextPart textPart)
    {
        Log.Debug("Decoding TextPart with charset: {charset}", textPart.ContentType.Charset);
        string charset = textPart.ContentType.Charset ?? "utf-8";

        using MemoryStream memoryStream = new();
        textPart.Content.DecodeTo(memoryStream);
        byte[] rawBytes = memoryStream.ToArray();

        try
        {
            Encoding encoding = Encoding.GetEncoding(charset);
            Log.Debug("Decoded using {charset}", charset);
            return encoding.GetString(rawBytes);
        }
        catch
        {
            Log.Warning("Failed to decode using charset {charset}. Falling back to ISO-8859-1.", charset);
            return Encoding.GetEncoding("ISO-8859-1").GetString(rawBytes);
        }
    }

    /// <summary>
    /// Decodes a raw string body with the given charset. Falls back to the raw string if decoding fails.
    /// </summary>
    private static string DecodeEmailBody(string rawBody, string charset)
    {
        Log.Debug("Decoding raw string body with charset: {charset}", charset);
        try
        {
            Encoding encoding = Encoding.GetEncoding(charset);
            return encoding.GetString(Encoding.Default.GetBytes(rawBody));
        }
        catch
        {
            Log.Warning("Failed to decode raw body using charset {charset}. Returning raw body.", charset);
            return rawBody;
        }
    }

    /// <summary>
    /// Converts inline CID images to Base64 data URIs in the HTML body.
    /// </summary>
    private static string ReplaceInlineImagesWithBase64(string htmlBody, IEnumerable<MimeEntity> bodyParts)
    {
        Log.Debug("Replacing inline images with Base64 in HTML body for input file.");
        foreach (MimeEntity part in bodyParts)
        {
            if (part is MimePart mimePart &&
                mimePart.ContentType.MediaType.Equals("image", StringComparison.OrdinalIgnoreCase))
            {
                string contentId = mimePart.ContentId?.Trim('<', '>') ?? string.Empty;
                if (string.IsNullOrEmpty(contentId))
                {
                    Log.Debug("Skipping image with empty Content-ID for input file.");
                    continue;
                }

                using MemoryStream memStream = new();
                mimePart.Content.DecodeTo(memStream);
                byte[] imageBytes = memStream.ToArray();
                string base64Data = Convert.ToBase64String(imageBytes);

                string mime = mimePart.ContentType.MimeType;
                string dataUri = $"data:{mime};base64,{base64Data}";

                htmlBody = htmlBody.Replace($"cid:{contentId}", dataUri);
                Log.Debug("Replaced inline image with CID: {contentId} for input file.", contentId);
            }
        }

        return htmlBody;
    }
}