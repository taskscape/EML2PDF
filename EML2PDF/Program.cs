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
            MimeMessage message = await MimeMessage.LoadAsync(emlFilePath);
            MimePart? pdfAttachment = message.Attachments
                .OfType<MimePart>()
                .FirstOrDefault(a => a.FileName != null && a.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase));
            if (pdfAttachment != null)
            {
                string outputFilePath = Path.ChangeExtension(emlFilePath, ".eml.pdf");
                if (File.Exists(outputFilePath))
                {
                    Log.Warning("Output file already exists: {outputFilePath}", outputFilePath);
                    Console.WriteLine($"File {outputFilePath} already exists.");
                    return 0;
                }
                Log.Information("Saving PDF attachment to {outputFilePath}", outputFilePath);
                await using (FileStream stream = File.Create(outputFilePath))
                {
                    await pdfAttachment.Content.DecodeToAsync(stream);
                }
                Log.Information("PDF attachment saved successfully: {outputFilePath}", outputFilePath);
                Console.WriteLine("RET-OUTPUT: " + outputFilePath);
                Console.WriteLine("PDF attachment extraction completed.");
                if (DeleteAfterwards)
                {
                    Log.Debug("DeleteAfterwards set to true, deleting {emlFilePath}", emlFilePath);
                    File.Delete(emlFilePath);
                }
                else
                {
                    string newPath = Path.Combine(
                        Path.GetDirectoryName(emlFilePath),
                        $"{DateTime.UtcNow:yyyyMMdd HHmm}_{Path.GetFileName(emlFilePath)}.bak");
                    Log.Debug("Moving original EML file to backup: {newPath}", newPath);
                    File.Move(emlFilePath, newPath);
                }
                await Log.CloseAndFlushAsync();
                return 0;
            }

            Log.Information("GetPDFFromAttachments is true, but no PDF attachment was found. Proceeding with HTML conversion.");
        }
            
        string htmlContent = ParseEmlToHtml(emlFilePath);
        Log.Information("Parsed EML file to HTML successfully. File: {file}", emlFilePath);

        if (!string.IsNullOrEmpty(htmlContent))
        {
            try
            {
                string outputFilePath = Path.ChangeExtension(emlFilePath, ".eml.pdf");
                if (File.Exists(outputFilePath))
                {
                    Log.Warning("Output file already exists: {outputFilePath}", outputFilePath);
                    Console.WriteLine($"File {outputFilePath} already exists.");
                    return 0;
                }
                Log.Information("Starting PDF generation for {outputFilePath}", outputFilePath);
                await SaveHtmlToPdf(htmlContent, outputFilePath);
                Log.Information("PDF generated successfully: {outputFilePath}", outputFilePath);
                Console.WriteLine("RET-OUTPUT: " + outputFilePath);
                Console.WriteLine("Rendering .eml file to PDF completed.");
                if (DeleteAfterwards)
                {
                    Log.Debug("DeleteAfterwards set to true, deleting {emlFilePath}", emlFilePath);
                    File.Delete(emlFilePath);
                }
                else
                {
                    string newPath = Path.Combine(
                        Path.GetDirectoryName(emlFilePath),
                        $"{DateTime.UtcNow:yyyyMMdd HHmm}_{Path.GetFileName(emlFilePath)}.bak");
                    Log.Debug("Moving original EML file to backup: {newPath}", newPath);
                    File.Move(emlFilePath, newPath);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error while saving HTML to PDF. File: {file}", emlFilePath);
                Console.WriteLine($"Error while saving HTML to PDF: {ex.Message}");
                await Log.CloseAndFlushAsync();
                return 1;
            }
        }
        else
        {
            Log.Error("Failed to parse .eml file: HTML content is empty. File: {file}", emlFilePath);
            Console.WriteLine("Failed to parse .eml file.");
            await Log.CloseAndFlushAsync();
            return 1;
        }
            
        Log.Information("Program completed successfully. File: {file}", emlFilePath);
        await Log.CloseAndFlushAsync();
        return 0;
    }

    /// <summary>
    /// Parse an EML file and return HTML content as a string
    /// </summary>
    private static string ParseEmlToHtml(string emlFilePath)
    {
        Log.Debug("Parsing EML file: {emlFilePath}", emlFilePath);
        MimeMessage message = MimeMessage.Load(emlFilePath);
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            
        TextPart? htmlPart = message.BodyParts
            .OfType<TextPart>()
            .FirstOrDefault(bp => bp.ContentType.MediaSubtype == "html");

        if (htmlPart != null)
        {
            Log.Debug("Found HTML part in EML file. File: {file}", emlFilePath);
            string htmlBody = DecodeEmailBody(htmlPart);
            htmlBody = ReplaceInlineImagesWithBase64(htmlBody, message.BodyParts);
            return htmlBody;
        }
            
        Log.Debug("No HTML part found, returning text body as <pre>. File: {file}", emlFilePath);
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
        Log.Debug("Downloaded browser");
        await using IBrowser browser = await Puppeteer.LaunchAsync(new LaunchOptions
        {
            Headless = true,
            Args = new[] { "--no-sandbox", "--disable-setuid-sandbox", "--disable-gpu" }
        });
        await using IPage page = await browser.NewPageAsync();
        Log.Debug("Browser launched and new page created");
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
        Log.Debug("Browser closed");
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
    private static string ReplaceInlineImagesWithBase64(
        string htmlBody,
        IEnumerable<MimeEntity> bodyParts)
    {
        Log.Debug("Replacing inline images with Base64 in HTML body");
        foreach (MimeEntity part in bodyParts)
        {
            if (part is MimePart mimePart
                && mimePart.ContentType.MediaType.Equals("image", StringComparison.OrdinalIgnoreCase))
            {
                string contentId = mimePart.ContentId?.Trim('<', '>');
                if (string.IsNullOrEmpty(contentId))
                {
                    Log.Debug("Skipping image with empty Content-ID");
                    continue;
                }

                using MemoryStream memStream = new MemoryStream();
                mimePart.Content.DecodeTo(memStream);
                byte[] imageBytes = memStream.ToArray();
                string base64Data = Convert.ToBase64String(imageBytes);

                string mime = mimePart.ContentType.MimeType;
                string dataUri = $"data:{mime};base64,{base64Data}";
                    
                htmlBody = htmlBody.Replace($"cid:{contentId}", dataUri);
                Log.Debug("Replaced inline image with CID: {contentId}", contentId);
            }
        }

        return htmlBody;
    }
}
