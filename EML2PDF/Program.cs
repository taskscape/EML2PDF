using System.Text;
using MimeKit;
using PuppeteerSharp;

namespace EML2PDF;

internal static class Program
{
    public static async Task<int> Main(string[] args)
    {
        string? emlFilePath;
            
        if (args.Length > 0)
        {
            emlFilePath = args[0].Trim().Trim('"');
        }
        else
        {
            Console.Write("Please enter the path to the EML file (you may include quotes if it has spaces): ");
            string userInput = Console.ReadLine() ?? string.Empty;
            emlFilePath = userInput.Trim().Trim('"');
        }
            
        if (string.IsNullOrWhiteSpace(emlFilePath) || !File.Exists(emlFilePath))
        {
            Console.WriteLine($"Invalid file path: \"{emlFilePath}\"");
            return 1;
        }
            
        string htmlContent = ParseEmlToHtml(emlFilePath);

        if (!string.IsNullOrEmpty(htmlContent))
        {
            try
            {
                string outputFilePath = Path.ChangeExtension(emlFilePath, ".pdf");
                if (File.Exists(outputFilePath))
                {
                    Console.WriteLine($"File {outputFilePath} already exists.");
                    return 0;
                }
                await SaveHtmlToPdf(htmlContent, outputFilePath);
                Console.WriteLine("Rendering .eml file to PDF completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while saving HTML to PDF: {ex.Message}");
                return 1;
            }
        }
        else
        {
            Console.WriteLine("Failed to parse .eml file.");
            return 1;
        }
            
        return 0;
    }

    /// <summary>
    /// Parse an EML file and return HTML content as a string
    /// </summary>
    private static string ParseEmlToHtml(string emlFilePath)
    {
        MimeMessage message = MimeMessage.Load(emlFilePath);
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            
        TextPart? htmlPart = message.BodyParts
            .OfType<TextPart>()
            .FirstOrDefault(bp => bp.ContentType.MediaSubtype == "html");

        if (htmlPart != null)
        {
            string htmlBody = DecodeEmailBody(htmlPart);
            htmlBody = ReplaceInlineImagesWithBase64(htmlBody, message.BodyParts);
            return htmlBody;
        }
            
        return $"<pre>{DecodeEmailBody(message.TextBody ?? string.Empty, "utf-8")}</pre>";
    }

    /// <summary>
    /// Saves the provided HTML content to a PDF file at the specified output path.
    /// Uses PuppeteerSharp in headless mode.
    /// </summary>
    private static async Task SaveHtmlToPdf(string htmlContent, string outputPath)
    {
        if (string.IsNullOrWhiteSpace(htmlContent))
            throw new ArgumentException("HTML content cannot be null or empty.", nameof(htmlContent));

        if (string.IsNullOrWhiteSpace(outputPath))
            throw new ArgumentException("Output path cannot be null or empty.", nameof(outputPath));
            
        await new BrowserFetcher().DownloadAsync();
        await using IBrowser browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
        await using IPage page = await browser.NewPageAsync();
        await page.SetContentAsync(htmlContent);
            
        int bodyHeight = await page.EvaluateExpressionAsync<int>("document.body.scrollHeight");
        await page.PdfAsync(outputPath, new PdfOptions
        {
            PrintBackground = true,
            Width = "8.5in",
            Height = $"{bodyHeight}px"
        });
        
        await browser.CloseAsync();
    }

    /// <summary>
    /// Decodes a MimeKit TextPart into a string using its declared charset or falls back to ISO-8859-1.
    /// </summary>
    private static string DecodeEmailBody(TextPart textPart)
    {
        string charset = textPart.ContentType.Charset ?? "utf-8";

        using MemoryStream memoryStream = new();
        textPart.Content.DecodeTo(memoryStream);
        byte[] rawBytes = memoryStream.ToArray();

        try
        {
            Encoding encoding = Encoding.GetEncoding(charset);
            return encoding.GetString(rawBytes);
        }
        catch
        {
            Console.WriteLine($"Warning: Failed to decode using charset {charset}. Falling back to ISO-8859-1.");
            return Encoding.GetEncoding("ISO-8859-1").GetString(rawBytes);
        }
    }

    /// <summary>
    /// Decodes a raw string body with the given charset. Falls back to the raw string if decoding fails.
    /// </summary>
    private static string DecodeEmailBody(string rawBody, string charset)
    {
        try
        {
            Encoding encoding = Encoding.GetEncoding(charset);
            return encoding.GetString(Encoding.Default.GetBytes(rawBody));
        }
        catch
        {
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
        foreach (MimeEntity part in bodyParts)
        {
            if (part is MimePart mimePart
                && mimePart.ContentType.MediaType.Equals("image", StringComparison.OrdinalIgnoreCase))
            {
                string contentId = mimePart.ContentId?.Trim('<', '>');
                if (string.IsNullOrEmpty(contentId))
                    continue;

                using MemoryStream memStream = new MemoryStream();
                mimePart.Content.DecodeTo(memStream);
                byte[] imageBytes = memStream.ToArray();
                string base64Data = Convert.ToBase64String(imageBytes);

                string mime = mimePart.ContentType.MimeType;
                string dataUri = $"data:{mime};base64,{base64Data}";
                    
                htmlBody = htmlBody.Replace($"cid:{contentId}", dataUri);
            }
        }

        return htmlBody;
    }
}