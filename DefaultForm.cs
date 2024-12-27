using System.Text;
using MimeKit;
using PuppeteerSharp;

namespace EML2PDF
{
    public partial class DefaultForm : Form
    {
        public DefaultForm()
        {
            InitializeComponent();

            string environment = string.Format("Environment: {0}, Runtime: {1}",
                System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString().ToLowerInvariant(),
                System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription);
            
            label1.Text =  environment;
        }
        
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // Prompt user to select an .eml file
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Email Files (*.eml)|*.eml",
                Title = "Select an EML File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string emlFilePath = openFileDialog.FileName;

                // Parse and render email content
                string htmlContent = ParseEmlToHtml(emlFilePath);

                if (!string.IsNullOrEmpty(htmlContent))
                {
                    SaveHtmlToPdf(htmlContent, "output.pdf");
                    Console.WriteLine("Rendering .eml file to PDF completed");
                }
                else
                {
                    Console.WriteLine("Failed to parse .eml file.");
                }
            }
        }

        static string ParseEmlToHtml(string emlFilePath)
        {
            var message = MimeMessage.Load(emlFilePath);
            string outputResourcesPath = "resources";
            Directory.CreateDirectory(outputResourcesPath);
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Access the HTML body or plain text body
            var htmlPart = message.BodyParts.OfType<TextPart>()
                .FirstOrDefault(bp => bp.ContentType.MediaSubtype == "html");

            string htmlBody = string.Empty;

            if (htmlPart != null)
            {
                // Decode the raw content of the HTML part
                htmlBody = DecodeEmailBody(htmlPart);

                // Replace CID references with local file paths
                htmlBody = ReplaceInlineImages(htmlBody, message.BodyParts, outputResourcesPath);
            }
            else
            {
                // Fallback to plain text if no HTML body is present
                var textPart = message.TextBody;
                htmlBody = $"<pre>{DecodeEmailBody(message.TextBody, "utf-8")}</pre>";
            }
            
            return htmlBody;
        }

        private static async Task SaveHtmlToPdf(string htmlContent, string outputPath)
        {
            if (string.IsNullOrWhiteSpace(htmlContent))
                throw new ArgumentException("HTML content cannot be null or empty.", nameof(htmlContent));

            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Output path cannot be null or empty.", nameof(outputPath));

            await new BrowserFetcher().DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();
            await page.SetContentAsync(htmlContent);
            await page.PdfAsync(outputPath);
            await page.CloseAsync();
        }

        static string DecodeEmailBody(TextPart textPart)
        {
            var charset = textPart.ContentType.Charset ?? "utf-8";

            using var memoryStream = new MemoryStream();
            textPart.Content.DecodeTo(memoryStream);
            var rawBytes = memoryStream.ToArray();

            try
            {
                var encoding = Encoding.GetEncoding(charset);
                return encoding.GetString(rawBytes);
            }
            catch
            {
                Console.WriteLine($"Warning: Failed to decode using charset {charset}. Falling back to ISO-8859-1.");
                return Encoding.GetEncoding("ISO-8859-1").GetString(rawBytes);
            }
        }

        static string ReplaceInlineImages(string htmlBody, IEnumerable<MimeEntity> bodyParts, string resourcesPath)
        {
            // Ensure the resources folder exists
            if (!Directory.Exists(resourcesPath))
            {
                Directory.CreateDirectory(resourcesPath);
                Console.WriteLine($"Created resources folder: {resourcesPath}");
            }

            foreach (var part in bodyParts)
            {
                if (part is MimePart mimePart && mimePart.ContentDisposition?.Disposition == ContentDisposition.Inline)
                {
                    // Get the content ID and remove angle brackets if present
                    string contentId = mimePart.ContentId?.Trim('<', '>');

                    if (!string.IsNullOrEmpty(contentId))
                    {
                        // Generate a unique file name for the image
                        string fileName = mimePart.FileName ?? Guid.NewGuid() + ".img";
                        string filePath = Path.Combine(resourcesPath, fileName);

                        // Save the image to the resources folder
                        try
                        {
                            using (var fileStream = File.Create(filePath))
                            {
                                mimePart.Content.DecodeTo(fileStream);
                            }
                            Console.WriteLine($"Saved resource: {filePath}");

                            // Replace the CID reference in the HTML with the local file path
                            htmlBody = htmlBody.Replace($"cid:{contentId}", fileName);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to save resource: {filePath}. Error: {ex.Message}");
                        }
                    }
                }
            }

            return htmlBody;
        }

        static string DecodeEmailBody(string rawBody, string charset)
        {
            try
            {
                var encoding = Encoding.GetEncoding(charset);
                return encoding.GetString(Encoding.Default.GetBytes(rawBody));
            }
            catch
            {
                return rawBody; // Return raw body if decoding fails
            }
        }
    }
}
