using CefSharp;
using CefSharp.DevTools.Page;
using CefSharp.WinForms;
using MimeKit;
using System.Text;

namespace EML2PDF6
{
    public partial class DefaultForm : Form
    {
        public DefaultForm()
        {
            InitializeComponent();
            string version = string.Format("Chromium: {0}, CEF: {1}, CefSharp: {2}", Cef.ChromiumVersion, Cef.CefVersion, Cef.CefSharpVersion);

            string environment = string.Format("Environment: {0}, Runtime: {1}",
                System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString().ToLowerInvariant(),
                System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription);

            toolStripStatusLabel1.Text = version + " " + environment;
        }

        private void OnBrowserLoadError(object sender, LoadErrorEventArgs e)
        {
            //Actions that trigger a download will raise an aborted error.
            //Aborted is generally safe to ignore
            if (e.ErrorCode == CefErrorCode.Aborted)
            {
                return;
            }

            string errorHtml =
                $"<html><body><h2>Failed to load URL {e.FailedUrl} with error {e.ErrorText} ({e.ErrorCode}).</h2></body></html>";

            _ = e.Browser.SetMainFrameDocumentContentAsync(errorHtml);

            //AddressChanged isn't called for failed Urls so we need to manually update the Url TextBox
            toolStripStatusLabel1.Text = e.FailedUrl;
        }


        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // Ensure the cache directory exists
            string cachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EML2PDF\\Cache");
            if (!Directory.Exists(cachePath))
            {
                Directory.CreateDirectory(cachePath);
            }

            // Initialize CefSharp
            CefSettings settings = new()
            {
                LogSeverity = LogSeverity.Disable, // Optional: Minimize logging
                CachePath = cachePath // Set the cache path
            };
            Cef.Initialize(settings);

            // Create and add the browser to the form
            using ChromiumWebBrowser browser = new("about:blank");
            this.Controls.Add(browser); // Add browser to the form's controls

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
                    SaveHtmlToPdf(htmlContent, "output.pdf", browser);
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
            // Load the .eml file
            MimeMessage? message = MimeMessage.Load(emlFilePath);

            // Get the HTML body
            TextPart? bodyPart = message.Body as TextPart;
            if (bodyPart == null && message.Body is Multipart multipart)
            {
                foreach (MimeEntity? part in multipart)
                {
                    if (part is TextPart { IsHtml: true } textPart)
                    {
                        bodyPart = textPart;
                        break;
                    }
                }
            }

            if (bodyPart == null)
            {
                Console.WriteLine("Failed to parse .eml file.");
                return null;
            }
            // Replace embedded images with data URIs
            string html = bodyPart.Text;
            if (message.Body is MultipartRelated multipartRelated)
            {
                foreach (MimeEntity? attachment in multipartRelated)
                {
                    if (attachment is MimePart { IsAttachment: true } imagePart)
                    {
                        using MemoryStream stream = new();
                        imagePart.Content.DecodeTo(stream);
                        string base64Image = Convert.ToBase64String(stream.ToArray());
                        string contentId = imagePart.ContentId;

                        html = html.Replace($"cid:{contentId}", $"data:{imagePart.ContentType.MimeType};base64,{base64Image}");
                    }
                }
            }

            return html;
        }

        static void SaveHtmlToPdf(string htmlContent, string outputPath, ChromiumWebBrowser browser)
        {
            //browser.IsBrowserInitializedChanged += async (s, args) =>
            //{
                if (browser.IsBrowserInitialized)
                {

                    browser.LoadingStateChanged += async (s, args) =>
                    {
                        if (!args.IsLoading)
                        {
                            PdfPrintSettings printSettings = new PdfPrintSettings();
                            printSettings.MarginTop = 100;   // 10mm
                            printSettings.MarginLeft  = 100; // 10mm
                            printSettings.PaperHeight = 29700; // A4
                            printSettings.PaperWidth = 21000;  // A4

                            bool success = await browser.PrintToPdfAsync(outputPath, printSettings);
                            if (success)
                            {
                                Console.WriteLine($"PDF saved to {outputPath}");
                            }
                            else
                            {
                                Console.WriteLine("Failed to save PDF.");
                            }
                        }
                    };

                    //string base64EncodedHtml = Convert.ToBase64String(Encoding.UTF8.GetBytes(htmlContent));
                    //browser.Load("data:text/html;base64," + base64EncodedHtml);
                    browser.LoadHtml(htmlContent);

            }
            //};
        }

    }
}
