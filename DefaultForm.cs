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
            var version = string.Format("Chromium: {0}, CEF: {1}, CefSharp: {2}",
Cef.ChromiumVersion, Cef.CefVersion, Cef.CefSharpVersion);

            var environment = string.Format("Environment: {0}, Runtime: {1}",
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

            var errorHtml = string.Format("<html><body><h2>Failed to load URL {0} with error {1} ({2}).</h2></body></html>",
                                              e.FailedUrl, e.ErrorText, e.ErrorCode);

            _ = e.Browser.SetMainFrameDocumentContentAsync(errorHtml);

            //AddressChanged isn't called for failed Urls so we need to manually update the Url TextBox
            toolStripStatusLabel1.Text = e.FailedUrl;
        }


        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // Initialize CefSharp
            var settings = new CefSettings
            {
                LogSeverity = LogSeverity.Disable, // Optional: Minimize logging
                CachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CefSharp\\Cache")
            };
            Cef.Initialize(settings);

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
            var message = MimeMessage.Load(emlFilePath);

            // Get the HTML body
            var bodyPart = message.Body as TextPart;
            if (bodyPart == null && message.Body is Multipart multipart)
            {
                foreach (var part in multipart)
                {
                    if (part is TextPart textPart && textPart.IsHtml)
                    {
                        bodyPart = textPart;
                        break;
                    }
                }
            }

            if (bodyPart == null) return null;

            // Replace embedded images with data URIs
            string html = bodyPart.Text;
            if (message.Body is MultipartRelated multipartRelated)
            {
                foreach (var attachment in multipartRelated)
                {
                    if (attachment is MimePart imagePart && imagePart.IsAttachment)
                    {
                        using var stream = new MemoryStream();
                        imagePart.Content.DecodeTo(stream);
                        string base64Image = Convert.ToBase64String(stream.ToArray());
                        string contentId = imagePart.ContentId;

                        html = html.Replace($"cid:{contentId}", $"data:{imagePart.ContentType.MimeType};base64,{base64Image}");
                    }
                }
            }

            return html;
        }

        static void SaveHtmlToPdf(string htmlContent, string outputPath)
        {

            using (var browser = new ChromiumWebBrowser("about:blank"))
            {
                // browser.CreateControl();


                browser.IsBrowserInitializedChanged += (sender, e) =>
                {
                    if (browser.IsBrowserInitialized)
                    {
                        var base64EncodedHtml = Convert.ToBase64String(Encoding.UTF8.GetBytes(htmlContent));
                        browser.Load("data:text/html;base64," + base64EncodedHtml);
                        //browser.LoadHtml(htmlContent);

                        browser.LoadingStateChanged += async (s, args) =>
                        {
                            if (!args.IsLoading)
                            {
                                var printSettings = new PdfPrintSettings();
                                printSettings.MarginTop = 100;   // 10mm
                                printSettings.MarginLeft  = 100; // 10mm
                                printSettings.PaperHeight = 29700; // A4
                                printSettings.PaperWidth = 21000;  // A4

                                var success = await browser.PrintToPdfAsync(outputPath, printSettings);
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
                    }
                };
            }

        }

    }
}
