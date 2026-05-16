using PuppeteerSharp;
using Serilog;

namespace EML2PDF;

internal static partial class Program
{
    /// <summary>
    /// Saves the provided HTML content to a PDF file at the specified output path.
    /// Uses PuppeteerSharp in headless mode.
    /// </summary>
    private static async Task SaveHtmlToPdf(string htmlContent, string outputPath, CancellationToken cancellationToken = default)
    {
        Log.Debug("SaveHtmlToPdf called with outputPath: {outputPath}", outputPath);
        if (string.IsNullOrWhiteSpace(htmlContent))
            throw new ArgumentException("HTML content cannot be null or empty.", nameof(htmlContent));

        if (string.IsNullOrWhiteSpace(outputPath))
            throw new ArgumentException("Output path cannot be null or empty.", nameof(outputPath));

        await new BrowserFetcher().DownloadAsync().WaitAsync(cancellationToken);
        Log.Debug("Downloaded browser for PDF conversion.");
        await using IBrowser browser = await Puppeteer.LaunchAsync(new LaunchOptions
        {
            Headless = true,
            Args = ["--no-sandbox", "--disable-setuid-sandbox", "--disable-gpu"]
        }).WaitAsync(cancellationToken);
        await using IPage page = await browser.NewPageAsync().WaitAsync(cancellationToken);
        Log.Debug("Browser launched and new page created for PDF conversion.");
        await page.SetContentAsync(htmlContent).WaitAsync(cancellationToken);

        int bodyHeight = await page.EvaluateExpressionAsync<int>("document.body.scrollHeight").WaitAsync(cancellationToken);
        Log.Debug("Measured body height: {bodyHeight}", bodyHeight);
        await page.PdfAsync(outputPath, new PdfOptions
        {
            PrintBackground = true,
            Width = "8.5in",
            Height = $"{bodyHeight}px"
        }).WaitAsync(cancellationToken);

        Log.Information("PDF file created at {outputPath}", outputPath);
        await browser.CloseAsync().WaitAsync(cancellationToken);
        Log.Debug("Browser closed after PDF conversion.");
    }
}
