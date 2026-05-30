using System.Diagnostics;
using Microsoft.Extensions.Configuration;
using MimeKit;
using Serilog;

namespace EML2PDF;

internal static partial class Program
{
    private static string SeqApiKey { get; set; }
    private static string SeqAddress { get; set; }

    public static async Task<int> Main(string[] args)
    {
        try
        {
            string? exePath = Process.GetCurrentProcess().MainModule?.FileName;
            string realExeDirectory = Path.GetDirectoryName(exePath);

            IConfiguration config = new ConfigurationBuilder()
                .SetBasePath(realExeDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            SeqAddress = config["Seq:ServerAddress"] ?? string.Empty;
            SeqApiKey = config["Seq:ApiKey"] ?? string.Empty;
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File(
                    path: $"{realExeDirectory}/logs/log-.txt",
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 7,
                    outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}"
                )
                .WriteTo.Seq(SeqAddress, apiKey: string.IsNullOrEmpty(SeqApiKey) ? null : SeqApiKey)
                .CreateLogger();

            Log.Information("Program start: EML2PDF conversion");
            Log.Debug("Executable directory: {realExeDirectory}", realExeDirectory);

            string? inputPath;

            if (args.Length > 0)
            {
                inputPath = args[0].Trim().Trim('"');
                Log.Debug("Received input path from arguments: {inputPath}", inputPath);
            }
            else
            {
                Console.Write("Please enter the path to an EML file or a directory (you may include quotes if it has spaces): ");
                string userInput = Console.ReadLine() ?? string.Empty;
                inputPath = userInput.Trim().Trim('"');
                Log.Debug("Received input path from console input: {inputPath}", inputPath);
            }

            bool isDirectory = !string.IsNullOrWhiteSpace(inputPath) && Directory.Exists(inputPath);
            bool isFile = !string.IsNullOrWhiteSpace(inputPath) && File.Exists(inputPath);

            if (!isDirectory && !isFile)
            {
                Log.Warning("Invalid input path (not an existing file or directory): {inputPath}", inputPath);
                Console.WriteLine($"Invalid path: \"{inputPath}\"");
                return 1;
            }

            List<string> emlFiles;
            string inputType;
            if (isDirectory)
            {
                inputType = "directory";
                emlFiles = Directory
                    .EnumerateFiles(inputPath, "*.eml", SearchOption.TopDirectoryOnly)
                    .OrderBy(static p => p, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                Log.Information("Processing directory {inputPath}. Found {count} .eml file(s).", inputPath, emlFiles.Count);
                Console.WriteLine($"Found {emlFiles.Count} .eml file(s) in directory: {inputPath}");
            }
            else
            {
                inputType = "file";
                emlFiles = [inputPath];
            }

            Stopwatch stopwatch = Stopwatch.StartNew();
            int processed = 0;
            int succeeded = 0;
            int failed = 0;

            foreach (string emlFilePath in emlFiles)
            {
                processed++;
                bool ok = await ProcessEmlFileAsync(emlFilePath);
                if (ok)
                {
                    succeeded++;
                }
                else
                {
                    failed++;
                }
            }

            stopwatch.Stop();

            Log.Information(
                "Processing summary: input {inputType} {inputPath} completed in {elapsed} ({elapsedMs} ms). " +
                "Files processed: {processed}, succeeded: {succeeded}, failed: {failed}.",
                inputType, inputPath, stopwatch.Elapsed, stopwatch.ElapsedMilliseconds, processed, succeeded, failed);
            Console.WriteLine(
                $"Done. Processed {processed} file(s) in {stopwatch.Elapsed}. Succeeded: {succeeded}, Failed: {failed}.");

            return failed > 0 ? 1 : 0;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Unhandled fatal error");
            Console.WriteLine($"Fatal error: {ex.Message}");
            return 1;
        }
        finally
        {
            await Log.CloseAndFlushAsync();
        }
    }

    private static async Task<bool> ProcessEmlFileAsync(string emlFilePath)
    {
        try
        {
            MimeMessage outerMessage = await MimeMessage.LoadAsync(emlFilePath);
            (List<MimePart> pdfAttachments, int _) = GetDeepestNestedPdfs(outerMessage);

            if (pdfAttachments.Count > 0)
            {
                Log.Information("Found {count} PDF attachment(s) in input file: {emlFilePath}. Extracting.", pdfAttachments.Count, emlFilePath);
                string? emlDirectory = Path.GetDirectoryName(emlFilePath);

                string emlBaseName = Path.GetFileNameWithoutExtension(emlFilePath);
                for (int i = 0; i < pdfAttachments.Count; i++)
                {
                    string outputFilePath = Path.Combine(emlDirectory, $"{emlBaseName}.{i + 1:D3}.pdf");

                    if (File.Exists(outputFilePath))
                    {
                        Log.Warning("Output file already exists: {outputFilePath} for input file: {emlFilePath}", outputFilePath, emlFilePath);
                        Console.WriteLine($"File {outputFilePath} already exists.");
                        continue;
                    }

                    Log.Information("Saving PDF attachment to {outputFilePath} for input file: {emlFilePath}", outputFilePath, emlFilePath);
                    await using (FileStream stream = File.Create(outputFilePath))
                    {
                        await pdfAttachments[i].Content.DecodeToAsync(stream);
                    }
                    Log.Information("PDF attachment saved successfully: {outputFilePath} for input file: {emlFilePath}", outputFilePath, emlFilePath);
                    Console.WriteLine("RET-OUTPUT: " + outputFilePath);
                }

                Console.WriteLine("PDF attachment extraction completed.");
                string backupPath1 = Path.Combine(
                    Path.GetDirectoryName(emlFilePath),
                    $"{DateTime.UtcNow:yyyyMMddHHmmss} - {Path.GetFileName(emlFilePath)}.bak");
                Log.Debug("Moving original EML file to backup: {backupPath1} for input file: {emlFilePath}", backupPath1, emlFilePath);
                File.Move(emlFilePath, backupPath1);
                return true;
            }

            Log.Information("No PDF attachments found in input file: {emlFilePath}. Proceeding with HTML body conversion.", emlFilePath);

            (MimeMessage deepestMessage, int depth) = GetDeepestNestedEml(outerMessage);
            MimeMessage messageToConvert = depth > 0 ? deepestMessage : outerMessage;
            if (depth > 0)
            {
                Log.Information("Found nested EML attachment(s) in input file: {emlFilePath}. Using the most deeply nested EML for conversion.", emlFilePath);
            }

            string htmlContent = ParseMimeMessageToHtml(messageToConvert);
            Log.Information("Parsed EML file to HTML successfully for input file: {emlFilePath}", emlFilePath);

            if (string.IsNullOrEmpty(htmlContent))
            {
                Log.Error("Failed to parse .eml file: HTML content is empty for input file: {emlFilePath}", emlFilePath);
                Console.WriteLine("Failed to parse .eml file.");
                return false;
            }

            string pdfOutputPath = Path.ChangeExtension(emlFilePath, ".eml.pdf");
            if (File.Exists(pdfOutputPath))
            {
                Log.Warning("Output file already exists: {pdfOutputPath} for input file: {emlFilePath}", pdfOutputPath, emlFilePath);
                Console.WriteLine($"File {pdfOutputPath} already exists.");
                return true;
            }
            Log.Information("Starting PDF generation for {pdfOutputPath} for input file: {emlFilePath}", pdfOutputPath, emlFilePath);
            using CancellationTokenSource renderCts = new(TimeSpan.FromMinutes(10));
            await SaveHtmlToPdf(htmlContent, pdfOutputPath, renderCts.Token);
            Log.Information("PDF generated successfully: {pdfOutputPath} for input file: {emlFilePath}", pdfOutputPath, emlFilePath);
            Console.WriteLine("RET-OUTPUT: " + pdfOutputPath);
            Console.WriteLine("Rendering .eml file to PDF completed.");
            string backupPath2 = Path.Combine(
                Path.GetDirectoryName(emlFilePath),
                $"{DateTime.UtcNow:yyyyMMddHHmmss} - {Path.GetFileName(emlFilePath)}.bak");
            Log.Debug("Moving original EML file to backup: {backupPath2} for input file: {emlFilePath}", backupPath2, emlFilePath);
            File.Move(emlFilePath, backupPath2);

            Log.Information("Program completed successfully for input file: {emlFilePath}", emlFilePath);
            return true;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to process input file: {emlFilePath}", emlFilePath);
            Console.WriteLine($"Error processing \"{emlFilePath}\": {ex.Message}");
            return false;
        }
    }
}
