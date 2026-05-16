using System.Diagnostics;
using System.Runtime.InteropServices;
using Xunit;

namespace EML2PDF.Tests;

public class EmlConversionIntegrationTests
{
    private static readonly string SolutionRoot;
    private static readonly string DataDirectory;
    private static readonly string EmlExePath;

    static EmlConversionIntegrationTests()
    {
        // AppContext.BaseDirectory: .../{Project}/bin/{Config}/{TFM}/
        string baseDir = AppContext.BaseDirectory.TrimEnd(
            Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        SolutionRoot  = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", ".."));
        DataDirectory = Path.Combine(SolutionRoot, "Data");

        // Extract build config and TFM from the test binary path so the derived
        // EML2PDF exe path matches the same build (Debug/Release, net10.0, etc.).
        string[] parts = baseDir.Split(Path.DirectorySeparatorChar);
        string tfm    = parts[^1];  // e.g. net10.0
        string config = parts[^2];  // e.g. Debug
        string exeName = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "EML2PDF.exe" : "EML2PDF";
        EmlExePath = Path.Combine(SolutionRoot, "EML2PDF", "bin", config, tfm, exeName);
    }

    /// <summary>
    /// Returns one test case per .eml file found in the Data directory.
    /// Yields nothing if the directory does not exist, which causes the theory to be skipped.
    /// </summary>
    public static IEnumerable<object[]> EmlFiles()
    {
        if (!Directory.Exists(DataDirectory))
            yield break;

        foreach (string file in Directory.GetFiles(DataDirectory, "*.eml"))
            yield return new object[] { file };
    }

    [Theory]
    [MemberData(nameof(EmlFiles))]
    public async Task Convert_ProducesOutputPdf_AndRestoresOriginalEml(string emlPath)
    {
        Assert.True(File.Exists(EmlExePath),
            $"EML2PDF executable not found at '{EmlExePath}'. " +
            "Ensure the full solution is built before running tests.");

        Assert.True(File.Exists(emlPath),
            $"EML file not found: '{emlPath}'.");

        string dataDir     = Path.GetDirectoryName(emlPath)!;
        string emlFileName = Path.GetFileName(emlPath);
        string baseName    = Path.GetFileNameWithoutExtension(emlPath);

        try
        {
            // ── Run conversion ───────────────────────────────────────────────
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName               = EmlExePath,
                    Arguments              = $"\"{emlPath}\"",
                    UseShellExecute        = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError  = true,
                    CreateNoWindow         = true
                }
            };

            process.Start();
            string stdout = await process.StandardOutput.ReadToEndAsync();
            string stderr = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            // ── Assertions ───────────────────────────────────────────────────
            Assert.True(process.ExitCode == 0,
                $"EML2PDF exited with code {process.ExitCode} for '{emlFileName}'." +
                $"{Environment.NewLine}stdout: {stdout}" +
                $"{Environment.NewLine}stderr: {stderr}");

            // Attachment PDFs  → {baseName}.001.pdf, {baseName}.002.pdf …  (3-char counter)
            // Body render PDF  → {baseName}.eml.pdf                        ("eml" = 3 chars)
            // Both are covered by the single wildcard {baseName}.???.pdf
            string[] outputPdfs = Directory.GetFiles(dataDir, $"{baseName}.???.pdf");
            Assert.True(outputPdfs.Length > 0,
                $"No output PDF produced for '{emlFileName}'." +
                $"{Environment.NewLine}stdout: {stdout}");
        }
        finally
        {
            // ── Cleanup: delete output PDFs ──────────────────────────────────
            foreach (string pdf in Directory.GetFiles(dataDir, $"{baseName}.???.pdf"))
                File.Delete(pdf);

            // ── Cleanup: restore the original .eml from its timestamped backup ──
            // Backup name pattern: "{yyyyMMddHHmmss} - {emlFileName}.bak"
            foreach (string bak in Directory.GetFiles(dataDir, $"* - {emlFileName}.bak"))
                File.Move(bak, emlPath, overwrite: true);
        }
    }
}
