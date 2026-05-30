# EML2PDF

EML2PDF is a .NET CLI utility that converts an email file (`.eml`) into a PDF.

## What the solution currently does

At runtime, the app:

1. Loads settings from `appsettings.json` located next to the executable.
2. Configures Serilog with:
	- rolling local file logs (`logs/log-*.txt`, daily rolling, last 7 files retained)
	- Seq sink (address and optional API key from config)
	- logs are flushed via `Log.CloseAndFlushAsync()` in a `finally` block at the end of every run
3. Reads the input path either:
	- from the first CLI argument, or
	- from console prompt input.
4. Validates that the path is an existing file or directory:
	- If it is a single `.eml` file, that file is processed.
	- If it is a directory, every top-level `*.eml` file in it is processed (non-recursive). A failure on one file is logged and counted but does not stop the rest of the batch.
5. For each file, searches for PDF attachments at the deepest nesting level:
	- If any PDF attachments are found, all of them are extracted and saved to the same folder, using each attachment's original filename.
	- If no PDF attachments are found, the deepest nested `.eml` body is rendered via headless Chromium (PuppeteerSharp) and exported to PDF as `<originalName>.eml.pdf`.
6. Writes the output in the same folder as input using `<originalName>.eml.pdf`.
7. If successful, renames the source `.eml` to `<yyyyMMddHHmmss> - <name>.eml.bak` in the same folder (UTC timestamp, no spaces, seconds precision).
8. When the run finishes, writes a summary log entry with the input type (file/directory), total elapsed processing time, and the number of files processed, succeeded, and failed.

## Usage

Download the latest release package. It contains the executable and `appsettings.json`.

Run one of:

- Interactive mode:

  ```bash
  EML2PDF.exe
  ```

- Direct argument mode (single file):

  ```bash
  EML2PDF.exe "C:\path\to\mail.eml"
  ```

- Direct argument mode (directory — processes every top-level `.eml` file):

  ```bash
  EML2PDF.exe "C:\path\to\folder"
  ```

## Configuration

`appsettings.json` keys currently used by the app:

- `Seq:ServerAddress`: Seq endpoint URL.
- `Seq:ApiKey`: API key used to authenticate with Seq. Leave empty to send events without authentication.

Example:

```json
{
  "Seq": {
	 "ServerAddress": "http://localhost:5341/",
	 "ApiKey": ""
  }
}
```

## Output and Exit Codes

- Success output: a line prefixed with `RET-OUTPUT: <pdfPath>`.
- When multiple PDF attachments are extracted, one `RET-OUTPUT:` line is printed per file.
- On completion a summary line is printed to the console and logged, e.g. `Done. Processed 5 file(s) in 00:00:12.3456789. Succeeded: 4, Failed: 1.`
- Exit code `0`: all processed files succeeded (or their output already existed).
- Exit code `1`: invalid input path, at least one file failed, or a runtime error occurred.

## Known resilience gaps (current implementation)

The following issues are currently present in code and should be considered known limitations:

1. ~~Startup errors are not globally guarded.~~ Fixed: all of `Main` is wrapped in a single `try/catch(Exception)/finally`; `CloseAndFlushAsync` is called unconditionally in the `finally` block.
2. Seq sink is always enabled when logger is built.
	- Bad/missing Seq endpoint can cause logging pipeline failures depending on environment/config.
4. `BrowserFetcher().DownloadAsync()` runs for every conversion.
	- Network or filesystem failures can break processing; cold-start latency is high.
5. Output file name is fixed to `<input>.eml.pdf`.
	- Existing file causes an early return instead of deterministic overwrite/versioning policy.
6. Source file cleanup (delete/move) is not wrapped in local error handling.
	- Permission/lock issues can fail after a successful conversion.
7. Backup naming has minute-level timestamp resolution.
	- Multiple runs in same minute can collide on `.bak` name.
8. MIME traversal/decoding assumes all nested payloads are valid.
	- Corrupt nested `.eml` or malformed parts can throw and abort.
9. HTML rendering does not sanitize/contain remote resources.
	- Embedded/linked assets may fail to load or behave inconsistently in headless rendering.
10. PDF page sizing uses `document.body.scrollHeight` as a single page height.
	 - Very long messages can create oversized single-page PDFs and stress rendering.
11. No cancellation or timeout boundaries.
	 - Stalls in MIME parsing, browser startup, network download, or PDF generation can hang the process.

## Dependencies

- .NET 10
- MimeKit
- PuppeteerSharp (Chromium download and headless rendering)
- Serilog + File sink + Seq sink