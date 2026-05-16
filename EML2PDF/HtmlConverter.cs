using System.Text;
using MimeKit;
using Serilog;

namespace EML2PDF;

internal static partial class Program
{
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
