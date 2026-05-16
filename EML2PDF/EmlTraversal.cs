using MimeKit;

namespace EML2PDF;

internal static partial class Program
{
    /// <summary>
    /// Recursively traverses the attachments in the given MimeMessage and returns the deepest nested EML.
    /// Returns a tuple containing the deepest message and the depth (0 means no nested EML was found).
    /// </summary>
    private static (MimeMessage message, int depth) GetDeepestNestedEml(MimeMessage message)
    {
        MimeMessage bestMessage = message;
        int bestDepth = 0;

        foreach (MimeEntity? attachment in message.Attachments)
        {
            switch (attachment)
            {
                case MessagePart mp:
                {
                    (MimeMessage nestedMessage, int depth) = GetDeepestNestedEml(mp.Message);
                    int candidateDepth = depth + 1;
                    if (candidateDepth > bestDepth)
                    {
                        bestMessage = nestedMessage;
                        bestDepth = candidateDepth;
                    }
                    break;
                }
                case MimePart part when
                    !string.IsNullOrEmpty(part.FileName) &&
                    part.FileName.EndsWith(".eml", StringComparison.OrdinalIgnoreCase):
                {
                    using MemoryStream ms = new MemoryStream();
                    part.Content.DecodeTo(ms);
                    ms.Position = 0;
                    MimeMessage nestedMessage = MimeMessage.Load(ms);
                    (MimeMessage candidateMessage, int depth) = GetDeepestNestedEml(nestedMessage);
                    int candidateDepth = depth + 1;
                    if (candidateDepth > bestDepth)
                    {
                        bestMessage = candidateMessage;
                        bestDepth = candidateDepth;
                    }
                    break;
                }
            }
        }

        return (bestMessage, bestDepth);
    }

    /// <summary>
    /// Recursively traverses the attachments in the given MimeMessage and returns all PDF attachments
    /// found at the deepest nesting level, along with that depth.
    /// Returns an empty list if no PDFs are found anywhere.
    /// </summary>
    private static (List<MimePart> pdfs, int depth) GetDeepestNestedPdfs(MimeMessage message)
    {
        List<MimePart> bestPdfs = [];
        int bestDepth = 0;

        foreach (MimeEntity? attachment in message.Attachments)
        {
            switch (attachment)
            {
                case MessagePart mp:
                {
                    (List<MimePart> nestedPdfs, int nestedDepth) = GetDeepestNestedPdfs(mp.Message);
                    int candidateDepth = nestedDepth + 1;
                    if (nestedPdfs.Count > 0)
                    {
                        if (candidateDepth > bestDepth)
                        {
                            bestPdfs = nestedPdfs;
                            bestDepth = candidateDepth;
                        }
                        else if (candidateDepth == bestDepth)
                        {
                            bestPdfs.AddRange(nestedPdfs);
                        }
                    }
                    break;
                }
                case MimePart part when
                    !string.IsNullOrEmpty(part.FileName) &&
                    part.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase):
                {
                    const int candidateDepth = 1;
                    if (candidateDepth > bestDepth)
                    {
                        bestPdfs = [part];
                        bestDepth = candidateDepth;
                    }
                    else if (candidateDepth == bestDepth)
                    {
                        bestPdfs.Add(part);
                    }
                    break;
                }
            }
        }

        return (bestPdfs, bestDepth);
    }
}
