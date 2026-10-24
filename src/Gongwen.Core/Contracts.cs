namespace Gongwen.Core;

public enum ParagraphRole { Preserve, Title, Recipient, Body, Heading1, Heading2, Heading3, Heading4, AttachmentNote, Signature, Date }
public sealed record ParagraphInfo(int Index, string Text, ParagraphRole SuggestedRole, bool NeedsReview, string Reason, bool CanEdit = false, bool HasAutomaticNumbering = false);
public sealed record DocumentAnalysis(string SourcePath, string SourceSha256, IReadOnlyList<ParagraphInfo> Paragraphs, IReadOnlyList<string> Warnings, int TableCount, int ImageCount, bool CanFormat, string? BlockingReason, string? SuggestedTemplateId = null, bool RequiresStructureReview = false, bool HasEditProtection = false);
public sealed record FormatTemplate(string Id, string Name, string TitleFont = "方正小标宋简体", string BodyFont = "仿宋", string Heading1Font = "黑体", string Heading2Font = "楷体", string LatinFont = "Times New Roman", double TopMarginCm = 3.7, double BottomMarginCm = 3.5, double LeftMarginCm = 2.8, double RightMarginCm = 2.6, double BodyFontPt = 16, double TitleFontPt = 22, int LinesPerPage = 22, int CharactersPerLine = 28);
public static class BuiltInTemplates {
 public static IReadOnlyList<FormatTemplate> All { get; } = [new("notice-v1", "通知文字稿"), new("request-v1", "请示文字稿"), new("report-v1", "报告文字稿")];
}
public sealed record FormatRequest(DocumentAnalysis Analysis, string OutputPath, FormatTemplate Template, IReadOnlyDictionary<int, ParagraphRole> Roles, bool AcknowledgedWarnings, IReadOnlyDictionary<int, string>? ParagraphEdits = null, bool AllowEditingProtectedCopy = false);
public sealed record FormatResult(string OutputPath, string SourceSha256, string OutputSha256, int FormattedParagraphCount, IReadOnlyList<string> Warnings, bool EditProtectionRemoved = false);
public sealed class DocumentProcessingException(string message) : Exception(message);
public interface IDocumentFormatter {
 DocumentAnalysis Analyze(string sourcePath);
 FormatResult Format(FormatRequest request);
}
