using System.ComponentModel;
using System.Runtime.CompilerServices;
using Gongwen.Core;

namespace Gongwen.Desktop;

public sealed record RoleChoice(ParagraphRole Role, string Label)
{
    public static IReadOnlyList<RoleChoice> All { get; } =
    [
        new(ParagraphRole.Preserve, "保留原格式"), new(ParagraphRole.Title, "主标题"),
        new(ParagraphRole.Recipient, "主送单位"), new(ParagraphRole.Body, "正文"),
        new(ParagraphRole.Heading1, "一级标题"), new(ParagraphRole.Heading2, "二级标题"),
        new(ParagraphRole.Heading3, "三级标题"), new(ParagraphRole.Heading4, "四级标题"),
        new(ParagraphRole.AttachmentNote, "附件说明"), new(ParagraphRole.Signature, "发文署名"),
        new(ParagraphRole.Date, "成文日期")
    ];
}

public sealed class ParagraphRow(ParagraphInfo paragraph) : INotifyPropertyChanged
{
    private ParagraphRole _role = paragraph.SuggestedRole;
    private bool _reviewed = !paragraph.NeedsReview;
    private bool _manualReview;
    private bool _roleChanged;
    private string _text = paragraph.Text;
    public int Index { get; } = paragraph.Index;
    public int DisplayIndex => Index + 1;
    public string OriginalText { get; } = paragraph.Text;
    public string Preview => ShortPreview(_text);
    public string FullText => _text;
    public bool IsEdited => !string.Equals(_text, OriginalText, StringComparison.Ordinal);
    public bool CanEdit { get; } = paragraph.CanEdit;
    public bool NeedsReview { get; } = paragraph.NeedsReview;
    public bool HasAutomaticNumbering { get; } = paragraph.HasAutomaticNumbering;
    public bool MustPreserve { get; } = paragraph.SuggestedRole == ParagraphRole.Preserve && paragraph.NeedsReview && !paragraph.CanEdit;
    public bool CanChangeRole => !MustPreserve;
    public string Reason { get; } = paragraph.Reason;
    public string ReviewLabel => HasAutomaticNumbering ? (MustPreserve ? "编号需保留" : IsEdited ? "编号文字已改" : "自动编号") : _manualReview ? (Reviewed ? "已确认" : "待确认") : IsEdited ? "已编辑" : MustPreserve ? "受保护" : _roleChanged ? "手动调整" : NeedsReview ? "建议复核" : "自动识别";
    public bool ManualReview
    {
        get => _manualReview;
        set { _manualReview = value; Changed(nameof(ReviewLabel)); }
    }
    public ParagraphRole Role
    {
        get => _role;
        set
        {
            if (_role == value || MustPreserve) return;
            _role = value;
            _roleChanged = true;
            Reviewed = false;
            Changed();
            Changed(nameof(ReviewLabel));
        }
    }
    public bool Reviewed
    {
        get => _reviewed;
        set { _reviewed = value; Changed(); Changed(nameof(ReviewLabel)); }
    }
    public void SetText(string text)
    {
        if (_text == text) return;
        _text = text;
        Reviewed = false;
        Changed(nameof(FullText)); Changed(nameof(Preview)); Changed(nameof(IsEdited)); Changed(nameof(ReviewLabel));
    }
    public event PropertyChangedEventHandler? PropertyChanged;
    private void Changed([CallerMemberName] string? property = null) => PropertyChanged?.Invoke(this, new(property));
    private static string ShortPreview(string text)
    {
        string singleLine = text.Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');
        return singleLine.Length <= 90 ? singleLine : singleLine[..90] + "…";
    }
}
