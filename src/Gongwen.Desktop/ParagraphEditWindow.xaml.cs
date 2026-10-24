using System.Windows;
using System.Xml;

namespace Gongwen.Desktop;

public partial class ParagraphEditWindow : Window
{
    public string EditedText => EditorText.Text;
    public ParagraphEditWindow(ParagraphRow paragraph)
    {
        InitializeComponent();
        HeadingText.Text = $"编辑第 {paragraph.DisplayIndex} 段";
        EditorText.Text = paragraph.FullText;
        Loaded += (_, _) => { EditorText.Focus(); EditorText.CaretIndex = EditorText.Text.Length; };
    }
    private void Save_Click(object sender, RoutedEventArgs e)
    {
        string text = EditorText.Text;
        if (text.Length > 100_000 || text.IndexOfAny(['\r', '\n', '\t', '\u2028', '\u2029']) >= 0)
        { ValidationText.Text = "一段最多 100,000 字符，不能包含换行、制表符或段落分隔符。请在 WPS 中进行多段调整。"; return; }
        try { XmlConvert.VerifyXmlChars(text); }
        catch (XmlException) { ValidationText.Text = "文字中含不受支持的控制字符，请去除后再保存。"; return; }
        DialogResult = true;
    }
}
