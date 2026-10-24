using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Gongwen.Core;
using Microsoft.Win32;

namespace Gongwen.Desktop;

public partial class MainWindow : Window
{
    private readonly IDocumentFormatter _formatter = new DocxFormatter();
    private readonly ObservableCollection<FormatTemplate> _templates;
    private readonly HashSet<string> _installedFonts;
    private readonly ObservableCollection<ParagraphRow> _rows = [];
    private readonly Dictionary<int, string> _paragraphEdits = [];
    private Dictionary<int, string> _exportedEdits = [];
    private DocumentAnalysis? _analysis;
    private string? _outputPath;
    private bool _busy;
    private bool _populatingTemplate;
    private bool _initialized;
    private bool _structuralRiskAcknowledged;
    private bool _allowClose;
    private bool _officeImportScheduled;
    private int _modalDepth;
    private bool IsManualReview => ManualReviewCheck.IsChecked == true;
    private bool HasPendingEdits => _paragraphEdits.Count != _exportedEdits.Count ||
        _paragraphEdits.Any(edit => !_exportedEdits.TryGetValue(edit.Key, out string? exported) || exported != edit.Value);

    public MainWindow()
    {
        InitializeComponent();
        var version = typeof(App).Assembly.GetName().Version;
        if (version is not null) Title = "公文一键排版 v" + version.ToString(3);
        _installedFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var family in Fonts.SystemFontFamilies)
        {
            _installedFonts.Add(family.Source);
            foreach (var alias in family.FamilyNames.Values) _installedFonts.Add(alias);
        }
        _templates = new ObservableCollection<FormatTemplate>(TemplateStore.Load(out bool configFailed));
        TemplatePicker.ItemsSource = _templates;
        ParagraphGrid.ItemsSource = _rows;
        TemplatePicker.SelectedIndex = 0;
        _initialized = true;
        UpdateReviewMode();
        Loaded += (_, _) => FitWindowToWorkArea();
        if (configFailed) SetStatus("部分模板设置无法读取，已恢复为默认设置。请核对字体和页边距。");
    }

    // Called on the UI dispatcher. Acknowledge the queued action before any user-facing save dialog starts.
    internal OfficeImportDelivery QueueOfficeImport(string inputPath)
    {
        if (!_initialized || !IsLoaded || !IsVisible || _allowClose) return OfficeImportDelivery.Rejected;
        if (!OfficeImportCommand.TryValidateDocumentPath(inputPath, out string path))
        { SetStatus("插件传入的路径无效，请使用已保存的本地或共享文件夹中的 DOCX。"); return OfficeImportDelivery.Rejected; }
        if (WindowState == WindowState.Minimized) WindowState = WindowState.Normal;
        Activate();
        if (_busy || _officeImportScheduled || _modalDepth > 0)
        {
            SetStatus("正在处理文档或等待当前对话框操作，请完成后再次点击公文排版按钮。");
            return OfficeImportDelivery.Busy;
        }
        _officeImportScheduled = true;
        _ = Dispatcher.BeginInvoke(new Action(async () =>
        {
            try { await LoadDocumentAsync(path); }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            { SetStatus("未能导入插件传入的文档。当前文稿及修改仍保留，请检查已保存文件后重试。"); }
            finally { _officeImportScheduled = false; }
        }));
        return OfficeImportDelivery.Accepted;
    }
    internal void ReportOfficeImportUnavailable() =>
        SetStatus("工具栏插件连接未能启动。请关闭并重新打开公文一键排版后重试。");
    private T ShowModal<T>(Func<T> show)
    {
        _modalDepth++;
        try { return show(); }
        finally { _modalDepth--; }
    }

    private void FitWindowToWorkArea()
    {
        // WPF exposes the desktop work area in device-independent units.
        var area = SystemParameters.WorkArea;
        double availableWidth = Math.Max(1, area.Width - 16);
        double availableHeight = Math.Max(1, area.Height - 16);
        MinWidth = Math.Min(MinWidth, availableWidth);
        MinHeight = Math.Min(MinHeight, availableHeight);
        Width = Math.Min(Width, availableWidth);
        Height = Math.Min(Height, availableHeight);
        double left = double.IsNaN(Left) ? area.Left + (area.Width - Width) / 2 : Left;
        double top = double.IsNaN(Top) ? area.Top + (area.Height - Height) / 2 : Top;
        Left = Math.Max(area.Left + 8, Math.Min(left, area.Right - Width - 8));
        Top = Math.Max(area.Top + 8, Math.Min(top, area.Bottom - Height - 8));
    }

    private void Template_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_populatingTemplate || TemplatePicker.SelectedItem is not FormatTemplate template) return;
        PopulateTemplate(template);
        AcknowledgeCheck.IsChecked = false;
        _structuralRiskAcknowledged = false;
    }
    private void PopulateTemplate(FormatTemplate t)
    {
        _populatingTemplate = true;
        FillFontPicker(TitleFontPicker, t.TitleFont);
        FillFontPicker(BodyFontPicker, t.BodyFont);
        FillFontPicker(Heading1FontPicker, t.Heading1Font);
        FillFontPicker(Heading2FontPicker, t.Heading2Font);
        FillFontPicker(LatinFontPicker, t.LatinFont);
        TitleSizeInput.Text = Number(t.TitleFontPt); BodySizeInput.Text = Number(t.BodyFontPt);
        TopInput.Text = Number(t.TopMarginCm); BottomInput.Text = Number(t.BottomMarginCm);
        LeftInput.Text = Number(t.LeftMarginCm); RightInput.Text = Number(t.RightMarginCm);
        LinesInput.Text = t.LinesPerPage.ToString(CultureInfo.InvariantCulture);
        CharactersInput.Text = t.CharactersPerLine.ToString(CultureInfo.InvariantCulture);
        _populatingTemplate = false;
    }
    private void FillFontPicker(ComboBox picker, string selected)
    {
        // Keep absent requested fonts visible; never silently substitute them.
        var choices = _installedFonts.Append(selected).Distinct(StringComparer.OrdinalIgnoreCase).Order(StringComparer.CurrentCultureIgnoreCase).ToList();
        picker.ItemsSource = choices;
        picker.SelectedItem = choices.First(name => name.Equals(selected, StringComparison.OrdinalIgnoreCase));
    }
    private static string Number(double value) => value.ToString("0.##", CultureInfo.InvariantCulture);
    private static bool TryNumber(TextBox box, out double value) => double.TryParse(box.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
    private bool TryReadTemplate(out FormatTemplate? result)
    {
        result = null;
        if (TemplatePicker.SelectedItem is not FormatTemplate selected ||
            !TryNumber(TitleSizeInput, out double titleSize) || !TryNumber(BodySizeInput, out double bodySize) ||
            !TryNumber(TopInput, out double top) || !TryNumber(BottomInput, out double bottom) ||
            !TryNumber(LeftInput, out double left) || !TryNumber(RightInput, out double right) ||
            !int.TryParse(LinesInput.Text, out int lines) || !int.TryParse(CharactersInput.Text, out int characters))
        { SetStatus("模板参数有误，请输入有效数字。"); return false; }
        var template = selected with
        {
            TitleFont = TitleFontPicker.SelectedItem as string ?? "", BodyFont = BodyFontPicker.SelectedItem as string ?? "",
            Heading1Font = Heading1FontPicker.SelectedItem as string ?? "", Heading2Font = Heading2FontPicker.SelectedItem as string ?? "",
            LatinFont = LatinFontPicker.SelectedItem as string ?? "", TitleFontPt = titleSize, BodyFontPt = bodySize,
            TopMarginCm = top, BottomMarginCm = bottom, LeftMarginCm = left, RightMarginCm = right,
            LinesPerPage = lines, CharactersPerLine = characters
        };
        if (!TemplateValidation.IsValid(template))
        { SetStatus("模板参数超出范围。正文 5–48 磅、标题 5–72 磅，行数 8–60、字数 10–60。各边距 0.5–8 厘米，左右合计须小于 16 厘米。"); return false; }
        result = template;
        return true;
    }
    private IReadOnlyList<string> MissingFonts(FormatTemplate template, IReadOnlyDictionary<int, ParagraphRole> roles)
    {
        var used = roles.Values.Where(r => r != ParagraphRole.Preserve).ToHashSet();
        var required = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (used.Count > 0) required.Add(template.LatinFont);
        if (used.Contains(ParagraphRole.Title)) required.Add(template.TitleFont);
        if (used.Contains(ParagraphRole.Heading1)) required.Add(template.Heading1Font);
        if (used.Contains(ParagraphRole.Heading2)) required.Add(template.Heading2Font);
        if (used.Any(r => r is not (ParagraphRole.Title or ParagraphRole.Heading1 or ParagraphRole.Heading2))) required.Add(template.BodyFont);
        return required.Where(font => !_installedFonts.Contains(font)).ToList();
    }
    private void SaveTemplate_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !TryReadTemplate(out var template) || template is null) return;
        int index = _templates.IndexOf(_templates.First(t => t.Id == template.Id));
        var previous = _templates[index];
        _templates[index] = template;
        TemplatePicker.SelectedItem = _templates[index];
        try { TemplateStore.Save(_templates); SetStatus("此模板设置已保存。正文和源文件名不会保存到设置中。"); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            _templates[index] = previous;
            TemplatePicker.SelectedItem = _templates[index];
            SetStatus("模板设置未能保存，请检查当前用户的文件夹权限。");
        }
    }
    private async void Import_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || _officeImportScheduled) { SetStatus("正在处理文档，请完成当前操作后重新导入。"); return; }
        var dialog = new OpenFileDialog { Filter = "Word 文档 (*.docx)|*.docx", Multiselect = false, CheckFileExists = true, Title = "选择已保存的 DOCX" };
        if (ShowModal(() => dialog.ShowDialog(this)) == true) await LoadDocumentAsync(dialog.FileName);
    }
    internal async Task<bool> LoadDocumentAsync(string path)
    {
        if (_busy) { SetStatus("正在处理文档，请完成当前操作后重新导入。"); return false; }
        if (!Path.GetExtension(path).Equals(".docx", StringComparison.OrdinalIgnoreCase))
        { SetStatus("请导入 DOCX。旧版 DOC 请先在 WPS 或 Word 中另存为 DOCX。"); return false; }
        if (!await ConfirmLeaveCurrentEditsAsync()) return false;
        SetBusy(true, "正在读取已保存文档并自动识别…");
        try
        {
            var analysis = await Task.Run(() => _formatter.Analyze(path));
            // Replace the current document only after the next one has been read successfully.
            _analysis = analysis;
            _outputPath = null; _paragraphEdits.Clear(); _exportedEdits.Clear(); _rows.Clear();
            AcknowledgeCheck.IsChecked = false; _structuralRiskAcknowledged = false;
            ProtectedCopyCheck.IsChecked = true;
            ProtectedCopyCheck.Visibility = analysis.HasEditProtection ? Visibility.Visible : Visibility.Collapsed;
            foreach (var paragraph in analysis.Paragraphs) _rows.Add(new ParagraphRow(paragraph) { ManualReview = IsManualReview });
            var suggested = _templates.FirstOrDefault(t => t.Id == analysis.SuggestedTemplateId);
            if (suggested is not null)
            {
                TemplatePicker.SelectedItem = suggested;
                TemplateHint.Text = "已自动匹配：" + suggested.Name + "，可手动更改";
            }
            else TemplateHint.Text = "未明确识别公文类型，请核对当前模板";
            SourceText.Text = Path.GetFileName(path);
            UpdateOutputLabel();
            UpdateDocumentSummary();
            List<string> messages = ["表格、图片、修订和特殊对象保持原样。输出后请在 WPS 或 Word 复核分页与打印效果。"];
            messages.AddRange(analysis.Warnings);
            if (!analysis.CanFormat) messages.Insert(0, analysis.BlockingReason ?? "此文档暂不支持自动排版。");
            int suggestions = analysis.Paragraphs.Count(p => p.NeedsReview && p.SuggestedRole != ParagraphRole.Preserve);
            if (suggestions > 0) messages.Add($"有 {suggestions} 段识别结果建议复核，可在列表中调整角色。");
            WarningText.Text = string.Join("\n", messages);
            if (_rows.Count > 0) ParagraphGrid.SelectedIndex = 0;
            SetStatus(analysis.CanFormat ? "自动识别完成，可编辑简单段落，或直接导出排好版 DOCX。" : "已停止自动排版，请先在 WPS 或 Word 中整理为普通文字稿。");
            return true;
        }
        catch (DocumentProcessingException ex)
        { SetStatus(ex.Message + (_analysis is not null ? " 当前文稿及修改已保留。" : "")); return false; }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        { SetStatus("读取失败，请确认 DOCX 有效、未加密且可读取。当前文稿及修改会保留。"); return false; }
        finally { SetBusy(false); }
    }
    private void ReviewMode_Changed(object sender, RoutedEventArgs e)
    { if (_initialized) UpdateReviewMode(); }
    private void UpdateReviewMode()
    {
        bool manual = IsManualReview;
        ManualReviewColumn.Visibility = manual ? Visibility.Visible : Visibility.Collapsed;
        AcknowledgeCheck.Visibility = manual ? Visibility.Visible : Visibility.Collapsed;
        AcknowledgeCheck.IsChecked = false;
        ModeDescription.Text = manual
            ? "逐段核对未确认项，再勾选提示确认。受保护段落保持原样。"
            : "普通文字稿可直接导出。建议复核处可调整角色；特殊对象与修订会保留。";
        foreach (var row in _rows) row.ManualReview = manual;
    }
    private void Topmost_Changed(object sender, RoutedEventArgs e) => Topmost = TopmostCheck.IsChecked == true;
    private void ParagraphSelection_Changed(object sender, SelectionChangedEventArgs e)
    { if (_initialized) UpdateEditorAvailability(); }
    private void UpdateEditorAvailability()
    {
        var selected = ParagraphGrid.SelectedItem as ParagraphRow;
        EditParagraphButton.IsEnabled = !_busy && _analysis?.CanFormat == true && selected?.CanEdit == true;
        EditHint.Text = selected is null ? "选择段落后可编辑纯文字；悬停预览可看全文。"
            : selected.CanEdit ? (selected.HasAutomaticNumbering ? $"第 {selected.DisplayIndex} 段文字可编辑，自动编号会保留。" : $"第 {selected.DisplayIndex} 段可编辑。修改后保留当前角色。")
            : "本段含对象、修订或其他复杂结构，请在 WPS 中修改并保存后重新拖入。";
    }
    private void EditParagraph_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || ParagraphGrid.SelectedItem is not ParagraphRow row || !row.CanEdit || _analysis?.CanFormat != true) return;
        var editor = new ParagraphEditWindow(row) { Owner = this };
        if (ShowModal(editor.ShowDialog) != true || editor.EditedText == row.FullText) return;
        row.SetText(editor.EditedText);
        if (row.IsEdited) _paragraphEdits[row.Index] = row.FullText;
        else _paragraphEdits.Remove(row.Index);
        AcknowledgeCheck.IsChecked = false;
        UpdateDocumentSummary(); UpdateOutputLabel();
        SetStatus("本段修改已暂存，导出后写入新 DOCX。当前角色已保留，可按需要调整。");
    }
    private void UpdateDocumentSummary()
    {
        if (_analysis is null) return;
        DocumentSummary.Text = $"{_analysis.Paragraphs.Count} 个段落 · {_analysis.TableCount} 个表格 · {_analysis.ImageCount} 张图片 · {_paragraphEdits.Count} 段已修改";
    }
    private void UpdateOutputLabel()
    {
        OutputText.Text = _outputPath ?? "";
        OutputHint.Text = _outputPath is null
            ? "尚未生成文件；导出成功后显示完整保存路径，可选中复制。"
            : HasPendingEdits ? "上次输出位置（尚未包含当前修改）；请再次导出后打开新文件。"
            : "输出位置（可选中复制），也可点击右侧打开文档或目录。";
    }
    private bool ConfirmStructuralRisk()
    {
        if (_analysis is null) return false;
        if (_structuralRiskAcknowledged) return true;
        if (!_analysis.RequiresStructureReview) return true;
        string details = _analysis.Warnings.Count > 0 ? string.Join("\n", _analysis.Warnings.Take(8)) : "文档含需要保留的特殊结构。";
        var choice = ShowModal(() => MessageBox.Show(this, details + "\n\n这些对象与修订将保持原样，程序不会自动接受修订。输出后需在 WPS 或 Word 核对布局。是否继续导出？",
            "特殊结构复核", MessageBoxButton.YesNo, MessageBoxImage.Warning));
        if (choice != MessageBoxResult.Yes) return false;
        _structuralRiskAcknowledged = true;
        return true;
    }
    private async void Format_Click(object sender, RoutedEventArgs e) => await ExportDocumentAsync();
    private async Task<bool> ExportDocumentAsync()
    {
        if (_busy) { SetStatus("正在处理文档，请等待当前操作完成。"); return false; }
        if (_analysis is null) { SetStatus("尚未导入文档。请先拖入或选择 DOCX，再点击导出。"); return false; }
        if (!_analysis.CanFormat) { SetStatus(_analysis.BlockingReason ?? "当前文档暂不支持自动排版，请在 WPS 或 Word 中整理后重新导入。"); return false; }
        if (_analysis.HasEditProtection && ProtectedCopyCheck.IsChecked != true)
        { SetStatus("当前文档有编辑限制。请勾选“仅在输出副本中解除编辑限制（原稿不变）”后导出。"); return false; }
        if (!TryReadTemplate(out var template) || template is null) return false;
        ParagraphGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        ParagraphGrid.CommitEdit(DataGridEditingUnit.Row, true);
        if (IsManualReview && _rows.Any(row => !row.Reviewed))
        { SetStatus("请核对未确认段落并勾选“已确认”，或切换到自动识别排版模式。"); return false; }
        if (IsManualReview && AcknowledgeCheck.IsChecked != true)
        { SetStatus("请勾选提示确认后导出。"); return false; }
        var roles = _rows.ToDictionary(row => row.Index, row => row.MustPreserve ? ParagraphRole.Preserve : row.Role);
        var missing = MissingFonts(template, roles);
        if (missing.Count > 0)
        { SetStatus("缺少所需字体：" + string.Join("、", missing) + "。请安装字体或明确选择本机已安装字体后继续。"); return false; }
        if (!IsManualReview && !ConfirmStructuralRisk()) return false;
        var save = new SaveFileDialog
        {
            Filter = "Word 文档 (*.docx)|*.docx", DefaultExt = ".docx", AddExtension = true,
            FileName = Path.GetFileNameWithoutExtension(_analysis.SourcePath) + "_排版.docx",
            Title = "导出排好版 DOCX（请选择新文件名）", OverwritePrompt = false
        };
        if (ShowModal(() => save.ShowDialog(this)) != true) return false;
        if (Path.GetFullPath(save.FileName).Equals(Path.GetFullPath(_analysis.SourcePath), StringComparison.OrdinalIgnoreCase))
        { SetStatus("不能覆盖原稿，请选择新的输出文件名。"); return false; }
        if (!Path.GetExtension(save.FileName).Equals(".docx", StringComparison.OrdinalIgnoreCase))
        { SetStatus("输出文件必须使用 .docx 扩展名，请重新选择文件名。"); return false; }
        if (File.Exists(save.FileName)) { SetStatus("输出文件已存在，请使用新的文件名；现有文件不会覆盖。"); return false; }
        var edits = new Dictionary<int, string>(_paragraphEdits);
        var request = new FormatRequest(_analysis, save.FileName, template, roles, true, edits, _analysis.HasEditProtection && ProtectedCopyCheck.IsChecked == true);
        SetBusy(true, "正在自动排版并导出新 DOCX…");
        try
        {
            var result = await Task.Run(() => _formatter.Format(request));
            _outputPath = result.OutputPath; _exportedEdits = edits;
            UpdateOutputLabel();
            SetStatus($"已导出 {result.FormattedParagraphCount} 个排版段落，含 {edits.Count} 段文字修改。" + (result.EditProtectionRemoved ? "已在输出副本解除编辑限制，原稿不变。" : "") + "请在 WPS 或 Word 检查分页与打印效果。");
            if (result.Warnings.Count > 0) WarningText.Text = string.Join("\n", result.Warnings);
            try { await AuditStore.WriteCompletedAsync(template, result, edits.Count); }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            { SetStatus("DOCX 已保存，但本机审计记录写入失败。请检查用户目录权限或磁盘空间，并用 WPS / Word 复核文档。"); }
            return true;
        }
        catch (DocumentProcessingException ex) { SetStatus(ex.Message); return false; }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        { SetStatus("导出未完成。当前修改仍保留，请检查目录权限、文件占用及模板参数后重试。"); return false; }
        finally { SetBusy(false); }
    }
    private async Task<bool> ConfirmLeaveCurrentEditsAsync()
    {
        if (!HasPendingEdits) return true;
        var choice = ShowModal(() => MessageBox.Show(this, "当前段落修改尚未导出。是否先导出为新 DOCX？\n\n选择“是”先导出；“否”放弃未导出的修改；“取消”继续编辑。",
            "保留当前修改", MessageBoxButton.YesNoCancel, MessageBoxImage.Question));
        if (choice == MessageBoxResult.Cancel) return false;
        if (choice == MessageBoxResult.No) return true;
        return await ExportDocumentAsync();
    }
    private void SetBusy(bool busy, string? message = null)
    {
        _busy = busy; ImportButton.IsEnabled = !busy; TemplatePanel.IsEnabled = !busy;
        ParagraphGrid.IsEnabled = !busy; AcknowledgeCheck.IsEnabled = !busy; ManualReviewCheck.IsEnabled = !busy;
        ProtectedCopyCheck.IsEnabled = !busy;
        // Keep the action visible and clickable for input errors; report the reason in the output area.
        FormatButton.IsEnabled = !busy;
        FormatButton.Content = busy ? "正在处理文档…" : "导出排好版 DOCX";
        OpenOutputButton.IsEnabled = !busy && _outputPath is not null;
        OpenFolderButton.IsEnabled = !busy && _outputPath is not null;
        CopyPathButton.IsEnabled = !busy && _outputPath is not null;
        if (_initialized) UpdateEditorAvailability();
        if (message is not null) SetStatus(message);
    }
    private void SetStatus(string message) => StatusText.Text = message;
    private void Window_DragOver(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        var files = e.Data.GetDataPresent(DataFormats.FileDrop) ? e.Data.GetData(DataFormats.FileDrop) as string[] : null;
        e.Effects = !_busy && !_officeImportScheduled && files?.Length == 1 && Path.GetExtension(files[0]).Equals(".docx", StringComparison.OrdinalIgnoreCase)
            ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }
    private async void Window_Drop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        e.Handled = true;
        if (_busy || _officeImportScheduled) { SetStatus("正在处理文档，请完成当前操作后重新导入。"); return; }
        if (e.Data.GetData(DataFormats.FileDrop) is not string[] files || files.Length != 1)
        { SetStatus("一次请只拖入一个 DOCX 文件。"); return; }
        await LoadDocumentAsync(files[0]);
    }
    private async void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (_allowClose) return;
        if (_busy || _officeImportScheduled) { e.Cancel = true; SetStatus("正在处理文档，请等待本次操作完成后关闭窗口。"); return; }
        if (!HasPendingEdits) return;
        e.Cancel = true;
        if (!await ConfirmLeaveCurrentEditsAsync()) return;
        _allowClose = true;
        _ = Dispatcher.BeginInvoke(new Action(Close));
    }
    private void CopyPath_Click(object sender, RoutedEventArgs e)
    {
        if (_outputPath is null) { SetStatus("尚未生成文件，请先导出 DOCX。"); return; }
        try { Clipboard.SetText(_outputPath); SetStatus("输出路径已复制。"); }
        catch (ExternalException) { SetStatus("暂时无法写入剪贴板，可在输出路径框中选中文字后复制。"); }
    }
    private void OpenOutput_Click(object sender, RoutedEventArgs e)
    {
        if (_outputPath is null) { SetStatus("尚未生成文件，请先导出 DOCX。"); return; }
        if (!File.Exists(_outputPath)) { SetStatus("输出文件已被移动或删除，请检查保存目录或重新导出。"); return; }
        OpenWithShell(_outputPath, "无法打开输出文档。请确认本机已安装 WPS 或 Word 并关联 DOCX。");
    }
    private void OpenFolder_Click(object sender, RoutedEventArgs e)
    {
        if (_outputPath is null) { SetStatus("尚未生成文件，请先导出 DOCX。"); return; }
        string folder = Path.GetDirectoryName(_outputPath)!;
        if (!Directory.Exists(folder)) { SetStatus("输出目录已被移动或删除，请检查保存位置。"); return; }
        OpenWithShell(folder, "无法打开输出目录，请使用文件资源管理器查看。");
    }
    private void OpenSamples_Click(object sender, RoutedEventArgs e) => OpenBundledFolder("samples", "未找到随附示例，请查看安装包交付文件。");
    private void OpenHelp_Click(object sender, RoutedEventArgs e) => OpenBundledFolder("docs", "未找到随附说明，请查看安装包交付文件。");
    private void OpenBundledFolder(string folder, string failure)
    {
        string path = Path.Combine(AppContext.BaseDirectory, folder);
        if (Directory.Exists(path)) OpenWithShell(path, failure); else SetStatus(failure);
    }
    private void OpenWithShell(string path, string failure)
    {
        try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
        catch (Exception ex) when (ex is Win32Exception or IOException or InvalidOperationException) { SetStatus(failure); }
    }
    internal async Task<DesktopSmokeResult> RunSmokeTestAsync(string input, string output)
    {
        if (!await LoadDocumentAsync(input) || _analysis is null || !_analysis.CanFormat)
            return new(false, true, false, false, [], "输入文档未通过分析", false);
        var template = BuiltInTemplates.All.FirstOrDefault(t => t.Id == _analysis.SuggestedTemplateId) ?? BuiltInTemplates.All[0];
        var roles = _rows.ToDictionary(row => row.Index, row => row.MustPreserve ? ParagraphRole.Preserve : row.Role);
        var missing = MissingFonts(template, roles);
        if (missing.Count > 0)
            return new(false, true, true, false, missing, "缺少默认模板所需字体，未执行排版", false);
        var request = new FormatRequest(_analysis, output, template, roles, true, AllowEditingProtectedCopy: _analysis.HasEditProtection);
        SetBusy(true, "正在运行 Windows 本机验证…");
        try
        {
            var result = await Task.Run(() => _formatter.Format(request));
            try { await AuditStore.WriteCompletedAsync(template, result); }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            { return new(false, true, true, File.Exists(result.OutputPath), [], "排版输出已保存，但审计记录写入失败", false); }
            return new(true, true, true, File.Exists(result.OutputPath), [], "WPF 启动、自动识别与格式引擎运行完成", false, true, result.EditProtectionRemoved);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        { return new(false, true, true, false, [], "格式引擎执行失败", false); }
        finally { SetBusy(false); }
    }
}

internal sealed record DesktopSmokeResult(bool Passed, bool WpfWindowLoaded, bool AnalysisCompleted, bool OutputCreated, IReadOnlyList<string> MissingFonts, string Status, bool WordWpsCompatibilityVerified, bool AuditWritten = false, bool EditProtectionRemoved = false);
