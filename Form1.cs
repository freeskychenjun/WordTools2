using WordTools2.Services;
using WordTools2.Models;

namespace WordTools2;

public partial class Form1 : Form
{
    private readonly DocumentService _documentService = new DocumentService();
    private readonly StyleConfig _styleConfig = new StyleConfig();
    private string? _currentFilePath;

    // 中文字体大小与数值大小的对应关系
    private readonly Dictionary<string, double> _fontSizeMap = new Dictionary<string, double>()
    {

         { "初号", 42 },
{ "小初", 36 },
{ "一号", 26 },
{ "小一", 24 },
{ "二号", 22 },
{ "小二", 18 },
{ "三号", 16 },
{ "小三", 15 },
{ "四号", 14 },
{ "小四", 12 },
{ "五号", 10.5 },
{ "小五", 9 },
{ "六号", 7.5 },
{ "小六", 6.5 },
{ "七号", 5.5 }
    };

    public Form1()
    {
        InitializeComponent();
        InitializeEventHandlers();
    }

    private void InitializeEventHandlers()
    {
        btnOpen.Click += BtnOpen_Click;
        btnApply.Click += BtnApply_Click;
        btnSave.Click += BtnSave_Click;

        this.DragEnter += Form1_DragEnter;
        this.DragDrop += Form1_DragDrop;
    }

    private void Form1_DragEnter(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
        {
            e.Effect = DragDropEffects.Copy;
        }
        else
        {
            e.Effect = DragDropEffects.None;
        }
    }

    private void Form1_DragDrop(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop)!;
            if (files.Length > 0)
            {
                LoadDocument(files[0]);
            }
        }
    }

    private void BtnOpen_Click(object? sender, EventArgs e)
    {
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            LoadDocument(openFileDialog.FileName);
        }
    }

    private void LoadDocument(string filePath)
    {
        try
        {
            if (_documentService.OpenDocument(filePath))
            {
                _currentFilePath = filePath;
                txtFileName.Text = filePath;
                lblStatus.Text = "已加载文档（只读）";
                LogMessage($"已打开文档: {Path.GetFileName(filePath)}");

                var stats = _documentService.GetDocumentStats();
                LogMessage("文档统计（基于大纲级别）:");
                LogMessage($"  正文文本 (大纲级别0): {stats["Normal"]}");
                LogMessage($"  一级标题 (大纲级别1): {stats["Heading1"]}");
                LogMessage($"  二级标题 (大纲级别2): {stats["Heading2"]}");
                LogMessage($"  三级标题 (大纲级别3): {stats["Heading3"]}");
                LogMessage($"  四级标题 (大纲级别4): {stats["Heading4"]}");

                if (stats["Level5"] > 0)
                    LogMessage($"  5级段落 (保持不变): {stats["Level5"]}");
                if (stats["Level6"] > 0)
                    LogMessage($"  6级段落 (保持不变): {stats["Level6"]}");
                if (stats["Level7"] > 0)
                    LogMessage($"  7级段落 (保持不变): {stats["Level7"]}");
                if (stats["Level8"] > 0)
                    LogMessage($"  8级段落 (保持不变): {stats["Level8"]}");
                if (stats["Level9"] > 0)
                    LogMessage($"  9级段落 (保持不变): {stats["Level9"]}");
                if (stats["NoLevel"] > 0)
                    LogMessage($"  无大纲级别: {stats["NoLevel"]}");

                LogMessage($"  说明: 仅对0-4级段落进行排版，5-9级段落维持原样");
                LogMessage("原始文档未被修改");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            LogMessage($"错误: {ex.Message}");
        }
    }

    private void BtnApply_Click(object? sender, EventArgs e)
    {
        if (_currentFilePath == null)
        {
            MessageBox.Show("请先打开一个 Word 文档", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        CollectStyleConfig();

        progressBar.Visible = true;
        progressBar.Value = 0;
        btnApply.Enabled = false;
        btnOpen.Enabled = false;
        btnSave.Enabled = false;

        Task.Run(() =>
        {
            try
            {
                _documentService.ApplyStyles(
                    _styleConfig,
                    progress => UpdateProgress(progress),
                    message => LogMessage(message)
                );

                this.Invoke((MethodInvoker)delegate
                {
                    MessageBox.Show("格式应用成功！原始文档未被修改，段落样式保持不变。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    lblStatus.Text = "处理完成（格式已应用，样式保持原样）";
                    lblProgress.Text = "";
                    progressBar.Visible = false;
                    btnApply.Enabled = true;
                    btnOpen.Enabled = true;
                    btnSave.Enabled = true;
                });
            }
            catch (Exception ex)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    MessageBox.Show($"应用样式失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LogMessage($"错误: {ex.Message}");
                    lblStatus.Text = "处理失败";
                    progressBar.Visible = false;
                    btnApply.Enabled = true;
                    btnOpen.Enabled = true;
                    btnSave.Enabled = true;
                });
            }
        });
    }

    private void BtnSave_Click(object? sender, EventArgs e)
    {
        if (_currentFilePath == null)
        {
            MessageBox.Show("请先打开一个 Word 文档", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        saveFileDialog.InitialDirectory = Path.GetDirectoryName(_currentFilePath);

        // 如果有工作文件，使用格式化后的文件名；否则使用原始文件名
        string baseFileName = Path.GetFileNameWithoutExtension(_currentFilePath);
        if (_documentService.HasWorkingFile())
        {
            saveFileDialog.FileName = baseFileName + "_formatted";
        }
        else
        {
            saveFileDialog.FileName = baseFileName;
        }

        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _documentService.SaveDocumentAs(saveFileDialog.FileName);
                LogMessage($"文档已保存: {Path.GetFileName(saveFileDialog.FileName)}");
                MessageBox.Show("文档保存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存文档失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogMessage($"错误: {ex.Message}");
            }
        }
    }

    private void CollectStyleConfig()
    {
        _styleConfig.Heading1.FontName = cmbHeading1Font.SelectedItem?.ToString() ?? "Microsoft YaHei UI";
        _styleConfig.Heading1.FontSize = _fontSizeMap[cmbHeading1Size.SelectedItem?.ToString() ?? "三号"];
        _styleConfig.Heading1.SpaceBefore = (double)nudHeading1SpaceBefore.Value;
        _styleConfig.Heading1.SpaceAfter = (double)nudHeading1SpaceAfter.Value;

        _styleConfig.Heading2.FontName = cmbHeading2Font.SelectedItem?.ToString() ?? "Microsoft YaHei UI";
        _styleConfig.Heading2.FontSize = _fontSizeMap[cmbHeading2Size.SelectedItem?.ToString() ?? "四号"];
        _styleConfig.Heading2.SpaceBefore = (double)nudHeading2SpaceBefore.Value;
        _styleConfig.Heading2.SpaceAfter = (double)nudHeading2SpaceAfter.Value;

        _styleConfig.Heading3.FontName = cmbHeading3Font.SelectedItem?.ToString() ?? "Microsoft YaHei UI";
        _styleConfig.Heading3.FontSize = _fontSizeMap[cmbHeading3Size.SelectedItem?.ToString() ?? "小四"];
        _styleConfig.Heading3.SpaceBefore = (double)nudHeading3SpaceBefore.Value;
        _styleConfig.Heading3.SpaceAfter = (double)nudHeading3SpaceAfter.Value;

        _styleConfig.Heading4.FontName = cmbHeading4Font.SelectedItem?.ToString() ?? "Microsoft YaHei UI";
        _styleConfig.Heading4.FontSize = _fontSizeMap[cmbHeading4Size.SelectedItem?.ToString() ?? "五号"];
        _styleConfig.Heading4.SpaceBefore = (double)nudHeading4SpaceBefore.Value;
        _styleConfig.Heading4.SpaceAfter = (double)nudHeading4SpaceAfter.Value;

        _styleConfig.Normal.FontName = cmbNormalFont.SelectedItem?.ToString() ?? "宋体";
        _styleConfig.Normal.FontSize = _fontSizeMap[cmbNormalSize.SelectedItem?.ToString() ?? "五号"];
        _styleConfig.Normal.LineSpacing = (double)nudNormalLineSpacing.Value;
    }

    private void UpdateProgress(string progress)
    {
        if (InvokeRequired)
        {
            Invoke((MethodInvoker)delegate { UpdateProgress(progress); });
            return;
        }

        lblProgress.Text = progress;
        if (progress.Contains('%'))
        {
            var match = System.Text.RegularExpressions.Regex.Match(progress, @"(\d+)%");
            if (match.Success && int.TryParse(match.Groups[1].Value, out int percent))
            {
                progressBar.Value = percent;
            }
        }
    }

    private void LogMessage(string message)
    {
        if (InvokeRequired)
        {
            Invoke((MethodInvoker)delegate { LogMessage(message); });
            return;
        }

        string timestamp = DateTime.Now.ToString("HH:mm:ss");
        txtLog.AppendText($"[{timestamp}] {message}\r\n");
        txtLog.ScrollToCaret();
    }

    private void Form1_Load(object sender, EventArgs e)
    {

    }
}
