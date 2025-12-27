using StyleConfigModel = WordTools2.Models.StyleConfig;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordTools2;

public partial class Form1 : Form
{
    private readonly WordTools2.Services.DocumentService _documentService = new WordTools2.Services.DocumentService();
    private WordTools2.Models.StyleConfig _styleConfig = new WordTools2.Models.StyleConfig();
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
        LoadConfig();
    }

    private void InitializeEventHandlers()
    {
        btnOpen.Click += BtnOpen_Click;
        btnApply.Click += BtnApply_Click;
        btnSave.Click += BtnSave_Click;

        this.DragEnter += Form1_DragEnter;
        this.DragDrop += Form1_DragDrop;
    }

    /// <summary>
    /// 从配置文件加载样式配置
    /// </summary>
    private void LoadConfig()
    {
        try
        {
            _styleConfig = WordTools2.Services.ConfigManager.LoadConfig();
            ApplyConfigToUI();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"加载配置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// 将配置应用到UI控件
    /// </summary>
    private void ApplyConfigToUI()
    {
        // 将配置应用到UI控件
        cmbHeading1Font.SelectedItem = _styleConfig.Heading1.FontName;
        cmbHeading1Size.SelectedItem = GetChineseFontSize(_styleConfig.Heading1.FontSize);
        nudHeading1SpaceBefore.Value = (decimal)_styleConfig.Heading1.SpaceBefore;
        nudHeading1SpaceAfter.Value = (decimal)_styleConfig.Heading1.SpaceAfter;
        nudHeading1LineSpacing.Value = (decimal)_styleConfig.Heading1.LineSpacing;
        chkHeading1Bold.Checked = _styleConfig.Heading1.Bold;

        cmbHeading2Font.SelectedItem = _styleConfig.Heading2.FontName;
        cmbHeading2Size.SelectedItem = GetChineseFontSize(_styleConfig.Heading2.FontSize);
        nudHeading2SpaceBefore.Value = (decimal)_styleConfig.Heading2.SpaceBefore;
        nudHeading2SpaceAfter.Value = (decimal)_styleConfig.Heading2.SpaceAfter;
        nudHeading2LineSpacing.Value = (decimal)_styleConfig.Heading2.LineSpacing;
        chkHeading2Bold.Checked = _styleConfig.Heading2.Bold;

        cmbHeading3Font.SelectedItem = _styleConfig.Heading3.FontName;
        cmbHeading3Size.SelectedItem = GetChineseFontSize(_styleConfig.Heading3.FontSize);
        nudHeading3SpaceBefore.Value = (decimal)_styleConfig.Heading3.SpaceBefore;
        nudHeading3SpaceAfter.Value = (decimal)_styleConfig.Heading3.SpaceAfter;
        nudHeading3LineSpacing.Value = (decimal)_styleConfig.Heading3.LineSpacing;
        chkHeading3Bold.Checked = _styleConfig.Heading3.Bold;

        cmbHeading4Font.SelectedItem = _styleConfig.Heading4.FontName;
        cmbHeading4Size.SelectedItem = GetChineseFontSize(_styleConfig.Heading4.FontSize);
        nudHeading4SpaceBefore.Value = (decimal)_styleConfig.Heading4.SpaceBefore;
        nudHeading4SpaceAfter.Value = (decimal)_styleConfig.Heading4.SpaceAfter;
        nudHeading4LineSpacing.Value = (decimal)_styleConfig.Heading4.LineSpacing;
        chkHeading4Bold.Checked = _styleConfig.Heading4.Bold;

        cmbNormalFont.SelectedItem = _styleConfig.Normal.FontName;
        cmbNormalSize.SelectedItem = GetChineseFontSize(_styleConfig.Normal.FontSize);
        nudNormalLineSpacing.Value = (decimal)_styleConfig.Normal.LineSpacing;

        // 表格标题配置
        cmbTableCaptionFont.SelectedItem = _styleConfig.TableCaption.FontName;
        cmbTableCaptionSize.SelectedItem = GetChineseFontSize(_styleConfig.TableCaption.FontSize);

        // 设置大纲级别
        int outlineLevel = _styleConfig.TableCaption.OutlineLevel;
        if (outlineLevel < 0) outlineLevel = 9; // 确保有效值
        cmbTableCaptionOutlineLevel.SelectedIndex = outlineLevel + 1; // +1因为第0项是"无(正文)"

        // 设置加粗
        chkTableCaptionBold.Checked = _styleConfig.TableCaption.Bold;

        // 图形标题配置
        cmbImageCaptionFont.SelectedItem = _styleConfig.ImageCaption.FontName;
        cmbImageCaptionSize.SelectedItem = GetChineseFontSize(_styleConfig.ImageCaption.FontSize);

        // 设置大纲级别
        int imageOutlineLevel = _styleConfig.ImageCaption.OutlineLevel;
        if (imageOutlineLevel < 0) imageOutlineLevel = 9; // 确保有效值
        cmbImageCaptionOutlineLevel.SelectedIndex = imageOutlineLevel + 1; // +1因为第0项是"无(正文)"

        // 设置加粗
        chkImageCaptionBold.Checked = _styleConfig.ImageCaption.Bold;
    }

    /// <summary>
    /// 根据数值大小获取对应的中文字体大小
    /// </summary>
    private string GetChineseFontSize(double fontSize)
    {
        // 根据数值大小获取对应的中文字体大小
        foreach (var kvp in _fontSizeMap)
        {
            if (Math.Abs(kvp.Value - fontSize) < 0.1)
            {
                return kvp.Key;
            }
        }
        return "小四"; // 默认返回小四
    }



    /// <summary>
    /// 保存配置到文件
    /// </summary>
    private void SaveConfig()
    {
        try
        {
            CollectStyleConfig();
            WordTools2.Services.ConfigManager.SaveConfig(_styleConfig);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"保存配置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
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
                LogMessage($"  一级标题: {stats["Heading1"]}");
                LogMessage($"  二级标题: {stats["Heading2"]}");
                LogMessage($"  三级标题: {stats["Heading3"]}");
                LogMessage($"  四级标题: {stats["Heading4"]}");
                LogMessage($"  表格标题: {stats["TableCaption"]}");

                if (stats["Other"] > 0)
                    LogMessage($"  其他格式: {stats["Other"]}");

                LogMessage("说明: 仅通过正则表达式识别段落类型并应用格式");

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
                    MessageBox.Show("格式应用成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    lblStatus.Text = "处理完成";
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
        _styleConfig.Heading1.LineSpacing = (double)nudHeading1LineSpacing.Value;
        _styleConfig.Heading1.Bold = chkHeading1Bold.Checked;

        _styleConfig.Heading2.FontName = cmbHeading2Font.SelectedItem?.ToString() ?? "Microsoft YaHei UI";
        _styleConfig.Heading2.FontSize = _fontSizeMap[cmbHeading2Size.SelectedItem?.ToString() ?? "四号"];
        _styleConfig.Heading2.SpaceBefore = (double)nudHeading2SpaceBefore.Value;
        _styleConfig.Heading2.SpaceAfter = (double)nudHeading2SpaceAfter.Value;
        _styleConfig.Heading2.LineSpacing = (double)nudHeading2LineSpacing.Value;
        _styleConfig.Heading2.Bold = chkHeading2Bold.Checked;

        _styleConfig.Heading3.FontName = cmbHeading3Font.SelectedItem?.ToString() ?? "Microsoft YaHei UI";
        _styleConfig.Heading3.FontSize = _fontSizeMap[cmbHeading3Size.SelectedItem?.ToString() ?? "小四"];
        _styleConfig.Heading3.SpaceBefore = (double)nudHeading3SpaceBefore.Value;
        _styleConfig.Heading3.SpaceAfter = (double)nudHeading3SpaceAfter.Value;
        _styleConfig.Heading3.LineSpacing = (double)nudHeading3LineSpacing.Value;
        _styleConfig.Heading3.Bold = chkHeading3Bold.Checked;

        _styleConfig.Heading4.FontName = cmbHeading4Font.SelectedItem?.ToString() ?? "Microsoft YaHei UI";
        _styleConfig.Heading4.FontSize = _fontSizeMap[cmbHeading4Size.SelectedItem?.ToString() ?? "五号"];
        _styleConfig.Heading4.SpaceBefore = (double)nudHeading4SpaceBefore.Value;
        _styleConfig.Heading4.SpaceAfter = (double)nudHeading4SpaceAfter.Value;
        _styleConfig.Heading4.LineSpacing = (double)nudHeading4LineSpacing.Value;
        _styleConfig.Heading4.Bold = chkHeading4Bold.Checked;

        _styleConfig.Normal.FontName = cmbNormalFont.SelectedItem?.ToString() ?? "宋体";
        _styleConfig.Normal.FontSize = _fontSizeMap[cmbNormalSize.SelectedItem?.ToString() ?? "五号"];
        _styleConfig.Normal.LineSpacing = (double)nudNormalLineSpacing.Value;

        // 表格标题配置
        _styleConfig.TableCaption.FontName = cmbTableCaptionFont.SelectedItem?.ToString() ?? "黑体";
        _styleConfig.TableCaption.FontSize = _fontSizeMap[cmbTableCaptionSize.SelectedItem?.ToString() ?? "五号"];

        // 设置大纲级别
        int outlineLevel = cmbTableCaptionOutlineLevel.SelectedIndex - 1; // -1因为第0项是"无(正文)"
        if (outlineLevel < 0) outlineLevel = 9; // 如果选择"无(正文)"，设置为9
        _styleConfig.TableCaption.OutlineLevel = outlineLevel;

        // 设置加粗
        _styleConfig.TableCaption.Bold = chkTableCaptionBold.Checked;

        // 图形标题配置
        _styleConfig.ImageCaption.FontName = cmbImageCaptionFont.SelectedItem?.ToString() ?? "黑体";
        _styleConfig.ImageCaption.FontSize = _fontSizeMap[cmbImageCaptionSize.SelectedItem?.ToString() ?? "五号"];

        // 设置大纲级别
        int imageOutlineLevel = cmbImageCaptionOutlineLevel.SelectedIndex - 1; // -1因为第0项是"无(正文)"
        if (imageOutlineLevel < 0) imageOutlineLevel = 9; // 如果选择"无(正文)"，设置为9
        _styleConfig.ImageCaption.OutlineLevel = imageOutlineLevel;

        // 设置加粗
        _styleConfig.ImageCaption.Bold = chkImageCaptionBold.Checked;
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

    private void btnSave_Click_1(object sender, EventArgs e)
    {

    }

    private void lblHeading4Before_Click(object sender, EventArgs e)
    {

    }
}
