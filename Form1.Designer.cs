using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordTools2;

partial class Form1
{
    private System.ComponentModel.IContainer components = null;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    private System.Windows.Forms.Button btnOpen;
    private System.Windows.Forms.Button btnApply;
    private System.Windows.Forms.Button btnSave;
    private System.Windows.Forms.GroupBox grpHeading1;
    private System.Windows.Forms.GroupBox grpHeading2;
    private System.Windows.Forms.GroupBox grpHeading3;
    private System.Windows.Forms.GroupBox grpHeading4;
    private System.Windows.Forms.GroupBox grpNormal;
    private System.Windows.Forms.GroupBox grpTableCaption;
    private System.Windows.Forms.TextBox txtFileName;
    private System.Windows.Forms.TextBox txtLog;
    private System.Windows.Forms.Label lblStatus;
    private System.Windows.Forms.Label lblProgress;
    private System.Windows.Forms.ProgressBar progressBar;

    // Heading 1 Controls
    private System.Windows.Forms.ComboBox cmbHeading1Font;
    private System.Windows.Forms.ComboBox cmbHeading1Size;
    private System.Windows.Forms.NumericUpDown nudHeading1SpaceBefore;
    private System.Windows.Forms.NumericUpDown nudHeading1SpaceAfter;
    private System.Windows.Forms.NumericUpDown nudHeading1LineSpacing = new NumericUpDown();
    private System.Windows.Forms.Label lblHeading1LineSpacing = new Label();

    // Heading 2 Controls
    private System.Windows.Forms.ComboBox cmbHeading2Font;
    private System.Windows.Forms.ComboBox cmbHeading2Size;
    private System.Windows.Forms.NumericUpDown nudHeading2SpaceBefore;
    private System.Windows.Forms.NumericUpDown nudHeading2SpaceAfter;
    private System.Windows.Forms.NumericUpDown nudHeading2LineSpacing = new NumericUpDown();
    private System.Windows.Forms.Label lblHeading2LineSpacing = new Label();

    // Heading 3 Controls
    private System.Windows.Forms.ComboBox cmbHeading3Font;
    private System.Windows.Forms.ComboBox cmbHeading3Size;
    private System.Windows.Forms.NumericUpDown nudHeading3SpaceBefore;
    private System.Windows.Forms.NumericUpDown nudHeading3SpaceAfter;
    private System.Windows.Forms.NumericUpDown nudHeading3LineSpacing = new NumericUpDown();
    private System.Windows.Forms.Label lblHeading3LineSpacing = new Label();

    // Heading 4 Controls
    private System.Windows.Forms.ComboBox cmbHeading4Font;
    private System.Windows.Forms.ComboBox cmbHeading4Size;
    private System.Windows.Forms.NumericUpDown nudHeading4SpaceBefore;
    private System.Windows.Forms.NumericUpDown nudHeading4SpaceAfter;
    private System.Windows.Forms.NumericUpDown nudHeading4LineSpacing = new NumericUpDown();
    private System.Windows.Forms.Label lblHeading4LineSpacing = new Label();

    // Normal Controls
    private System.Windows.Forms.ComboBox cmbNormalFont;
    private System.Windows.Forms.ComboBox cmbNormalSize;
    private System.Windows.Forms.NumericUpDown nudNormalLineSpacing;
    private System.Windows.Forms.Label lblNormalLineSpacing;

    // Table Caption Controls
    private System.Windows.Forms.ComboBox cmbTableCaptionFont;
    private System.Windows.Forms.ComboBox cmbTableCaptionSize;
    private System.Windows.Forms.ComboBox cmbTableCaptionOutlineLevel;
    private System.Windows.Forms.CheckBox chkTableCaptionBold;
    private System.Windows.Forms.Label lblTableCaptionFont;
    private System.Windows.Forms.Label lblTableCaptionSize;
    private System.Windows.Forms.Label lblTableCaptionOutlineLevel;
    private System.Windows.Forms.Label lblTableCaptionBold;
    private System.Windows.Forms.GroupBox grpImageCaption;

    // Image Caption Controls
    private System.Windows.Forms.ComboBox cmbImageCaptionFont;
    private System.Windows.Forms.ComboBox cmbImageCaptionSize;
    private System.Windows.Forms.ComboBox cmbImageCaptionOutlineLevel;
    private System.Windows.Forms.CheckBox chkImageCaptionBold;
    private System.Windows.Forms.Label lblImageCaptionFont;
    private System.Windows.Forms.Label lblImageCaptionSize;
    private System.Windows.Forms.Label lblImageCaptionOutlineLevel;
    private System.Windows.Forms.Label lblImageCaptionBold;

    // Open/Save Dialogs
    private System.Windows.Forms.OpenFileDialog openFileDialog;
    private System.Windows.Forms.SaveFileDialog saveFileDialog;

    private void InitializeComponent()
    {
        btnOpen = new Button();
        btnApply = new Button();
        btnSave = new Button();
        grpHeading1 = new GroupBox();
        lblHeading1Font = new Label();
        cmbHeading1Font = new ComboBox();
        lblHeading1Size = new Label();
        cmbHeading1Size = new ComboBox();
        lblHeading1Before = new Label();
        nudHeading1SpaceBefore = new NumericUpDown();
        lblHeading1After = new Label();
        nudHeading1SpaceAfter = new NumericUpDown();
        lblHeading1LineSpacing = new Label();
        nudHeading1LineSpacing = new NumericUpDown();
        grpHeading2 = new GroupBox();
        lblHeading2Font = new Label();
        cmbHeading2Font = new ComboBox();
        lblHeading2Size = new Label();
        cmbHeading2Size = new ComboBox();
        lblHeading2Before = new Label();
        nudHeading2SpaceBefore = new NumericUpDown();
        lblHeading2After = new Label();
        nudHeading2SpaceAfter = new NumericUpDown();
        lblHeading2LineSpacing = new Label();
        nudHeading2LineSpacing = new NumericUpDown();
        grpHeading3 = new GroupBox();
        lblHeading3Font = new Label();
        cmbHeading3Font = new ComboBox();
        lblHeading3Size = new Label();
        cmbHeading3Size = new ComboBox();
        lblHeading3Before = new Label();
        nudHeading3SpaceBefore = new NumericUpDown();
        lblHeading3After = new Label();
        nudHeading3SpaceAfter = new NumericUpDown();
        lblHeading3LineSpacing = new Label();
        nudHeading3LineSpacing = new NumericUpDown();
        grpHeading4 = new GroupBox();
        lblHeading4Font = new Label();
        cmbHeading4Font = new ComboBox();
        lblHeading4Size = new Label();
        cmbHeading4Size = new ComboBox();
        lblHeading4Before = new Label();
        nudHeading4SpaceBefore = new NumericUpDown();
        lblHeading4After = new Label();
        nudHeading4SpaceAfter = new NumericUpDown();
        lblHeading4LineSpacing = new Label();
        nudHeading4LineSpacing = new NumericUpDown();
        grpNormal = new GroupBox();
        lblNormalFont = new Label();
        cmbNormalFont = new ComboBox();
        lblNormalSize = new Label();
        cmbNormalSize = new ComboBox();
        lblNormalLineSpacing = new Label();
        nudNormalLineSpacing = new NumericUpDown();
        grpTableCaption = new GroupBox();
        lblTableCaptionFont = new Label();
        cmbTableCaptionFont = new ComboBox();
        lblTableCaptionSize = new Label();
        cmbTableCaptionSize = new ComboBox();
        lblTableCaptionOutlineLevel = new Label();
        cmbTableCaptionOutlineLevel = new ComboBox();
        lblTableCaptionBold = new Label();
        chkTableCaptionBold = new CheckBox();
        txtFileName = new TextBox();
        txtLog = new TextBox();
        lblStatus = new Label();
        lblProgress = new Label();
        progressBar = new ProgressBar();
        openFileDialog = new OpenFileDialog();
        saveFileDialog = new SaveFileDialog();
        topPanel = new Panel();
        stylePanel = new Panel();
        grpImageCaption = new GroupBox();
        lblImageCaptionFont = new Label();
        cmbImageCaptionFont = new ComboBox();
        lblImageCaptionSize = new Label();
        cmbImageCaptionSize = new ComboBox();
        lblImageCaptionOutlineLevel = new Label();
        cmbImageCaptionOutlineLevel = new ComboBox();
        lblImageCaptionBold = new Label();
        chkImageCaptionBold = new CheckBox();
        statusPanel = new Panel();
        lblStatusTitle = new Label();
        logPanel = new Panel();
        grpHeading1.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)nudHeading1SpaceBefore).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading1SpaceAfter).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading1LineSpacing).BeginInit();
        grpHeading2.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)nudHeading2SpaceBefore).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading2SpaceAfter).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading2LineSpacing).BeginInit();
        grpHeading3.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)nudHeading3SpaceBefore).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading3SpaceAfter).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading3LineSpacing).BeginInit();
        grpHeading4.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)nudHeading4SpaceBefore).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading4SpaceAfter).BeginInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading4LineSpacing).BeginInit();
        grpNormal.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)nudNormalLineSpacing).BeginInit();
        grpTableCaption.SuspendLayout();
        topPanel.SuspendLayout();
        stylePanel.SuspendLayout();
        grpImageCaption.SuspendLayout();
        statusPanel.SuspendLayout();
        logPanel.SuspendLayout();
        SuspendLayout();
        // 
        // btnOpen
        // 
        btnOpen.BackColor = Color.FromArgb(0, 120, 212);
        btnOpen.Cursor = Cursors.Hand;
        btnOpen.FlatStyle = FlatStyle.Flat;
        btnOpen.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
        btnOpen.ForeColor = Color.White;
        btnOpen.Location = new Point(12, 12);
        btnOpen.Name = "btnOpen";
        btnOpen.Size = new Size(100, 35);
        btnOpen.TabIndex = 0;
        btnOpen.Text = "打开文档";
        btnOpen.UseVisualStyleBackColor = false;
        // 
        // btnApply
        // 
        btnApply.BackColor = Color.FromArgb(16, 110, 190);
        btnApply.Cursor = Cursors.Hand;
        btnApply.FlatStyle = FlatStyle.Flat;
        btnApply.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
        btnApply.ForeColor = Color.White;
        btnApply.Location = new Point(118, 12);
        btnApply.Name = "btnApply";
        btnApply.Size = new Size(100, 35);
        btnApply.TabIndex = 1;
        btnApply.Text = "应用格式";
        btnApply.UseVisualStyleBackColor = false;
        // 
        // btnSave
        // 
        btnSave.BackColor = Color.FromArgb(16, 142, 234);
        btnSave.Cursor = Cursors.Hand;
        btnSave.FlatStyle = FlatStyle.Flat;
        btnSave.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
        btnSave.ForeColor = Color.White;
        btnSave.Location = new Point(224, 12);
        btnSave.Name = "btnSave";
        btnSave.Size = new Size(100, 35);
        btnSave.TabIndex = 2;
        btnSave.Text = "保存文档";
        btnSave.UseVisualStyleBackColor = false;
        btnSave.Click += btnSave_Click_1;
        // 
        // grpHeading1
        // 
        grpHeading1.Controls.Add(lblHeading1Font);
        grpHeading1.Controls.Add(cmbHeading1Font);
        grpHeading1.Controls.Add(lblHeading1Size);
        grpHeading1.Controls.Add(cmbHeading1Size);
        grpHeading1.Controls.Add(lblHeading1Before);
        grpHeading1.Controls.Add(nudHeading1SpaceBefore);
        grpHeading1.Controls.Add(lblHeading1After);
        grpHeading1.Controls.Add(nudHeading1SpaceAfter);
        grpHeading1.Controls.Add(lblHeading1LineSpacing);
        grpHeading1.Controls.Add(nudHeading1LineSpacing);
        grpHeading1.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
        grpHeading1.Location = new Point(12, 12);
        grpHeading1.Name = "grpHeading1";
        grpHeading1.Size = new Size(956, 65);
        grpHeading1.TabIndex = 0;
        grpHeading1.TabStop = false;
        grpHeading1.Text = "一级标题";
        // 
        // lblHeading1Font
        // 
        lblHeading1Font.Location = new Point(12, 25);
        lblHeading1Font.Name = "lblHeading1Font";
        lblHeading1Font.Size = new Size(50, 20);
        lblHeading1Font.TabIndex = 0;
        lblHeading1Font.Text = "字体:";
        // 
        // cmbHeading1Font
        // 
        cmbHeading1Font.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbHeading1Font.Items.AddRange(new object[] { "宋体", "黑体", "仿宋", "楷体", "Arial", "Times New Roman" });
        cmbHeading1Font.Location = new Point(68, 22);
        cmbHeading1Font.Name = "cmbHeading1Font";
        cmbHeading1Font.Size = new Size(150, 25);
        cmbHeading1Font.TabIndex = 1;
        // 
        // lblHeading1Size
        // 
        lblHeading1Size.Location = new Point(230, 25);
        lblHeading1Size.Name = "lblHeading1Size";
        lblHeading1Size.Size = new Size(50, 20);
        lblHeading1Size.TabIndex = 2;
        lblHeading1Size.Text = "字号:";
        // 
        // cmbHeading1Size
        // 
        cmbHeading1Size.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbHeading1Size.Items.AddRange(new object[] { "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六", "七号" });
        cmbHeading1Size.Location = new Point(286, 22);
        cmbHeading1Size.Name = "cmbHeading1Size";
        cmbHeading1Size.Size = new Size(80, 25);
        cmbHeading1Size.TabIndex = 3;
        // 
        // lblHeading1Before
        // 
        lblHeading1Before.Location = new Point(442, 24);
        lblHeading1Before.Name = "lblHeading1Before";
        lblHeading1Before.Size = new Size(67, 20);
        lblHeading1Before.TabIndex = 4;
        lblHeading1Before.Text = "段前(磅):";
        // 
        // nudHeading1SpaceBefore
        // 
        nudHeading1SpaceBefore.Location = new Point(520, 21);
        nudHeading1SpaceBefore.Maximum = new decimal(new int[] { 72, 0, 0, 0 });
        nudHeading1SpaceBefore.Name = "nudHeading1SpaceBefore";
        nudHeading1SpaceBefore.Size = new Size(45, 23);
        nudHeading1SpaceBefore.TabIndex = 5;
        nudHeading1SpaceBefore.Value = new decimal(new int[] { 24, 0, 0, 0 });
        // 
        // lblHeading1After
        // 
        lblHeading1After.Location = new Point(602, 23);
        lblHeading1After.Name = "lblHeading1After";
        lblHeading1After.Size = new Size(60, 20);
        lblHeading1After.TabIndex = 6;
        lblHeading1After.Text = "段后(磅):";
        // 
        // nudHeading1SpaceAfter
        // 
        nudHeading1SpaceAfter.Location = new Point(668, 21);
        nudHeading1SpaceAfter.Maximum = new decimal(new int[] { 72, 0, 0, 0 });
        nudHeading1SpaceAfter.Name = "nudHeading1SpaceAfter";
        nudHeading1SpaceAfter.Size = new Size(45, 23);
        nudHeading1SpaceAfter.TabIndex = 7;
        nudHeading1SpaceAfter.Value = new decimal(new int[] { 24, 0, 0, 0 });
        // 
        // lblHeading1LineSpacing
        // 
        lblHeading1LineSpacing.Location = new Point(775, 22);
        lblHeading1LineSpacing.Name = "lblHeading1LineSpacing";
        lblHeading1LineSpacing.Size = new Size(60, 20);
        lblHeading1LineSpacing.TabIndex = 8;
        lblHeading1LineSpacing.Text = "行距(磅):";
        // 
        // nudHeading1LineSpacing
        // 
        nudHeading1LineSpacing.DecimalPlaces = 1;
        nudHeading1LineSpacing.Increment = new decimal(new int[] { 5, 0, 0, 65536 });
        nudHeading1LineSpacing.Location = new Point(841, 21);
        nudHeading1LineSpacing.Name = "nudHeading1LineSpacing";
        nudHeading1LineSpacing.Size = new Size(60, 23);
        nudHeading1LineSpacing.TabIndex = 9;
        nudHeading1LineSpacing.Value = new decimal(new int[] { 24, 0, 0, 0 });
        // 
        // grpHeading2
        // 
        grpHeading2.Controls.Add(lblHeading2Font);
        grpHeading2.Controls.Add(cmbHeading2Font);
        grpHeading2.Controls.Add(lblHeading2Size);
        grpHeading2.Controls.Add(cmbHeading2Size);
        grpHeading2.Controls.Add(lblHeading2Before);
        grpHeading2.Controls.Add(nudHeading2SpaceBefore);
        grpHeading2.Controls.Add(lblHeading2After);
        grpHeading2.Controls.Add(nudHeading2SpaceAfter);
        grpHeading2.Controls.Add(lblHeading2LineSpacing);
        grpHeading2.Controls.Add(nudHeading2LineSpacing);
        grpHeading2.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
        grpHeading2.Location = new Point(12, 83);
        grpHeading2.Name = "grpHeading2";
        grpHeading2.Size = new Size(956, 65);
        grpHeading2.TabIndex = 1;
        grpHeading2.TabStop = false;
        grpHeading2.Text = "二级标题";
        // 
        // lblHeading2Font
        // 
        lblHeading2Font.Location = new Point(12, 25);
        lblHeading2Font.Name = "lblHeading2Font";
        lblHeading2Font.Size = new Size(50, 20);
        lblHeading2Font.TabIndex = 0;
        lblHeading2Font.Text = "字体:";
        // 
        // cmbHeading2Font
        // 
        cmbHeading2Font.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbHeading2Font.Items.AddRange(new object[] { "宋体", "黑体", "仿宋", "楷体", "Arial", "Times New Roman" });
        cmbHeading2Font.Location = new Point(68, 22);
        cmbHeading2Font.Name = "cmbHeading2Font";
        cmbHeading2Font.Size = new Size(150, 25);
        cmbHeading2Font.TabIndex = 1;
        // 
        // lblHeading2Size
        // 
        lblHeading2Size.Location = new Point(230, 25);
        lblHeading2Size.Name = "lblHeading2Size";
        lblHeading2Size.Size = new Size(50, 20);
        lblHeading2Size.TabIndex = 2;
        lblHeading2Size.Text = "字号:";
        // 
        // cmbHeading2Size
        // 
        cmbHeading2Size.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbHeading2Size.Items.AddRange(new object[] { "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六", "七号" });
        cmbHeading2Size.Location = new Point(286, 22);
        cmbHeading2Size.Name = "cmbHeading2Size";
        cmbHeading2Size.Size = new Size(80, 25);
        cmbHeading2Size.TabIndex = 3;
        // 
        // lblHeading2Before
        // 
        lblHeading2Before.Location = new Point(442, 22);
        lblHeading2Before.Name = "lblHeading2Before";
        lblHeading2Before.Size = new Size(67, 20);
        lblHeading2Before.TabIndex = 4;
        lblHeading2Before.Text = "段前(磅):";
        // 
        // nudHeading2SpaceBefore
        // 
        nudHeading2SpaceBefore.Location = new Point(520, 21);
        nudHeading2SpaceBefore.Maximum = new decimal(new int[] { 72, 0, 0, 0 });
        nudHeading2SpaceBefore.Name = "nudHeading2SpaceBefore";
        nudHeading2SpaceBefore.Size = new Size(45, 23);
        nudHeading2SpaceBefore.TabIndex = 5;
        nudHeading2SpaceBefore.Value = new decimal(new int[] { 12, 0, 0, 0 });
        // 
        // lblHeading2After
        // 
        lblHeading2After.Location = new Point(602, 23);
        lblHeading2After.Name = "lblHeading2After";
        lblHeading2After.Size = new Size(58, 20);
        lblHeading2After.TabIndex = 6;
        lblHeading2After.Text = "段后(磅):";
        // 
        // nudHeading2SpaceAfter
        // 
        nudHeading2SpaceAfter.Location = new Point(668, 20);
        nudHeading2SpaceAfter.Maximum = new decimal(new int[] { 72, 0, 0, 0 });
        nudHeading2SpaceAfter.Name = "nudHeading2SpaceAfter";
        nudHeading2SpaceAfter.Size = new Size(45, 23);
        nudHeading2SpaceAfter.TabIndex = 7;
        nudHeading2SpaceAfter.Value = new decimal(new int[] { 12, 0, 0, 0 });
        // 
        // lblHeading2LineSpacing
        // 
        lblHeading2LineSpacing.Location = new Point(775, 23);
        lblHeading2LineSpacing.Name = "lblHeading2LineSpacing";
        lblHeading2LineSpacing.Size = new Size(59, 20);
        lblHeading2LineSpacing.TabIndex = 8;
        lblHeading2LineSpacing.Text = "行距(磅):";
        // 
        // nudHeading2LineSpacing
        // 
        nudHeading2LineSpacing.DecimalPlaces = 1;
        nudHeading2LineSpacing.Increment = new decimal(new int[] { 5, 0, 0, 65536 });
        nudHeading2LineSpacing.Location = new Point(841, 21);
        nudHeading2LineSpacing.Name = "nudHeading2LineSpacing";
        nudHeading2LineSpacing.Size = new Size(60, 23);
        nudHeading2LineSpacing.TabIndex = 9;
        nudHeading2LineSpacing.Value = new decimal(new int[] { 24, 0, 0, 0 });
        // 
        // grpHeading3
        // 
        grpHeading3.Controls.Add(lblHeading3Font);
        grpHeading3.Controls.Add(cmbHeading3Font);
        grpHeading3.Controls.Add(lblHeading3Size);
        grpHeading3.Controls.Add(cmbHeading3Size);
        grpHeading3.Controls.Add(lblHeading3Before);
        grpHeading3.Controls.Add(nudHeading3SpaceBefore);
        grpHeading3.Controls.Add(lblHeading3After);
        grpHeading3.Controls.Add(nudHeading3SpaceAfter);
        grpHeading3.Controls.Add(lblHeading3LineSpacing);
        grpHeading3.Controls.Add(nudHeading3LineSpacing);
        grpHeading3.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
        grpHeading3.Location = new Point(12, 154);
        grpHeading3.Name = "grpHeading3";
        grpHeading3.Size = new Size(956, 65);
        grpHeading3.TabIndex = 2;
        grpHeading3.TabStop = false;
        grpHeading3.Text = "三级标题 ";
        // 
        // lblHeading3Font
        // 
        lblHeading3Font.Location = new Point(12, 25);
        lblHeading3Font.Name = "lblHeading3Font";
        lblHeading3Font.Size = new Size(50, 20);
        lblHeading3Font.TabIndex = 0;
        lblHeading3Font.Text = "字体:";
        // 
        // cmbHeading3Font
        // 
        cmbHeading3Font.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbHeading3Font.Items.AddRange(new object[] { "宋体", "黑体", "仿宋", "楷体", "Arial", "Times New Roman" });
        cmbHeading3Font.Location = new Point(68, 22);
        cmbHeading3Font.Name = "cmbHeading3Font";
        cmbHeading3Font.Size = new Size(150, 25);
        cmbHeading3Font.TabIndex = 1;
        // 
        // lblHeading3Size
        // 
        lblHeading3Size.Location = new Point(230, 25);
        lblHeading3Size.Name = "lblHeading3Size";
        lblHeading3Size.Size = new Size(50, 20);
        lblHeading3Size.TabIndex = 2;
        lblHeading3Size.Text = "字号:";
        // 
        // cmbHeading3Size
        // 
        cmbHeading3Size.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbHeading3Size.Items.AddRange(new object[] { "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六", "七号" });
        cmbHeading3Size.Location = new Point(286, 22);
        cmbHeading3Size.Name = "cmbHeading3Size";
        cmbHeading3Size.Size = new Size(80, 25);
        cmbHeading3Size.TabIndex = 3;
        // 
        // lblHeading3Before
        // 
        lblHeading3Before.Location = new Point(442, 24);
        lblHeading3Before.Name = "lblHeading3Before";
        lblHeading3Before.Size = new Size(60, 20);
        lblHeading3Before.TabIndex = 4;
        lblHeading3Before.Text = "段前(磅):";
        // 
        // nudHeading3SpaceBefore
        // 
        nudHeading3SpaceBefore.Location = new Point(520, 23);
        nudHeading3SpaceBefore.Maximum = new decimal(new int[] { 72, 0, 0, 0 });
        nudHeading3SpaceBefore.Name = "nudHeading3SpaceBefore";
        nudHeading3SpaceBefore.Size = new Size(45, 23);
        nudHeading3SpaceBefore.TabIndex = 5;
        nudHeading3SpaceBefore.Value = new decimal(new int[] { 12, 0, 0, 0 });
        // 
        // lblHeading3After
        // 
        lblHeading3After.Location = new Point(604, 27);
        lblHeading3After.Name = "lblHeading3After";
        lblHeading3After.Size = new Size(58, 20);
        lblHeading3After.TabIndex = 6;
        lblHeading3After.Text = "段后(磅):";
        // 
        // nudHeading3SpaceAfter
        // 
        nudHeading3SpaceAfter.Location = new Point(668, 25);
        nudHeading3SpaceAfter.Maximum = new decimal(new int[] { 72, 0, 0, 0 });
        nudHeading3SpaceAfter.Name = "nudHeading3SpaceAfter";
        nudHeading3SpaceAfter.Size = new Size(45, 23);
        nudHeading3SpaceAfter.TabIndex = 7;
        nudHeading3SpaceAfter.Value = new decimal(new int[] { 12, 0, 0, 0 });
        // 
        // lblHeading3LineSpacing
        // 
        lblHeading3LineSpacing.Location = new Point(775, 27);
        lblHeading3LineSpacing.Name = "lblHeading3LineSpacing";
        lblHeading3LineSpacing.Size = new Size(60, 20);
        lblHeading3LineSpacing.TabIndex = 8;
        lblHeading3LineSpacing.Text = "行距(磅):";
        // 
        // nudHeading3LineSpacing
        // 
        nudHeading3LineSpacing.DecimalPlaces = 1;
        nudHeading3LineSpacing.Increment = new decimal(new int[] { 5, 0, 0, 65536 });
        nudHeading3LineSpacing.Location = new Point(841, 25);
        nudHeading3LineSpacing.Name = "nudHeading3LineSpacing";
        nudHeading3LineSpacing.Size = new Size(60, 23);
        nudHeading3LineSpacing.TabIndex = 9;
        nudHeading3LineSpacing.Value = new decimal(new int[] { 24, 0, 0, 0 });
        // 
        // grpHeading4
        // 
        grpHeading4.Controls.Add(lblHeading4Font);
        grpHeading4.Controls.Add(cmbHeading4Font);
        grpHeading4.Controls.Add(lblHeading4Size);
        grpHeading4.Controls.Add(cmbHeading4Size);
        grpHeading4.Controls.Add(lblHeading4Before);
        grpHeading4.Controls.Add(nudHeading4SpaceBefore);
        grpHeading4.Controls.Add(lblHeading4After);
        grpHeading4.Controls.Add(nudHeading4SpaceAfter);
        grpHeading4.Controls.Add(lblHeading4LineSpacing);
        grpHeading4.Controls.Add(nudHeading4LineSpacing);
        grpHeading4.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
        grpHeading4.Location = new Point(12, 225);
        grpHeading4.Name = "grpHeading4";
        grpHeading4.Size = new Size(956, 65);
        grpHeading4.TabIndex = 3;
        grpHeading4.TabStop = false;
        grpHeading4.Text = "四级标题 ";
        // 
        // lblHeading4Font
        // 
        lblHeading4Font.Location = new Point(12, 25);
        lblHeading4Font.Name = "lblHeading4Font";
        lblHeading4Font.Size = new Size(50, 20);
        lblHeading4Font.TabIndex = 0;
        lblHeading4Font.Text = "字体:";
        // 
        // cmbHeading4Font
        // 
        cmbHeading4Font.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbHeading4Font.Items.AddRange(new object[] { "宋体", "黑体", "仿宋", "楷体", "Arial", "Times New Roman" });
        cmbHeading4Font.Location = new Point(68, 22);
        cmbHeading4Font.Name = "cmbHeading4Font";
        cmbHeading4Font.Size = new Size(150, 25);
        cmbHeading4Font.TabIndex = 1;
        // 
        // lblHeading4Size
        // 
        lblHeading4Size.Location = new Point(230, 25);
        lblHeading4Size.Name = "lblHeading4Size";
        lblHeading4Size.Size = new Size(50, 20);
        lblHeading4Size.TabIndex = 2;
        lblHeading4Size.Text = "字号:";
        // 
        // cmbHeading4Size
        // 
        cmbHeading4Size.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbHeading4Size.Items.AddRange(new object[] { "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六", "七号" });
        cmbHeading4Size.Location = new Point(286, 22);
        cmbHeading4Size.Name = "cmbHeading4Size";
        cmbHeading4Size.Size = new Size(80, 25);
        cmbHeading4Size.TabIndex = 3;
        // 
        // lblHeading4Before
        // 
        lblHeading4Before.Location = new Point(442, 24);
        lblHeading4Before.Name = "lblHeading4Before";
        lblHeading4Before.Size = new Size(60, 20);
        lblHeading4Before.TabIndex = 4;
        lblHeading4Before.Text = "段前(磅):";
        lblHeading4Before.Click += lblHeading4Before_Click;
        // 
        // nudHeading4SpaceBefore
        // 
        nudHeading4SpaceBefore.Location = new Point(520, 22);
        nudHeading4SpaceBefore.Maximum = new decimal(new int[] { 72, 0, 0, 0 });
        nudHeading4SpaceBefore.Name = "nudHeading4SpaceBefore";
        nudHeading4SpaceBefore.Size = new Size(45, 23);
        nudHeading4SpaceBefore.TabIndex = 5;
        // 
        // lblHeading4After
        // 
        lblHeading4After.Location = new Point(599, 25);
        lblHeading4After.Name = "lblHeading4After";
        lblHeading4After.Size = new Size(61, 20);
        lblHeading4After.TabIndex = 6;
        lblHeading4After.Text = "段后(磅):";
        // 
        // nudHeading4SpaceAfter
        // 
        nudHeading4SpaceAfter.Location = new Point(668, 23);
        nudHeading4SpaceAfter.Maximum = new decimal(new int[] { 72, 0, 0, 0 });
        nudHeading4SpaceAfter.Name = "nudHeading4SpaceAfter";
        nudHeading4SpaceAfter.Size = new Size(45, 23);
        nudHeading4SpaceAfter.TabIndex = 7;
        // 
        // lblHeading4LineSpacing
        // 
        lblHeading4LineSpacing.Location = new Point(775, 27);
        lblHeading4LineSpacing.Name = "lblHeading4LineSpacing";
        lblHeading4LineSpacing.Size = new Size(60, 20);
        lblHeading4LineSpacing.TabIndex = 8;
        lblHeading4LineSpacing.Text = "行距(磅):";
        // 
        // nudHeading4LineSpacing
        // 
        nudHeading4LineSpacing.DecimalPlaces = 1;
        nudHeading4LineSpacing.Increment = new decimal(new int[] { 5, 0, 0, 65536 });
        nudHeading4LineSpacing.Location = new Point(841, 25);
        nudHeading4LineSpacing.Name = "nudHeading4LineSpacing";
        nudHeading4LineSpacing.Size = new Size(60, 23);
        nudHeading4LineSpacing.TabIndex = 9;
        nudHeading4LineSpacing.Value = new decimal(new int[] { 24, 0, 0, 0 });
        // 
        // grpNormal
        // 
        grpNormal.Controls.Add(lblNormalFont);
        grpNormal.Controls.Add(cmbNormalFont);
        grpNormal.Controls.Add(lblNormalSize);
        grpNormal.Controls.Add(cmbNormalSize);
        grpNormal.Controls.Add(lblNormalLineSpacing);
        grpNormal.Controls.Add(nudNormalLineSpacing);
        grpNormal.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
        grpNormal.Location = new Point(12, 435);
        grpNormal.Name = "grpNormal";
        grpNormal.Size = new Size(956, 50);
        grpNormal.TabIndex = 4;
        grpNormal.TabStop = false;
        grpNormal.Text = "正文段落";
        // 
        // lblNormalFont
        // 
        lblNormalFont.Location = new Point(12, 22);
        lblNormalFont.Name = "lblNormalFont";
        lblNormalFont.Size = new Size(50, 20);
        lblNormalFont.TabIndex = 0;
        lblNormalFont.Text = "字体:";
        // 
        // cmbNormalFont
        // 
        cmbNormalFont.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbNormalFont.Items.AddRange(new object[] { "宋体", "黑体", "仿宋", "楷体", "Arial", "Times New Roman" });
        cmbNormalFont.Location = new Point(68, 18);
        cmbNormalFont.Name = "cmbNormalFont";
        cmbNormalFont.Size = new Size(150, 25);
        cmbNormalFont.TabIndex = 1;
        // 
        // lblNormalSize
        // 
        lblNormalSize.Location = new Point(230, 22);
        lblNormalSize.Name = "lblNormalSize";
        lblNormalSize.Size = new Size(50, 20);
        lblNormalSize.TabIndex = 2;
        lblNormalSize.Text = "字号:";
        // 
        // cmbNormalSize
        // 
        cmbNormalSize.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbNormalSize.Items.AddRange(new object[] { "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六", "七号" });
        cmbNormalSize.Location = new Point(286, 17);
        cmbNormalSize.Name = "cmbNormalSize";
        cmbNormalSize.Size = new Size(80, 25);
        cmbNormalSize.TabIndex = 3;
        // 
        // lblNormalLineSpacing
        // 
        lblNormalLineSpacing.Location = new Point(441, 19);
        lblNormalLineSpacing.Name = "lblNormalLineSpacing";
        lblNormalLineSpacing.Size = new Size(58, 20);
        lblNormalLineSpacing.TabIndex = 4;
        lblNormalLineSpacing.Text = "行距(磅):";
        // 
        // nudNormalLineSpacing
        // 
        nudNormalLineSpacing.DecimalPlaces = 1;
        nudNormalLineSpacing.Increment = new decimal(new int[] { 5, 0, 0, 65536 });
        nudNormalLineSpacing.Location = new Point(518, 17);
        nudNormalLineSpacing.Name = "nudNormalLineSpacing";
        nudNormalLineSpacing.Size = new Size(60, 23);
        nudNormalLineSpacing.TabIndex = 5;
        nudNormalLineSpacing.Value = new decimal(new int[] { 24, 0, 0, 0 });
        // 
        // grpTableCaption
        // 
        grpTableCaption.Controls.Add(lblTableCaptionFont);
        grpTableCaption.Controls.Add(cmbTableCaptionFont);
        grpTableCaption.Controls.Add(lblTableCaptionSize);
        grpTableCaption.Controls.Add(cmbTableCaptionSize);
        grpTableCaption.Controls.Add(lblTableCaptionOutlineLevel);
        grpTableCaption.Controls.Add(cmbTableCaptionOutlineLevel);
        grpTableCaption.Controls.Add(lblTableCaptionBold);
        grpTableCaption.Controls.Add(chkTableCaptionBold);
        grpTableCaption.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
        grpTableCaption.Location = new Point(12, 297);
        grpTableCaption.Name = "grpTableCaption";
        grpTableCaption.Size = new Size(956, 65);
        grpTableCaption.TabIndex = 5;
        grpTableCaption.TabStop = false;
        grpTableCaption.Text = "表格标题";
        // 
        // lblTableCaptionFont
        // 
        lblTableCaptionFont.Location = new Point(12, 25);
        lblTableCaptionFont.Name = "lblTableCaptionFont";
        lblTableCaptionFont.Size = new Size(50, 20);
        lblTableCaptionFont.TabIndex = 0;
        lblTableCaptionFont.Text = "字体:";
        // 
        // cmbTableCaptionFont
        // 
        cmbTableCaptionFont.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbTableCaptionFont.Items.AddRange(new object[] { "黑体", "宋体", "仿宋", "楷体", "Arial", "Times New Roman" });
        cmbTableCaptionFont.Location = new Point(68, 22);
        cmbTableCaptionFont.Name = "cmbTableCaptionFont";
        cmbTableCaptionFont.Size = new Size(150, 25);
        cmbTableCaptionFont.TabIndex = 1;
        // 
        // lblTableCaptionSize
        // 
        lblTableCaptionSize.Location = new Point(230, 25);
        lblTableCaptionSize.Name = "lblTableCaptionSize";
        lblTableCaptionSize.Size = new Size(50, 20);
        lblTableCaptionSize.TabIndex = 2;
        lblTableCaptionSize.Text = "字号:";
        // 
        // cmbTableCaptionSize
        // 
        cmbTableCaptionSize.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbTableCaptionSize.Items.AddRange(new object[] { "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六", "七号" });
        cmbTableCaptionSize.Location = new Point(286, 22);
        cmbTableCaptionSize.Name = "cmbTableCaptionSize";
        cmbTableCaptionSize.Size = new Size(80, 25);
        cmbTableCaptionSize.TabIndex = 3;
        // 
        // lblTableCaptionOutlineLevel
        // 
        lblTableCaptionOutlineLevel.Location = new Point(442, 25);
        lblTableCaptionOutlineLevel.Name = "lblTableCaptionOutlineLevel";
        lblTableCaptionOutlineLevel.Size = new Size(70, 20);
        lblTableCaptionOutlineLevel.TabIndex = 4;
        lblTableCaptionOutlineLevel.Text = "大纲级别:";
        // 
        // cmbTableCaptionOutlineLevel
        // 
        cmbTableCaptionOutlineLevel.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbTableCaptionOutlineLevel.Items.AddRange(new object[] { "无(正文)", "1级", "2级", "3级", "4级", "5级", "6级", "7级", "8级", "9级" });
        cmbTableCaptionOutlineLevel.Location = new Point(518, 22);
        cmbTableCaptionOutlineLevel.Name = "cmbTableCaptionOutlineLevel";
        cmbTableCaptionOutlineLevel.Size = new Size(80, 25);
        cmbTableCaptionOutlineLevel.TabIndex = 5;
        // 
        // lblTableCaptionBold
        // 
        lblTableCaptionBold.AutoSize = true;
        lblTableCaptionBold.Location = new Point(614, 25);
        lblTableCaptionBold.Name = "lblTableCaptionBold";
        lblTableCaptionBold.Size = new Size(35, 17);
        lblTableCaptionBold.TabIndex = 6;
        lblTableCaptionBold.Text = "加粗:";
        // 
        // chkTableCaptionBold
        // 
        chkTableCaptionBold.AutoSize = true;
        chkTableCaptionBold.Location = new Point(661, 27);
        chkTableCaptionBold.Name = "chkTableCaptionBold";
        chkTableCaptionBold.Size = new Size(15, 14);
        chkTableCaptionBold.TabIndex = 7;
        chkTableCaptionBold.UseVisualStyleBackColor = true;
        // 
        // txtFileName
        // 
        txtFileName.BackColor = Color.FromArgb(240, 240, 240);
        txtFileName.BorderStyle = BorderStyle.FixedSingle;
        txtFileName.Location = new Point(332, 18);
        txtFileName.Name = "txtFileName";
        txtFileName.ReadOnly = true;
        txtFileName.Size = new Size(659, 23);
        txtFileName.TabIndex = 3;
        txtFileName.Text = "未选择文件";
        // 
        // txtLog
        // 
        txtLog.BackColor = Color.FromArgb(250, 250, 250);
        txtLog.BorderStyle = BorderStyle.FixedSingle;
        txtLog.Font = new Font("Consolas", 8.5F);
        txtLog.Location = new Point(3, 35);
        txtLog.Multiline = true;
        txtLog.Name = "txtLog";
        txtLog.ReadOnly = true;
        txtLog.ScrollBars = ScrollBars.Vertical;
        txtLog.Size = new Size(978, 141);
        txtLog.TabIndex = 0;
        // 
        // lblStatus
        // 
        lblStatus.ForeColor = Color.FromArgb(16, 142, 234);
        lblStatus.Location = new Point(12, 782);
        lblStatus.Name = "lblStatus";
        lblStatus.Size = new Size(290, 20);
        lblStatus.TabIndex = 1;
        lblStatus.Text = "就绪";
        // 
        // lblProgress
        // 
        lblProgress.Location = new Point(12, 20);
        lblProgress.Name = "lblProgress";
        lblProgress.Size = new Size(290, 20);
        lblProgress.TabIndex = 2;
        // 
        // progressBar
        // 
        progressBar.Location = new Point(3, 8);
        progressBar.Name = "progressBar";
        progressBar.Size = new Size(978, 17);
        progressBar.Style = ProgressBarStyle.Continuous;
        progressBar.TabIndex = 3;
        progressBar.Visible = false;
        // 
        // openFileDialog
        // 
        openFileDialog.Filter = "Word 文档|*.docx|所有文件|*.*";
        openFileDialog.Title = "选择 Word 文档";
        // 
        // saveFileDialog
        // 
        saveFileDialog.Filter = "Word 文档|*.docx|所有文件|*.*";
        saveFileDialog.Title = "保存 Word 文档";
        // 
        // topPanel
        // 
        topPanel.BackColor = Color.FromArgb(255, 255, 255);
        topPanel.BorderStyle = BorderStyle.FixedSingle;
        topPanel.Controls.Add(btnOpen);
        topPanel.Controls.Add(btnApply);
        topPanel.Controls.Add(btnSave);
        topPanel.Controls.Add(txtFileName);
        topPanel.Location = new Point(12, 12);
        topPanel.Name = "topPanel";
        topPanel.Size = new Size(1004, 60);
        topPanel.TabIndex = 0;
        // 
        // stylePanel
        // 
        stylePanel.AutoScroll = true;
        stylePanel.BackColor = Color.FromArgb(255, 255, 255);
        stylePanel.BorderStyle = BorderStyle.FixedSingle;
        stylePanel.Controls.Add(grpHeading1);
        stylePanel.Controls.Add(grpHeading2);
        stylePanel.Controls.Add(grpHeading3);
        stylePanel.Controls.Add(grpHeading4);
        stylePanel.Controls.Add(grpNormal);
        stylePanel.Controls.Add(grpTableCaption);
        stylePanel.Controls.Add(grpImageCaption);
        stylePanel.Location = new Point(12, 78);
        stylePanel.Name = "stylePanel";
        stylePanel.Size = new Size(1004, 501);
        stylePanel.TabIndex = 1;
        // 
        // grpImageCaption
        // 
        grpImageCaption.Controls.Add(lblImageCaptionFont);
        grpImageCaption.Controls.Add(cmbImageCaptionFont);
        grpImageCaption.Controls.Add(lblImageCaptionSize);
        grpImageCaption.Controls.Add(cmbImageCaptionSize);
        grpImageCaption.Controls.Add(lblImageCaptionOutlineLevel);
        grpImageCaption.Controls.Add(cmbImageCaptionOutlineLevel);
        grpImageCaption.Controls.Add(lblImageCaptionBold);
        grpImageCaption.Controls.Add(chkImageCaptionBold);
        grpImageCaption.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
        grpImageCaption.Location = new Point(13, 363);
        grpImageCaption.Name = "grpImageCaption";
        grpImageCaption.Size = new Size(955, 65);
        grpImageCaption.TabIndex = 6;
        grpImageCaption.TabStop = false;
        grpImageCaption.Text = "图形标题";
        // 
        // lblImageCaptionFont
        // 
        lblImageCaptionFont.Location = new Point(12, 25);
        lblImageCaptionFont.Name = "lblImageCaptionFont";
        lblImageCaptionFont.Size = new Size(50, 20);
        lblImageCaptionFont.TabIndex = 0;
        lblImageCaptionFont.Text = "字体:";
        // 
        // cmbImageCaptionFont
        // 
        cmbImageCaptionFont.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbImageCaptionFont.Items.AddRange(new object[] { "黑体", "宋体", "仿宋", "楷体", "Arial", "Times New Roman" });
        cmbImageCaptionFont.Location = new Point(68, 22);
        cmbImageCaptionFont.Name = "cmbImageCaptionFont";
        cmbImageCaptionFont.Size = new Size(150, 25);
        cmbImageCaptionFont.TabIndex = 1;
        // 
        // lblImageCaptionSize
        // 
        lblImageCaptionSize.Location = new Point(230, 25);
        lblImageCaptionSize.Name = "lblImageCaptionSize";
        lblImageCaptionSize.Size = new Size(50, 20);
        lblImageCaptionSize.TabIndex = 2;
        lblImageCaptionSize.Text = "字号:";
        // 
        // cmbImageCaptionSize
        // 
        cmbImageCaptionSize.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbImageCaptionSize.Items.AddRange(new object[] { "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "小六", "七号" });
        cmbImageCaptionSize.Location = new Point(286, 22);
        cmbImageCaptionSize.Name = "cmbImageCaptionSize";
        cmbImageCaptionSize.Size = new Size(80, 25);
        cmbImageCaptionSize.TabIndex = 3;
        // 
        // lblImageCaptionOutlineLevel
        // 
        lblImageCaptionOutlineLevel.Location = new Point(441, 23);
        lblImageCaptionOutlineLevel.Name = "lblImageCaptionOutlineLevel";
        lblImageCaptionOutlineLevel.Size = new Size(70, 20);
        lblImageCaptionOutlineLevel.TabIndex = 4;
        lblImageCaptionOutlineLevel.Text = "大纲级别:";
        // 
        // cmbImageCaptionOutlineLevel
        // 
        cmbImageCaptionOutlineLevel.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbImageCaptionOutlineLevel.Items.AddRange(new object[] { "无(正文)", "1级", "2级", "3级", "4级", "5级", "6级", "7级", "8级", "9级" });
        cmbImageCaptionOutlineLevel.Location = new Point(517, 20);
        cmbImageCaptionOutlineLevel.Name = "cmbImageCaptionOutlineLevel";
        cmbImageCaptionOutlineLevel.Size = new Size(80, 25);
        cmbImageCaptionOutlineLevel.TabIndex = 5;
        // 
        // lblImageCaptionBold
        // 
        lblImageCaptionBold.Location = new Point(613, 23);
        lblImageCaptionBold.Name = "lblImageCaptionBold";
        lblImageCaptionBold.Size = new Size(45, 20);
        lblImageCaptionBold.TabIndex = 6;
        lblImageCaptionBold.Text = "加粗:";
        // 
        // chkImageCaptionBold
        // 
        chkImageCaptionBold.Location = new Point(660, 25);
        chkImageCaptionBold.Name = "chkImageCaptionBold";
        chkImageCaptionBold.Size = new Size(15, 14);
        chkImageCaptionBold.TabIndex = 7;
        chkImageCaptionBold.UseVisualStyleBackColor = true;
        // 
        // statusPanel
        // 
        statusPanel.BackColor = Color.FromArgb(255, 255, 255);
        statusPanel.BorderStyle = BorderStyle.FixedSingle;
        statusPanel.Controls.Add(lblStatusTitle);
        statusPanel.Controls.Add(lblProgress);
        statusPanel.Location = new Point(12, 444);
        statusPanel.Name = "statusPanel";
        statusPanel.Size = new Size(1004, 70);
        statusPanel.TabIndex = 2;
        // 
        // lblStatusTitle
        // 
        lblStatusTitle.Location = new Point(0, 0);
        lblStatusTitle.Name = "lblStatusTitle";
        lblStatusTitle.Size = new Size(100, 23);
        lblStatusTitle.TabIndex = 0;
        // 
        // logPanel
        // 
        logPanel.BackColor = Color.FromArgb(255, 255, 255);
        logPanel.BorderStyle = BorderStyle.FixedSingle;
        logPanel.Controls.Add(progressBar);
        logPanel.Controls.Add(txtLog);
        logPanel.Location = new Point(13, 585);
        logPanel.Name = "logPanel";
        logPanel.Size = new Size(1003, 182);
        logPanel.TabIndex = 3;
        // 
        // Form1
        // 
        AllowDrop = true;
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        BackColor = Color.FromArgb(243, 243, 243);
        ClientSize = new Size(1057, 808);
        Controls.Add(topPanel);
        Controls.Add(lblStatus);
        Controls.Add(stylePanel);
        Controls.Add(statusPanel);
        Controls.Add(logPanel);
        Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
        MaximizeBox = false;
        MinimumSize = new Size(900, 800);
        Name = "Form1";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Word 文档排版工具";
        Load += Form1_Load;
        grpHeading1.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)nudHeading1SpaceBefore).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading1SpaceAfter).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading1LineSpacing).EndInit();
        grpHeading2.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)nudHeading2SpaceBefore).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading2SpaceAfter).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading2LineSpacing).EndInit();
        grpHeading3.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)nudHeading3SpaceBefore).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading3SpaceAfter).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading3LineSpacing).EndInit();
        grpHeading4.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)nudHeading4SpaceBefore).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading4SpaceAfter).EndInit();
        ((System.ComponentModel.ISupportInitialize)nudHeading4LineSpacing).EndInit();
        grpNormal.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)nudNormalLineSpacing).EndInit();
        grpTableCaption.ResumeLayout(false);
        grpTableCaption.PerformLayout();
        topPanel.ResumeLayout(false);
        topPanel.PerformLayout();
        stylePanel.ResumeLayout(false);
        grpImageCaption.ResumeLayout(false);
        statusPanel.ResumeLayout(false);
        logPanel.ResumeLayout(false);
        logPanel.PerformLayout();
        ResumeLayout(false);
    }

    #endregion

    private Label lblHeading1Font;
    private Label lblHeading1Size;
    private Label lblHeading1Before;
    private Label lblHeading1After;
    private Label lblHeading2Font;
    private Label lblHeading2Size;
    private Label lblHeading2Before;
    private Label lblHeading2After;
    private Label lblHeading3Font;
    private Label lblHeading3Size;
    private Label lblHeading3Before;
    private Label lblHeading3After;
    private Label lblHeading4Font;
    private Label lblHeading4Size;
    private Label lblHeading4Before;
    private Label lblHeading4After;
    private Label lblNormalFont;
    private Label lblNormalSize;
    private Panel topPanel;
    private Panel stylePanel;
    private Panel statusPanel;
    private Label lblStatusTitle;
    private Panel logPanel;
}
