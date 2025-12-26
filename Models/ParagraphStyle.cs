namespace WordTools2.Models
{
    /// <summary>
    /// 段落样式配置类
    /// </summary>
    public class ParagraphStyle
    {
        /// <summary>
        /// 字体名称
        /// </summary>
        public string FontName { get; set; } = "Microsoft YaHei UI";

        /// <summary>
        /// 字号大小（半点单位）
        /// </summary>
        public double FontSize { get; set; } = 10.5;

        /// <summary>
        /// 段前间距（磅）
        /// </summary>
        public double SpaceBefore { get; set; } = 6;

        /// <summary>
        /// 段后间距（磅）
        /// </summary>
        public double SpaceAfter { get; set; } = 6;
        
        /// <summary>
        /// 行距（磅）
        /// </summary>
        public double LineSpacing { get; set; } = 0;

        /// <summary>
        /// 大纲级别（0-9，其中9表示正文级别）
        /// </summary>
        public int OutlineLevel { get; set; } = 9;

        /// <summary>
        /// 是否加粗
        /// </summary>
        public bool Bold { get; set; } = false;

        public ParagraphStyle() { }

        public ParagraphStyle(string fontName, double fontSize, double spaceBefore = 0, double spaceAfter = 0, double lineSpacing = 0, int outlineLevel = 9, bool bold = false)
        {
            FontName = fontName;
            FontSize = fontSize;
            SpaceBefore = spaceBefore;
            SpaceAfter = spaceAfter;
            LineSpacing = lineSpacing;
            OutlineLevel = outlineLevel;
            Bold = bold;
        }
    }
}
