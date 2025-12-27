namespace WordTools2.Models
{
    /// <summary>
    /// 样式配置类 - 包含所有段落样式的配置
    /// </summary>
    public class StyleConfig
    {
        /// <summary>
        /// 一级标题样式
        /// </summary>
        public ParagraphStyle Heading1 { get; set; } = new ParagraphStyle("黑体", 16, 24, 24, 24, 0, false);

        /// <summary>
        /// 二级标题样式
        /// </summary>
        public ParagraphStyle Heading2 { get; set; } = new ParagraphStyle("楷体", 14, 12, 12, 24, 1, false);

        /// <summary>
        /// 三级标题样式
        /// </summary>
        public ParagraphStyle Heading3 { get; set; } = new ParagraphStyle("宋体", 13, 12, 12, 24, 2, false);

        /// <summary>
        /// 四级标题样式
        /// </summary>
        public ParagraphStyle Heading4 { get; set; } = new ParagraphStyle("宋体", 12, 0, 0, 24, 3, false);

        /// <summary>
        /// 正文样式
        /// </summary>
        public ParagraphStyle Normal { get; set; } = new ParagraphStyle("宋体", 10.5, 0, 0, 24);

        /// <summary>
        /// 表格标题样式
        /// </summary>
        public ParagraphStyle TableCaption { get; set; } = new ParagraphStyle("黑体", 10.5, 0, 0, 24, 8, false);

        /// <summary>
        /// 图形标题样式
        /// </summary>
        public ParagraphStyle ImageCaption { get; set; } = new ParagraphStyle("黑体", 10.5, 0, 0, 24, 6, false);

        public StyleConfig() { }
    }
}