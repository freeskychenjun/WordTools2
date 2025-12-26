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
        public ParagraphStyle Heading1 { get; set; } = new ParagraphStyle("黑体", 16, 16, 8, 20);

        /// <summary>
        /// 二级标题样式
        /// </summary>
        public ParagraphStyle Heading2 { get; set; } = new ParagraphStyle("楷体", 14, 12, 6, 18);

        /// <summary>
        /// 三级标题样式
        /// </summary>
        public ParagraphStyle Heading3 { get; set; } = new ParagraphStyle("宋体", 13, 10, 4, 16);

        /// <summary>
        /// 四级标题样式
        /// </summary>
        public ParagraphStyle Heading4 { get; set; } = new ParagraphStyle("宋体", 12, 8, 3, 15);

        /// <summary>
        /// 正文样式
        /// </summary>
        public ParagraphStyle Normal { get; set; } = new ParagraphStyle("宋体", 10.5, 0, 0, 15);

        /// <summary>
        /// 表格标题样式
        /// </summary>
        public ParagraphStyle TableCaption { get; set; } = new ParagraphStyle("黑体", 10.5, 0, 0, 15, 8, false);

        /// <summary>
        /// 图形标题样式
        /// </summary>
        public ParagraphStyle ImageCaption { get; set; } = new ParagraphStyle("黑体", 10.5, 0, 0, 15, 6, false);

        public StyleConfig() { }
    }
}