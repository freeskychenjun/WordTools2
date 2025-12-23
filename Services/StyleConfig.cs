using WordTools2.Models;

namespace WordTools2.Services
{
    /// <summary>
    /// 样式配置管理类
    /// </summary>
    public class StyleConfig
    {
        /// <summary>
        /// 一级标题样式
        /// </summary>
        public ParagraphStyle Heading1 { get; set; }

        /// <summary>
        /// 二级标题样式
        /// </summary>
        public ParagraphStyle Heading2 { get; set; }

        /// <summary>
        /// 三级标题样式
        /// </summary>
        public ParagraphStyle Heading3 { get; set; }

        /// <summary>
        /// 四级标题样式
        /// </summary>
        public ParagraphStyle Heading4 { get; set; }

        /// <summary>
        /// 正文样式
        /// </summary>
        public ParagraphStyle Normal { get; set; }

        public StyleConfig()
        {
            // 初始化默认样式
            Heading1 = new ParagraphStyle("黑体", 16, 24, 24);       // 黑体，三号，段前24磅，段后24磅
            Heading2 = new ParagraphStyle("楷体", 13, 12, 12);       // 楷体，小四，段前12磅，段后12磅
            Heading3 = new ParagraphStyle("宋体", 13, 12, 12);       // 宋体，小四，段前12磅，段后12磅
            Heading4 = new ParagraphStyle("宋体", 13, 0, 0);         // 宋体，小四，段前0磅，段后0磅
            Normal = new ParagraphStyle("宋体", 13, 0, 0, 25);       // 宋体，小四，行距25磅
        }

        /// <summary>
        /// 根据样式名称获取对应的样式配置
        /// </summary>
        public ParagraphStyle GetStyle(string styleName)
        {
            return styleName switch
            {
                "Heading1" => Heading1,
                "Heading2" => Heading2,
                "Heading3" => Heading3,
                "Heading4" => Heading4,
                "Normal" => Normal,
                _ => Normal
            };
        }
    }
}
