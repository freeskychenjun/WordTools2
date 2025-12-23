using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordTools2.Models;
using System.Linq;

namespace WordTools2.Services
{
    /// <summary>
    /// 文档服务类 - 负责文档的读取、处理和保存
    /// </summary>
    public class DocumentService
    {
        private string? _originalFilePath; // 原始文件路径（永远不被修改）
        private string? _workingFilePath;   // 工作文件路径（可以修改）
        private WordprocessingDocument? _document;

        public DocumentService() { }

        /// <summary>
        /// 打开 Word 文档
        /// </summary>
        public bool OpenDocument(string filePath)
        {
            try
            {
                CloseDocument();
                _originalFilePath = filePath;
                _workingFilePath = null; // 初始时没有工作文件
                _document = WordprocessingDocument.Open(filePath, false);
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception($"无法打开文档: {ex.Message}");
            }
        }

        /// <summary>
        /// 关闭当前文档
        /// </summary>
        public void CloseDocument()
        {
            _document?.Dispose();
            _document = null;
            
            // 清理工作文件（如果存在）
            if (_workingFilePath != null && File.Exists(_workingFilePath))
            {
                try
                {
                    File.Delete(_workingFilePath);
                }
                catch
                {
                    // 忽略删除临时文件的错误
                }
                _workingFilePath = null;
            }
        }

        /// <summary>
        /// 应用样式配置到文档
        /// </summary>
        public void ApplyStyles(StyleConfig config, Action<string> updateProgress, Action<string> logMessage)
        {
            if (_originalFilePath == null)
                throw new Exception("请先打开一个文档");

            try
            {
                // 清理之前的工作文件
                if (_workingFilePath != null && File.Exists(_workingFilePath))
                {
                    File.Delete(_workingFilePath);
                }

                // 创建新的工作文件（基于原始文件）
                _workingFilePath = Path.GetTempFileName();
                File.Copy(_originalFilePath, _workingFilePath, true);
                
                try
                {
                    // 关闭原始文档的只读引用
                    _document?.Dispose();
                    _document = null;

                    using (var doc = WordprocessingDocument.Open(_workingFilePath, true))
                    {
                        var mainPart = doc.MainDocumentPart;
                        if (mainPart == null)
                            throw new Exception("无法获取文档主体部分");

                        var stylesPart = mainPart.StyleDefinitionsPart;
                        if (stylesPart == null)
                        {
                            stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                            stylesPart.Styles = new Styles();
                        }

                        var styles = stylesPart.Styles;
                        if (styles == null)
                            throw new Exception("无法获取样式定义");

                        UpdateOrCreateStyle(styles, "Heading1", config.Heading1);
                        UpdateOrCreateStyle(styles, "Heading2", config.Heading2);
                        UpdateOrCreateStyle(styles, "Heading3", config.Heading3);
                        UpdateOrCreateStyle(styles, "Heading4", config.Heading4);
                        UpdateOrCreateStyle(styles, "Normal", config.Normal);

                        stylesPart.Styles?.Save();
                        logMessage("样式定义已更新");

                        var paragraphs = mainPart.Document.Descendants<Paragraph>().ToList();
                        int total = paragraphs.Count;
                        int processed = 0;

                        foreach (var paragraph in paragraphs)
                        {
                            // 首先尝试根据大纲级别推断样式
                            var inferredStyle = InferStyleFromParagraph(paragraph, config);
                            
                            // 如果大纲级别在0-4范围内，应用对应的格式（不修改样式）
                            if (!string.IsNullOrEmpty(inferredStyle))
                            {
                                var style = GetParagraphStyle(inferredStyle, config);
                                if (style != null)
                                {
                                    ApplyStyleToParagraph(paragraph, style);
                                    logMessage($"处理段落格式: {inferredStyle} (大纲级别 {GetOutlineLevelFromStyleName(inferredStyle)})");
                                }
                            }
                            // 如果返回null，说明大纲级别是5-9级，保持不变
                            else
                            {
                                logMessage($"跳过段落: 大纲级别5-9级，维持原样");
                            }

                            processed++;
                            if (processed % 10 == 0)
                            {
                                var percent = (int)((double)processed / total * 100);
                                updateProgress($"处理中... {percent}%");
                            }
                        }

                        mainPart.Document.Save();
                        logMessage($"已处理 {total} 个段落");
                    }

                    // 重新打开处理后的工作文件（只读模式）
                    _document = WordprocessingDocument.Open(_workingFilePath, false);

                    updateProgress("处理完成");
                    logMessage("样式应用成功（原始文档未被修改）");
                }
                catch
                {
                    // 如果出错，清理工作文件
                    if (_workingFilePath != null && File.Exists(_workingFilePath))
                    {
                        File.Delete(_workingFilePath);
                        _workingFilePath = null;
                    }
                    throw;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"应用样式失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 保存文档到新路径
        /// </summary>
        public void SaveDocumentAs(string newPath)
        {
            if (_originalFilePath == null)
                throw new Exception("没有可保存的文档");

            try
            {
                string sourcePath;
                
                // 如果有工作文件，保存工作文件；否则保存原始文件
                if (_workingFilePath != null && File.Exists(_workingFilePath))
                {
                    sourcePath = _workingFilePath;
                }
                else
                {
                    sourcePath = _originalFilePath;
                }

                // 关闭文档以释放文件锁
                _document?.Dispose();
                _document = null;
                
                // 复制源文件到新位置
                File.Copy(sourcePath, newPath, true);
                
                // 重新打开源文件（只读模式）
                if (_workingFilePath != null && File.Exists(_workingFilePath))
                {
                    _document = WordprocessingDocument.Open(_workingFilePath, false);
                }
                else
                {
                    _document = WordprocessingDocument.Open(_originalFilePath, false);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"保存文档失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取文档统计信息（基于大纲级别，显示0-9级统计）
        /// </summary>
        public Dictionary<string, int> GetDocumentStats()
        {
            var stats = new Dictionary<string, int>
            {
                { "Normal", 0 },      // 正文文本（大纲级别0）
                { "Heading1", 0 },    // 1级标题
                { "Heading2", 0 },    // 2级标题
                { "Heading3", 0 },    // 3级标题
                { "Heading4", 0 },    // 4级标题
                { "Level5", 0 },      // 5级（保持不变）
                { "Level6", 0 },      // 6级（保持不变）
                { "Level7", 0 },      // 7级（保持不变）
                { "Level8", 0 },      // 8级（保持不变）
                { "Level9", 0 },      // 9级（保持不变）
                { "NoLevel", 0 }      // 无大纲级别
            };

            if (_document == null || _document.MainDocumentPart == null)
                return stats;

            // 获取样式定义，用于解析继承的大纲级别
            var styleDefinitions = new Dictionary<string, Style>();
            var stylesPart = _document.MainDocumentPart.StyleDefinitionsPart;
            if (stylesPart != null && stylesPart.Styles != null)
            {
                foreach (var style in stylesPart.Styles.Elements<Style>())
                {
                    if (style.StyleId != null && style.StyleId.Value != null)
                    {
                        styleDefinitions[style.StyleId.Value] = style;
                    }
                }
            }

            foreach (var paragraph in _document.MainDocumentPart.Document.Descendants<Paragraph>())
            {
                // 1. 先检查直接的大纲级别
                int? outlineLevel = null;
                var paragraphProperties = paragraph.ParagraphProperties;
                
                if (paragraphProperties != null)
                {
                    var outlineLevelElement = paragraphProperties.GetFirstChild<OutlineLevel>();
                    if (outlineLevelElement != null && outlineLevelElement.Val != null)
                    {
                        if (int.TryParse(outlineLevelElement.Val.InnerText, out int level))
                        {
                            outlineLevel = level;
                        }
                    }
                    
                    // 2. 如果没有直接的大纲级别，检查样式
                    if (outlineLevel == null)
                    {
                        var styleId = paragraphProperties.ParagraphStyleId?.Val?.Value;
                        if (styleId != null && styleDefinitions.TryGetValue(styleId, out Style style))
                        {
                            // 检查样式名称是否包含"heading"（不区分大小写）
                            var styleName = style.GetFirstChild<StyleName>()?.Val?.Value ?? "";
                            if (styleName.IndexOf("heading", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // 根据样式名称推断大纲级别
                                if (styleName.IndexOf("1", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    outlineLevel = 1;
                                }
                                else if (styleName.IndexOf("2", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    outlineLevel = 2;
                                }
                                else if (styleName.IndexOf("3", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    outlineLevel = 3;
                                }
                                else if (styleName.IndexOf("4", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    outlineLevel = 4;
                                }
                                else
                                {
                                    // 默认视为1级标题
                                    outlineLevel = 1;
                                }
                            }
                            else
                            {
                                // 检查样式是否有OutlineLevel设置
                                var styleParagraphProps = style.GetFirstChild<ParagraphProperties>();
                                if (styleParagraphProps != null)
                                {
                                    var styleOutlineLevel = styleParagraphProps.GetFirstChild<OutlineLevel>();
                                    if (styleOutlineLevel != null && styleOutlineLevel.Val != null)
                                    {
                                        if (int.TryParse(styleOutlineLevel.Val.InnerText, out int level))
                                        {
                                            outlineLevel = level;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // 3. 根据大纲级别统计
                if (outlineLevel.HasValue)
                {
                    string category = outlineLevel.Value switch
                    {
                        0 => "Normal",    // 正文文本
                        1 => "Heading1",  // 1级标题
                        2 => "Heading2",  // 2级标题
                        3 => "Heading3",  // 3级标题
                        4 => "Heading4",  // 4级标题
                        5 => "Level5",    // 5级（保持不变）
                        6 => "Level6",    // 6级（保持不变）
                        7 => "Level7",    // 7级（保持不变）
                        8 => "Level8",    // 8级（保持不变）
                        9 => "Level9",    // 9级（保持不变）
                        _ => "Normal"     // 其他级别视为正文
                    };
                    
                    stats[category]++;
                }
                else
                {
                    // 没有大纲级别的
                    stats["NoLevel"]++;
                }
            }

            return stats;
        }

        /// <summary>
        /// 获取原始文件路径
        /// </summary>
        public string? GetOriginalFilePath()
        {
            return _originalFilePath;
        }

        /// <summary>
        /// 检查是否有已处理的工作文件
        /// </summary>
        public bool HasWorkingFile()
        {
            return _workingFilePath != null && File.Exists(_workingFilePath);
        }



        /// <summary>
        /// 更新或创建样式定义
        /// </summary>
        private void UpdateOrCreateStyle(Styles styles, string styleId, ParagraphStyle styleConfig)
        {
            var style = styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);

            if (style == null)
            {
                style = new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId
                };

                var styleName = new StyleName() { Val = styleId };
                style.Append(styleName);

                var basedOn = new BasedOn() { Val = "Normal" };
                style.Append(basedOn);

                var nextParagraphStyle = new NextParagraphStyle() { Val = "Normal" };
                style.Append(nextParagraphStyle);

                styles.Append(style);
            }

            var runProperties = style.GetFirstChild<RunProperties>();
            if (runProperties == null)
            {
                runProperties = new RunProperties();
                style.Append(runProperties);
            }
            
            var runFonts = runProperties.GetFirstChild<RunFonts>();
            if (runFonts == null)
            {
                runFonts = new RunFonts();
                runProperties.Append(runFonts);
            }
            runFonts.Ascii = styleConfig.FontName;
            runFonts.HighAnsi = styleConfig.FontName;
            runFonts.EastAsia = styleConfig.FontName;

            var fontSize = runProperties.GetFirstChild<FontSize>();
            if (fontSize == null)
            {
                fontSize = new FontSize();
                runProperties.Append(fontSize);
            }
            fontSize.Val = ConvertFontSizeToHalfPoints(styleConfig.FontSize);

            var paragraphProperties = style.GetFirstChild<ParagraphProperties>();
            if (paragraphProperties == null)
            {
                paragraphProperties = new ParagraphProperties();
                style.Append(paragraphProperties);
            }

            var spacing = paragraphProperties.GetFirstChild<SpacingBetweenLines>();
            if (spacing == null)
            {
                spacing = new SpacingBetweenLines();
                paragraphProperties.Append(spacing);
            }
            spacing.Before = ConvertPointsToTwips(styleConfig.SpaceBefore);
            spacing.After = ConvertPointsToTwips(styleConfig.SpaceAfter);
        }

        /// <summary>
        /// 应用格式到单个段落（不修改样式，直接应用格式属性）
        /// </summary>
        private void ApplyStyleToParagraph(Paragraph paragraph, ParagraphStyle styleConfig)
        {
            var properties = paragraph.ParagraphProperties ?? new ParagraphProperties();

            // 不修改样式ID，保持原有样式不变
            // var styleId = properties.ParagraphStyleId ?? new ParagraphStyleId();
            // styleId.Val = GetStyleIdForParagraphStyle(styleConfig);
            // properties.ParagraphStyleId = styleId;

            // 不设置大纲级别，保持原有大纲级别
            // var outlineLevel = properties.GetFirstChild<OutlineLevel>();
            // if (outlineLevel == null)
            // {
            //     outlineLevel = new OutlineLevel();
            //     properties.Append(outlineLevel);
            // }
            // 
            // // 根据样式设置对应的大纲级别
            // var targetLevel = GetOutlineLevelForStyle(styleId.Val?.Value ?? "Normal");
            // outlineLevel.Val = targetLevel;

            var spacing = properties.GetFirstChild<SpacingBetweenLines>();
            if (spacing == null)
            {
                spacing = new SpacingBetweenLines();
                properties.Append(spacing);
            }
            spacing.Before = ConvertPointsToTwips(styleConfig.SpaceBefore);
            spacing.After = ConvertPointsToTwips(styleConfig.SpaceAfter);
            
            // 设置行距
            if (styleConfig.LineSpacing > 0)
            {
                spacing.LineRule = LineSpacingRuleValues.Exact;
                spacing.Line = ConvertPointsToTwips(styleConfig.LineSpacing);
            }

            paragraph.ParagraphProperties = properties;

            foreach (var run in paragraph.Elements<Run>())
            {
                var runProperties = run.RunProperties ?? new RunProperties();

                var runFonts = runProperties.GetFirstChild<RunFonts>();
                if (runFonts == null)
                {
                    runFonts = new RunFonts();
                    runProperties.Append(runFonts);
                }
                runFonts.Ascii = styleConfig.FontName;
                runFonts.HighAnsi = styleConfig.FontName;
                runFonts.EastAsia = styleConfig.FontName;

                var fontSize = runProperties.GetFirstChild<FontSize>();
                if (fontSize == null)
                {
                    fontSize = new FontSize();
                    runProperties.Append(fontSize);
                }
                fontSize.Val = ConvertFontSizeToHalfPoints(styleConfig.FontSize);

                run.RunProperties = runProperties;
            }
        }

        /// <summary>
        /// 根据段落的大纲级别、样式名称、字体大小或文本格式推断样式类型（仅对0-4级进行处理，其他级别保持不变）
        /// </summary>
        private string? InferStyleFromParagraph(Paragraph paragraph, StyleConfig config)
        {
            var paragraphProperties = paragraph.ParagraphProperties;
            
            // 获取段落文本
            string text = string.Join("", paragraph.Descendants<Run>().Select(r => r.InnerText)).Trim();
            
            if (paragraphProperties != null)
            {
                // 1. 先检查直接的大纲级别
                int? outlineLevel = null;
                var outlineLevelElement = paragraphProperties.GetFirstChild<OutlineLevel>();
                if (outlineLevelElement != null)
                {
                    if (int.TryParse(outlineLevelElement.Val?.InnerText, out int level))
                    {
                        outlineLevel = level;
                    }
                }
                
                // 2. 如果没有直接的大纲级别，检查样式定义
                if (outlineLevel == null)
                {
                    var styleId = paragraphProperties.ParagraphStyleId?.Val?.Value;
                    if (!string.IsNullOrEmpty(styleId))
                    {
                        // 检查样式定义
                        var stylesPart = _document?.MainDocumentPart?.StyleDefinitionsPart;
                        if (stylesPart != null && stylesPart.Styles != null)
                        {
                            var style = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
                            if (style != null)
                            {
                                // 检查样式名称是否包含"heading"（不区分大小写）
                                var styleName = style.GetFirstChild<StyleName>()?.Val?.Value ?? "";
                                if (styleName.IndexOf("heading", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    // 根据样式名称推断标题级别
                                    if (styleName.IndexOf("1", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        return "Heading1";
                                    }
                                    else if (styleName.IndexOf("2", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        return "Heading2";
                                    }
                                    else if (styleName.IndexOf("3", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        return "Heading3";
                                    }
                                    else if (styleName.IndexOf("4", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        return "Heading4";
                                    }
                                    else
                                    {
                                        // 默认视为一级标题
                                        return "Heading1";
                                    }
                                }
                            }
                        }
                    }
                }
                
                // 3. 根据大纲级别返回结果
                if (outlineLevel != null)
                {
                    return outlineLevel switch
                    {
                        0 => "Normal",    // 正文文本
                        1 => "Heading1",  // 1级标题
                        2 => "Heading2",  // 2级标题
                        3 => "Heading3",  // 3级标题
                        4 => "Heading4",  // 4级标题
                        _ => null         // 5-9级保持不变，返回null
                    };
                }
                
                // 4. 如果没有大纲级别和标准样式名称，尝试根据字体大小推断
                // 获取段落的字体大小
                double fontSize = GetParagraphFontSize(paragraph);
                
                // 根据字体大小推断标题级别（字体越大，标题级别越高）
                if (fontSize >= 16) return "Heading1";  // 一级标题：16pt及以上
                if (fontSize >= 14) return "Heading2";  // 二级标题：14pt-15.5pt
                if (fontSize >= 13) return "Heading3";  // 三级标题：13pt-13.5pt
                if (fontSize >= 12) return "Heading4";  // 四级标题：12pt
            }
            
            // 5. 如果没有大纲级别、标准样式名称和大字体，尝试根据文本格式推断
            // 检查文本是否符合标题格式，如"4 工程任务和规模"、"4.1 可行性研究报告"等
            if (!string.IsNullOrEmpty(text))
            {
                // 匹配一级标题格式：数字+空格+文本（如"4 工程任务和规模"）
                if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+\s+[\u4e00-\u9fa5\w]"))
                {
                    return "Heading1";
                }
                // 匹配二级标题格式：数字.数字+空格+文本（如"4.1 可行性研究报告"）
                else if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+\.\d+\s+[\u4e00-\u9fa5\w]"))
                {
                    return "Heading2";
                }
                // 匹配三级标题格式：数字.数字.数字+空格+文本（如"4.1.1 可行性研究报告"）
                else if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+\.\d+\.\d+\s+[\u4e00-\u9fa5\w]"))
                {
                    return "Heading3";
                }
                // 匹配四级标题格式：数字.数字.数字.数字+空格+文本（如"4.1.1.1 可行性研究报告"）
                else if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\d+\.\d+\.\d+\.\d+\s+[\u4e00-\u9fa5\w]"))
                {
                    return "Heading4";
                }
            }

            // 默认视为正文文本
            return "Normal";
        }
        
        /// <summary>
        /// 获取段落的字体名称
        /// </summary>
        private string GetParagraphFontName(Paragraph paragraph)
        {
            // 默认字体名称
            string defaultFontName = "宋体";
            
            // 1. 首先从Run元素中获取字体名称
            string mostCommonFont = defaultFontName;
            int fontCount = 0;
            
            foreach (var run in paragraph.Elements<Run>())
            {
                var runProperties = run.RunProperties;
                if (runProperties == null)
                {
                    continue;
                }
                
                var runFonts = runProperties.GetFirstChild<RunFonts>();
                if (runFonts == null)
                {
                    continue;
                }
                
                // 优先使用EastAsia字体（中文），然后是HighAnsi，最后是Ascii
                string fontName = runFonts.EastAsia?.Value ?? 
                                 runFonts.HighAnsi?.Value ?? 
                                 runFonts.Ascii?.Value ?? 
                                 defaultFontName;
                
                // 记录最常见的字体名称
                mostCommonFont = fontName;
                fontCount++;
                
                // 如果找到至少一个字体名称，就返回
                if (fontCount > 0)
                {
                    return mostCommonFont;
                }
            }
            
            // 2. 如果没有从Run中获取到字体名称，尝试从样式中获取
            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties != null)
            {
                var styleId = paragraphProperties.ParagraphStyleId?.Val?.Value;
                if (!string.IsNullOrEmpty(styleId))
                {
                    // 查找样式定义
                    var stylesPart = _document?.MainDocumentPart?.StyleDefinitionsPart;
                    if (stylesPart != null && stylesPart.Styles != null)
                    {
                        var style = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
                        if (style != null)
                        {
                            var runProperties = style.GetFirstChild<RunProperties>();
                            if (runProperties != null)
                            {
                                var runFonts = runProperties.GetFirstChild<RunFonts>();
                                if (runFonts != null)
                                {
                                    return runFonts.EastAsia?.Value ?? 
                                           runFonts.HighAnsi?.Value ?? 
                                           runFonts.Ascii?.Value ?? 
                                           defaultFontName;
                                }
                            }
                        }
                    }
                }
            }
            
            // 3. 默认返回宋体
            return defaultFontName;
        }
        
        /// <summary>
        /// 获取段落的字体大小（磅）
        /// 优先级：直接设置的字体大小 > 样式推断的字体大小 > 默认字体大小
        /// </summary>
        private double GetParagraphFontSize(Paragraph paragraph)
        {
            // 默认字体大小
            double defaultFontSize = 10.5;
            
            // 1. 首先从Run元素中获取直接设置的字体大小
            double maxDirectFontSize = 0;
            
            foreach (var run in paragraph.Elements<Run>())
            {
                var runProperties = run.RunProperties;
                if (runProperties != null)
                {
                    var fontSizeElement = runProperties.GetFirstChild<FontSize>();
                    if (fontSizeElement != null && fontSizeElement.Val != null)
                    {
                        // FontSize的Val值是半点单位（1pt = 2半点）
                        if (int.TryParse(fontSizeElement.Val.InnerText, out int halfPoints))
                        {
                            double fontSize = halfPoints / 2.0;
                            if (fontSize > maxDirectFontSize)
                            {
                                maxDirectFontSize = fontSize;
                            }
                        }
                    }
                }
            }
            
            // 2. 如果有直接设置的字体大小，直接返回
            if (maxDirectFontSize > 0)
            {
                return maxDirectFontSize;
            }
            
            // 3. 否则，从样式推断字体大小
            double styleBasedFontSize = 0;
            
            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties != null)
            {
                var styleId = paragraphProperties.ParagraphStyleId?.Val?.Value;
                if (!string.IsNullOrEmpty(styleId))
                {
                    // 查找样式定义
                    var stylesPart = _document?.MainDocumentPart?.StyleDefinitionsPart;
                    if (stylesPart != null && stylesPart.Styles != null)
                    {
                        // 构建样式字典
                        var styleDict = new Dictionary<string, Style>();
                        foreach (var style in stylesPart.Styles.Elements<Style>())
                        {
                            if (style.StyleId != null && style.StyleId.Value != null)
                            {
                                styleDict[style.StyleId.Value] = style;
                            }
                        }
                        
                        // 递归查找样式名称
                        string currentStyleId = styleId;
                        string styleName = "";
                        
                        while (!string.IsNullOrEmpty(currentStyleId) && styleDict.TryGetValue(currentStyleId, out Style style))
                        {
                            var styleNameElement = style.GetFirstChild<StyleName>();
                            if (styleNameElement != null && styleNameElement.Val != null)
                            {
                                styleName = styleNameElement.Val.Value;
                                break; // 找到后停止递归
                            }
                            
                            // 检查是否有基于样式，继续递归
                            var basedOn = style.GetFirstChild<BasedOn>();
                            if (basedOn != null && basedOn.Val != null)
                            {
                                currentStyleId = basedOn.Val.Value;
                            }
                            else
                            {
                                break;
                            }
                        }
                        
                        // 根据样式名称推断字体大小
                        if (!string.IsNullOrEmpty(styleName))
                        {
                            // 不区分大小写比较，使用更灵活的匹配
                            if (styleName.IndexOf("heading", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // 根据样式名称中的数字推断字体大小
                                if (styleName.IndexOf("1", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    styleBasedFontSize = 18.0; // 小二
                                }
                                else if (styleName.IndexOf("2", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    styleBasedFontSize = 16.0; // 三号
                                }
                                else if (styleName.IndexOf("3", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    styleBasedFontSize = 15.0; // 小三
                                }
                                else if (styleName.IndexOf("4", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    styleBasedFontSize = 14.0; // 四号
                                }
                                else
                                {
                                    styleBasedFontSize = 16.0; // 默认标题大小
                                }
                            }
                        }
                    }
                }
            }
            
            // 4. 返回样式推断的字体大小或默认字体大小
            return styleBasedFontSize > 0 ? styleBasedFontSize : defaultFontSize;
        }

        /// <summary>
        /// 根据样式名称获取对应的段落样式配置
        /// </summary>
        private ParagraphStyle? GetParagraphStyle(string styleName, StyleConfig config)
        {
            return styleName switch
            {
                "Heading1" => config.Heading1,
                "Heading2" => config.Heading2,
                "Heading3" => config.Heading3,
                "Heading4" => config.Heading4,
                "Normal" => config.Normal,
                _ => config.Normal
            };
        }

        /// <summary>
        /// 根据段落样式配置获取样式 ID
        /// </summary>
        private string GetStyleIdForParagraphStyle(ParagraphStyle style)
        {
            if (style.FontSize >= 16) return "Heading1";
            if (style.FontSize >= 14) return "Heading2";
            if (style.FontSize >= 13) return "Heading3";
            if (style.FontSize >= 12) return "Heading4";

            return "Normal";
        }

        /// <summary>
        /// 根据样式ID获取对应的大纲级别
        /// </summary>
        private int GetOutlineLevelForStyle(string styleName)
        {
            return styleName switch
            {
                "Normal" => 0,    // 正文文本
                "Heading1" => 1,  // 一级标题
                "Heading2" => 2,  // 二级标题
                "Heading3" => 3,  // 三级标题
                "Heading4" => 4,  // 四级标题
                _ => 9            // 其他级别
            };
        }

        /// <summary>
        /// 根据样式名称获取对应的大纲级别（用于日志显示）
        /// </summary>
        private int GetOutlineLevelFromStyleName(string styleName)
        {
            return styleName switch
            {
                "Normal" => 0,    // 正文文本
                "Heading1" => 1,  // 一级标题
                "Heading2" => 2,  // 二级标题
                "Heading3" => 3,  // 三级标题
                "Heading4" => 4,  // 四级标题
                _ => -1           // 未知级别
            };
        }

        /// <summary>
        /// 将字号（磅）转换为半点单位（1磅=2半点）
        /// </summary>
        private string ConvertFontSizeToHalfPoints(double fontSize)
        {
            return ((int)(fontSize * 2)).ToString();
        }

        /// <summary>
        /// 将半点单位转换为字号（磅）
        /// </summary>
        private double ConvertHalfPointsToFontSize(string? halfPoints)
        {
            if (int.TryParse(halfPoints ?? "", out int hp))
            {
                return hp / 2.0;
            }
            return 10.5;
        }

        /// <summary>
        /// 将磅转换为缇（1磅=20缇）
        /// </summary>
        private string ConvertPointsToTwips(double points)
        {
            return ((int)(points * 20)).ToString();
        }
    }
}
