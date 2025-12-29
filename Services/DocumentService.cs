using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordTools2.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace WordTools2.Services
{
    /// <summary>
    /// 文档服务类 - 使用 Open XML SDK 处理 Word 文档
    /// </summary>
    public class DocumentService
    {
        private string? _originalFilePath; // 原始文件路径（永远不被修改）
        private string? _workingFilePath;   // 工作文件路径（可以修改）
        private WordprocessingDocument? _document;
        private Document? _docBody;         // 文档主体引用

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
                _docBody = _document.MainDocumentPart?.Document;

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
            if (_document != null)
            {
                try
                {
                    _document.Dispose();
                }
                catch { }
                _document = null;
                _docBody = null;
            }

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
        public void ApplyStyles(Models.StyleConfig config, Action<string> updateProgress, Action<string> logMessage)
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
                    if (_document != null)
                    {
                        _document.Dispose();
                        _document = null;
                        _docBody = null;
                    }

                    // 使用可读写模式打开文档
                    using (var wordDoc = WordprocessingDocument.Open(_workingFilePath, true))
                    {
                        var doc = wordDoc.MainDocumentPart?.Document;
                        if (doc == null)
                            throw new Exception("无法读取文档内容");

                        var body = doc.Body;
                        if (body == null)
                            throw new Exception("文档没有主体内容");

                        logMessage("开始直接修改段落格式（优先使用样式大纲级别判断，样式无法判断时使用正则表达式补充）");

                        // 诊断：检查原始文档的样式定义
                        CheckOriginalDocumentStyles(wordDoc, logMessage);

                        // 计算需要跳过的段落数量（基于分页符）
                        int skipParagraphs = 0;
                        if (config.SkipPages > 0)
                        {
                            skipParagraphs = CountParagraphsToSkip(wordDoc, body, config.SkipPages);
                            if (skipParagraphs > 0)
                            {
                                logMessage($"跳过前 {config.SkipPages} 页，共 {skipParagraphs} 个段落");
                            }
                            else
                            {
                                logMessage($"警告：未能检测到 {config.SkipPages} 个分页符/分节符，可能存在自然分页。已设置为不跳过任何段落，请手动调整跳过页数或检查文档中的分页符/分节符。");
                            }
                        }

                        var paragraphs = body.Elements<Paragraph>().ToList();
                        int total = paragraphs.Count - skipParagraphs;
                        int processed = 0;

                        // 从指定位置开始处理段落
                        int paragraphIndex = 0;
                        foreach (var paragraph in paragraphs)
                        {
                            // 跳过前几页的段落
                            if (paragraphIndex < skipParagraphs)
                            {
                                paragraphIndex++;
                                continue;
                            }

                            // 通过 Open XML SDK 获取实际样式名称
                            var actualStyleInfo = GetParagraphStyleInfo(wordDoc, paragraph);

                            // 获取包含自动编号的完整文本
                            var paragraphTextWithNumbering = GetParagraphTextWithNumbering(wordDoc, paragraph);

                            // 对表格标题段落应用专用样式
                            if (actualStyleInfo?.StyleName == "TableCaption" ||
                                IsTableCaptionByPattern(paragraphTextWithNumbering))
                            {
                                var style = config.TableCaption;
                                ApplyStyleToParagraph(wordDoc, paragraph, style, "TableCaption", logMessage, config);
                                logMessage($"处理表格标题：{paragraph.InnerText.Trim()}");
                                processed++;
                                if (processed % 10 == 0)
                                {
                                    var percent = (int)((double)processed / total * 100);
                                    updateProgress($"处理中... {percent}%");
                                }
                                continue;
                            }

                            // 对图形标题段落应用专用样式
                            if (actualStyleInfo?.StyleName == "ImageCaption" ||
                                IsImageCaptionByPattern(paragraphTextWithNumbering))
                            {
                                var style = config.ImageCaption;
                                ApplyStyleToParagraph(wordDoc, paragraph, style, "ImageCaption", logMessage, config);
                                logMessage($"处理图形标题：{paragraph.InnerText.Trim()}");
                                processed++;
                                if (processed % 10 == 0)
                                {
                                    var percent = (int)((double)processed / total * 100);
                                    updateProgress($"处理中... {percent}%");
                                }
                                continue;
                            }

                            // 对图片段落应用专用样式
                            if (IsImageParagraph(paragraph))
                            {
                                ApplyImageParagraphStyle(paragraph, logMessage);
                                logMessage("处理图片段落：应用单倍行距和0间距");
                                processed++;
                                if (processed % 10 == 0)
                                {
                                    var percent = (int)((double)processed / total * 100);
                                    updateProgress($"处理中... {percent}%");
                                }
                                continue; // 继续处理下一个段落
                            }

                            // 优先使用样式的大纲级别判断，只有在样式判断为正文时才使用正则表达式补充判断
                            string styleName = InferStyleFromParagraph(wordDoc, paragraph, paragraph.InnerText);

                            if (!string.IsNullOrEmpty(styleName))
                            {
                                var style = GetParagraphStyle(styleName, config);
                                if (style != null)
                                {
                                    ApplyStyleToParagraph(wordDoc, paragraph, style, styleName, logMessage, config);
                                    logMessage($"处理段落格式: {styleName}");
                                }
                            }

                            processed++;
                            if (processed % 10 == 0)
                            {
                                var percent = (int)((double)processed / total * 100);
                                updateProgress($"处理中... {percent}%");
                            }

                            paragraphIndex++;
                        }

                        // 处理表格格式设置
                        ProcessTablesInDocument(wordDoc, config, logMessage);

                        // 确保所有更改都提交到文档
                        wordDoc.MainDocumentPart?.Document?.Save();

                        logMessage($"已处理 {total} 个段落");
                        logMessage("文档已保存，所有格式更改已提交");
                    } // end using (wordDoc)

                    // 重新打开处理后的工作文件（只读模式）
                    _document = WordprocessingDocument.Open(_workingFilePath, false);
                    _docBody = _document.MainDocumentPart?.Document;

                    updateProgress("处理完成 100%");
                    logMessage("样式应用成功（原始文档未被修改）");
                } // end inner try
                catch
                {
                    // 如果出错，清理工作文件
                    if (_workingFilePath != null && File.Exists(_workingFilePath))
                    {
                        File.Delete(_workingFilePath);
                        _workingFilePath = null;
                    }
                    throw;
                } // end catch of inner try
            } // end outer try
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
                if (_document != null)
                {
                    _document.Dispose();
                    _document = null;
                    _docBody = null;
                }

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
                _docBody = _document.MainDocumentPart?.Document;
            }
            catch (Exception ex)
            {
                throw new Exception($"保存文档失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取文档统计信息（基于文本格式识别）
        /// </summary>
        /// <summary>
        /// 获取文档统计信息（优先使用样式大纲级别，样式无法判断时使用正则表达式）
        /// </summary>
        public Dictionary<string, int> GetDocumentStats()
        {
            var stats = new Dictionary<string, int>
            {
                { "Normal", 0 },      // 正文文本
                { "Heading1", 0 },    // 1级标题
                { "Heading2", 0 },    // 2级标题
                { "Heading3", 0 },    // 3级标题
                { "Heading4", 0 },    // 4级标题
                { "TableCaption", 0 }, // 表格标题
                { "ImageCaption", 0 }, // 图形标题
                { "Image", 0 },       // 图片段落
                { "Other", 0 }        // 其他格式
            };

            if (_docBody == null)
                return stats;

            var body = _docBody.Body;
            if (body == null)
                return stats;

            foreach (var paragraph in body.Elements<Paragraph>())
            {
                // 首先检查是否为图片段落
                if (IsImageParagraph(paragraph))
                {
                    stats["Image"]++;
                    continue;
                }

                // 获取段落文本
                var text = paragraph.InnerText.Trim();
                
                // 优先使用样式的大纲级别判断
                string styleType = InferStyleFromParagraph(_document, paragraph, text);

                // 基于样式类型统计
                if (string.IsNullOrEmpty(text))
                {
                    stats["Normal"]++;
                }
                else
                {
                    switch (styleType)
                    {
                        case "TableCaption":
                            stats["TableCaption"]++;
                            break;
                        case "ImageCaption":
                            stats["ImageCaption"]++;
                            break;
                        case "Heading1":
                            stats["Heading1"]++;
                            break;
                        case "Heading2":
                            stats["Heading2"]++;
                            break;
                        case "Heading3":
                            stats["Heading3"]++;
                            break;
                        case "Heading4":
                            stats["Heading4"]++;
                            break;
                        case "Normal":
                            stats["Normal"]++;
                            break;
                        default:
                            stats["Other"]++;
                            break;
                    }
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
        /// 直接修改段落格式（不修改样式定义）
        /// 只修改当前段落的字体、字号、间距、大纲级别等属性
        /// </summary>
        private void ApplyStyleToParagraph(WordprocessingDocument wordDoc, Paragraph paragraph, ParagraphStyle styleConfig, string styleName, Action<string> logMessage, Models.StyleConfig? config = null)
        {
            // 首先获取段落文本用于调试
            var text = paragraph.InnerText.Trim();

            // 跳过图片段落，不进行任何格式设置
            if (styleName == "Image")
            {
                logMessage("跳过图片段落，不应用任何格式设置");
                return;
            }

            // 设置大纲级别和格式（表格标题需要特殊处理）
            SetOutlineLevelForParagraph(paragraph, styleName, logMessage, config);

            // 获取或创建段落属性
            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties == null)
            {
                paragraphProperties = new ParagraphProperties();
                paragraph.AppendChild(paragraphProperties);
            }

            // 设置段落间距
            var spacing = paragraphProperties.SpacingBetweenLines;
            if (spacing == null)
            {
                spacing = new SpacingBetweenLines();
                paragraphProperties.AppendChild(spacing);
            }

            spacing.Before = ConvertPointsToTwips(styleConfig.SpaceBefore).ToString();
            spacing.After = ConvertPointsToTwips(styleConfig.SpaceAfter).ToString();

            // 对标题和正文段落都设置行距
            if (styleConfig.LineSpacing > 0)
            {
                // 将磅值转换为缇单位（1磅 = 20缇）
                spacing.Line = ((int)(styleConfig.LineSpacing * 20)).ToString();
                spacing.LineRule = LineSpacingRuleValues.Exact;
            }

            // 设置段落中所有运行的字体属性
            var runs = paragraph.Elements<Run>().ToList();
            logMessage($"段落包含 {runs.Count} 个 Run 元素");

            foreach (var run in runs)
            {
                var runText = run.InnerText?.Trim() ?? "";
                logMessage($"  - Run 文本: \"{runText}\" (长度: {runText.Length})");

                var runProperties = run.RunProperties;
                if (runProperties == null)
                {
                    runProperties = new RunProperties();
                    run.AppendChild(runProperties);
                }

                // 设置字体
                var runFonts = runProperties.Elements<RunFonts>().FirstOrDefault();
                if (runFonts == null)
                {
                    runFonts = new RunFonts();
                    runProperties.AppendChild(runFonts);
                }

                // 所有段落（包括标题和正文）都采用相同的中英文字体处理方式
                // 中文字符保持用户设置的字体，英文和数字使用Times New Roman
                runFonts.EastAsia = styleConfig.FontName;      // 中文字符使用用户设置的字体
                runFonts.Ascii = "Times New Roman";              // 英文字符使用Times New Roman
                runFonts.HighAnsi = "Times New Roman";              // 英文字符使用Times New Roman
                runFonts.ComplexScript = "Times New Roman";                 // 复杂脚本使用Times New Roman

                // 记录字体设置信息
                string styleType = GetStyleTypeFromConfig(styleConfig);
                logMessage($"  - {styleType}字体设置：中文={styleConfig.FontName}, 英文/数字=Times New Roman");

                // 设置字体大小
                var fontSize = runProperties.Elements<FontSize>().FirstOrDefault();
                if (fontSize == null)
                {
                    fontSize = new FontSize();
                    runProperties.AppendChild(fontSize);
                }
                fontSize.Val = ConvertFontSizeToHalfPoints(styleConfig.FontSize);

                var fontSizeCs = runProperties.Elements<FontSizeComplexScript>().FirstOrDefault();
                if (fontSizeCs == null)
                {
                    fontSizeCs = new FontSizeComplexScript();
                    runProperties.AppendChild(fontSizeCs);
                }
                fontSizeCs.Val = ConvertFontSizeToHalfPoints(styleConfig.FontSize);

                // 设置加粗
                var bold = runProperties.Elements<Bold>().FirstOrDefault();
                if (bold == null)
                {
                    bold = new Bold();
                    runProperties.AppendChild(bold);
                }
                bold.Val = styleConfig.Bold;

                // 清除下划线（我们不希望保留原始文档中的下划线）
                var underline = runProperties.Elements<Underline>().FirstOrDefault();
                if (underline != null)
                {
                    runProperties.RemoveChild(underline);
                }
            }

            // 为自动编号设置字体格式
            SetNumberingFontFormat(wordDoc, paragraph, styleConfig, logMessage);

            // 调试信息：记录应用的格式（简化正文段落显示）
            string finalStyleType = GetStyleTypeFromConfig(styleConfig);
            string displayText = text;

            // 对于正文段落，只显示前50个字符以减少日志负担
            if (styleName == "Normal" && text.Length > 50)
            {
                displayText = text.Substring(0, 50) + "...";
            }

            logMessage($"应用格式到段落: [{finalStyleType}] 文字:\"{displayText}\" 字体:{styleConfig.FontName} 字号:{styleConfig.FontSize} 段前距:{styleConfig.SpaceBefore} 段后距:{styleConfig.SpaceAfter}");
        }

        /// <summary>
        /// 为图片段落应用专用样式：段前间距和段后间距都设置为0磅，行距设为单倍行距
        /// </summary>
        private void ApplyImageParagraphStyle(Paragraph paragraph, Action<string> logMessage)
        {
            var text = paragraph.InnerText.Trim();

            // 获取或创建段落属性
            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties == null)
            {
                paragraphProperties = new ParagraphProperties();
                paragraph.AppendChild(paragraphProperties);
            }

            // 设置段落间距
            var spacing = paragraphProperties.SpacingBetweenLines;
            if (spacing == null)
            {
                spacing = new SpacingBetweenLines();
                paragraphProperties.AppendChild(spacing);
            }

            // 图片段落专用设置：段前间距和段后间距都设置为0磅
            spacing.Before = "0"; // 段前间距0磅
            spacing.After = "0";  // 段后间距0磅

            // 强制设置单倍行距 - 清除任何可能影响行距的设置
            spacing.Line = "240"; // 单倍行距（240缇，即12磅）
            spacing.LineRule = LineSpacingRuleValues.Auto;

            // 图片段落不设置大纲级别
            logMessage($"应用图片段落样式: 段前距=0磅, 段后距=0磅, 行距=单倍行距");

            // 对于包含图片的段落，通常不需要设置字体属性，因为图片本身没有字体
            // 但为了保持一致性，我们仍然记录日志
            string displayText = text.Length > 30 ? text.Substring(0, 30) + "..." : text;
            logMessage($"图片段落处理完成: 包含图片的段落，文本内容: \"{displayText}\"");
        }

        /// <summary>
        /// 计算需要跳过的段落数量（基于分页符和分节符）
        /// </summary>
        private int CountParagraphsToSkip(WordprocessingDocument wordDoc, Body body, int skipPages)
        {
            if (skipPages <= 0)
                return 0;

            int pageBreaksFound = 0;
            int paragraphsToSkip = 0;

            // 遍历段落，统计分页符和分节符
            foreach (var paragraph in body.Elements<Paragraph>())
            {
                bool hasPageBreak = false;

                // 检查段落是否包含分页符
                if (ContainsPageBreak(paragraph))
                {
                    hasPageBreak = true;
                    pageBreaksFound++;

                    // 当找到足够的分页符时停止
                    if (pageBreaksFound >= skipPages)
                    {
                        // 跳过分页符所在的段落及其之后的所有段落
                        paragraphsToSkip++;
                        break;
                    }
                }

                // 检查段落是否是分节符开始（新的节从下一页开始）
                var sectionBreak = GetSectionBreak(paragraph);
                if (sectionBreak != null)
                {
                    // 检查是否是"下一页"分节符
                    if (IsNextPageSectionBreak(sectionBreak))
                    {
                        pageBreaksFound++;

                        // 当找到足够的分页符时停止
                        if (pageBreaksFound >= skipPages)
                        {
                            // 跳过分节符所在的段落及其之后的所有段落
                            paragraphsToSkip++;
                            break;
                        }
                    }
                }

                paragraphsToSkip++;
            }

            // 如果文档中没有足够的分页符/分节符（可能是自然分页），返回0不跳过
            // 让用户自己判断是否需要调整
            if (pageBreaksFound < skipPages)
            {
                // 如果检测到的分页符/分节符少于用户设置的跳过页数
                // 可能存在自然分页，返回0避免跳过多内容
                return 0;
            }

            return paragraphsToSkip;
        }

        /// <summary>
        /// 检查段落是否包含分页符
        /// </summary>
        private bool ContainsPageBreak(Paragraph paragraph)
        {
            try
            {
                var paragraphProperties = paragraph.ParagraphProperties;
                if (paragraphProperties == null)
                    return false;

                // 检查段落属性中的分页符
                var runs = paragraph.Elements<Run>().ToList();
                foreach (var run in runs)
                {
                    // 检查运行中的分页符
                    var breakElements = run.Elements<Break>().ToList();
                    foreach (var breakElement in breakElements)
                    {
                        // 检查是否为分页符（Type属性为null或"page"）
                        if (breakElement.Type == null || breakElement.Type.Value == BreakValues.Page)
                        {
                            return true;
                        }
                    }
                }

                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 获取段落的分节符信息
        /// </summary>
        private SectionProperties? GetSectionBreak(Paragraph paragraph)
        {
            try
            {
                var paragraphProperties = paragraph.ParagraphProperties;
                if (paragraphProperties == null)
                    return null;

                var sectionProperties = paragraphProperties.SectionProperties;
                if (sectionProperties == null)
                    return null;

                return sectionProperties;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 判断分节符是否是"下一页"类型（相当于分页符）
        /// </summary>
        private bool IsNextPageSectionBreak(SectionProperties sectionBreak)
        {
            try
            {
                // 检查 SectionType 子元素
                var sectionType = sectionBreak.Elements<SectionType>().FirstOrDefault();
                if (sectionType == null || sectionType.Val == null)
                    return true; // 默认认为是分页

                return sectionType.Val.Value == SectionMarkValues.NextPage;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 检查原始文档的样式定义，诊断大纲级别问题
        /// </summary>
        private void CheckOriginalDocumentStyles(WordprocessingDocument wordDoc, Action<string> logMessage)
        {
            try
            {
                var styleDefinitionsPart = wordDoc.MainDocumentPart?.StyleDefinitionsPart;

                logMessage("=== 原始文档样式诊断 ===");

                if (styleDefinitionsPart == null)
                {
                    logMessage("文档中没有样式定义部分");
                    logMessage("=== 样式诊断完成 ===");
                    return;
                }

                var styles = styleDefinitionsPart.Styles;
                if (styles == null)
                {
                    logMessage("样式定义部分为空");
                    logMessage("=== 样式诊断完成 ===");
                    return;
                }

                int styleCount = 0;
                foreach (var style in styles.Elements<Style>())
                {
                    if (style.StyleId != null && style.Type != null && style.StyleName != null)
                    {
                        styleCount++;
                        logMessage($"样式: {style.StyleName} (ID: {style.StyleId}, 类型: {style.Type})");
                    }
                }

                logMessage($"共发现 {styleCount} 个样式定义");
                logMessage("=== 样式诊断完成 ===");
            }
            catch (Exception ex)
            {
                logMessage($"样式诊断出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取段落的样式信息（使用 Open XML SDK 的真实样式）
        /// </summary>
        private ParagraphStyleInfo? GetParagraphStyleInfo(WordprocessingDocument wordDoc, Paragraph paragraph)
        {
            try
            {
                var paragraphProperties = paragraph.ParagraphProperties;
                if (paragraphProperties == null)
                    return null;

                var paragraphStyleId = paragraphProperties.ParagraphStyleId;
                if (paragraphStyleId == null || paragraphStyleId.Val == null)
                    return null;

                string styleId = paragraphStyleId.Val.Value;
                string? styleName = GetStyleNameFromStyleId(wordDoc, styleId);

                return new ParagraphStyleInfo
                {
                    StyleId = styleId,
                    StyleName = styleName,
                    ParagraphText = paragraph.InnerText
                };
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 根据 styleId 从样式定义中获取样式名称
        /// </summary>
        private string? GetStyleNameFromStyleId(WordprocessingDocument wordDoc, string styleId)
        {
            try
            {
                var styleDefinitionsPart = wordDoc.MainDocumentPart?.StyleDefinitionsPart;
                if (styleDefinitionsPart == null)
                    return null;

                var styles = styleDefinitionsPart.Styles;
                if (styles == null)
                    return null;

                // 查找匹配的样式
                var style = styles.Elements<Style>()
                    .FirstOrDefault(s => s.StyleId?.Value == styleId);

                if (style == null)
                    return null;

                // 方法1: 检查是否有 Aliases 元素（别名列表）
                var aliases = style.Elements<Aliases>().FirstOrDefault();
                if (aliases != null && !string.IsNullOrEmpty(aliases.Val?.Value))
                {
                    return aliases.Val.Value;
                }

                // 方法2: 检查是否有 Name 元素
                var nameElement = style.Elements<Name>().FirstOrDefault();
                if (nameElement != null && !string.IsNullOrEmpty(nameElement.Val?.Value))
                {
                    return nameElement.Val.Value;
                }

                // 方法3: 使用 StyleName 元素
                var styleName = style.StyleName;
                if (styleName != null && styleName.Val != null)
                {
                    return styleName.Val.Value;
                }

                return null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 检测段落是否包含图片
        /// </summary>
        private bool IsImageParagraph(Paragraph paragraph)
        {
            try
            {
                // 检查段落是否包含drawing元素
                foreach (var item in paragraph.ChildElements)
                {
                    if (item is Drawing || item is Picture)
                    {
                        return true;
                    }
                }

                // 检查段落中的运行元素是否包含图片
                foreach (var run in paragraph.Elements<Run>())
                {
                    foreach (var item in run.ChildElements)
                    {
                        if (item is Drawing || item is Picture || item is EmbeddedObject)
                        {
                            return true;
                        }
                    }
                }

                // 特别处理：如果段落文本为空但包含运行元素，可能是图片
                if (string.IsNullOrEmpty(paragraph.InnerText?.Trim()) && paragraph.Elements<Run>().Any())
                {
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                // 如果检测出错，记录错误并保守处理，不认为是图片段落
                System.Diagnostics.Debug.WriteLine($"检测图片段落时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 根据段落样式和文本推断样式类型
        /// 优先使用样式的大纲级别判断，只有在样式判断为正文时才使用正则表达式补充判断
        /// </summary>
        private string InferStyleFromParagraph(WordprocessingDocument? wordDoc, Paragraph paragraph, string text)
        {
            var trimmedText = text.Trim();
            if (string.IsNullOrEmpty(trimmedText))
            {
                return "Normal"; // 空段落视为正文
            }

            // 优先使用样式的大纲级别判断
            var outlineLevelStyle = InferStyleFromOutlineLevel(wordDoc, paragraph);
            
            // 如果样式已经判断为有级别的标题，则采用样式判断结果，不再使用正则表达式
            if (outlineLevelStyle != "Normal" && outlineLevelStyle != null)
            {
                return outlineLevelStyle;
            }

            // 如果通过样式判断为正文文本段落，则使用正则表达式补充判断
            // 获取包含自动编号的完整文本
            var paragraphTextWithNumbering = GetParagraphTextWithNumbering(wordDoc, paragraph);

            // 识别表格标题：以"表"开头，可能包含编号的段落
            if (IsTableCaptionByPattern(paragraphTextWithNumbering))
            {
                return "TableCaption";
            }

            // 识别图形标题：以"图"开头，可能包含编号的段落
            if (IsImageCaptionByPattern(paragraphTextWithNumbering))
            {
                return "ImageCaption";
            }

            // 通过正则表达式识别标题格式
            // 一级标题：格式如 "1 标题" 或 "2 水文" 等
            if (Regex.IsMatch(trimmedText, @"^[1-9]\d?\s+[\u4e00-\u9fa5]+$"))
            {
                return "Heading1";
            }

            // 二级标题：格式如 "1.1 标题" 或 "2.1 设计洪水" 等
            if (Regex.IsMatch(trimmedText, @"^[1-9]\d?\.[1-9]\d?\s*[\u4e00-\u9fa5]+$"))
            {
                return "Heading2";
            }

            // 三级标题：格式如 "1.1.1 标题" 等
            if (Regex.IsMatch(trimmedText, @"^[1-9]\d?\.[1-9]\d?\.[1-9]\d?\s*[\u4e00-\u9fa5]+$"))
            {
                return "Heading3";
            }

            // 四级标题：格式如 "1.1.1.1 标题" 等
            if (Regex.IsMatch(trimmedText, @"^[1-9]\d?\.[1-9]\d?\.[1-9]\d?\.[1-9]\d?\s*[\u4e00-\u9fa5]+$"))
            {
                return "Heading4";
            }

            // 其他格式都视为正文
            return "Normal";
        }

        /// <summary>
        /// 从段落样式中获取大纲级别，并根据大纲级别返回样式类型
        /// 大纲级别 1 = Heading1, 2 = Heading2, 3 = Heading3, 4 = Heading4
        /// 正文或其他 = Normal
        /// </summary>
        private string? InferStyleFromOutlineLevel(WordprocessingDocument? wordDoc, Paragraph paragraph)
        {
            if (wordDoc == null)
                return null;

            try
            {
                var paragraphProperties = paragraph.ParagraphProperties;
                if (paragraphProperties == null)
                    return null;

                // 首先尝试从段落属性直接读取大纲级别
                var outlineLvl = paragraphProperties.OutlineLevel;
                int level = -1;

                if (outlineLvl != null && outlineLvl.Val != null)
                {
                    // 段落直接设置了大纲级别
                    if (int.TryParse(outlineLvl.Val.InnerText, out level))
                    {
                        // Word大纲级别：0=一级标题, 1=二级标题, 2=三级标题, 3=四级标题
                        return level switch
                        {
                            0 => "Heading1",
                            1 => "Heading2",
                            2 => "Heading3",
                            3 => "Heading4",
                            _ => "Normal"
                        };
                    }
                }
                else
                {
                    // 段落未直接设置大纲级别，尝试从样式定义中获取
                    var paragraphStyleId = paragraphProperties.ParagraphStyleId;
                    if (paragraphStyleId != null && paragraphStyleId.Val != null)
                    {
                        string styleId = paragraphStyleId.Val.Value;
                        int? styleOutlineLevel = GetOutlineLevelFromStyleDefinition(wordDoc, styleId);

                        if (styleOutlineLevel.HasValue)
                        {
                            level = styleOutlineLevel.Value;
                            // Word大纲级别：0=一级标题, 1=二级标题, 2=三级标题, 3=四级标题
                            return level switch
                            {
                                0 => "Heading1",
                                1 => "Heading2",
                                2 => "Heading3",
                                3 => "Heading4",
                                _ => "Normal"
                            };
                        }
                    }
                }

                return null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 从样式定义中获取大纲级别
        /// </summary>
        private static int? GetOutlineLevelFromStyleDefinition(WordprocessingDocument wordDoc, string styleId)
        {
            try
            {
                var styleDefinitionsPart = wordDoc.MainDocumentPart?.StyleDefinitionsPart;
                if (styleDefinitionsPart == null)
                    return null;

                var styles = styleDefinitionsPart.Styles;
                if (styles == null)
                    return null;

                // 查找匹配的样式
                var style = styles.Elements<Style>()
                    .FirstOrDefault(s => s.StyleId?.Value == styleId);

                if (style == null)
                    return null;

                // 从样式的段落属性中获取大纲级别
                var styleParagraphProperties = style.StyleParagraphProperties;
                if (styleParagraphProperties != null)
                {
                    var outlineLvl = styleParagraphProperties.OutlineLevel;
                    if (outlineLvl != null && outlineLvl.Val != null)
                    {
                        if (int.TryParse(outlineLvl.Val.InnerText, out int level))
                        {
                            return level;
                        }
                    }
                }

                return null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 根据段落文本格式推断样式类型（通过正则表达式识别标题编号）
        /// 此方法仅作为后备方案，当样式无法判断时使用
        /// </summary>
        [Obsolete("请使用 InferStyleFromParagraph(WordprocessingDocument, Paragraph, string) 方法")]
        private string InferStyleFromParagraphLegacy(string text)
        {
            var trimmedText = text.Trim();
            if (string.IsNullOrEmpty(trimmedText))
            {
                return "Normal"; // 空段落视为正文
            }

            // 识别表格标题：以"表"开头，可能包含编号的段落
            if (IsTableCaptionByPattern(trimmedText))
            {
                return "TableCaption";
            }

            // 识别图形标题：以"图"开头，可能包含编号的段落
            if (IsImageCaptionByPattern(trimmedText))
            {
                return "ImageCaption";
            }

            // 通过正则表达式识别标题格式（更严格的规则，避免误识别正文中的一般数字）

            // 一级标题：格式如 "1 标题" 或 "2 水文" 等
            if (Regex.IsMatch(trimmedText, @"^[1-9]\d?\s+[\u4e00-\u9fa5]+$"))
            {
                return "Heading1";
            }

            // 二级标题：格式如 "1.1 标题" 或 "2.1 设计洪水" 等
            if (Regex.IsMatch(trimmedText, @"^[1-9]\d?\.[1-9]\d?\s*[\u4e00-\u9fa5]+$"))
            {
                return "Heading2";
            }

            // 三级标题：格式如 "1.1.1 标题" 等
            if (Regex.IsMatch(trimmedText, @"^[1-9]\d?\.[1-9]\d?\.[1-9]\d?\s*[\u4e00-\u9fa5]+$"))
            {
                return "Heading3";
            }

            // 四级标题：格式如 "1.1.1.1 标题" 等
            if (Regex.IsMatch(trimmedText, @"^[1-9]\d?\.[1-9]\d?\.[1-9]\d?\.[1-9]\d?\s*[\u4e00-\u9fa5]+$"))
            {
                return "Heading4";
            }

            // 其他格式都视为正文
            return "Normal";
        }

        /// <summary>
        /// 获取段落的完整文本(包括自动编号)
        /// </summary>
        private string GetParagraphTextWithNumbering(WordprocessingDocument? wordDoc, Paragraph paragraph)
        {
            var baseText = paragraph.InnerText?.Trim() ?? "";
            
            // 如果没有文档或段落为空,直接返回
            if (wordDoc == null || string.IsNullOrEmpty(baseText))
            {
                return baseText;
            }

            try
            {
                // 检查段落是否有编号属性
                var paragraphProperties = paragraph.ParagraphProperties;
                if (paragraphProperties == null)
                {
                    return baseText;
                }

                // 获取编号引用
                var numberingProperties = paragraphProperties.NumberingProperties;
                if (numberingProperties == null)
                {
                    // 如果没有编号属性,检查是否包含域代码
                    return ProcessTextWithFieldCodes(baseText);
                }

                var numberingId = numberingProperties.NumberingId?.Val?.Value;
                var indentationLevel = numberingProperties.NumberingLevelReference?.Val?.Value;

                if (numberingId == null)
                {
                    // 如果没有编号ID,检查是否包含域代码
                    return ProcessTextWithFieldCodes(baseText);
                }

                // 从NumberingDefinitionsPart获取编号定义
                var numberingPart = wordDoc.MainDocumentPart?.NumberingDefinitionsPart;
                if (numberingPart == null)
                {
                    // 如果没有编号定义,检查是否包含域代码
                    return ProcessTextWithFieldCodes(baseText);
                }

                var numbering = numberingPart.Numbering;
                if (numbering == null)
                {
                    // 如果没有编号,检查是否包含域代码
                    return ProcessTextWithFieldCodes(baseText);
                }

                // 查找对应的编号实例
                var numberingInstance = numbering.Elements<DocumentFormat.OpenXml.Wordprocessing.NumberingInstance>()
                    .FirstOrDefault(ni => ni.NumberID?.Value == numberingId);
                
                if (numberingInstance == null)
                {
                    // 如果没有编号实例,检查是否包含域代码
                    return ProcessTextWithFieldCodes(baseText);
                }

                // 获取编号文本(如果存在)
                var levelOverride = numberingInstance.Elements<LevelOverride>()
                    .FirstOrDefault(lo => lo.LevelIndex?.Value == indentationLevel);

                string numberingText = "";
                
                if (levelOverride != null)
                {
                    var levelOverrideLevel = levelOverride.Elements<Level>().FirstOrDefault();
                    if (levelOverrideLevel != null)
                    {
                        var lvlText = levelOverrideLevel.LevelText?.Val?.Value;
                        
                        if (lvlText != null && !string.IsNullOrEmpty(lvlText))
                        {
                            numberingText = ExtractNumberingPrefixFromLevelText(lvlText);
                        }
                    }
                }
                
                // 如果没有在override中找到,尝试从abstractNum中查找
                if (string.IsNullOrEmpty(numberingText))
                {
                    var abstractNumId = numberingInstance.AbstractNumId?.Val?.Value;
                    if (abstractNumId != null)
                    {
                        var abstractNum = numbering.Elements<AbstractNum>()
                            .FirstOrDefault(an => an.AbstractNumberId?.Value == abstractNumId);
                        
                        if (abstractNum != null)
                        {
                            var level = abstractNum.Elements<Level>()
                                .FirstOrDefault(l => l.LevelIndex?.Value == indentationLevel);
                            
                            if (level != null)
                            {
                                var lvlText = level.LevelText?.Val?.Value;
                                if (lvlText != null && !string.IsNullOrEmpty(lvlText))
                                {
                                    numberingText = ExtractNumberingPrefixFromLevelText(lvlText);
                                }
                            }
                        }
                    }
                }

                // 如果成功提取到编号前缀,拼接编号前缀和文本
                if (!string.IsNullOrEmpty(numberingText))
                {
                    return $"{numberingText}{baseText}";
                }
            }
            catch (Exception ex)
            {
                // 如果提取失败,记录错误并返回处理后的文本
                System.Diagnostics.Debug.WriteLine($"提取编号失败: {ex.Message}");
            }

            // 最后尝试处理域代码
            return ProcessTextWithFieldCodes(baseText);
        }

        /// <summary>
        /// 处理包含域代码的文本,提取实际内容
        /// 例如: "表  STYLEREF 3 \s 1.3.1 SEQ 表 \* ARABIC \s 3 1 白塔河流..." -> "表  白塔河流..."
        /// </summary>
        private string ProcessTextWithFieldCodes(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            // 检查是否包含域代码标记
            if (!text.Contains("STYLEREF") && !text.Contains("SEQ") && !text.Contains("REF"))
            {
                return text;
            }

            // 如果文本以"表 "开头,保留"表 "并移除域代码
            if (text.StartsWith("表 "))
            {
                // 移除所有域代码标记和参数
                var cleanedText = Regex.Replace(text, @"表\s+(STYLEREF|SEQ|REF|TOC|DATE|TIME|AUTHOR|TITLE|PAGE|NUMPAGES|SECTIONPAGES)[^\u4e00-\u9fa5]*", "表 ");
                return cleanedText.Trim();
            }

            // 如果文本以"图 "开头,保留"图 "并移除域代码
            if (text.StartsWith("图 "))
            {
                var cleanedText = Regex.Replace(text, @"图\s+(STYLEREF|SEQ|REF|TOC|DATE|TIME|AUTHOR|TITLE|PAGE|NUMPAGES|SECTIONPAGES)[^\u4e00-\u9fa5]*", "图 ");
                return cleanedText.Trim();
            }

            // 对于其他情况,移除所有域代码,但保留开头的"表"或"图"
            var result = Regex.Replace(text, @"\s+(STYLEREF|SEQ|REF|TOC|DATE|TIME|AUTHOR|TITLE|PAGE|NUMPAGES|SECTIONPAGES)[^\u4e00-\u9fa5]*", " ");
            return result.Trim();
        }

        /// <summary>
        /// 从LevelText中提取编号前缀文本
        /// 例如: "表 %1.%2" -> "表 "
        /// </summary>
        private string ExtractNumberingPrefixFromLevelText(string levelText)
        {
            // LevelText通常包含占位符,如 "%1.%2.%3" 或 "表 %1.%2"
            // 我们需要提取占位符前面的文本前缀
            if (string.IsNullOrEmpty(levelText))
            {
                return "";
            }

            // 查找第一个占位符 "%数字" 的位置
            var match = Regex.Match(levelText, @"%\d+");
            if (match.Success)
            {
                // 返回占位符前面的文本
                return levelText.Substring(0, match.Index).TrimEnd();
            }

            // 如果没有找到占位符,返回空字符串
            return "";
        }

        /// <summary>
        /// 通过文本模式判断是否为表格标题
        /// </summary>
        private bool IsTableCaptionByPattern(string text)
        {
            return Regex.IsMatch(text, @"^\s*表(?=\s|\d)\s*\d*(\.\d+)*(-\s*\d+)*\s+.*$");
        }

        /// <summary>
        /// 通过文本模式判断是否为图形标题
        /// </summary>
        private bool IsImageCaptionByPattern(string text)
        {
            return Regex.IsMatch(text, @"^\s*图(?=\s|\d)\s*\d*(\.\d+)*(-\s*\d+)*\s+.*$");
        }

        /// <summary>
        /// 获取段落的详细信息
        /// </summary>
        public Dictionary<string, object> GetParagraphDetails(Paragraph paragraph)
        {
            var details = new Dictionary<string, object>();

            // 获取段落文本
            string text = paragraph.InnerText.Trim();
            details["Text"] = text;

            // 尝试从文档中获取实际样式
            if (_document != null)
            {
                var styleInfo = GetParagraphStyleInfo(_document, paragraph);
                // 优先使用样式的大纲级别判断，只有在样式判断为正文时才使用正则表达式补充判断
                details["StyleType"] = InferStyleFromParagraph(_document, paragraph, text);
                details["StyleId"] = styleInfo?.StyleName ?? "无";
            }
            else
            {
                // 如果文档未打开，通过正则表达式推断样式类型
                details["StyleType"] = InferStyleFromParagraph(null, paragraph, text);
                details["StyleId"] = "无";
            }

            // 获取字体名称
            string fontName = GetParagraphFontName(paragraph);
            details["FontName"] = fontName;

            // 获取字体大小
            double fontSize = GetParagraphFontSize(paragraph);
            details["FontSize"] = fontSize;

            // 获取段前间距
            double spaceBefore = GetParagraphSpaceBefore(paragraph);
            details["SpaceBefore"] = spaceBefore;

            // 获取段后间距
            double spaceAfter = GetParagraphSpaceAfter(paragraph);
            details["SpaceAfter"] = spaceAfter;

            return details;
        }

        /// <summary>
        /// 获取段落的字体名称
        /// </summary>
        private string GetParagraphFontName(Paragraph paragraph)
        {
            // 默认字体名称
            string defaultFontName = "宋体";

            // 使用字典统计各字体出现次数，返回出现次数最多的字体
            var fontCounts = new Dictionary<string, int>();

            foreach (var run in paragraph.Elements<Run>())
            {
                var runProperties = run.RunProperties;
                if (runProperties != null)
                {
                    var runFonts = runProperties.Elements<RunFonts>().FirstOrDefault();
                    if (runFonts != null)
                    {
                        // 优先使用EastAsia字体（中文），然后是HighAnsi，最后是Ascii
                        string fontName = !string.IsNullOrEmpty(runFonts.EastAsia?.Value) ? runFonts.EastAsia.Value :
                                         !string.IsNullOrEmpty(runFonts.HighAnsi?.Value) ? runFonts.HighAnsi.Value :
                                         !string.IsNullOrEmpty(runFonts.Ascii?.Value) ? runFonts.Ascii.Value :
                                         defaultFontName;

                        if (!string.IsNullOrEmpty(fontName) && fontName != defaultFontName)
                        {
                            if (fontCounts.ContainsKey(fontName))
                            {
                                fontCounts[fontName]++;
                            }
                            else
                            {
                                fontCounts[fontName] = 1;
                            }
                        }
                    }
                }
            }

            // 如果Run中有字体设置，返回出现次数最多的字体
            if (fontCounts.Any())
            {
                return fontCounts.OrderByDescending(kvp => kvp.Value).First().Key;
            }

            // 默认返回宋体
            return defaultFontName;
        }

        /// <summary>
        /// 获取段落的字体大小（磅）
        /// </summary>
        private double GetParagraphFontSize(Paragraph paragraph)
        {
            // 默认字体大小
            double defaultFontSize = 10.5;

            // 使用字典统计各字体大小出现次数，返回出现次数最多的字体大小
            var fontSizeCounts = new Dictionary<double, int>();

            foreach (var run in paragraph.Elements<Run>())
            {
                var runProperties = run.RunProperties;
                if (runProperties != null)
                {
                    var fontSize = runProperties.Elements<FontSize>().FirstOrDefault();
                    if (fontSize != null && fontSize.Val != null && !string.IsNullOrEmpty(fontSize.Val.Value))
                    {
                        if (double.TryParse(fontSize.Val.Value, out double fontSizeValue))
                        {
                            double fontSizeInPoints = fontSizeValue / 2.0; // 转换为磅

                            if (fontSizeCounts.ContainsKey(fontSizeInPoints))
                            {
                                fontSizeCounts[fontSizeInPoints]++;
                            }
                            else
                            {
                                fontSizeCounts[fontSizeInPoints] = 1;
                            }
                        }
                    }
                }
            }

            // 如果Run中有字体大小设置，返回出现次数最多的字体大小
            if (fontSizeCounts.Any())
            {
                return fontSizeCounts.OrderByDescending(kvp => kvp.Value).First().Key;
            }

            // 默认返回10.5磅
            return defaultFontSize;
        }

        /// <summary>
        /// 获取段落的段前间距（磅）
        /// </summary>
        private double GetParagraphSpaceBefore(Paragraph paragraph)
        {
            // 默认段前间距
            double defaultSpaceBefore = 0;

            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties != null && paragraphProperties.SpacingBetweenLines != null)
            {
                var spacing = paragraphProperties.SpacingBetweenLines;
                if (!string.IsNullOrEmpty(spacing.Before))
                {
                    if (int.TryParse(spacing.Before, out int beforeValue))
                    {
                        return beforeValue / 20.0; // 转换为磅（1磅=20缇）
                    }
                }
            }

            // 默认返回0磅
            return defaultSpaceBefore;
        }

        /// <summary>
        /// 获取段落的段后间距（磅）
        /// </summary>
        private double GetParagraphSpaceAfter(Paragraph paragraph)
        {
            // 默认段后间距
            double defaultSpaceAfter = 0;

            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties != null && paragraphProperties.SpacingBetweenLines != null)
            {
                var spacing = paragraphProperties.SpacingBetweenLines;
                if (!string.IsNullOrEmpty(spacing.After))
                {
                    if (int.TryParse(spacing.After, out int afterValue))
                    {
                        return afterValue / 20.0; // 转换为磅（1磅=20缇）
                    }
                }
            }

            // 默认返回0磅
            return defaultSpaceAfter;
        }

        /// <summary>
        /// 根据样式名称获取对应的段落样式配置
        /// </summary>
        private ParagraphStyle? GetParagraphStyle(string styleName, Models.StyleConfig config)
        {
            return styleName switch
            {
                "Heading1" => config.Heading1,
                "Heading2" => config.Heading2,
                "Heading3" => config.Heading3,
                "Heading4" => config.Heading4,
                "Normal" => config.Normal,
                "TableCaption" => config.TableCaption,
                "Image" => CreateImageParagraphStyle(),
                _ => config.Normal
            };
        }

        /// <summary>
        /// 创建图片段落的专用样式配置
        /// </summary>
        private ParagraphStyle CreateImageParagraphStyle()
        {
            return new ParagraphStyle
            {
                FontName = "宋体",
                FontSize = 10.5,
                SpaceBefore = 0,
                SpaceAfter = 0,
                LineSpacing = 1.0
            };
        }

        /// <summary>
        /// 根据样式配置获取样式类型描述
        /// </summary>
        private string GetStyleTypeFromConfig(ParagraphStyle styleConfig)
        {
            if (styleConfig.FontSize >= 16) return "一级标题";
            if (styleConfig.FontSize >= 14) return "二级标题";
            if (styleConfig.FontSize >= 13) return "三级标题";
            if (styleConfig.FontSize >= 12) return "四级标题";
            return "正文文本";
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
            if (!string.IsNullOrEmpty(halfPoints) && double.TryParse(halfPoints, out double value))
            {
                return value / 2.0;
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

        /// <summary>
        /// 根据样式名称获取对应的大纲级别
        /// 注意：Word大纲级别从0开始，0=最高级标题，1=二级标题，以此类推
        /// </summary>
        private int GetOutlineLevelForStyle(string styleName, Models.StyleConfig? config = null)
        {
            return styleName switch
            {
                "Normal" => 9,
                "Image" => 9,
                "TableCaption" => config?.TableCaption.OutlineLevel ?? 9,
                "ImageCaption" => config?.ImageCaption.OutlineLevel ?? 9,
                "Heading1" => 0,
                "Heading2" => 1,
                "Heading3" => 2,
                "Heading4" => 3,
                _ => 9
            };
        }

        /// <summary>
        /// 为段落设置大纲级别
        /// </summary>
        private void SetOutlineLevelForParagraph(Paragraph paragraph, string styleName, Action<string> logMessage, Models.StyleConfig? config = null)
        {
            if (styleName == "Normal" || styleName == "Image")
            {
                if (styleName == "Normal")
                {
                    logMessage($"正文段落：不设置大纲级别");
                }
                else
                {
                    logMessage($"图片段落：不设置大纲级别");
                }
                return;
            }

            int targetLevel = GetOutlineLevelForStyle(styleName, config);

            // 获取或创建段落属性
            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties == null)
            {
                paragraphProperties = new ParagraphProperties();
                paragraph.AppendChild(paragraphProperties);
            }

            // 设置大纲级别
            var outlineLvl = paragraphProperties.OutlineLevel;
            if (outlineLvl == null)
            {
                outlineLvl = new OutlineLevel();
                paragraphProperties.AppendChild(outlineLvl);
            }
            outlineLvl.Val = targetLevel;

            logMessage($"设置大纲级别: 样式名={styleName}, 目标级别={targetLevel}");
        }

        /// <summary>
        /// 尝试从样式定义中查找并设置自动编号的字体格式
        /// </summary>
        private bool TrySetStyleNumberingFormat(WordprocessingDocument wordDoc, int styleIdValue, ParagraphStyle styleConfig, Action<string> logMessage)
        {
            try
            {
                var styleDefinitionsPart = wordDoc.MainDocumentPart?.StyleDefinitionsPart;
                if (styleDefinitionsPart == null)
                {
                    logMessage("    文档中没有样式定义部分");
                    return false;
                }

                var styles = styleDefinitionsPart.Styles;
                if (styles == null)
                {
                    logMessage("    样式定义部分为空");
                    return false;
                }

                // 查找对应的样式
                var style = styles.Elements<Style>()
                    .FirstOrDefault(s => s.StyleId?.Value == styleIdValue.ToString());

                if (style == null)
                {
                    logMessage($"    未找到样式ID={styleIdValue}");
                    return false;
                }

                logMessage($"    找到样式: {style.StyleName?.Val?.Value ?? "无名称"} (ID={style.StyleId?.Value})");

                // 检查样式的段落属性中是否有自动编号
                var styleParagraphProperties = style.StyleParagraphProperties;
                if (styleParagraphProperties == null)
                {
                    logMessage("    样式没有段落属性");
                    return false;
                }

                var styleNumbering = styleParagraphProperties.NumberingProperties;
                if (styleNumbering == null)
                {
                    logMessage("    样式段落属性中没有自动编号");
                    return false;
                }

                var numberingId = styleNumbering.NumberingId?.Val?.Value;
                var indentationLevel = styleNumbering.NumberingLevelReference?.Val?.Value;

                logMessage($"    样式包含自动编号: 编号ID={numberingId}, 级别={indentationLevel}");

                if (numberingId == null)
                    return false;

                if (!numberingId.HasValue)
                {
                    logMessage("    编号ID为null，无法继续");
                    return false;
                }

                // 现在查找NumberingDefinitionsPart中的编号定义并设置字体
                var numberingPart = wordDoc.MainDocumentPart?.NumberingDefinitionsPart;
                if (numberingPart == null)
                {
                    logMessage("    文档中没有编号定义部分");
                    return false;
                }

                var numbering = numberingPart.Numbering;
                if (numbering == null)
                {
                    logMessage("    编号定义为空");
                    return false;
                }

                // 查找对应的编号实例
                var numberingInstanceId = numberingId.Value;
                var numberingInstanceIdStr = numberingInstanceId.ToString();
                var numberingInstance = numbering.Elements<DocumentFormat.OpenXml.Wordprocessing.NumberingInstance>()
                    .FirstOrDefault(ni => ni.NumberID != null && ni.NumberID.InnerText == numberingInstanceIdStr);

                if (numberingInstance == null)
                {
                    logMessage($"    未找到编号实例：ID={numberingId}");
                    return false;
                }

                // 查找或创建LevelOverride
                var levelOverride = numberingInstance.Elements<LevelOverride>()
                    .FirstOrDefault(lo => lo.LevelIndex?.Value == indentationLevel);

                Level levelToModify;

                if (levelOverride == null)
                {
                    logMessage($"    创建新的LevelOverride：级别={indentationLevel}");
                    levelOverride = new LevelOverride();
                    levelOverride.LevelIndex = indentationLevel.Value;
                    numberingInstance.AppendChild(levelOverride);

                    // 从AbstractNum中复制原始Level定义
                    var abstractNumId = numberingInstance.AbstractNumId?.Val?.Value;
                    if (abstractNumId != null)
                    {
                        var abstractNum = numbering.Elements<AbstractNum>()
                            .FirstOrDefault(an => an.AbstractNumberId?.Value == abstractNumId);

                        if (abstractNum != null)
                        {
                            var originalLevel = abstractNum.Elements<Level>()
                                .FirstOrDefault(l => l.LevelIndex?.Value == indentationLevel);

                            if (originalLevel != null)
                            {
                                levelToModify = (Level)originalLevel.CloneNode(true);
                                levelOverride.AppendChild(levelToModify);
                            }
                            else
                            {
                                levelToModify = new Level();
                                levelToModify.LevelIndex = indentationLevel.Value;
                                levelOverride.AppendChild(levelToModify);
                            }
                        }
                        else
                        {
                            levelToModify = new Level();
                            levelToModify.LevelIndex = indentationLevel.Value;
                            levelOverride.AppendChild(levelToModify);
                        }
                    }
                    else
                    {
                        levelToModify = new Level();
                        levelToModify.LevelIndex = indentationLevel.Value;
                        levelOverride.AppendChild(levelToModify);
                    }
                }
                else
                {
                    levelToModify = levelOverride.Elements<Level>().FirstOrDefault();
                    if (levelToModify == null)
                    {
                        levelToModify = new Level();
                        levelToModify.LevelIndex = indentationLevel.Value;
                        levelOverride.AppendChild(levelToModify);
                    }
                    logMessage($"    使用现有的LevelOverride：级别={indentationLevel}");
                }

                // 获取或创建Level的RunProperties
                var runProperties = levelToModify.Elements<RunProperties>().FirstOrDefault();
                if (runProperties == null)
                {
                    runProperties = new RunProperties();
                    levelToModify.AppendChild(runProperties);
                    logMessage("    创建新的RunProperties");
                }

                // 设置字体
                var runFonts = runProperties.Elements<RunFonts>().FirstOrDefault();
                if (runFonts == null)
                {
                    runFonts = new RunFonts();
                    runProperties.AppendChild(runFonts);
                }

                // 设置编号字体：中文字符使用用户设置的字体，英文使用Times New Roman
                runFonts.EastAsia = styleConfig.FontName;
                runFonts.Ascii = "Times New Roman";
                runFonts.HighAnsi = "Times New Roman";
                runFonts.ComplexScript = "Times New Roman";

                // 设置字体大小
                var fontSize = runProperties.Elements<FontSize>().FirstOrDefault();
                if (fontSize == null)
                {
                    fontSize = new FontSize();
                    runProperties.AppendChild(fontSize);
                }
                fontSize.Val = ConvertFontSizeToHalfPoints(styleConfig.FontSize);

                var fontSizeCs = runProperties.Elements<FontSizeComplexScript>().FirstOrDefault();
                if (fontSizeCs == null)
                {
                    fontSizeCs = new FontSizeComplexScript();
                    runProperties.AppendChild(fontSizeCs);
                }
                fontSizeCs.Val = ConvertFontSizeToHalfPoints(styleConfig.FontSize);

                // 设置加粗
                var bold = runProperties.Elements<Bold>().FirstOrDefault();
                if (bold == null)
                {
                    bold = new Bold();
                    runProperties.AppendChild(bold);
                }
                bold.Val = styleConfig.Bold;

                logMessage($"    已设置样式自动编号字体：字体={styleConfig.FontName}, 字号={styleConfig.FontSize}, 加粗={styleConfig.Bold}");
                return true;
            }
            catch (Exception ex)
            {
                logMessage($"    处理样式自动编号时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 为段落的自动编号设置字体格式
        /// </summary>
        /// <param name="wordDoc">Word文档对象</param>
        /// <param name="paragraph">段落对象</param>
        /// <param name="styleConfig">段落样式配置</param>
        /// <param name="logMessage">日志记录回调</param>
        private void SetNumberingFontFormat(WordprocessingDocument wordDoc, Paragraph paragraph, ParagraphStyle styleConfig, Action<string> logMessage)
        {
            try
            {
                var paragraphProperties = paragraph.ParagraphProperties;
                if (paragraphProperties == null)
                {
                    logMessage("段落没有 ParagraphProperties");
                    return;
                }

                // 诊断：打印段落属性的详细信息
                logMessage("=== 段落属性诊断 ===");
                var paragraphStyleId = paragraphProperties.ParagraphStyleId?.Val?.Value;
                logMessage($"  StyleId: {paragraphStyleId ?? "无"}");

                var numberingProperties = paragraphProperties.NumberingProperties;
                int? numberingId = null;
                int? indentationLevel = null;

                if (numberingProperties == null)
                {
                    logMessage("  NumberingProperties: 无");
                }
                else
                {
                    numberingId = numberingProperties.NumberingId?.Val?.Value;
                    indentationLevel = numberingProperties.NumberingLevelReference?.Val?.Value;
                    logMessage($"  NumberingProperties: 有 (编号ID={numberingId}, 级别={indentationLevel})");
                }

                // 如果没有NumberingProperties，但段落有StyleId，尝试从样式定义中查找自动编号
                if (numberingProperties == null && paragraphStyleId != null)
                {
                    logMessage($"  尝试从样式定义 (StyleId={paragraphStyleId}) 中查找自动编号...");
                    int? styleIdInt = null;
                    if (int.TryParse(paragraphStyleId, out int id))
                    {
                        styleIdInt = id;
                    }
                    else
                    {
                        logMessage("  无法将StyleId转换为整数");
                        return;
                    }
                    bool styleHasNumbering = TrySetStyleNumberingFormat(wordDoc, styleIdInt.Value, styleConfig, logMessage);
                    if (styleHasNumbering)
                    {
                        logMessage("  成功从样式定义中设置自动编号格式");
                        return;
                    }
                    else
                    {
                        logMessage("  样式定义中没有自动编号或设置失败");
                    }
                }

                if (numberingProperties == null)
                {
                    logMessage("段落没有使用自动编号，跳过编号格式设置");
                    return;
                }

                if (numberingId == null)
                {
                    logMessage("编号ID为空，跳过");
                    return;
                }

                logMessage($"检测到自动编号：编号ID={numberingId}, 级别={indentationLevel}");

                var numberingPart = wordDoc.MainDocumentPart?.NumberingDefinitionsPart;
                if (numberingPart == null)
                {
                    logMessage("文档中没有编号定义部分");
                    return;
                }

                var numbering = numberingPart.Numbering;
                if (numbering == null)
                {
                    logMessage("编号定义为空");
                    return;
                }

                // 查找对应的编号实例
                var numberingInstance = numbering.Elements<DocumentFormat.OpenXml.Wordprocessing.NumberingInstance>()
                    .FirstOrDefault(ni => ni.NumberID?.Value == numberingId);

                if (numberingInstance == null)
                {
                    logMessage($"未找到编号实例：ID={numberingId}");
                    return;
                }

                logMessage($"找到编号实例：ID={numberingId}");
                logMessage("=== 诊断结束 ===");

                // 查找或创建LevelOverride
                var levelOverride = numberingInstance.Elements<LevelOverride>()
                    .FirstOrDefault(lo => lo.LevelIndex?.Value == indentationLevel);

                Level levelToModify;

                if (levelOverride == null)
                {
                    // 如果没有LevelOverride，创建一个新的
                    logMessage($"创建新的LevelOverride：级别={indentationLevel}");
                    levelOverride = new LevelOverride();
                    levelOverride.LevelIndex = indentationLevel.Value;
                    numberingInstance.AppendChild(levelOverride);

                    // 从AbstractNum中复制原始Level定义
                    var abstractNumId = numberingInstance.AbstractNumId?.Val?.Value;
                    if (abstractNumId != null)
                    {
                        var abstractNum = numbering.Elements<AbstractNum>()
                            .FirstOrDefault(an => an.AbstractNumberId?.Value == abstractNumId);

                        if (abstractNum != null)
                        {
                            var originalLevel = abstractNum.Elements<Level>()
                                .FirstOrDefault(l => l.LevelIndex?.Value == indentationLevel);

                            if (originalLevel != null)
                            {
                                // 克隆原始Level到LevelOverride
                                levelToModify = (Level)originalLevel.CloneNode(true);
                                levelOverride.AppendChild(levelToModify);
                            }
                            else
                            {
                                // 如果没有找到原始Level，创建一个新的
                                levelToModify = new Level();
                                levelToModify.LevelIndex = indentationLevel.Value;
                                levelOverride.AppendChild(levelToModify);
                            }
                        }
                        else
                        {
                            levelToModify = new Level();
                            levelToModify.LevelIndex = indentationLevel.Value;
                            levelOverride.AppendChild(levelToModify);
                        }
                    }
                    else
                    {
                        levelToModify = new Level();
                        levelToModify.LevelIndex = indentationLevel.Value;
                        levelOverride.AppendChild(levelToModify);
                    }
                }
                else
                {
                    // 如果已有LevelOverride，获取其中的Level
                    levelToModify = levelOverride.Elements<Level>().FirstOrDefault();
                    if (levelToModify == null)
                    {
                        levelToModify = new Level();
                        levelToModify.LevelIndex = indentationLevel.Value;
                        levelOverride.AppendChild(levelToModify);
                    }
                    logMessage($"使用现有的LevelOverride：级别={indentationLevel}");
                }

                // 获取或创建Level的RunProperties
                var runProperties = levelToModify.Elements<RunProperties>().FirstOrDefault();
                if (runProperties == null)
                {
                    runProperties = new RunProperties();
                    levelToModify.AppendChild(runProperties);
                    logMessage("创建新的RunProperties");
                }

                // 设置字体
                var runFonts = runProperties.Elements<RunFonts>().FirstOrDefault();
                if (runFonts == null)
                {
                    runFonts = new RunFonts();
                    runProperties.AppendChild(runFonts);
                }

                // 设置编号字体：中文字符使用用户设置的字体，英文使用Times New Roman
                runFonts.EastAsia = styleConfig.FontName;
                runFonts.Ascii = "Times New Roman";
                runFonts.HighAnsi = "Times New Roman";
                runFonts.ComplexScript = "Times New Roman";

                // 设置字体大小
                var fontSize = runProperties.Elements<FontSize>().FirstOrDefault();
                if (fontSize == null)
                {
                    fontSize = new FontSize();
                    runProperties.AppendChild(fontSize);
                }
                fontSize.Val = ConvertFontSizeToHalfPoints(styleConfig.FontSize);

                var fontSizeCs = runProperties.Elements<FontSizeComplexScript>().FirstOrDefault();
                if (fontSizeCs == null)
                {
                    fontSizeCs = new FontSizeComplexScript();
                    runProperties.AppendChild(fontSizeCs);
                }
                fontSizeCs.Val = ConvertFontSizeToHalfPoints(styleConfig.FontSize);

                // 设置加粗
                var bold = runProperties.Elements<Bold>().FirstOrDefault();
                if (bold == null)
                {
                    bold = new Bold();
                    runProperties.AppendChild(bold);
                }
                bold.Val = styleConfig.Bold;

                logMessage($"已设置自动编号字体：字体={styleConfig.FontName}, 字号={styleConfig.FontSize}, 加粗={styleConfig.Bold}");
            }
            catch (Exception ex)
            {
                logMessage($"设置自动编号字体格式时出错: {ex.Message}");
                logMessage($"错误堆栈: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 处理文档中的所有表格格式设置
        /// </summary>
        private void ProcessTablesInDocument(WordprocessingDocument wordDoc, Models.StyleConfig config, Action<string> logMessage)
        {
            try
            {
                var body = wordDoc.MainDocumentPart?.Document?.Body;
                if (body == null)
                    return;

                var tables = body.Elements<Table>().ToList();
                if (tables.Count == 0)
                {
                    logMessage("文档中没有发现表格");
                    return;
                }

                logMessage($"发现 {tables.Count} 个表格，开始处理表格格式");

                int processedTables = 0;
                foreach (var table in tables)
                {
                    ProcessSingleTable(wordDoc, table, config, logMessage);
                    processedTables++;

                    if (processedTables % 5 == 0)
                    {
                        logMessage($"已处理 {processedTables}/{tables.Count} 个表格");
                    }
                }

                logMessage($"表格处理完成，共处理 {processedTables} 个表格");
            }
            catch (Exception ex)
            {
                logMessage($"处理表格时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理单个表格的格式设置
        /// </summary>
        private void ProcessSingleTable(WordprocessingDocument wordDoc, Table table, Models.StyleConfig config, Action<string> logMessage)
        {
            try
            {
                // 获取正文段落的字体名称（表格要保持一致）
                string bodyFontName = config.Normal.FontName;

                // 处理表格中的所有行和单元格
                foreach (var row in table.Elements<TableRow>())
                {
                    foreach (var cell in row.Elements<TableCell>())
                    {
                        // 处理单元格中的所有段落
                        foreach (var paragraph in cell.Elements<Paragraph>())
                        {
                            // 检查是否为表格内部的标题段落
                            if (IsTableCaptionParagraph(wordDoc, paragraph))
                            {
                                // 对表格内部的标题应用表格标题格式
                                ApplyStyleToParagraph(wordDoc, paragraph, config.TableCaption, "TableCaption", logMessage, config);
                                logMessage($"处理表格内部标题：{paragraph.InnerText.Trim()}");
                            }
                            else
                            {
                                // 对普通表格段落应用标准表格格式
                                ProcessTableParagraph(paragraph, bodyFontName, logMessage);
                            }
                        }
                    }
                }

                logMessage($"表格处理完成：字体名称={bodyFontName}，英文=Times New Roman，字号保持不变");
            }
            catch (Exception ex)
            {
                logMessage($"处理单个表格时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 检查段落是否为表格标题(包括表格内部)
        /// </summary>
        private bool IsTableCaptionParagraph(WordprocessingDocument wordDoc, Paragraph paragraph)
        {
            // 获取段落文本
            var text = paragraph.InnerText.Trim();
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            // 获取包含自动编号的完整文本
            var paragraphTextWithNumbering = GetParagraphTextWithNumbering(wordDoc, paragraph);

            // 识别表格标题:以"表"开头,可能包含编号的段落
            return IsTableCaptionByPattern(paragraphTextWithNumbering);
        }

        /// <summary>
        /// 处理表格段落的格式设置
        /// </summary>
        private void ProcessTableParagraph(Paragraph paragraph, string bodyFontName, Action<string> logMessage)
        {
            try
            {
                // 跳过空段落
                if (string.IsNullOrWhiteSpace(paragraph.InnerText))
                {
                    return;
                }

                // 设置段落中所有运行的字体属性
                foreach (var run in paragraph.Elements<Run>())
                {
                    var runProperties = run.RunProperties;
                    if (runProperties == null)
                    {
                        runProperties = new RunProperties();
                        run.AppendChild(runProperties);
                    }

                    // 设置字体
                    var runFonts = runProperties.Elements<RunFonts>().FirstOrDefault();
                    if (runFonts == null)
                    {
                        runFonts = new RunFonts();
                        runProperties.AppendChild(runFonts);
                    }

                    // 表格字体设置规则：
                    // 1. 字体名称与正文段落保持一致
                    // 2. 中文字符使用正文字体，非中文字符使用Times New Roman
                    runFonts.EastAsia = bodyFontName;    // 中文字符使用正文字体
                    runFonts.Ascii = "Times New Roman";   // 英文字符使用Times New Roman
                    runFonts.HighAnsi = "Times New Roman";   // 英文字符使用Times New Roman
                    runFonts.ComplexScript = "Times New Roman";      // 复杂脚本使用Times New Roman

                    // 注意：不修改字体大小，保持表格原有的字号

                    // 清除表格文本中的下划线
                    var underline = runProperties.Elements<Underline>().FirstOrDefault();
                    if (underline != null)
                    {
                        runProperties.RemoveChild(underline);
                    }
                }

                // 记录处理信息
                string displayText = paragraph.InnerText.Trim();
                if (displayText.Length > 30)
                {
                    displayText = displayText.Substring(0, 30) + "...";
                }
                logMessage($"表格段落处理: \"{displayText}\" - 中文字体={bodyFontName}, 英文字体=Times New Roman");
            }
            catch (Exception ex)
            {
                logMessage($"处理表格段落时出错: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 段落样式信息类
    /// </summary>
    public class ParagraphStyleInfo
    {
        public string StyleId { get; set; } = "";
        public string? StyleName { get; set; }
        public string ParagraphText { get; set; } = "";
    }
}
