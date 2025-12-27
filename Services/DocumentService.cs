using NPOI.XWPF.UserModel;
using WordTools2.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace WordTools2.Services
{
    /// <summary>
    /// 文档服务类 - 负责文档的读取、处理和保存
    /// </summary>
    public class DocumentService
    {
        private string? _originalFilePath; // 原始文件路径（永远不被修改）
        private string? _workingFilePath;   // 工作文件路径（可以修改）
        private XWPFDocument? _document;

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
                
                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    _document = new XWPFDocument(fileStream);
                }
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
                    _document.Close();
                }
                catch { }
                _document = null;
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
                        _document.Close();
                        _document = null;
                    }

                    using (var fileStream = new FileStream(_workingFilePath, FileMode.Open, FileAccess.ReadWrite))
                    {
                        var doc = new XWPFDocument(fileStream);
                        
                        logMessage("开始直接修改段落格式（不修改样式定义）");
                        
                        // 诊断：检查原始文档的样式定义
                        CheckOriginalDocumentStyles(logMessage);

                        var paragraphs = doc.Paragraphs.ToList();
                        int total = paragraphs.Count;
                        int processed = 0;

                        foreach (var paragraph in paragraphs)
                        {
                            // 通过正则表达式识别段落样式类型
                            var inferredStyle = InferStyleFromParagraph(paragraph, config);
                            
                            // 对表格标题段落应用专用样式
                            if (inferredStyle == "TableCaption")
                            {
                                var style = config.TableCaption;
                                ApplyStyleToParagraph(paragraph, style, "TableCaption", logMessage, config);
                                logMessage($"处理表格标题：{paragraph.ParagraphText.Trim()}");
                                processed++;
                                if (processed % 10 == 0)
                                {
                                    var percent = (int)((double)processed / total * 100);
                                    updateProgress($"处理中... {percent}%");
                                }
                                continue;
                            }
                            
                            // 对图形标题段落应用专用样式
                            if (inferredStyle == "ImageCaption")
                            {
                                var style = config.ImageCaption;
                                ApplyStyleToParagraph(paragraph, style, "ImageCaption", logMessage, config);
                                logMessage($"处理图形标题：{paragraph.ParagraphText.Trim()}");
                                processed++;
                                if (processed % 10 == 0)
                                {
                                    var percent = (int)((double)processed / total * 100);
                                    updateProgress($"处理中... {percent}%");
                                }
                                continue;
                            }
                            
                            // 对图片段落应用专用样式
                            if (inferredStyle == "Image")
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
                            
                            // 应用识别出的样式格式（仅对非图片段落）
                            if (!string.IsNullOrEmpty(inferredStyle))
                            {
                                var style = GetParagraphStyle(inferredStyle, config);
                                if (style != null)
                                {
                                    ApplyStyleToParagraph(paragraph, style, inferredStyle, logMessage, config);
                                    logMessage($"处理段落格式: {inferredStyle}");
                                }
                            }

                            processed++;
                            if (processed % 10 == 0)
                            {
                                var percent = (int)((double)processed / total * 100);
                                updateProgress($"处理中... {percent}%");
                            }
                        }

                        // 处理表格格式设置
                        ProcessTablesInDocument(doc, config, logMessage);
                        
                        // 确保所有更改都提交到文档
                        using (var output = new FileStream(_workingFilePath, FileMode.Create, FileAccess.Write))
                        {
                            doc.Write(output);
                        }
                        doc.Close();
                        
                        logMessage($"已处理 {total} 个段落");
                        logMessage("文档已保存，所有格式更改已提交");
                    }

                    // 重新打开处理后的工作文件（只读模式）
                    using (var fileStream = new FileStream(_workingFilePath, FileMode.Open, FileAccess.Read))
                    {
                        _document = new XWPFDocument(fileStream);
                    }

                    updateProgress("处理完成 100%");
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
                if (_document != null)
                {
                    _document.Close();
                    _document = null;
                }
                
                // 复制源文件到新位置
                File.Copy(sourcePath, newPath, true);
                
                // 重新打开源文件（只读模式）
                if (_workingFilePath != null && File.Exists(_workingFilePath))
                {
                    using (var fileStream = new FileStream(_workingFilePath, FileMode.Open, FileAccess.Read))
                    {
                        _document = new XWPFDocument(fileStream);
                    }
                }
                else
                {
                    using (var fileStream = new FileStream(_originalFilePath, FileMode.Open, FileAccess.Read))
                    {
                        _document = new XWPFDocument(fileStream);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"保存文档失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取文档统计信息（基于文本格式识别）
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

            if (_document == null)
                return stats;

            foreach (var paragraph in _document.Paragraphs)
            {
                // 首先检查是否为图片段落
                if (IsImageParagraph(paragraph))
                {
                    stats["Image"]++;
                    continue;
                }
                
                // 获取段落文本
                var text = paragraph.ParagraphText.Trim();
                
                // 基于文本格式判断
                if (string.IsNullOrEmpty(text))
                {
                    stats["Normal"]++;
                }
                // 识别表格标题：以"表"开头，可能包含编号的段落
                else if (Regex.IsMatch(text, @"^表\d*(\.\d+)*(-\d+)*\s*[\u4e00-\u9fa5].*$"))
                {
                    stats["TableCaption"]++; // 表格标题
                }
                // 识别图形标题：以"图"开头，可能包含编号的段落
                else if (Regex.IsMatch(text, @"^图\d*(\.\d+)*(-\d+)*\s+.*$"))
                {
                    stats["ImageCaption"]++; // 图形标题
                }
                else if (Regex.IsMatch(text, @"^[1-9]\d?\s+[\u4e00-\u9fa5]+$"))
                {
                    stats["Heading1"]++; // 一级标题格式：单个数字+空格+中文内容
                }
                else if (Regex.IsMatch(text, @"^[1-9]\d?\.[1-9]\d?\s*[\u4e00-\u9fa5]+$"))
                {
                    stats["Heading2"]++; // 二级标题格式：数字.数字+空格/中文内容
                }
                else if (Regex.IsMatch(text, @"^[1-9]\d?\.[1-9]\d?\.[1-9]\d?\s*[\u4e00-\u9fa5]+$"))
                {
                    stats["Heading3"]++; // 三级标题格式：数字.数字.数字+空格/中文内容
                }
                else if (Regex.IsMatch(text, @"^[1-9]\d?\.[1-9]\d?\.[1-9]\d?\.[1-9]\d?\s*[\u4e00-\u9fa5]+$"))
                {
                    stats["Heading4"]++; // 四级标题格式：数字.数字.数字.数字+空格/中文内容
                }
                else
                {
                    stats["Other"]++; // 其他格式
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


        /// <summary>
        /// 直接修改段落格式（不修改样式定义）
        /// 只修改当前段落的字体、字号、间距、大纲级别等属性
        /// </summary>
        private void ApplyStyleToParagraph(XWPFParagraph paragraph, ParagraphStyle styleConfig, string styleName, Action<string> logMessage, Models.StyleConfig? config = null)
        {
            // 首先获取段落文本用于调试
            var text = paragraph.ParagraphText.Trim();

            // 跳过图片段落，不进行任何格式设置
            if (styleName == "Image")
            {
                logMessage("跳过图片段落，不应用任何格式设置");
                return;
            }

            // 设置大纲级别和格式（表格标题需要特殊处理）
            SetOutlineLevelForParagraph(paragraph, styleName, logMessage, config);

            // 设置段落间距
            var paraCTP2 = paragraph.GetCTP();
            if (paraCTP2.pPr == null)
            {
                paraCTP2.AddNewPPr();
            }
            if (paraCTP2.pPr.spacing == null)
            {
                paraCTP2.pPr.spacing = new NPOI.OpenXmlFormats.Wordprocessing.CT_Spacing();
            }

            paraCTP2.pPr.spacing.before = (ulong)(styleConfig.SpaceBefore * 20); // NPOI使用缇单位
            paraCTP2.pPr.spacing.after = (ulong)(styleConfig.SpaceAfter * 20); // NPOI使用缇单位
            
            // 只对正文段落设置行距，标题段落保持原有行距
            if (styleName == "Normal" && styleConfig.LineSpacing > 0)
            {
                paraCTP2.pPr.spacing.line = ((int)(styleConfig.LineSpacing * 20)).ToString(); // NPOI使用缇单位
                paraCTP2.pPr.spacing.lineRule = NPOI.OpenXmlFormats.Wordprocessing.ST_LineSpacingRule.exact;
            }

            // 设置段落中所有运行的字体属性
            foreach (var run in paragraph.Runs)
            {
                var runCTR = run.GetCTR();
                if (runCTR.rPr == null)
                {
                    runCTR.AddNewRPr();
                }

                // 设置字体
                if (runCTR.rPr.rFonts == null)
                {
                    runCTR.rPr.rFonts = new NPOI.OpenXmlFormats.Wordprocessing.CT_Fonts();
                }
                
                // 所有段落（包括标题和正文）都采用相同的中英文字体处理方式
                // 中文字符保持用户设置的字体，英文和数字使用Times New Roman
                runCTR.rPr.rFonts.eastAsia = styleConfig.FontName;      // 中文字符使用用户设置的字体
                runCTR.rPr.rFonts.ascii = "Times New Roman";              // 英文字符使用Times New Roman
                runCTR.rPr.rFonts.hAnsi = "Times New Roman";              // 英文字符使用Times New Roman
                runCTR.rPr.rFonts.cs = "Times New Roman";                 // 复杂脚本使用Times New Roman
                
                // 记录字体设置信息
                string styleType = GetStyleTypeFromConfig(styleConfig);
                logMessage($"{styleType}字体设置：中文={styleConfig.FontName}, 英文/数字=Times New Roman");

                // 设置字体大小
                if (runCTR.rPr.sz == null)
                {
                    runCTR.rPr.sz = new NPOI.OpenXmlFormats.Wordprocessing.CT_HpsMeasure();
                }
                runCTR.rPr.sz.val = (ulong)(styleConfig.FontSize * 2); // NPOI使用半点单位

                if (runCTR.rPr.szCs == null)
                {
                    runCTR.rPr.szCs = new NPOI.OpenXmlFormats.Wordprocessing.CT_HpsMeasure();
                }
                runCTR.rPr.szCs.val = (ulong)(styleConfig.FontSize * 2); // NPOI使用半点单位

                // 设置加粗
                if (runCTR.rPr.b == null)
                {
                    runCTR.rPr.b = new NPOI.OpenXmlFormats.Wordprocessing.CT_OnOff();
                }
                runCTR.rPr.b.val = styleConfig.Bold;
            }

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
        private void ApplyImageParagraphStyle(XWPFParagraph paragraph, Action<string> logMessage)
        {
            var text = paragraph.ParagraphText.Trim();
            
            // 确保段落属性存在
            var paraCTP = paragraph.GetCTP();
            if (paraCTP.pPr == null)
            {
                paraCTP.AddNewPPr();
            }
            if (paraCTP.pPr.spacing == null)
            {
                paraCTP.pPr.spacing = new NPOI.OpenXmlFormats.Wordprocessing.CT_Spacing();
            }

            // 图片段落专用设置：段前间距和段后间距都设置为0磅
            paraCTP.pPr.spacing.before = 0; // 段前间距0磅
            paraCTP.pPr.spacing.after = 0;  // 段后间距0磅
            
            // 强制设置单倍行距 - 清除任何可能影响行距的设置
            paraCTP.pPr.spacing.line = "240"; // 单倍行距（240缇，即12磅）
            paraCTP.pPr.spacing.lineRule = NPOI.OpenXmlFormats.Wordprocessing.ST_LineSpacingRule.auto;

            // 确保行距设置为单倍行距
            if (paraCTP.pPr.spacing.line != "240")
            {
                paraCTP.pPr.spacing.line = "240"; // 确保行距为单倍行距
            }
            if (paraCTP.pPr.spacing.lineRule != NPOI.OpenXmlFormats.Wordprocessing.ST_LineSpacingRule.auto)
            {
                paraCTP.pPr.spacing.lineRule = NPOI.OpenXmlFormats.Wordprocessing.ST_LineSpacingRule.auto; // 确保行距规则为自动
            }

            // 图片段落不设置大纲级别
            logMessage($"应用图片段落样式: 段前距=0磅, 段后距=0磅, 行距=单倍行距");
            
            // 对于包含图片的段落，通常不需要设置字体属性，因为图片本身没有字体
            // 但为了保持一致性，我们仍然记录日志
            string displayText = text.Length > 30 ? text.Substring(0, 30) + "..." : text;
            logMessage($"图片段落处理完成: 包含图片的段落，文本内容: \"{displayText}\"");
        }

        /// <summary>
        /// 检查原始文档的样式定义，诊断大纲级别问题
        /// </summary>
        private void CheckOriginalDocumentStyles(Action<string> logMessage)
        {
            try
            {
                if (_document != null)
                {
                    logMessage("=== 原始文档样式诊断 ===");
                    logMessage("NPOI不直接提供样式定义访问接口，跳过样式诊断");
                    logMessage("=== 样式诊断完成 ===");
                }
            }
            catch (Exception ex)
            {
                logMessage($"样式诊断出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 检测段落是否包含图片
        /// </summary>
        private bool IsImageParagraph(XWPFParagraph paragraph)
        {
            try
            {
                // 获取段落的CT_P对象
                var paragraphCTP = paragraph.GetCTP();
                
                // 检查段落是否包含drawing元素
                if (paragraphCTP.Items != null)
                {
                    foreach (var item in paragraphCTP.Items)
                    {
                        if (item.GetType().Name.Contains("Drawing") || 
                            item.GetType().Name.Contains("Picture") ||
                            item.GetType().Name.Contains("Graphic"))
                        {
                            return true;
                        }
                    }
                }
                
                // 检查段落中的运行元素是否包含图片
                foreach (var run in paragraph.Runs)
                {
                    // 获取运行元素的CT_R对象
                    var runCTR = run.GetCTR();
                    
                    // 检查Run是否包含drawing元素
                    if (runCTR.Items != null)
                    {
                        foreach (var item in runCTR.Items)
                        {
                            if (item.GetType().Name.Contains("Drawing") || 
                                item.GetType().Name.Contains("Picture") ||
                                item.GetType().Name.Contains("Graphic") ||
                                item.GetType().Name.Contains("Object"))
                            {
                                return true;
                            }
                        }
                    }
                }
                
                // 特别处理：如果段落文本为空但包含运行元素，可能是图片
                if (string.IsNullOrEmpty(paragraph.ParagraphText?.Trim()) && paragraph.Runs.Count > 0)
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
        /// 根据段落文本格式推断样式类型（通过正则表达式识别标题编号）
        /// </summary>
        private string InferStyleFromParagraph(XWPFParagraph paragraph, Models.StyleConfig config)
        {
            // 首先检查是否为图片段落
            if (IsImageParagraph(paragraph))
            {
                return "Image"; // 图片段落
            }
            
            // 获取段落文本
            var text = paragraph.ParagraphText.Trim();
            if (string.IsNullOrEmpty(text))
            {
                return "Normal"; // 空段落视为正文
            }

            // 识别表格标题：以"表"开头，可能包含编号的段落
            // 规则：以"表"开头，后跟数字编号（可选），可能包含连字符编号，然后是空格和中文内容
            if (Regex.IsMatch(text, @"^表\d*(\.\d+)*(-\d+)*\s*[\u4e00-\u9fa5].*$"))
            {
                return "TableCaption"; // 表格标题
            }
            
            // 识别图形标题：以"图"开头，可能包含编号的段落
            // 规则：以"图"开头，后跟数字编号（可选），然后是空格和任意内容
            if (Regex.IsMatch(text, @"^图\d*(\.\d+)*(-\d+)*\s+.*$"))
            {
                return "ImageCaption"; // 图形标题
            }

            // 通过正则表达式识别标题格式（更严格的规则，避免误识别正文中的一般数字）
            
            // 一级标题：格式如 "1 标题" 或 "2 水文" 等
            // 规则：单个数字（1-99）+ 空格 + 中文内容（至少1个字符）
            if (Regex.IsMatch(text, @"^[1-9]\d?\s+[\u4e00-\u9fa5]+$"))
            {
                return "Heading1"; // 一级标题
            }
            
            // 二级标题：格式如 "1.1 标题" 或 "2.1 设计洪水" 等
            // 规则：数字.数字（如1.1, 2.3, 10.5等，避免匹配年份、小数等）+ 空格/中文内容
            if (Regex.IsMatch(text, @"^[1-9]\d?\.[1-9]\d?\s*[\u4e00-\u9fa5]+$"))
            {
                return "Heading2"; // 二级标题
            }
            
            // 三级标题：格式如 "1.1.1 标题" 等
            // 规则：数字.数字.数字 + 空格/中文内容
            if (Regex.IsMatch(text, @"^[1-9]\d?\.[1-9]\d?\.[1-9]\d?\s*[\u4e00-\u9fa5]+$"))
            {
                return "Heading3"; // 三级标题
            }
            
            // 四级标题：格式如 "1.1.1.1 标题" 等
            // 规则：数字.数字.数字.数字 + 空格/中文内容
            if (Regex.IsMatch(text, @"^[1-9]\d?\.[1-9]\d?\.[1-9]\d?\.[1-9]\d?\s*[\u4e00-\u9fa5]+$"))
            {
                return "Heading4"; // 四级标题
            }

            // 其他格式都视为正文
            return "Normal";
        }
        
        /// <summary>
        /// 获取段落的字体名称
        /// </summary>
        private string GetParagraphFontName(XWPFParagraph paragraph)
        {
            // 默认字体名称
            string defaultFontName = "宋体";
            
            // 1. 首先从Run元素中获取字体名称
            // 使用字典统计各字体出现次数，返回出现次数最多的字体
            var fontCounts = new Dictionary<string, int>();
            
            foreach (var run in paragraph.Runs)
            {
                var runCTR = run.GetCTR();
                if (runCTR.rPr != null && runCTR.rPr.rFonts != null)
                {
                    // 优先使用EastAsia字体（中文），然后是HighAnsi，最后是Ascii
                    string fontName = !string.IsNullOrEmpty(runCTR.rPr.rFonts.eastAsia) ? runCTR.rPr.rFonts.eastAsia :
                                     !string.IsNullOrEmpty(runCTR.rPr.rFonts.hAnsi) ? runCTR.rPr.rFonts.hAnsi :
                                     !string.IsNullOrEmpty(runCTR.rPr.rFonts.ascii) ? runCTR.rPr.rFonts.ascii :
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
            
            // 如果Run中有字体设置，返回出现次数最多的字体
            if (fontCounts.Any())
            {
                return fontCounts.OrderByDescending(kvp => kvp.Value).First().Key;
            }
            
            // 3. 默认返回宋体
            return defaultFontName;
        }

        /// <summary>
        /// 获取段落的字体大小（磅）
        /// </summary>
        private double GetParagraphFontSize(XWPFParagraph paragraph)
        {
            // 默认字体大小
            double defaultFontSize = 10.5;
            
            // 1. 从Run元素中获取字体大小
            // 使用字典统计各字体大小出现次数，返回出现次数最多的字体大小
            var fontSizeCounts = new Dictionary<double, int>();
            
            foreach (var run in paragraph.Runs)
            {
                var runCTR = run.GetCTR();
                if (runCTR.rPr != null && runCTR.rPr.sz != null && runCTR.rPr.sz.val != null)
                {
                    if (double.TryParse(runCTR.rPr.sz.val.ToString(), out double fontSizeValue))
                    {
                        double fontSize = fontSizeValue / 2.0; // 转换为磅
                        
                        if (fontSizeCounts.ContainsKey(fontSize))
                        {
                            fontSizeCounts[fontSize]++;
                        }
                        else
                        {
                            fontSizeCounts[fontSize] = 1;
                        }
                    }
                }
            }
            
            // 如果Run中有字体大小设置，返回出现次数最多的字体大小
            if (fontSizeCounts.Any())
            {
                return fontSizeCounts.OrderByDescending(kvp => kvp.Value).First().Key;
            }
            
            // 3. 默认返回10.5磅
            return defaultFontSize;
        }

        /// <summary>
        /// 获取段落的段前间距（磅）
        /// </summary>
        private double GetParagraphSpaceBefore(XWPFParagraph paragraph)
        {
            // 默认段前间距
            double defaultSpaceBefore = 0;
            
            // 从段落属性中获取间距
            var paraCTP = paragraph.GetCTP();
            if (paraCTP.pPr != null && paraCTP.pPr.spacing != null && paraCTP.pPr.spacing.@before.HasValue)
            {
                if (int.TryParse(paraCTP.pPr.spacing.@before.ToString(), out int beforeValue))
                {
                    return beforeValue / 20.0; // 转换为磅（1磅=20缇）
                }
            }
            
            // 默认返回0磅
            return defaultSpaceBefore;
        }

        /// <summary>
        /// 获取段落的段后间距（磅）
        /// </summary>
        private double GetParagraphSpaceAfter(XWPFParagraph paragraph)
        {
            // 默认段后间距
            double defaultSpaceAfter = 0;
            
            // 从段落属性中获取间距
            var paraCTP = paragraph.GetCTP();
            if (paraCTP.pPr != null && paraCTP.pPr.spacing != null && paraCTP.pPr.spacing.@after.HasValue)
            {
                if (int.TryParse(paraCTP.pPr.spacing.@after.ToString(), out int afterValue))
                {
                    return afterValue / 20.0; // 转换为磅（1磅=20缇）
                }
            }
            
            // 默认返回0磅
            return defaultSpaceAfter;
        }
        

        
        /// <summary>
        /// 获取段落的详细信息
        /// </summary>
        public Dictionary<string, object> GetParagraphDetails(XWPFParagraph paragraph)
        {
            var details = new Dictionary<string, object>();
            
            // 获取段落文本
            string text = paragraph.ParagraphText.Trim();
            details["Text"] = text;
            
            // 通过正则表达式推断样式类型
            string inferredStyle = InferStyleFromParagraph(paragraph, new Models.StyleConfig());
            details["StyleType"] = inferredStyle;
            
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
                FontName = "宋体",        // 图片段落不需要特殊字体
                FontSize = 10.5,          // 默认字体大小
                SpaceBefore = 0,          // 段前间距0磅
                SpaceAfter = 0,           // 段后间距0磅
                LineSpacing = 1.0         // 单倍行距
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
        private double ConvertHalfPointsToFontSize(ulong? halfPoints)
        {
            if (halfPoints.HasValue)
            {
                return halfPoints.Value / 2.0;
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
                "Normal" => 9,        // 正文文本不参与大纲
                "Image" => 9,         // 图片段落不参与大纲
                "TableCaption" => config?.TableCaption.OutlineLevel ?? 9,   // 表格标题使用用户设置
                "ImageCaption" => config?.ImageCaption.OutlineLevel ?? 9,   // 图形标题使用用户设置
                "Heading1" => 0,      // 一级标题 - 对应Word大纲级别0（最高级）
                "Heading2" => 1,      // 二级标题 - 对应Word大纲级别1
                "Heading3" => 2,      // 三级标题 - 对应Word大纲级别2
                "Heading4" => 3,      // 四级标题 - 对应Word大纲级别3
                _ => 9                // 其他级别
            };
        }

        /// <summary>
        /// 为段落设置大纲级别
        /// </summary>
        private void SetOutlineLevelForParagraph(XWPFParagraph paragraph, string styleName, Action<string> logMessage, Models.StyleConfig? config = null)
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
            
            // 设置大纲级别 - 使用XWPFParagraph API
            var paraCTP = paragraph.GetCTP();
            if (paraCTP.pPr == null)
            {
                paraCTP.AddNewPPr();
            }
            if (paraCTP.pPr.outlineLvl == null)
            {
                paraCTP.pPr.outlineLvl = new NPOI.OpenXmlFormats.Wordprocessing.CT_DecimalNumber();
            }
            paraCTP.pPr.outlineLvl.val = targetLevel.ToString();
            
            logMessage($"设置大纲级别: 样式名={styleName}, 目标级别={targetLevel}");
        }



        /// <summary>
        /// 处理文档中的所有表格格式设置
        /// </summary>
        private void ProcessTablesInDocument(XWPFDocument doc, Models.StyleConfig config, Action<string> logMessage)
        {
            try
            {
                var tables = doc.Tables.ToList();
                if (tables.Count == 0)
                {
                    logMessage("文档中没有发现表格");
                    return;
                }

                logMessage($"发现 {tables.Count} 个表格，开始处理表格格式");
                
                int processedTables = 0;
                foreach (var table in tables)
                {
                    ProcessSingleTable(table, config, logMessage);
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
        private void ProcessSingleTable(XWPFTable table, Models.StyleConfig config, Action<string> logMessage)
        {
            try
            {
                // 获取正文段落的字体名称（表格要保持一致）
                string bodyFontName = config.Normal.FontName;
                
                // 处理表格中的所有行和单元格
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.GetTableCells())
                    {
                        // 处理单元格中的所有段落
                        foreach (var paragraph in cell.Paragraphs)
                        {
                            // 检查是否为表格内部的标题段落
                            if (IsTableCaptionParagraph(paragraph))
                            {
                                // 对表格内部的标题应用表格标题格式
                                ApplyStyleToParagraph(paragraph, config.TableCaption, "TableCaption", logMessage, config);
                                logMessage($"处理表格内部标题：{paragraph.ParagraphText.Trim()}");
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
        /// 检查段落是否为表格标题（包括表格内部）
        /// </summary>
        private bool IsTableCaptionParagraph(XWPFParagraph paragraph)
        {
            // 获取段落文本
            var text = paragraph.ParagraphText.Trim();
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            // 识别表格标题：以"表"开头，可能包含编号的段落
            // 规则：以"表"开头，后跟数字编号（可选），可能包含连字符编号，然后是空格和中文内容
            return Regex.IsMatch(text, @"^表\d*(\.\d+)*(-\d+)*\s*[\u4e00-\u9fa5].*$");
        }

        /// <summary>
        /// 处理表格段落的格式设置
        /// </summary>
        private void ProcessTableParagraph(XWPFParagraph paragraph, string bodyFontName, Action<string> logMessage)
        {
            try
            {
                // 跳过空段落
                if (string.IsNullOrWhiteSpace(paragraph.ParagraphText))
                {
                    return;
                }

                // 设置段落中所有运行的字体属性
                foreach (var run in paragraph.Runs)
                {
                    var runCTR = run.GetCTR();
                    if (runCTR.rPr == null)
                    {
                        runCTR.AddNewRPr();
                    }

                    // 设置字体
                    if (runCTR.rPr.rFonts == null)
                    {
                        runCTR.rPr.rFonts = new NPOI.OpenXmlFormats.Wordprocessing.CT_Fonts();
                    }
                    
                    // 表格字体设置规则：
                    // 1. 字体名称与正文段落保持一致
                    // 2. 中文字符使用正文字体，非中文字符使用Times New Roman
                    runCTR.rPr.rFonts.eastAsia = bodyFontName;    // 中文字符使用正文字体
                    runCTR.rPr.rFonts.ascii = "Times New Roman";   // 英文字符使用Times New Roman
                    runCTR.rPr.rFonts.hAnsi = "Times New Roman";   // 英文字符使用Times New Roman
                    runCTR.rPr.rFonts.cs = "Times New Roman";      // 复杂脚本使用Times New Roman
                    
                    // 注意：不修改字体大小，保持表格原有的字号
                }
                
                // 记录处理信息
                string displayText = paragraph.ParagraphText.Trim();
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
}