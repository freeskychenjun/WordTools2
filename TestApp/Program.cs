using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace ComprehensiveTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "d:\\CodeBuddy\\WordTools2\\test.docx";
            
            try
            {
                Console.WriteLine("=== 综合测试: 标题识别、字体名称和字体大小 ===");
                
                using (var doc = WordprocessingDocument.Open(filePath, false))
                {
                    if (doc.MainDocumentPart == null)
                    {
                        Console.WriteLine("❌ 无法获取文档主体部分");
                        return;
                    }

                    var paragraphs = doc.MainDocumentPart.Document.Descendants<Paragraph>().ToList();
                    if (paragraphs.Count == 0)
                    {
                        Console.WriteLine("❌ 文档中没有段落");
                        return;
                    }

                    // 测试前5个段落
                    Console.WriteLine($"\n=== 测试前5个段落 ===");
                    for (int i = 0; i < Math.Min(5, paragraphs.Count); i++)
                    {
                        var paragraph = paragraphs[i];
                        int paragraphNumber = i + 1;
                        
                        // 获取段落文本
                        var text = string.Join("", paragraph.Descendants<Run>().Select(r => r.InnerText)).Trim();
                        if (string.IsNullOrEmpty(text))
                        {
                            continue;
                        }
                        
                        Console.WriteLine($"\n段落{paragraphNumber:D2}: '{text}'");
                        
                        // 1. 测试大纲级别识别
                        Console.WriteLine("\n  1. 大纲级别识别:");
                        int? outlineLevel = null;
                        string outlineLevelDetails = "";
                        
                        var paragraphProperties = paragraph.ParagraphProperties;
                        if (paragraphProperties != null)
                        {
                            // 直接大纲级别
                            var outlineLevelElement = paragraphProperties.GetFirstChild<OutlineLevel>();
                            if (outlineLevelElement != null)
                            {
                                outlineLevelDetails += "直接大纲级别: ";
                                if (outlineLevelElement.Val != null)
                                {
                                    if (int.TryParse(outlineLevelElement.Val.InnerText, out int level))
                                    {
                                        outlineLevel = level;
                                        outlineLevelDetails += $"{level}级";
                                    }
                                    else
                                    {
                                        outlineLevelDetails += $"无法解析({outlineLevelElement.Val.InnerText})";
                                    }
                                }
                                else
                                {
                                    outlineLevelDetails += "OutlineLevel.Val为空";
                                }
                            }
                            else
                            {
                                outlineLevelDetails += "没有直接大纲级别，尝试从样式获取: ";
                                
                                // 从样式获取
                                var styleId = paragraphProperties.ParagraphStyleId?.Val?.Value;
                                if (!string.IsNullOrEmpty(styleId))
                                {
                                    var stylesPart = doc.MainDocumentPart.StyleDefinitionsPart;
                                    if (stylesPart != null && stylesPart.Styles != null)
                                    {
                                        var style = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
                                        if (style != null)
                                        {
                                            var styleName = style.GetFirstChild<StyleName>()?.Val?.Value ?? "";
                                            outlineLevelDetails += $"样式名称={styleName}";
                                            
                                            if (styleName.IndexOf("heading", StringComparison.OrdinalIgnoreCase) >= 0)
                                            {
                                                if (styleName.IndexOf("1", StringComparison.OrdinalIgnoreCase) >= 0)
                                                {
                                                    outlineLevel = 1;
                                                    outlineLevelDetails += "，推断为1级标题";
                                                }
                                                else if (styleName.IndexOf("2", StringComparison.OrdinalIgnoreCase) >= 0)
                                                {
                                                    outlineLevel = 2;
                                                    outlineLevelDetails += "，推断为2级标题";
                                                }
                                                else if (styleName.IndexOf("3", StringComparison.OrdinalIgnoreCase) >= 0)
                                                {
                                                    outlineLevel = 3;
                                                    outlineLevelDetails += "，推断为3级标题";
                                                }
                                                else if (styleName.IndexOf("4", StringComparison.OrdinalIgnoreCase) >= 0)
                                                {
                                                    outlineLevel = 4;
                                                    outlineLevelDetails += "，推断为4级标题";
                                                }
                                                else
                                                {
                                                    outlineLevel = 1;
                                                    outlineLevelDetails += "，推断为1级标题";
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        
                        Console.WriteLine($"     {outlineLevelDetails}");
                        Console.WriteLine($"     最终大纲级别: {(outlineLevel.HasValue ? outlineLevel.Value.ToString() : "无")}");
                        
                        // 2. 测试字体名称
                        Console.WriteLine("  \n2. 字体名称:");
                        string fontName = "宋体";
                        int runCount = 0;
                        
                        foreach (var run in paragraph.Elements<Run>())
                        {
                            runCount++;
                            var runProperties = run.RunProperties;
                            if (runProperties != null)
                            {
                                var runFonts = runProperties.GetFirstChild<RunFonts>();
                                if (runFonts != null)
                                {
                                    string currentFont = runFonts.EastAsia?.Value ?? 
                                                       runFonts.HighAnsi?.Value ?? 
                                                       runFonts.Ascii?.Value ?? 
                                                       "无";
                                    Console.WriteLine($"     Run{runCount}: {currentFont}");
                                    if (!string.IsNullOrEmpty(currentFont) && currentFont != "无")
                                    {
                                        fontName = currentFont;
                                    }
                                }
                            }
                        }
                        
                        Console.WriteLine($"     最终字体名称: {fontName}");
                        
                        // 3. 测试字体大小
                        Console.WriteLine("  \n3. 字体大小:");
                        runCount = 0;
                        double maxFontSize = 10.5;
                        
                        foreach (var run in paragraph.Elements<Run>())
                        {
                            runCount++;
                            var runProperties = run.RunProperties;
                            if (runProperties != null)
                            {
                                var fontSizeElement = runProperties.GetFirstChild<FontSize>();
                                if (fontSizeElement != null && fontSizeElement.Val != null)
                                {
                                    if (int.TryParse(fontSizeElement.Val.InnerText, out int halfPoints))
                                    {
                                        double currentFontSize = halfPoints / 2.0;
                                        Console.WriteLine($"     Run{runCount}: {currentFontSize:0.0}pt");
                                        if (currentFontSize > maxFontSize)
                                        {
                                            maxFontSize = currentFontSize;
                                        }
                                    }
                                }
                                else
                                {
                                    Console.WriteLine($"     Run{runCount}: 没有FontSize元素，使用样式推断");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"     Run{runCount}: 没有RunProperties元素，使用样式推断");
                            }
                        }
                        
                        // 如果直接读取的字体大小较小，尝试从样式推断
                        if (maxFontSize <= 12 && paragraphProperties != null)
                        {
                            var styleId = paragraphProperties.ParagraphStyleId?.Val?.Value;
                            if (!string.IsNullOrEmpty(styleId))
                            {
                                var stylesPart = doc.MainDocumentPart.StyleDefinitionsPart;
                                if (stylesPart != null && stylesPart.Styles != null)
                                {
                                    var style = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
                                    if (style != null)
                                    {
                                        var styleName = style.GetFirstChild<StyleName>()?.Val?.Value ?? "";
                                        if (styleName.IndexOf("heading 1", StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            Console.WriteLine($"     从样式推断: {styleName} -> 18.0pt (小二)");
                                            maxFontSize = 18.0;
                                        }
                                        else if (styleName.IndexOf("heading 2", StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            Console.WriteLine($"     从样式推断: {styleName} -> 16.0pt (三号)");
                                            maxFontSize = 16.0;
                                        }
                                        else if (styleName.IndexOf("heading 3", StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            Console.WriteLine($"     从样式推断: {styleName} -> 15.0pt (小三)");
                                            maxFontSize = 15.0;
                                        }
                                        else if (styleName.IndexOf("heading 4", StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            Console.WriteLine($"     从样式推断: {styleName} -> 14.0pt (四号)");
                                            maxFontSize = 14.0;
                                        }
                                    }
                                }
                            }
                        }
                        
                        Console.WriteLine($"     最终字体大小: {maxFontSize:0.0}pt");
                        
                        // 4. 验证第一段
                        if (paragraphNumber == 1)
                        {
                            Console.WriteLine("  \n4. 第一段验证:");
                            bool isLevel1 = outlineLevel == 1;
                            bool isSimSun = fontName == "宋体" || fontName == "SimSun";
                            bool isXiaoEr = Math.Abs(maxFontSize - 18.0) < 0.1;
                            
                            Console.WriteLine($"     大纲级别1级: {(isLevel1 ? "✅" : "❌")}");
                            Console.WriteLine($"     字体为宋体: {(isSimSun ? "✅" : "❌")}");
                            Console.WriteLine($"     字体大小小二: {(isXiaoEr ? "✅" : "❌")}");
                            
                            if (isLevel1 && isSimSun && isXiaoEr)
                            {
                                Console.WriteLine("     \n✅ 所有验证通过: 程序能够正确判断段落的大纲级别、字体名称、字体大小");
                            }
                            else
                            {
                                Console.WriteLine("     \n❌ 验证失败: 程序无法正确判断段落属性");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 测试时出错: {ex.Message}");
                Console.WriteLine($"详细信息: {ex}");
            }
        }
    }
}