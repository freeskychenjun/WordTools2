using System;
using WordTools2.Services;
using WordTools2.Models;
using System.IO;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;

namespace ComprehensiveTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string testFilePath = @"d:\CodeBuddy\WordTools2\test.docx";
            
            // 如果测试文件不存在，创建一个
            if (!File.Exists(testFilePath))
            {
                Console.WriteLine("测试文件不存在，正在创建测试文档...");
                CreateTestDocument(testFilePath);
                Console.WriteLine("测试文档已创建");
            }
            
            try
            {
                Console.WriteLine("=== 使用NPOI库测试: 标题识别、字体名称和字体大小 ===");
                
                var documentService = new DocumentService();
                
                if (documentService.OpenDocument(testFilePath))
                {
                    Console.WriteLine("✅ 成功打开文档");
                    
                    // 获取文档统计信息
                    var stats = documentService.GetDocumentStats();
                    Console.WriteLine("\n=== 文档统计信息 ===");
                    foreach (var stat in stats)
                    {
                        if (stat.Value > 0)
                        {
                            Console.WriteLine($"{stat.Key}: {stat.Value}个段落");
                        }
                    }
                    
                    // 获取文档中的前几个段落进行测试
                    // 由于DocumentService的内部文档对象是私有的，我们通过应用样式来间接测试
                    var config = new StyleConfig
                    {
                        Heading1 = new ParagraphStyle
                        {
                            FontName = "宋体",
                            FontSize = 18.0,
                            SpaceBefore = 12.0,
                            SpaceAfter = 6.0,
                            LineSpacing = 1.0
                        },
                        Heading2 = new ParagraphStyle
                        {
                            FontName = "宋体",
                            FontSize = 16.0,
                            SpaceBefore = 10.0,
                            SpaceAfter = 4.0,
                            LineSpacing = 1.0
                        },
                        Heading3 = new ParagraphStyle
                        {
                            FontName = "宋体",
                            FontSize = 15.0,
                            SpaceBefore = 8.0,
                            SpaceAfter = 4.0,
                            LineSpacing = 1.0
                        },
                        Heading4 = new ParagraphStyle
                        {
                            FontName = "宋体",
                            FontSize = 14.0,
                            SpaceBefore = 6.0,
                            SpaceAfter = 2.0,
                            LineSpacing = 1.0
                        },
                        Normal = new ParagraphStyle
                        {
                            FontName = "宋体",
                            FontSize = 10.5,
                            SpaceBefore = 0.0,
                            SpaceAfter = 0.0,
                            LineSpacing = 1.0
                        }
                    };
                    
                    Console.WriteLine("\n=== 测试样式应用功能 ===");
                    documentService.ApplyStyles(config, 
                        (progress) => Console.WriteLine($"进度: {progress}"), 
                        (log) => Console.WriteLine($"日志: {log}"));
                    
                    Console.WriteLine("\n✅ 样式应用测试完成");
                    
                    // 保存测试文档
                    string outputFilePath = @"d:\CodeBuddy\WordTools2\output_test.docx";
                    documentService.SaveDocumentAs(outputFilePath);
                    Console.WriteLine($"✅ 文档已保存到: {outputFilePath}");
                    
                    // 关闭文档
                    documentService.CloseDocument();
                }
                else
                {
                    Console.WriteLine("❌ 无法打开文档");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 测试时出错: {ex.Message}");
                Console.WriteLine($"详细信息: {ex}");
            }
        }
        
        static void CreateTestDocument(string filePath)
        {
            using (var document = new XWPFDocument())
            {
                // 创建标题1
                var para1 = document.CreateParagraph();
                para1.Alignment = ParagraphAlignment.CENTER;
                var run1 = para1.CreateRun();
                run1.SetText("测试文档标题");
                
                // 设置字体
                var rPr1 = run1.GetCTR().rPr;
                if (rPr1 == null)
                {
                    rPr1 = run1.GetCTR().AddNewRPr();
                }
                
                if (rPr1.rFonts == null)
                {
                    rPr1.rFonts = new CT_Fonts();
                }
                rPr1.rFonts.ascii = "宋体";
                rPr1.rFonts.hAnsi = "宋体";
                rPr1.rFonts.eastAsia = "宋体";
                
                if (rPr1.sz == null)
                {
                    rPr1.sz = new CT_HpsMeasure();
                }
                rPr1.sz.val = (ulong)(18 * 2); // 18磅 = 36半点
                
                // 设置大纲级别
                var pPr1 = para1.GetCTP().pPr;
                if (pPr1 == null)
                {
                    pPr1 = para1.GetCTP().AddNewPPr();
                }
                
                if (pPr1.outlineLvl == null)
                {
                    pPr1.outlineLvl = new CT_DecimalNumber();
                }
                pPr1.outlineLvl.val = "1"; // 大纲级别1
                
                // 创建标题2
                var para2 = document.CreateParagraph();
                var run2 = para2.CreateRun();
                run2.SetText("1 水文气象");
                
                // 设置字体
                var rPr2 = run2.GetCTR().rPr;
                if (rPr2 == null)
                {
                    rPr2 = run2.GetCTR().AddNewRPr();
                }
                
                if (rPr2.rFonts == null)
                {
                    rPr2.rFonts = new CT_Fonts();
                }
                rPr2.rFonts.ascii = "宋体";
                rPr2.rFonts.hAnsi = "宋体";
                rPr2.rFonts.eastAsia = "宋体";
                
                if (rPr2.sz == null)
                {
                    rPr2.sz = new CT_HpsMeasure();
                }
                rPr2.sz.val = (ulong)(16 * 2); // 16磅 = 32半点
                
                // 设置大纲级别
                var pPr2 = para2.GetCTP().pPr;
                if (pPr2 == null)
                {
                    pPr2 = para2.GetCTP().AddNewPPr();
                }
                
                if (pPr2.outlineLvl == null)
                {
                    pPr2.outlineLvl = new CT_DecimalNumber();
                }
                pPr2.outlineLvl.val = "2"; // 大纲级别2
                
                // 创建标题3
                var para3 = document.CreateParagraph();
                var run3 = para3.CreateRun();
                run3.SetText("1.1 气象要素");
                
                // 设置字体
                var rPr3 = run3.GetCTR().rPr;
                if (rPr3 == null)
                {
                    rPr3 = run3.GetCTR().AddNewRPr();
                }
                
                if (rPr3.rFonts == null)
                {
                    rPr3.rFonts = new CT_Fonts();
                }
                rPr3.rFonts.ascii = "宋体";
                rPr3.rFonts.hAnsi = "宋体";
                rPr3.rFonts.eastAsia = "宋体";
                
                if (rPr3.sz == null)
                {
                    rPr3.sz = new CT_HpsMeasure();
                }
                rPr3.sz.val = (ulong)(15 * 2); // 15磅 = 30半点
                
                // 设置大纲级别
                var pPr3 = para3.GetCTP().pPr;
                if (pPr3 == null)
                {
                    pPr3 = para3.GetCTP().AddNewPPr();
                }
                
                if (pPr3.outlineLvl == null)
                {
                    pPr3.outlineLvl = new CT_DecimalNumber();
                }
                pPr3.outlineLvl.val = "3"; // 大纲级别3
                
                // 创建正文
                var para4 = document.CreateParagraph();
                var run4 = para4.CreateRun();
                run4.SetText("这是正文内容，用于测试文档处理功能。");
                
                // 设置字体
                var rPr4 = run4.GetCTR().rPr;
                if (rPr4 == null)
                {
                    rPr4 = run4.GetCTR().AddNewRPr();
                }
                
                if (rPr4.rFonts == null)
                {
                    rPr4.rFonts = new CT_Fonts();
                }
                rPr4.rFonts.ascii = "宋体";
                rPr4.rFonts.hAnsi = "宋体";
                rPr4.rFonts.eastAsia = "宋体";
                
                if (rPr4.sz == null)
                {
                    rPr4.sz = new CT_HpsMeasure();
                }
                rPr4.sz.val = (ulong)(10.5 * 2); // 10.5磅 = 21半点
                
                // 保存文档
                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    document.Write(fileStream);
                }
            }
        }
    }
}