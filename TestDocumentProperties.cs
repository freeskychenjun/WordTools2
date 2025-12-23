using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace TestDocumentProperties;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("测试DocumentService识别文档属性...");
        Console.WriteLine("======================================");
        
        try
        {
            // 创建DocumentService实例
            var documentService = new WordTools2.Services.DocumentService();
            
            // 打开测试文档
            string filePath = "test.docx";
            Console.WriteLine($"正在读取文档: {filePath}");
            
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                // 获取文档的主文档部分
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                if (mainPart == null)
                {
                    Console.WriteLine("错误: 文档没有主文档部分");
                    return;
                }
                
                // 获取文档的body
                Body body = mainPart.Document.Body;
                if (body == null)
                {
                    Console.WriteLine("错误: 文档没有body");
                    return;
                }
                
                // 获取第一个段落
                Paragraph firstParagraph = body.Elements<Paragraph>().FirstOrDefault();
                if (firstParagraph == null)
                {
                    Console.WriteLine("错误: 文档没有段落");
                    return;
                }
                
                Console.WriteLine("\n=== 第一段属性识别结果 ===");
                
                // 1. 测试大纲级别识别
                int outlineLevel = documentService.GetParagraphOutlineLevel(firstParagraph, mainPart);
                Console.WriteLine($"大纲级别: {outlineLevel}");
                
                // 2. 测试字体名称识别
                string fontName = documentService.GetParagraphFontName(firstParagraph, mainPart);
                Console.WriteLine($"字体名称: {fontName}");
                
                // 3. 测试字体大小识别
                double fontSize = documentService.GetParagraphFontSize(firstParagraph, mainPart);
                Console.WriteLine($"字体大小: {fontSize}pt");
                
                Console.WriteLine("\n=== 测试完成 ===");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"测试过程中发生错误: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
        
        Console.WriteLine("\n按任意键退出...");
        Console.ReadKey();
    }
}