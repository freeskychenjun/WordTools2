using System;
using System.Text.RegularExpressions;

namespace WordTools2
{
    class TestRegex
    {
        static void Main(string[] args)
        {
            // 测试修改后的正则表达式
            string pattern = @"^\s*表(?=\s|\d)\s*\d*(\.\d+)*(-\s*\d+)*\s+.*$";
            
            // 测试用例
            string[] testCases = {
                "表5.6.1-1 GIS图层列",
                "表1 测试表格",
                "表2.1 中文标题",
                "表3.1.1-2 混合 Title",
                "  表4  带空格的标题",
                "正文文本，不是表格标题",
                "图1 测试图片"
            };
            
            Console.WriteLine("=== 表格标题识别测试 ===");
            Console.WriteLine($"正则表达式: {pattern}");
            Console.WriteLine("------------------------");
            
            foreach (string testCase in testCases)
            {
                bool isMatch = Regex.IsMatch(testCase, pattern);
                Console.WriteLine($"'{testCase}' -> {isMatch}");
            }
            
            Console.WriteLine("\n=== 测试完成 ===");
            Console.WriteLine("按任意键退出...");
            Console.ReadKey();
        }
    }
}