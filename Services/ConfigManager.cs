using System;
using System.Text.Json;
using System.IO;
using WordTools2.Models;

namespace WordTools2.Services
{
    /// <summary>
    /// 配置文件管理类，用于保存和读取配置
    /// </summary>
    public class ConfigManager
    {
        private static string _configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
        
        /// <summary>
        /// 保存配置到文件
        /// </summary>
        public static void SaveConfig(Models.StyleConfig config)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
                };
                
                string jsonString = JsonSerializer.Serialize(config, options);
                File.WriteAllText(_configPath, jsonString);
            }
            catch (Exception ex)
            {
                throw new Exception($"保存配置失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 从文件读取配置
        /// </summary>
        public static Models.StyleConfig LoadConfig()
        {
            try
            {
                if (File.Exists(_configPath))
                {
                    string jsonString = File.ReadAllText(_configPath);
                    var config = JsonSerializer.Deserialize<Models.StyleConfig>(jsonString);
                    return config ?? new Models.StyleConfig();
                }
                else
                {
                    // 如果配置文件不存在，返回默认配置
                    var defaultConfig = new Models.StyleConfig();
                    SaveConfig(defaultConfig);
                    return defaultConfig;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"读取配置失败: {ex.Message}");
            }
        }
    }
}