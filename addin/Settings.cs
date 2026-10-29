using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace OfficeAi
{
    /// <summary>设置持久化：%LOCALAPPDATA%\OfficeAi\settings.json</summary>
    internal static class Settings
    {
        private static string Dir
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OfficeAi");
            }
        }

        private static string FilePath { get { return Path.Combine(Dir, "settings.json"); } }

        public static Dictionary<string, object> Defaults()
        {
            return new Dictionary<string, object>
            {
                { "apiKey", "" },
                { "baseUrl", "https://api.deepseek.com" },
                { "model", "deepseek-chat" },
                { "temperature", 0.7 },
                { "method", "type" },          // 写入方式：type=模拟输入, vba=VBA
                { "theme", "light" },          // 界面主题：light=浅色, dark=深色
                { "autoOpenPane", true },      // 第5条：打开文件时自动开侧边栏
                { "autoRead", true },          // 第4条：是否自动读取文档
                { "autoReadIntervalMin", 0 },  // 第4条：自动重读间隔（分钟），0=仅打开时读一次
                // 崩溃取证（T2）
                { "watchdog", false },         // 仅排障时显式启用，日常使用不附加原生调试器
                { "crashDumpAllAv", false },   // true=连「宿主可能自行处理掉的探测型 AV」也写 full-memory dump
                                               //      （会白写几百 MB 并卡几秒；只在怀疑漏抓时临时打开）
                // ima openapi（第3条）
                { "imaClientId", "" },
                { "imaApiKey", "" },
                { "imaBaseUrl", "https://ima.qq.com" }
            };
        }

        public static Dictionary<string, object> Load()
        {
            var def = Defaults();
            try
            {
                if (!File.Exists(FilePath)) return def;
                string json = File.ReadAllText(FilePath, Encoding.UTF8);
                var d = new JavaScriptSerializer().DeserializeObject(json) as Dictionary<string, object>;
                if (d == null) return def;
                foreach (var kv in d) def[kv.Key] = kv.Value;
            }
            catch (Exception e)
            {
                Connect.Log("Settings.Load 失败: " + e.Message);
            }
            return def;
        }

        /// <summary>返回给前端的 JSON（含 apiKey / ima 凭据，只在本机；前端设置面板需要回显）</summary>
        public static string LoadJson()
        {
            try { return new JavaScriptSerializer().Serialize(Load()); }
            catch (Exception e) { Connect.Log("Settings.LoadJson 失败: " + e.Message); return "{}"; }
        }

        public static bool SaveJson(string json)
        {
            try
            {
                var ser = new JavaScriptSerializer();
                var d = ser.DeserializeObject(json) as Dictionary<string, object>;
                if (d == null) return false;
                var cur = Load();
                foreach (var kv in d) cur[kv.Key] = kv.Value;
                Directory.CreateDirectory(Dir);
                File.WriteAllText(FilePath, ser.Serialize(cur), new UTF8Encoding(false));
                Connect.Log("设置已保存（字段数 " + d.Count + "）");
                return true;
            }
            catch (Exception e)
            {
                Connect.Log("Settings.SaveJson 失败: " + e.Message);
                return false;
            }
        }

        public static string Get(string key, string fallback)
        {
            try
            {
                var d = Load();
                if (d.ContainsKey(key) && d[key] != null)
                {
                    string s = Convert.ToString(d[key]);
                    return string.IsNullOrEmpty(s) ? fallback : s;
                }
            }
            catch { }
            return fallback;
        }

        public static bool GetBool(string key, bool fallback)
        {
            try
            {
                var d = Load();
                if (d.ContainsKey(key) && d[key] != null) return Convert.ToBoolean(d[key]);
            }
            catch { }
            return fallback;
        }

        public static int GetInt(string key, int fallback)
        {
            try
            {
                var d = Load();
                if (d.ContainsKey(key) && d[key] != null) return Convert.ToInt32(d[key]);
            }
            catch { }
            return fallback;
        }
    }
}
