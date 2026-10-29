using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Web.Script.Serialization;

namespace OfficeAi
{
    /// <summary>
    /// 从当前打开的文档里抓取所有内嵌图片。
    /// Word/Excel/PPT 的内嵌图存在 OOXML 包的 */media/ 里，本身没有独立磁盘路径，
    /// 因此提取时把它们落盘到缓存目录，作为可供“复制/另存/打开”的真实文件；
    /// 预览一律用同目录下的缩略图，不依赖真实文件。
    /// </summary>
    internal static class ImageExtractor
    {
        private static readonly string[] ImgExt =
            { ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf" };

        /// <summary>缓存根目录：%LOCALAPPDATA%\OfficeAi\imgcache</summary>
        public static string CacheRoot
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OfficeAi", "imgcache");
            }
        }

        private const int ThumbMax = 220;

        /// <summary>
        /// 提取文档图片。返回 JSON：
        /// {"ok":true,"count":N,"docDir":"<相对 imgcache 的子目录>","items":[{"name","orig","thumb","w","h","size"}]}
        /// orig / thumb 都是相对 imgcache 的正斜杠路径，前端拼 https://officeai-img.local/ 前缀即可预览。
        /// </summary>
        /// </summary>
        public static string Extract(object app, Connect.HostKind kind)
        {
            try
            {
                string docPath = GetDocPath(app, kind);
                if (string.IsNullOrEmpty(docPath) || !File.Exists(docPath))
                    return Fail("请先保存当前文档（磁盘上还没有文件，无法提取内嵌图片）。");

                string docDir = MakeDocDir(docPath);
                string origDir = Path.Combine(docDir, "orig");
                string thumbDir = Path.Combine(docDir, "thumbs");
                if (Directory.Exists(origDir)) Directory.Delete(origDir, true);
                if (Directory.Exists(thumbDir)) Directory.Delete(thumbDir, true);
                Directory.CreateDirectory(origDir);
                Directory.CreateDirectory(thumbDir);

                var items = new List<Dictionary<string, object>>();
                // docDirRel = 相对 CacheRoot 的子目录（正斜杠），如 "报告_1A2B3C4D"
                string docDirRel = docDir.Substring(CacheRoot.Length).Trim('\\', '/').Replace('\\', '/');
                // 文档正被 Word/Excel/PPT 打开时，ZipFile.OpenRead 会以独占方式打开而失败
                // （报「正由另一进程使用」）。这里用 FileShare.ReadWrite 共享读取，才能读打开中的文档。
                using (var fs = new FileStream(docPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    int idx = 0;
                    foreach (var entry in zip.Entries)
                    {
                        string n = entry.FullName.ToLowerInvariant();
                        bool isMedia = n.Contains("/media/") || n.StartsWith("media/");
                        if (!isMedia) continue;
                        if (!HasImageExt(n)) continue;
                        if (entry.Length == 0) continue;

                        idx++;
                        string baseName = idx.ToString("D3");
                        string ext = Path.GetExtension(entry.Name);

                        string origRel = docDirRel + "/orig/" + baseName + ext;
                        string origAbs = Path.Combine(CacheRoot, origRel.Replace('/', '\\'));
                        entry.ExtractToFile(origAbs, true);

                        // 缩略图
                        string thumbRel = "";
                        int w = 0, h = 0;
                        try
                        {
                            using (var img = Image.FromFile(origAbs))
                            {
                                w = img.Width; h = img.Height;
                                double scale = Math.Min(1.0, (double)ThumbMax / Math.Max(w, h));
                                int tw = Math.Max(1, (int)(w * scale));
                                int th = Math.Max(1, (int)(h * scale));
                                using (var bmp = new Bitmap(tw, th))
                                using (var g = Graphics.FromImage(bmp))
                                {
                                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    g.DrawImage(img, 0, 0, tw, th);
                                    thumbRel = docDirRel + "/thumbs/" + baseName + ".png";
                                    bmp.Save(Path.Combine(CacheRoot, thumbRel.Replace('/', '\\')), ImageFormat.Png);
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            // EMF/WMF 等矢量格式可能无法生成缩略图，保留原图即可
                            Connect.Log("缩略图失败 " + baseName + ext + ": " + e.Message);
                        }

                        var item = new Dictionary<string, object>
                        {
                            { "name", baseName + ext },
                            { "orig", origRel },
                            { "thumb", thumbRel },
                            { "w", w }, { "h", h },
                            { "size", new FileInfo(origAbs).Length },
                            { "abs", origAbs.Replace('\\', '/') }
                        };
                        items.Add(item);
                    }
                }

                var res = new Dictionary<string, object>
                {
                    { "ok", true },
                    { "count", items.Count },
                    { "docDir", docDirRel },
                    { "docName", Path.GetFileName(docPath) },
                    { "items", items }
                };
                Connect.Log("ImageExtractor: 提取 " + items.Count + " 张图（" + Path.GetFileName(docPath) + "）");
                return new JavaScriptSerializer().Serialize(res);
            }
            catch (Exception e)
            {
                Connect.Log("ImageExtractor 异常: " + e);
                return Fail(e.Message);
            }
        }

        /// <summary>把图片保存到用户指定位置（供“另存为”）。</summary>
        public static string SaveTo(string srcAbs, string destAbs)
        {
            try
            {
                if (!File.Exists(srcAbs)) return Fail("源图片不存在：" + srcAbs);
                Directory.CreateDirectory(Path.GetDirectoryName(destAbs));
                File.Copy(srcAbs, destAbs, true);
                return "{\"ok\":true,\"path\":\"" + Esc(destAbs) + "\"}";
            }
            catch (Exception e) { return Fail(e.Message); }
        }

        private static string GetDocPath(object app, Connect.HostKind kind)
        {
            try
            {
                dynamic a = app;
                switch (kind)
                {
                    case Connect.HostKind.Word:
                        return (string)a.ActiveDocument.FullName;
                    case Connect.HostKind.Excel:
                        return (string)a.ActiveWorkbook.FullName;
                    case Connect.HostKind.PowerPoint:
                        return (string)a.ActivePresentation.FullName;
                }
            }
            catch { }
            return null;
        }

        private static bool HasImageExt(string lower)
        {
            foreach (var e in ImgExt)
                if (lower.EndsWith(e)) return true;
            return false;
        }

        private static string MakeDocDir(string docPath)
        {
            string name = Path.GetFileNameWithoutExtension(docPath);
            // 用路径哈希避免同名不同文档互相覆盖
            int h = docPath.ToLowerInvariant().GetHashCode();
            string dir = Path.Combine(CacheRoot, name + "_" + Math.Abs(h).ToString("X8"));
            Directory.CreateDirectory(dir);
            return dir;
        }

        private static string Fail(string msg)
        {
            return "{\"ok\":false,\"count\":0,\"error\":\"" + Esc(msg) + "\"}";
        }

        private static string Esc(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n");
        }
    }
}
