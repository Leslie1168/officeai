using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeAi
{
    /// <summary>
    /// 精准定位与替换（第1条）。
    /// 目标：不整篇替换，而是「针对选区改写」或「全文查找替换」，并且能处理 Word 表格单元格。
    /// </summary>
    internal static class Precise
    {
        #region 读取当前选区

        /// <summary>
        /// 返回当前选区信息 JSON：
        /// {"ok":true,"kind":"word|excel|ppt","range":"A1:B2","text":"...","isTableCell":bool,
        ///  "cellRow":r,"cellCol":c,"tableIndex":i,"count":字符数}
        /// </summary>
        public static string GetSelection(object app, Connect.HostKind kind)
        {
            try
            {
                if (app == null) return "{\"ok\":false,\"message\":\"没有宿主对象\"}";
                dynamic a = app;
                switch (kind)
                {
                    case Connect.HostKind.Word: return GetSelWord(a);
                    case Connect.HostKind.Excel: return GetSelExcel(a);
                    case Connect.HostKind.PowerPoint: return GetSelPpt(a);
                }
                return "{\"ok\":false,\"message\":\"认不出当前宿主\"}";
            }
            catch (Exception e)
            {
                Connect.Log("GetSelection 异常: " + e);
                return "{\"ok\":false,\"message\":\"" + Esc(e.Message) + "\"}";
            }
        }

        private static string GetSelWord(dynamic a)
        {
            dynamic sel = a.Selection;
            if (sel == null) return "{\"ok\":false,\"message\":\"没有光标\"}";
            string text = Safe(() => (string)sel.Text, "");
            bool inTable = Safe(() => (bool)sel.Information[12] /*wdWithInTable*/, false);
            int row = 0, col = 0;
            if (inTable)
            {
                row = Safe(() => (int)sel.Cells[1].RowIndex, 0);
                col = Safe(() => (int)sel.Cells[1].ColumnIndex, 0);
            }
            // 选区在文档中的字符区间，便于“定位改写”
            int start = Safe(() => (int)sel.Start, -1);
            int end = Safe(() => (int)sel.End, -1);
            var d = new Dictionary<string, object>
            {
                { "ok", true }, { "kind", "word" },
                { "text", text }, { "count", text.Length },
                { "isTableCell", inTable }, { "cellRow", row }, { "cellCol", col },
                { "start", start }, { "end", end },
                { "document", Safe(() => (string)a.ActiveDocument.FullName, "") },
                { "range", (start >= 0 && end >= start) ? (start + "-" + end) : "" }
            };
            return Json(d);
        }

        private static string GetSelExcel(dynamic a)
        {
            dynamic sel = a.Selection;
            if (sel == null) return "{\"ok\":false,\"message\":\"没有活动选区\"}";
            string addr = Safe(() => (string)sel.Address, "");
            string sheet = Safe(() => (string)a.ActiveSheet.Name, "");
            // 读取选区内容（含多单元格）
            string text = "";
            int rows = Safe(() => (int)sel.Rows.Count, 1);
            int cols = Safe(() => (int)sel.Columns.Count, 1);
            object v = null;
            try { v = sel.Value2; } catch { }
            var sb = new StringBuilder();
            object[,] arr = v as object[,];
            if (arr != null)
            {
                for (int r = 1; r <= arr.GetLength(0); r++)
                {
                    var line = new StringBuilder();
                    for (int c = 1; c <= arr.GetLength(1); c++)
                    {
                        if (c > 1) line.Append('\t');
                        object cell = arr[r, c];
                        line.Append(cell == null ? "" : Convert.ToString(cell));
                    }
                    sb.Append(line.ToString().TrimEnd('\t')).Append('\n');
                }
                text = sb.ToString();
            }
            else if (v != null) text = Convert.ToString(v);
            var d = new Dictionary<string, object>
            {
                { "ok", true }, { "kind", "excel" },
                { "sheet", sheet }, { "range", addr },
                { "text", text }, { "count", text.Length },
                { "rows", rows }, { "cols", cols }
            };
            return Json(d);
        }

        private static string GetSelPpt(dynamic a)
        {
            dynamic win = a.ActiveWindow;
            if (win == null) return "{\"ok\":false,\"message\":\"没有活动窗口\"}";
            int slideIdx = Safe(() => (int)win.View.Slide.SlideIndex, 0);
            string text = "";
            bool hasTextSel = false;
            try
            {
                dynamic sel = win.Selection;
                if ((int)sel.Type == 3 /*ppSelectionText*/)
                {
                    text = Safe(() => (string)sel.TextRange.Text, "");
                    hasTextSel = text.Length > 0;
                }
                else if ((int)sel.Type == 4 /*ppSelectionShape*/)
                {
                    text = Safe(() => (string)sel.ShapeRange[1].TextFrame.TextRange.Text, "");
                    hasTextSel = text.Length > 0;
                }
            }
            catch { }
            var d = new Dictionary<string, object>
            {
                { "ok", true }, { "kind", "ppt" },
                { "slide", slideIdx }, { "text", text }, { "count", text.Length },
                { "hasTextSel", hasTextSel },
                { "range", "第 " + slideIdx + " 张" }
            };
            return Json(d);
        }

        #endregion

        #region 改写选区

        /// <summary>把当前选区内容替换为 newText（保留其余文档不变）。</summary>
        public static string ReplaceSelection(object app, Connect.HostKind kind, string newText)
        {
            try
            {
                if (app == null) return "{\"ok\":false,\"message\":\"没有宿主对象\"}";
                dynamic a = app;
                string msg;
                bool ok;
                switch (kind)
                {
                    case Connect.HostKind.Word: ok = ReplWord(a, newText, out msg); break;
                    case Connect.HostKind.Excel: ok = ReplExcel(a, newText, out msg); break;
                    case Connect.HostKind.PowerPoint: ok = ReplPpt(a, newText, out msg); break;
                    default: ok = false; msg = "认不出当前宿主"; break;
                }
                Connect.Log("ReplaceSelection(" + kind + ") ok=" + ok + " " + msg);
                return ok ? "{\"ok\":true,\"message\":\"已替换选区\"}"
                          : "{\"ok\":false,\"message\":\"" + Esc(msg) + "\"}";
            }
            catch (Exception e)
            {
                Connect.Log("ReplaceSelection 异常: " + e);
                return "{\"ok\":false,\"message\":\"" + Esc(e.Message) + "\"}";
            }
        }

        private static bool ReplWord(dynamic a, string newText, out string msg)
        {
            msg = "";
            try
            {
                dynamic sel = a.Selection;
                if (sel == null) { msg = "没有光标"; return false; }
                sel.Text = newText.Replace("\r\n", "\r").Replace("\n", "\r");
                return true;
            }
            catch (Exception e) { msg = e.Message; return false; }
        }

        private static bool ReplExcel(dynamic a, string newText, out string msg)
        {
            msg = "";
            try
            {
                dynamic sel = a.Selection;
                if (sel == null) { msg = "没有活动单元格"; return false; }
                string[] lines = newText.Replace("\r\n", "\n").Split('\n');
                if (lines.Length == 1)
                {
                    sel.Value2 = lines[0];           // 单个单元格直接写
                }
                else
                {
                    dynamic ws = a.ActiveSheet;
                    int r0 = (int)sel.Cells[1, 1].Row;
                    int c0 = (int)sel.Cells[1, 1].Column;
                    for (int i = 0; i < lines.Length; i++)
                        ws.Cells(r0 + i, c0).Value2 = lines[i];
                }
                return true;
            }
            catch (Exception e) { msg = e.Message; return false; }
        }

        private static bool ReplPpt(dynamic a, string newText, out string msg)
        {
            msg = "";
            try
            {
                dynamic sel = a.ActiveWindow.Selection;
                if ((int)sel.Type == 3) { sel.TextRange.Text = newText; return true; }
                if ((int)sel.Type == 4) { sel.ShapeRange[1].TextFrame.TextRange.Text = newText; return true; }
                msg = "请先在幻灯片里选中一个文本框";
                return false;
            }
            catch (Exception e) { msg = e.Message; return false; }
        }

        #endregion

        #region 全文查找替换

        /// <summary>
        /// 全文查找替换。scope=all 替换所有；scope=first 只替换第一个。
        /// 返回 {"ok":true,"replaced":N}
        /// </summary>
        public static string FindReplace(object app, Connect.HostKind kind,
                                         string find, string replace, bool all)
        {
            try
            {
                if (app == null) return "{\"ok\":false,\"message\":\"没有宿主对象\"}";
                if (string.IsNullOrEmpty(find)) return "{\"ok\":false,\"message\":\"查找内容不能为空\"}";
                dynamic a = app;
                int n;
                string msg;
                switch (kind)
                {
                    case Connect.HostKind.Word: n = FrWord(a, find, replace, all, out msg); break;
                    case Connect.HostKind.Excel: n = FrExcel(a, find, replace, all, out msg); break;
                    case Connect.HostKind.PowerPoint: n = FrPpt(a, find, replace, all, out msg); break;
                    default: return "{\"ok\":false,\"message\":\"认不出当前宿主\"}";
                }
                Connect.Log("FindReplace(" + kind + ") 替换 " + n + " 处 " + msg);
                if (n < 0) return "{\"ok\":false,\"message\":\"" + Esc(msg) + "\"}";
                return "{\"ok\":true,\"replaced\":" + n + "}";
            }
            catch (Exception e)
            {
                Connect.Log("FindReplace 异常: " + e);
                return "{\"ok\":false,\"message\":\"" + Esc(e.Message) + "\"}";
            }
        }

        private static int FrWord(dynamic a, string find, string replace, bool all, out string msg)
        {
            msg = "";
            try
            {
                dynamic doc = a.ActiveDocument;
                dynamic content = doc.Content;
                string body = (string)content.Text;
                string needle = find.Replace("\r\n", "\r").Replace("\n", "\r");
                string replacement = replace.Replace("\r\n", "\r").Replace("\n", "\r");
                var hits = new List<int>();
                int offset = 0, at;
                while ((at = body.IndexOf(needle, offset, StringComparison.Ordinal)) >= 0)
                {
                    hits.Add(at);
                    if (!all) break;
                    offset = at + needle.Length;
                }
                int origin = (int)content.Start;
                for (int i = hits.Count - 1; i >= 0; i--)
                    doc.Range(origin + hits[i], origin + hits[i] + needle.Length).Text = replacement;
                return hits.Count;
            }
            catch (Exception e) { msg = e.Message; return -1; }
        }

        private static int FrExcel(dynamic a, string find, string replace, bool all, out string msg)
        {
            msg = "";
            try
            {
                dynamic ws = a.ActiveSheet;
                dynamic used = ws.UsedRange;
                if ((double)used.CountLarge > 200000) { msg = "工作表范围过大，请缩小数据区域后再替换。"; return -1; }
                int count = 0;
                for (int r = 1; r <= (int)used.Rows.Count; r++)
                    for (int c = 1; c <= (int)used.Columns.Count; c++)
                    {
                        dynamic cell = used.Cells[r, c];
                        if ((bool)cell.HasFormula) continue; // never flatten formulas into values
                        object value = cell.Value2;
                        if (!(value is string)) continue;
                        string text = (string)value;
                        if (!text.Contains(find)) continue;
                        if (all)
                        {
                            int at = 0, hit;
                            while ((hit = text.IndexOf(find, at, StringComparison.Ordinal)) >= 0) { count++; at = hit + find.Length; }
                            cell.Value2 = text.Replace(find, replace);
                        }
                        else { cell.Value2 = ReplaceFirst(text, find, replace); return 1; }
                    }
                return count;
            }
            catch (Exception e) { msg = e.Message; return -1; }
        }

        private static int FrPpt(dynamic a, string find, string replace, bool all, out string msg)
        {
            msg = "";
            try
            {
                dynamic pres = a.ActivePresentation;
                int count = 0;
                int slides = (int)pres.Slides.Count;
                for (int i = 1; i <= slides; i++)
                {
                    dynamic slide = pres.Slides[i];
                    int shp = (int)slide.Shapes.Count;
                    for (int j = 1; j <= shp; j++)
                    {
                        dynamic sh;
                        try { sh = slide.Shapes[j]; } catch { continue; }
                        string t;
                        try { t = Convert.ToString(sh.TextFrame.TextRange.Text); }
                        catch { continue; }
                        if (string.IsNullOrEmpty(t) || !t.Contains(find)) continue;
                        string nt = all ? t.Replace(find, replace)
                                        : ReplaceFirst(t, find, replace);
                        try { sh.TextFrame.TextRange.Text = nt; count++; } catch { }
                        if (!all && count >= 1) break;
                    }
                    if (!all && count >= 1) break;
                }
                return count;
            }
            catch (Exception e) { msg = e.Message; return -1; }
        }

        private static string ReplaceFirst(string s, string find, string replace)
        {
            int i = s.IndexOf(find, StringComparison.Ordinal);
            if (i < 0) return s;
            return s.Substring(0, i) + replace + s.Substring(i + find.Length);
        }

        #endregion

        #region 辅助

        private static T Safe<T>(Func<T> f, T fb)
        {
            try { return f(); } catch { return fb; }
        }

        private static string Json(Dictionary<string, object> d)
        {
            return new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(d);
        }

        private static string Esc(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"")
                    .Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
        }

        #endregion
    }
}
