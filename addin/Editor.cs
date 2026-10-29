using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeAi
{
    /// <summary>
    /// 编辑计划执行引擎（第1条）。
    /// 目标：写作模式不再“整篇重生成”，而是把用户指令转成一组结构化编辑动作，
    /// 由这里在当前文档里真正定位并执行。
    /// 支持的 op：
    ///   replace     查找替换（find -> replace，可选 all 全文 / first 首处）
    ///   append      在光标处插入（等于原来“模拟输入”，用于“续写”时把光标之后继续）
    ///   insertAfter 定位到 anchor 文本之后续写（真正“定位到我说的地方”）
    ///   rewrite     覆盖当前选区
    /// 注意：执行仍由用户点「确认执行」触发，AI 侧不会自动落到文档。
    /// </summary>
    internal static class Editor
    {
        /// <summary>
        /// 执行一个编辑动作。op 为 Dictionary。
        /// 返回 {"ok":bool,"message":"...","detail":"..."}
        /// </summary>
        public static string Execute(object app, Connect.HostKind kind, Dictionary<string, object> op, string writeMethod)
        {
            string action = op != null && op.ContainsKey("op") ? Convert.ToString(op["op"]) : "";
            switch (action)
            {
                case "replace":   return DoReplace(app, kind, op);
                case "append":    return DoAppend(app, kind, op, writeMethod);
                case "insertAfter": return DoInsertAfter(app, kind, op, writeMethod);
                case "rewrite":   return DoRewrite(app, kind, op);
                default:
                    return Json(false, "未知编辑动作：" + action);
            }
        }

        /// <summary>执行整组编辑，返回 {"ok":bool,"message":"...","detail":"共执行 N 步"} 或失败。</summary>
        public static string ExecutePlan(object app, Connect.HostKind kind, List<object> ops, string writeMethod)
        {
            if (ops == null || ops.Count == 0)
                return Json(false, "没有可执行的编辑动作。");
            foreach (var item in ops)
            {
                var op = item as Dictionary<string, object>;
                if (op == null) return Json(false, "编辑动作格式无效。");
                string name = Get(op, "op");
                if (name != "replace" && name != "append" && name != "insertAfter" && name != "rewrite")
                    return Json(false, "无法识别的编辑动作：" + name);
                if (name == "replace" && (Get(op, "find").Length == 0 || !op.ContainsKey("replace")))
                    return Json(false, "替换动作缺少原文或新内容。");
                if (name != "replace" && Get(op, "text").Length == 0)
                    return Json(false, "编辑内容为空。");
                if (name == "insertAfter" && Get(op, "anchor").Length == 0)
                    return Json(false, "定位续写缺少原文锚点。");
                if (op.ContainsKey("all") && !(op["all"] is bool)) return Json(false, "替换范围格式无效。");
            }
            // Word plans form one undo record, allowing a failed later action to roll back.
            dynamic word = app;
            bool grouped = false;
            string original = null;
            if (kind == Connect.HostKind.Word)
            {
                try
                {
                    original = (string)word.ActiveDocument.Content.Text;
                    word.UndoRecord.StartCustomRecord("OfficeAI 修改原文");
                    grouped = true;
                }
                catch { return Json(false, "当前 Word 不支持安全编辑记录，未修改文档。"); }
            }
            var sb = new StringBuilder();
            int done = 0;
            string failure = null;
            try
            {
                foreach (var item in ops)
                {
                    var result = ParseJson(Execute(app, kind, (Dictionary<string, object>)item, writeMethod));
                    if (result == null || !result.ContainsKey("ok") || !Convert.ToBoolean(result["ok"]))
                    {
                        failure = "第 " + (done + 1) + " 步失败：" +
                            (result != null && result.ContainsKey("message") ? Convert.ToString(result["message"]) : "执行失败");
                        break;
                    }
                    done++;
                    if (result.ContainsKey("detail")) sb.AppendLine(done + ". " + Convert.ToString(result["detail"]));
                }
            }
            catch (Exception ex) { failure = "执行失败：" + ex.Message; }
            finally { if (grouped) word.UndoRecord.EndCustomRecord(); }
            if (failure != null)
            {
                if (grouped && original != (string)word.ActiveDocument.Content.Text)
                {
                    bool undone = word.ActiveDocument.Undo(1);
                    bool restored = undone && original == (string)word.ActiveDocument.Content.Text;
                    return Json(false, failure + (restored ? "；本次修改已全部撤回，请重新生成。" : "；自动撤回失败，请检查文档。"));
                }
                return Json(false, failure + (done > 0 ? "；前 " + done + " 步已执行，请检查后重新生成。" : "；未修改原文。"));
            }
            return "{\"ok\":true,\"message\":\"已执行 " + done + " 个编辑动作，可用 Ctrl+Z 撤销\",\"detail\":\"" + Esc(sb.ToString().Trim()) + "\"}";
        }

        #region 各动作

        private static string DoReplace(object app, Connect.HostKind kind, Dictionary<string, object> op)
        {
            string find = Get(op, "find");
            string replace = Get(op, "replace");
            if (string.IsNullOrEmpty(find))
                return Json(false, "替换缺少「查找内容」find。");
            bool all = !op.ContainsKey("all") || Convert.ToBoolean(op["all"]);
            // 复用 Precise 的全文查找替换
            string r = Precise.FindReplace(app, kind, find, replace, all);
            var d = ParseJson(r);
            if (d == null || !d.ContainsKey("ok") || !Convert.ToBoolean(d["ok"]))
            {
                string m = d != null && d.ContainsKey("message") ? Convert.ToString(d["message"]) : "替换失败";
                return Json(false, "替换失败：" + m);
            }
            int n = d.ContainsKey("replaced") ? Convert.ToInt32(d["replaced"]) : 0;
            if (n == 0) return Json(false, "未找到要替换的原文，请重新生成预览。");
            return Ok("已替换" + (all ? "全部" : "首处") + "：" + find + " → " + replace +
                      (all ? "（" + n + " 处）" : ""));
        }

        private static string DoAppend(object app, Connect.HostKind kind, Dictionary<string, object> op, string writeMethod)
        {
            string text = Get(op, "text");
            if (string.IsNullOrEmpty(text)) return Json(false, "续写内容为空。");
            // 续写 = 在当前光标处继续输入（不覆盖原文）
            return Writer.Apply(app, kind, Normalize(text), writeMethod);
        }

        private static string DoInsertAfter(object app, Connect.HostKind kind, Dictionary<string, object> op, string writeMethod)
        {
            string anchor = Get(op, "anchor");
            string text = Get(op, "text");
            if (string.IsNullOrWhiteSpace(anchor)) return Json(false, "定位续写缺少锚点 anchor。");
            if (string.IsNullOrEmpty(text)) return Json(false, "续写内容为空。");
            string err;
            bool ok;
            switch (kind)
            {
                case Connect.HostKind.Word:   ok = InsertAfterWord(app, anchor, text, out err); break;
                case Connect.HostKind.Excel:  ok = InsertAfterExcel(app, anchor, text, out err); break;
                case Connect.HostKind.PowerPoint: ok = InsertAfterPpt(app, anchor, text, out err); break;
                default: return Json(false, "不支持的宿主。");
            }
            Connect.Log("insertAfter(" + kind + ") ok=" + ok + " anchor=" + anchor + " " + err);
            return ok ? Ok("已在「" + AnchorPreview(anchor) + "」之后续写。")
                      : Json(false, "定位失败：" + err);
        }

        private static string DoRewrite(object app, Connect.HostKind kind, Dictionary<string, object> op)
        {
            string text = Get(op, "text");
            if (string.IsNullOrEmpty(text)) return Json(false, "改写内容为空。");
            return Precise.ReplaceSelection(app, kind, Normalize(text));
        }

        #endregion

        #region 定位续写（第1条：真正定位到指定位置）

        private static bool InsertAfterWord(object app, string anchor, string text, out string err)
        {
            err = "";
            try
            {
                dynamic a = app;
                dynamic doc = a.ActiveDocument;
                dynamic content = doc.Content;
                string body = (string)content.Text;
                string needle = ToWordText(anchor);
                int at = body.IndexOf(needle, StringComparison.Ordinal);
                if (at < 0) { err = "没有找到锚点「" + AnchorPreview(anchor) + "」"; return false; }
                if (body.IndexOf(needle, at + needle.Length, StringComparison.Ordinal) >= 0)
                { err = "锚点有多处，请提供更完整的原文以唯一定位。"; return false; }
                int end = (int)content.Start + at + needle.Length;
                doc.Range(end, end).Text = ToWordText(text);
                return true;
            }
            catch (Exception e) { err = e.Message; return false; }
        }

        private static bool InsertAfterExcel(object app, string anchor, string text, out string err)
        {
            err = "";
            try
            {
                dynamic a = app;
                dynamic ws = a.ActiveSheet;
                // 找到第一个包含 anchor 的单元格，在它下方单元格写入续写
                dynamic used = ws.UsedRange;
                int rows = (int)used.Rows.Count;
                int cols = (int)used.Columns.Count;
                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        object v = null;
                        try { v = used.Cells(r, c).Value2; } catch { }
                        string s = v == null ? "" : Convert.ToString(v);
                        if (s.Contains(anchor))
                        {
                            int tr = used.Row + r - 1;  // 绝对行
                            int tc = used.Column + c - 1;
                            ws.Cells(tr + 1, tc).Value2 = text.Replace("\r\n", "\n");
                            return true;
                        }
                    }
                }
                err = "没有找到包含「" + AnchorPreview(anchor) + "」的单元格";
                return false;
            }
            catch (Exception e) { err = e.Message; return false; }
        }

        private static bool InsertAfterPpt(object app, string anchor, string text, out string err)
        {
            err = "";
            try
            {
                dynamic a = app;
                dynamic pres = a.ActivePresentation;
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
                        try { t = Convert.ToString(sh.TextFrame.TextRange.Text); } catch { continue; }
                        if (string.IsNullOrEmpty(t)) continue;
                        int idx = t.IndexOf(anchor, StringComparison.Ordinal);
                        if (idx < 0) continue;
                        string nt = t.Substring(0, idx + anchor.Length) + text + t.Substring(idx + anchor.Length);
                        sh.TextFrame.TextRange.Text = nt;
                        return true;
                    }
                }
                err = "没有找到包含锚点的文本框";
                return false;
            }
            catch (Exception e) { err = e.Message; return false; }
        }

        #endregion

        #region 辅助

        private static string Normalize(string text) => text;

        private static string ToWordText(string text)
            => text.Replace("\r\n", "\n").Replace("\n", "\r");

        private static string AnchorPreview(string s)
        {
            s = (s ?? "").Replace("\r", " ").Replace("\n", " ");
            return s.Length > 20 ? s.Substring(0, 20) + "…" : s;
        }

        private static string Get(Dictionary<string, object> d, string k)
            => d != null && d.ContainsKey(k) && d[k] != null ? Convert.ToString(d[k]) : "";

        private static Dictionary<string, object> ParseJson(string s)
        {
            try { return new System.Web.Script.Serialization.JavaScriptSerializer().DeserializeObject(s) as Dictionary<string, object>; }
            catch { return null; }
        }

        private static string Ok(string msg)
            => "{\"ok\":true,\"message\":\"" + Esc(msg) + "\",\"detail\":\"" + Esc(msg) + "\"}";

        private static string Json(bool ok, string msg)
            => "{\"ok\":" + (ok ? "true" : "false") + ",\"message\":\"" + Esc(msg) + "\"}";

        private static string Esc(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"")
                    .Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
        }

        #endregion
    }
}
