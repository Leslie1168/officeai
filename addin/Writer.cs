using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeAi
{
    /// <summary>
    /// 两种写入通道：
    ///   type —— 模拟输入：走宿主自己的“在光标处输入”接口，等于用户手打，格式与撤销栈都自然。
    ///            （不能用 SendKeys：它对中文无效。）
    ///   vba  —— 生成 VBA 代码，尝试注入临时模块并执行；
    ///            若宿主没开“信任对 VBA 工程对象模型的访问”，则退回对象模型执行，
    ///            并把生成的代码回传前端供用户查看/复制。
    /// 注意：写入永远由用户点「确认导入」触发，AI 侧不会自动落到文档上。
    /// </summary>
    internal static class Writer
    {
        private const string MacroName = "OfficeAi_WriteText";

        public static string Apply(object app, Connect.HostKind kind, string text, string method)
        {
            if (app == null) return Fail("没取到宿主程序对象。");
            if (string.IsNullOrEmpty(text)) return Fail("没有要写入的内容。");

            bool wantVba = string.Equals(method, "vba", StringComparison.OrdinalIgnoreCase);
            string vba = BuildVba(kind, text);

            if (wantVba)
            {
                string runErr;
                if (TryRunVba(app, vba, out runErr))
                {
                    Connect.Log("写入成功（VBA）");
                    return Ok("已通过 VBA 写入。", "vba", vba);
                }

                // VBA 通道不可用（多半是没开信任设置）→ 退回对象模型，同时把代码给用户
                Connect.Log("VBA 通道不可用，回退对象模型: " + runErr);
                string msg;
                bool ok = ApplyNative(app, kind, text, out msg);
                if (!ok) return Fail("VBA 不可用且对象模型写入也失败：" + msg);
                return Ok("宿主未开启「信任对 VBA 工程对象模型的访问」，已改用对象模型写入（效果相同）。" +
                          "需要真正跑 VBA 的话，在 文件→选项→信任中心→宏设置 里勾选该项。",
                          "vba-fallback", vba);
            }

            string m2;
            if (!ApplyNative(app, kind, text, out m2)) return Fail(m2);
            Connect.Log("写入成功（模拟输入）");
            return Ok("已写入。", "type", null);
        }

        #region 模拟输入

        private static bool ApplyNative(object app, Connect.HostKind kind, string text, out string err)
        {
            err = "";
            try
            {
                dynamic a = app;
                switch (kind)
                {
                    case Connect.HostKind.Word:
                        {
                            dynamic sel = a.Selection;
                            if (sel == null) { err = "没有光标位置。"; return false; }
                            sel.TypeText(ToWordText(text));
                            return true;
                        }
                    case Connect.HostKind.Excel:
                        {
                            dynamic cell = a.ActiveCell;
                            if (cell == null) { err = "没有活动单元格。"; return false; }
                            int r0 = (int)cell.Row;
                            int c0 = (int)cell.Column;
                            string[] lines = text.Replace("\r\n", "\n").Split('\n');
                            dynamic ws = a.ActiveSheet;
                            for (int i = 0; i < lines.Length; i++)
                                ws.Cells(r0 + i, c0).Value2 = lines[i];
                            return true;
                        }
                    case Connect.HostKind.PowerPoint:
                        {
                            dynamic win = a.ActiveWindow;
                            if (win == null) { err = "没有活动窗口。"; return false; }
                            dynamic sel = win.Selection;
                            // 有选中的文本框就写进去；否则在幻灯片上新建一个文本框
                            try
                            {
                                if ((int)sel.Type == 3 /*ppSelectionText*/)
                                {
                                    sel.TextRange.Text = text.Replace("\r\n", "\n").Replace("\n", "\r");
                                    return true;
                                }
                            }
                            catch { }
                            dynamic slide = win.View.Slide;
                            float w = (float)slide.Master.Width;
                            float h = (float)slide.Master.Height;
                            dynamic box = slide.Shapes.AddTextbox(1 /*msoTextOrientationHorizontal*/,
                                                                 40, 40, w - 80, h - 80);
                            box.TextFrame.TextRange.Text = text;
                            return true;
                        }
                    default:
                        err = "认不出当前宿主程序。";
                        return false;
                }
            }
            catch (Exception e)
            {
                err = e.Message;
                return false;
            }
        }

        /// <summary>Word 的段落分隔用 \r，不是 \n</summary>
        private static string ToWordText(string text)
        {
            return text.Replace("\r\n", "\n").Replace("\n", "\r");
        }

        #endregion

        #region VBA

        private static bool TryRunVba(object app, string code, out string err)
        {
            err = "";
            dynamic comp = null;
            dynamic proj = null;
            try
            {
                dynamic a = app;
                dynamic vbe = a.VBE;
                if (vbe == null) { err = "取不到 VBE。"; return false; }
                proj = vbe.ActiveVBProject;
                if (proj == null) { err = "取不到 VBProject（通常是没有开启信任设置）。"; return false; }

                comp = proj.VBComponents.Add(1); // vbext_ct_StdModule
                comp.CodeModule.AddFromString(code);
                a.Run(MacroName);
                return true;
            }
            catch (Exception e)
            {
                err = e.Message;
                return false;
            }
            finally
            {
                try { if (proj != null && comp != null) proj.VBComponents.Remove(comp); }
                catch { }
            }
        }

        /// <summary>生成 VBA 代码。每行一个字符串字面量，用 Join 拼接，避免超长语句。</summary>
        private static string BuildVba(Connect.HostKind kind, string text)
        {
            string[] lines = text.Replace("\r\n", "\n").Split('\n');
            var sb = new StringBuilder();
            sb.Append("Sub ").Append(MacroName).AppendLine("()");
            sb.AppendLine("    Dim t As String");
            sb.Append("    t = Join(Array(");
            for (int i = 0; i < lines.Length; i++)
            {
                if (i > 0) sb.Append(", ");
                sb.Append('"').Append(lines[i].Replace("\"", "\"\"")).Append('"');
            }
            sb.AppendLine("), vbCr)");

            switch (kind)
            {
                case Connect.HostKind.Word:
                    sb.AppendLine("    Selection.TypeText Text:=t");
                    break;
                case Connect.HostKind.Excel:
                    sb.AppendLine("    Dim a As Variant");
                    sb.AppendLine("    a = Split(t, vbCr)");
                    sb.AppendLine("    Dim i As Long");
                    sb.AppendLine("    For i = 0 To UBound(a)");
                    sb.AppendLine("        ActiveSheet.Cells(ActiveCell.Row + i, ActiveCell.Column).Value = a(i)");
                    sb.AppendLine("    Next i");
                    break;
                case Connect.HostKind.PowerPoint:
                    sb.AppendLine("    ActiveWindow.Selection.TextRange.Text = t");
                    break;
                default:
                    sb.AppendLine("    ' 未知宿主，未生成写入语句");
                    break;
            }
            sb.AppendLine("End Sub");
            return sb.ToString();
        }

        #endregion

        #region JSON

        private static string Ok(string msg, string usedMethod, string vba)
        {
            var sb = new StringBuilder();
            sb.Append("{\"ok\":true,\"message\":\"").Append(Esc(msg))
              .Append("\",\"method\":\"").Append(Esc(usedMethod)).Append("\"");
            if (!string.IsNullOrEmpty(vba))
                sb.Append(",\"vba\":\"").Append(Esc(vba)).Append("\"");
            sb.Append("}");
            return sb.ToString();
        }

        private static string Fail(string msg)
        {
            return "{\"ok\":false,\"message\":\"" + Esc(msg) + "\"}";
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
