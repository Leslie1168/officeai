using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace OfficeAi
{
    internal static class VbaFormatPlan
    {
        internal static string Json(object value) => new JavaScriptSerializer {MaxJsonLength=4194304}.Serialize(value);
        internal static string Hash(string text)
        { using(var hash=SHA256.Create()) return Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(text??""))); }
        internal static string Get(Dictionary<string,object> d,string key,string fallback="") => d.ContainsKey(key) && d[key]!=null ? Convert.ToString(d[key],CultureInfo.InvariantCulture) : fallback;
        internal static bool HasFormat(object[] ops)
        { if(ops==null)return false; foreach(var item in ops) if(item is Dictionary<string,object> d && Get(d,"op")=="format") return true; return false; }
        internal sealed class Plan {internal string Code,Summary;internal int Count;}
        internal static Plan Build(object[] ops,string original,int selectionStart,int selectionEnd)
        {
            if(ops==null || ops.Length==0 || ops.Length>12) throw new ArgumentException("一次请指定 1—12 个格式操作。");
            var code=new StringBuilder();var summary=new StringBuilder();
            code.AppendLine("Option Explicit\nPrivate changed As Boolean\nPublic Sub OfficeAI_Apply(ByVal d As Object)");
            code.AppendLine("  Dim oldScreen As Boolean, opened As Boolean, n As Long, msg As String\n  Dim r As Object, p As Object, role As String\n  On Error GoTo Failed\n  oldScreen = Application.ScreenUpdating\n  changed = False\n  Application.UndoRecord.StartCustomRecord \"OfficeAI VBA 格式\"\n  opened = True\n  Application.ScreenUpdating = False");
            int count=0;
            foreach(var item in ops)
            {
                var op=item as Dictionary<string,object>;
                if(op==null || Get(op,"op")!="format") throw new ArgumentException("格式与正文改写请分两次提交，避免范围变化。");
                string scope=Get(op,"scope");int start,end;
                if(scope=="document") {start=0;end=original.Length;summary.AppendLine("全文格式");}
                else if(scope=="selection")
                {start=selectionStart;end=selectionEnd;if(end<=start)throw new ArgumentException("请先选中文字后重新发送要求。");summary.AppendLine("生成方案时锁定的选区");}
                else if(scope=="match")
                {
                    string find=Get(op,"find").Replace("\r\n","\r").Replace("\n","\r");
                    if(string.IsNullOrWhiteSpace(find))throw new ArgumentException("指定部分必须提供原文。");
                    start=original.IndexOf(find,StringComparison.Ordinal);
                    if(start<0)throw new ArgumentException("没有找到指定原文，请重新描述或选中那一段。");
                    if(original.IndexOf(find,start+find.Length,StringComparison.Ordinal)>=0)throw new ArgumentException("原文有多处相同，请选中具体部分或提供更完整的原文。");
                    end=start+find.Length;summary.AppendLine("指定原文："+find.Replace('\r',' ').Substring(0,Math.Min(find.Length,45)));
                }
                else throw new ArgumentException("范围应为 document、selection 或 match。");
                if(start<0 || end>original.Length || end<=start) throw new ArgumentException("格式范围无效。");
                string preset=Get(op,"preset","none");
                if(preset!="none" && preset!="business" && preset!="academic")throw new ArgumentException("未知全文格式预设。");
                var properties=op.ContainsKey("properties") ? op["properties"] as Dictionary<string,object> : null;
                properties=properties ?? new Dictionary<string,object>();Validate(properties);
                if(preset=="none" && properties.Count==0)throw new ArgumentException("请说明需要更改的字号、字体或间距。");
                bool keep=op.ContainsKey("preserveHeadings") && Convert.ToBoolean(op["preserveHeadings"]);
                if(keep && scope!="document")throw new ArgumentException("保留标题层级只用于全文预设。");
                code.AppendLine("  Set r = d.Range("+start+", "+end+")");
                code.AppendLine("  If r.Paragraphs.Count > 2000 Then Err.Raise 5, , \"一次最多处理 2000 段，请分部分处理\"");
                if(keep)
                {
                    code.AppendLine("  For Each p In r.Paragraphs\n    Set r = p.Range\n    role = \"body\"\n    If CStr(r.Style) = CStr(d.Styles(-63).NameLocal) Or p.OutlineLevel = 1 Then role = \"title\"\n    If CStr(r.Style) = CStr(d.Styles(-75).NameLocal) Or (p.OutlineLevel >= 2 And p.OutlineLevel <= 9) Then role = \"subtitle\"\n    Select Case role");
                    foreach(string role in new[]{"title","subtitle","body"})
                    {code.AppendLine("    Case \""+role+"\"");var values=Defaults(preset,role);foreach(var kv in properties)values[kv.Key]=kv.Value;Assignments(code,values,"      ");}
                    code.AppendLine("    End Select\n  Next p");
                }
                else {var values=preset=="none" ? new Dictionary<string,object>() : Defaults(preset,"body");foreach(var kv in properties)values[kv.Key]=kv.Value;Assignments(code,values,"  ");}
                count++;
            }
            code.AppendLine("  If opened Then Application.UndoRecord.EndCustomRecord\n  Application.ScreenUpdating = oldScreen\n  Exit Sub\nFailed:\n  n = Err.Number: msg = Err.Description\n  On Error Resume Next\n  If opened Then Application.UndoRecord.EndCustomRecord\n  If changed Then d.Undo 1\n  Application.ScreenUpdating = oldScreen\n  On Error GoTo 0\n  Err.Raise n, \"OfficeAI VBA\", msg\nEnd Sub");
            return new Plan {Code=code.ToString(),Summary=summary.ToString().Trim(),Count=count};
        }
        private static Dictionary<string,object> Defaults(string preset,string role)
        {
            bool title=role=="title", sub=role=="subtitle";
            return new Dictionary<string,object>{{"fontName",title||sub?"黑体":"宋体"},{"fontSize",title?22:sub?16:12},{"bold",title||sub},{"italic",false},{"alignment",title?"center":sub?"left":"justify"},{"lineSpacing",title?1.2:sub?1.15:preset=="academic"?1.75:1.5},{"spaceBefore",title?12:sub?8:0},{"spaceAfter",title?10:sub?6:6},{"firstLineIndentChars",title||sub?0:2}};
        }
        private static void Validate(Dictionary<string,object> p)
        {
            foreach(var kv in p)
            {
                if(kv.Key=="fontName") {string s=Convert.ToString(kv.Value);if(string.IsNullOrWhiteSpace(s)||s.Length>64)throw new ArgumentException("字体名称无效。");foreach(char c in s)if(char.IsControl(c))throw new ArgumentException("字体名称不能包含控制字符。");}
                else if(kv.Key=="bold"||kv.Key=="italic") {if(!(kv.Value is bool))throw new ArgumentException("加粗/斜体应为布尔值。");}
                else if(kv.Key=="alignment") {string s=Convert.ToString(kv.Value);if(s!="left"&&s!="center"&&s!="right"&&s!="justify")throw new ArgumentException("对齐方式无效。");}
                else
                {
                    double min=0,max=200;
                    if(kv.Key=="fontSize"){min=5;max=96;}
                    else if(kv.Key=="lineSpacing"){min=.5;max=5;}
                    else if(kv.Key=="firstLineIndentChars"){min=-4;max=10;}
                    else if(kv.Key!="spaceBefore"&&kv.Key!="spaceAfter")throw new ArgumentException("暂不支持该格式字段："+kv.Key);
                    if(kv.Value==null||kv.Value is bool)throw new ArgumentException("格式数值无效："+kv.Key);
                    double n=Convert.ToDouble(kv.Value,CultureInfo.InvariantCulture);if(double.IsNaN(n)||double.IsInfinity(n)||n<min||n>max)throw new ArgumentException("格式数值超出范围："+kv.Key);
                }
            }
        }
        private static void Assignments(StringBuilder code,Dictionary<string,object> p,string pad)
        {
            foreach(var kv in p)
            {
                string field,value;
                switch(kv.Key)
                {
                    case "fontName": field="Font.Name";value="\""+Convert.ToString(kv.Value).Replace("\"","\"\"")+"\"";break;
                    case "fontSize":field="Font.Size";value=Number(kv.Value);break;
                    case "bold":field="Font.Bold";value=(bool)kv.Value?"-1":"0";break;
                    case "italic":field="Font.Italic";value=(bool)kv.Value?"-1":"0";break;
                    case "alignment":field="ParagraphFormat.Alignment";value=Array.IndexOf(new[]{"left","center","right","justify"},Convert.ToString(kv.Value)).ToString();break;
                    case "lineSpacing":
                        Set(code,pad,"ParagraphFormat.DisableLineHeightGrid","True");Set(code,pad,"ParagraphFormat.LineSpacingRule","5");Set(code,pad,"ParagraphFormat.LineSpacing",Number(Convert.ToDouble(kv.Value,CultureInfo.InvariantCulture)*12));continue;
                    case "spaceBefore":Set(code,pad,"ParagraphFormat.SpaceBeforeAuto","0");field="ParagraphFormat.SpaceBefore";value=Number(kv.Value);break;
                    case "spaceAfter":Set(code,pad,"ParagraphFormat.SpaceAfterAuto","0");field="ParagraphFormat.SpaceAfter";value=Number(kv.Value);break;
                    default:field="ParagraphFormat.CharacterUnitFirstLineIndent";value=Number(kv.Value);break;
                }
                Set(code,pad,field,value);
            }
        }
        private static string Number(object value)=>Convert.ToDouble(value,CultureInfo.InvariantCulture).ToString("0.####",CultureInfo.InvariantCulture);
        private static void Set(StringBuilder code,string pad,string field,string value)
        {code.AppendLine(pad+"If r."+field+" <> "+value+" Then\n"+pad+"  r."+field+" = "+value+"\n"+pad+"  changed = True\n"+pad+"End If");}
    }
}
