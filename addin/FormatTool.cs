using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Web.Script.Serialization;

namespace OfficeAi
{
    // Read-only preparation freezes actual targets; writing runs later on the Office STA.
    internal static class FormatTool
    {
        internal sealed class Target
        {
            internal object Range;
            internal string Level, Text;
        }
        internal sealed class Session
        {
            internal object Document;
            internal Connect.HostKind Kind;
            internal string OriginalText, Scope;
            internal readonly List<Target> Targets = new List<Target>();
            internal readonly Dictionary<string, Dictionary<string, object>> Styles = new Dictionary<string, Dictionary<string, object>>();
            internal string Token = Guid.NewGuid().ToString("N");
            internal DateTime Created = DateTime.UtcNow;
            internal string Preview()
            {
                var counts = new Dictionary<string, int> { {"title",0},{"subtitle",0},{"body",0} };
                foreach (var t in Targets) counts[t.Level]++;
                string unit = Kind == Connect.HostKind.Word ? "个非空段落" : Kind == Connect.HostKind.Excel ? "个单元格区域" : "个文本区域";
                return Json(new { ok=true, token=Token, count=Targets.Count, counts,
                    message="已锁定 " + Targets.Count + " " + unit + "。检查下面的参数后再应用。",
                    undo=Kind == Connect.HostKind.Word ? "可用 Ctrl+Z 一次撤销本次排版" : "仅调整已选区域的格式",
                    first=Targets[0].Text.Length>60 ? Targets[0].Text.Substring(0,60)+"…" : Targets[0].Text });
            }
        }
        internal static Session Prepare(object app, Connect.HostKind kind, Dictionary<string, object> request)
        {
            if (app == null) throw new InvalidOperationException("没有可用的 Office 文档。");
            var session = new Session { Kind=kind, Scope=Get(request,"scope","selection") };
            string strategy=Get(request,"strategy","structure"), level=Get(request,"level","body");
            if(strategy!="structure" && strategy!="uniform") throw new ArgumentException("请选择排版方式。");
            if(level!="title" && level!="subtitle" && level!="body") throw new ArgumentException("样式类型无效。");
            var items=request.ContainsKey("items") ? request["items"] as object[] : null;
            if(items==null || items.Length==0 || items.Length>3) throw new ArgumentException("排版参数不完整。");
            foreach(var item in items)
            {
                var style=item as Dictionary<string,object>;
                if(style==null) throw new ArgumentException("样式格式无效。");
                string name=Get(style,"level","");
                if(name!="title" && name!="subtitle" && name!="body") throw new ArgumentException("样式类型无效。");
                Validate(style,"fontSize",5,96); Validate(style,"spaceBefore",0,200);
                Validate(style,"spaceAfter",0,200); Validate(style,"lineSpacing",0.5,5);
                if(!style.ContainsKey("bold") || !(style["bold"] is bool)) throw new ArgumentException("加粗参数无效。");
                if(session.Styles.ContainsKey(name)) throw new ArgumentException("样式类型重复。");
                session.Styles[name]=style;
            }
            if(strategy=="structure" && kind!=Connect.HostKind.Word) throw new ArgumentException("按标题层级排版仅适用于 Word。请选择统一样式。");
            dynamic a=app;
            if(kind==Connect.HostKind.Word)
            {
                dynamic doc=a.ActiveDocument;
                if((bool)doc.ReadOnly || (int)doc.ProtectionType != -1) throw new InvalidOperationException("文档为只读或受保护，未进行排版。");
                session.Document=doc; session.OriginalText=(string)doc.Content.Text;
                dynamic range;
                if(session.Scope=="document") range=doc.Content.Duplicate;
                else if(session.Scope=="paragraph") range=a.Selection.Paragraphs[1].Range.Duplicate;
                else if(session.Scope=="selection")
                {
                    if((int)a.Selection.Start==(int)a.Selection.End) throw new InvalidOperationException("没有选中文字。先在文档中选中段落，或把范围改为“光标所在段落 / 整篇文档”。");
                    range=a.Selection.Range.Duplicate;
                }
                else throw new ArgumentException("排版范围无效。");
                if((int)range.StoryType!=1) throw new InvalidOperationException("本次仅支持文档正文，页眉、页脚和文本框请单独处理。");
                int count=(int)range.Paragraphs.Count;
                if(count>2000) throw new InvalidOperationException("本次超过 2000 段，请分段处理。");
                // Snapshot the collection before any font or layout mutation.
                int end=(int)range.End;
                for(int i=1;i<=count;i++)
                {
                    dynamic paragraph=range.Paragraphs[i]; dynamic r=paragraph.Range.Duplicate;
                    if(session.Scope=="selection" && (int)r.Start>=end) continue;
                    string text=(string)r.Text;
                    if(string.IsNullOrWhiteSpace(text.Replace("\a",""))) continue;
                    string role=strategy=="uniform" ? level : WordLevel(doc,r);
                    session.Targets.Add(new Target { Range=(object)r, Level=role, Text=text.Trim() });
                }
            }
            else if(kind==Connect.HostKind.Excel)
            {
                if(session.Scope!="selection") throw new ArgumentException("Excel 请使用“当前选区”。");
                session.Document=a.ActiveWorkbook;
                dynamic r=a.Selection;
                if((double)r.CountLarge>20000 || (int)r.Rows.Count>2000) throw new InvalidOperationException("选区过大，请选择实际需要排版的单元格。");
                if((int)r.Areas.Count!=1) throw new InvalidOperationException("请先选择一个连续的单元格区域。");
                if((bool)r.Worksheet.ProtectContents) throw new InvalidOperationException("工作表已保护，未进行排版。");
                session.Targets.Add(new Target {Range=(object)r,Level=level,Text=(string)r.Worksheet.Name+"!"+(string)r.Address});
            }
            else if(kind==Connect.HostKind.PowerPoint)
            {
                if(session.Scope!="selection") throw new ArgumentException("PowerPoint 请使用“当前选区”。");
                session.Document=a.ActivePresentation;
                dynamic selection=a.ActiveWindow.Selection;
                int type=(int)selection.Type;
                if(type==3) AddPptTarget(session,selection.TextRange,level);
                else if(type==2)
                {
                    int count=(int)selection.ShapeRange.Count;
                    if(count>100) throw new InvalidOperationException("一次最多处理 100 个文本框。");
                    for(int i=1;i<=count;i++)
                    {
                        dynamic shape=selection.ShapeRange[i];
                        if((int)shape.HasTextFrame==-1 && (int)shape.TextFrame.HasText==-1) AddPptTarget(session,shape.TextFrame.TextRange,level);
                    }
                }
                else throw new InvalidOperationException("请先选中幻灯片文字或含文字的文本框。");
            }
            else throw new InvalidOperationException("当前宿主不支持排版。");
            if(session.Targets.Count==0) throw new InvalidOperationException("范围内没有可排版的文字。");
            foreach(var t in session.Targets) if(!session.Styles.ContainsKey(t.Level)) throw new ArgumentException("缺少 "+t.Level+" 样式参数。");
            return session;
        }
        private static void AddPptTarget(Session session, object range, string level)
        {
            dynamic r=range; string text=(string)r.Text;
            if(!string.IsNullOrWhiteSpace(text)) session.Targets.Add(new Target { Range=range, Level=level, Text=text });
        }
        private static string WordLevel(object document, object range)
        {
            dynamic d=document, r=range;
            string style=(string)r.Style.NameLocal;
            if(style==(string)d.Styles[-63].NameLocal) return "title"; // wdStyleTitle
            if(style==(string)d.Styles[-75].NameLocal) return "subtitle"; // wdStyleSubtitle
            int outline=(int)r.ParagraphFormat.OutlineLevel;
            return outline==1 ? "title" : outline>=2 && outline<=9 ? "subtitle" : "body";
        }
        internal static string Execute(object app, Session session)
        {
            if(session==null) return Fail("请先预览并锁定排版范围。");
            bool record=false, screenWas=true, screenTouched=false;
            string before=null;
            dynamic a=app;
            try
            {
                if(DateTime.UtcNow-session.Created>TimeSpan.FromMinutes(10)) return Fail("预览已过期，请重新预览。");
                object current=session.Kind==Connect.HostKind.Word ? (object)a.ActiveDocument : session.Kind==Connect.HostKind.Excel ? (object)a.ActiveWorkbook : (object)a.ActivePresentation;
                if(!Equals(current,session.Document)) return Fail("文档已切换，请回到原文档或重新预览。");
                if(session.Kind==Connect.HostKind.Word)
                {
                    dynamic doc=session.Document;
                    if((string)doc.Content.Text!=session.OriginalText) return Fail("原文已变化，请重新预览。");
                    if((bool)doc.ReadOnly || (int)doc.ProtectionType!=-1) return Fail("文档为只读或受保护。");
                    before=Fingerprint(session);
                    screenWas=(bool)a.ScreenUpdating; screenTouched=true;
                    a.UndoRecord.StartCustomRecord("OfficeAI 排版"); record=true;
                    a.ScreenUpdating=false;
                }
                foreach(var target in session.Targets)
                {
                    var style=session.Styles[target.Level];
                    if(session.Kind==Connect.HostKind.Word) ApplyWord(target.Range,style);
                    else if(session.Kind==Connect.HostKind.Excel) ApplyExcel(target.Range,style);
                    else ApplyPpt(target.Range,style);
                }
                if(session.Kind==Connect.HostKind.Word && (string)((dynamic)session.Document).Content.Text!=session.OriginalText)
                    throw new InvalidOperationException("排版期间检测到正文变化，已中止。");
                string unit=session.Kind==Connect.HostKind.Word ? "个段落" : session.Kind==Connect.HostKind.Excel ? "个单元格区域" : "个文本区域";
                return Json(new {ok=true,count=session.Targets.Count,message="已排版 "+session.Targets.Count+" "+unit+(session.Kind==Connect.HostKind.Word ? "，可用 Ctrl+Z 一次撤销。" : "。")});
            }
            catch(Exception ex)
            {
                Connect.Log("FormatTool.Execute 异常: "+ex);
                if(record)
                {
                    try
                    {
                        a.UndoRecord.EndCustomRecord();record=false;
                        if(before!=Fingerprint(session))
                        {
                            bool undone=((dynamic)session.Document).Undo(1);
                            return Fail("排版失败："+ex.Message+(undone && before==Fingerprint(session) ? "；本次格式修改已撤回。" : "；请检查文档格式。"));
                        }
                    }
                    catch(Exception undoError) {Connect.Log("排版回退失败："+undoError.Message);}
                }
                return Fail("排版失败："+ex.Message);
            }
            finally
            {
                if(record) try {a.UndoRecord.EndCustomRecord();} catch(Exception ex) {Connect.Log("结束排版记录失败："+ex.Message);}
                if(screenTouched) try {a.ScreenUpdating=screenWas;} catch(Exception ex) {Connect.Log("恢复屏幕刷新失败："+ex.Message);}
            }
        }
        private static void ApplyWord(object range, Dictionary<string,object> style)
        {
            dynamic r=range, f=r.Font, p=r.ParagraphFormat;
            f.Size=N(style,"fontSize"); f.Bold=(bool)style["bold"] ? -1 : 0;
            p.SpaceBeforeAuto=0; p.SpaceAfterAuto=0;
            p.SpaceBefore=N(style,"spaceBefore"); p.SpaceAfter=N(style,"spaceAfter");
            p.DisableLineHeightGrid=true;
            p.LineSpacingRule=5; // wdLineSpaceMultiple (4 is wdLineSpaceExactly)
            p.LineSpacing=N(style,"lineSpacing")*12;
        }
        private static void ApplyExcel(object range, Dictionary<string,object> style)
        {
            dynamic r=range;
            r.Font.Size=N(style,"fontSize");r.Font.Bold=(bool)style["bold"];r.WrapText=true;
            r.Rows.RowHeight=Math.Min(409.0f,N(style,"fontSize")*1.5f*N(style,"lineSpacing"));
        }
        private static void ApplyPpt(object range, Dictionary<string,object> style)
        {
            dynamic r=range;
            r.Font.Size=N(style,"fontSize");r.Font.Bold=(bool)style["bold"] ? -1 : 0;
            dynamic p=r.ParagraphFormat;
            p.LineRuleWithin=-1; p.SpaceWithin=N(style,"lineSpacing");
            p.LineRuleBefore=0; p.SpaceBefore=N(style,"spaceBefore");
            p.LineRuleAfter=0; p.SpaceAfter=N(style,"spaceAfter");
        }
        private static string Fingerprint(Session session)
        {
            var values=new List<object>();
            foreach(var target in session.Targets)
            {
                dynamic r=target.Range,p=r.ParagraphFormat,f=r.Font;
                values.Add(new object[] {(object)f.Size,(object)f.Bold,(object)p.SpaceBeforeAuto,(object)p.SpaceAfterAuto,
                    (object)p.SpaceBefore,(object)p.SpaceAfter,(object)p.LineSpacingRule,(object)p.LineSpacing,(object)p.DisableLineHeightGrid});
            }
            return Json(values);
        }
        private static string Get(Dictionary<string,object> d,string name,string fallback)=>d.ContainsKey(name) ? Convert.ToString(d[name]) : fallback;
        private static float N(Dictionary<string,object> d,string name)=>Convert.ToSingle(d[name],CultureInfo.InvariantCulture);
        private static void Validate(Dictionary<string,object> style,string name,double min,double max)
        {
            if(!style.ContainsKey(name) || style[name]==null || style[name] is bool) throw new ArgumentException("缺少排版参数："+name);
            double value=Convert.ToDouble(style[name],CultureInfo.InvariantCulture);
            if(double.IsNaN(value)||double.IsInfinity(value)||value<min||value>max) throw new ArgumentException("排版参数超出范围："+name);
        }
        internal static string Json(object value)=>new JavaScriptSerializer {MaxJsonLength=4194304}.Serialize(value);
        internal static string Fail(string message)=>Json(new {ok=false,message});
    }
}
