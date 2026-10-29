using System;using System.IO;using System.Text;using System.Runtime.InteropServices;using System.Collections.Generic;using System.Web.Script.Serialization;
namespace OfficeAi
{
 internal static class Program
 {
  private delegate bool EnumProc(IntPtr hwnd,IntPtr lParam);
  [DllImport("user32.dll")]private static extern bool EnumChildWindows(IntPtr hwnd,EnumProc callback,IntPtr param);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)]private static extern int GetClassName(IntPtr hwnd,StringBuilder name,int max);
  [DllImport("user32.dll")]private static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
  [DllImport("oleacc.dll")]private static extern int AccessibleObjectFromWindow(IntPtr hwnd,uint id,ref Guid iid,[MarshalAs(UnmanagedType.IDispatch)]out object value);
  [STAThread]private static int Main(string[] args)
  {
   if(args.Length!=2)return 2;
   object result;dynamic project=null,component=null;bool ran=false;string phase="读取任务";
   try
   {
    var request=(Dictionary<string,object>)new JavaScriptSerializer {MaxJsonLength=4194304}.DeserializeObject(File.ReadAllText(args[0],Encoding.UTF8));
    IntPtr hwnd=new IntPtr(Convert.ToInt64(request["hwnd"]));uint pid;GetWindowThreadProcessId(hwnd,out pid);
    if(pid!=Convert.ToUInt32(request["pid"]))throw new InvalidOperationException("原 Office 窗口已关闭，请重新生成方案。");
    object native=null;phase="连接 Word 窗口";
    EnumChildWindows(hwnd,(child,param)=>{
     var name=new StringBuilder(128);GetClassName(child,name,128);if(name.ToString()!="_WwG")return true;
     Guid iid=new Guid("00020400-0000-0000-C000-000000000046");object obj;
     if(AccessibleObjectFromWindow(child,unchecked((uint)-16),ref iid,out obj)==0 && obj!=null){native=obj;return false;}return true;
    },IntPtr.Zero);
    if(native==null)throw new InvalidOperationException("无法连接指定的 Word 文档窗口。");
    phase="获取 Word 应用对象";dynamic app=((dynamic)native).Application;dynamic doc=null;
    string nameExpected=Convert.ToString(request["document"]);
    phase="定位目标文档";for(int i=1;i<=(int)app.Documents.Count;i++)if((string)app.Documents[i].FullName==nameExpected){doc=app.Documents[i];break;}
    if(doc==null)throw new InvalidOperationException("原文档已关闭。");
    if((string)app.ActiveDocument.FullName!=nameExpected)throw new InvalidOperationException("活动文档已切换，请回到原文档重新生成方案。");
    string original=(string)doc.Content.Text;
    if(VbaFormatPlan.Hash(original)!=Convert.ToString(request["hash"]))throw new InvalidOperationException("原文已变化，请重新生成方案。");
    if((bool)doc.ReadOnly || (int)doc.ProtectionType!=-1)throw new InvalidOperationException("文档为只读或受保护。");
    var ops=(object[])request["ops"];
    var plan=VbaFormatPlan.Build(ops,original,Convert.ToInt32(request["start"]),Convert.ToInt32(request["end"]));
    phase="访问文档 VBA 工程";try {project=doc.VBProject;}catch(Exception ex){throw new InvalidOperationException("Word 未允许 VBA 工程访问，尚未修改文档。"+ex.Message);}
    string module="OfficeAI_"+Guid.NewGuid().ToString("N").Substring(0,8);
    phase="创建临时 VBA 模块";component=project.VBComponents.Add(1);component.Name=module;component.CodeModule.AddFromString(plan.Code);
    phase="运行 VBA";
    app.Run(module+".OfficeAI_Apply",doc);
    ran=true;
    if((string)doc.Content.Text!=original)throw new InvalidOperationException("格式执行后正文发生变化，请检查文档。");
    result=new {ok=true,method="vba",message="已通过 VBA 执行 "+plan.Count+" 个格式操作，可用 Ctrl+Z 一次撤销。"};
   }
   catch(Exception ex){result=new {ok=false,method="vba",message=phase+"："+ex.Message};}
   finally
   {
    if(project!=null && component!=null)try {project.VBComponents.Remove(component);}catch(Exception ex){File.WriteAllText(args[1]+".cleanup-warning.txt","临时 VBA 模块未能清理："+ex.Message,Encoding.UTF8);}
   }
   File.WriteAllText(args[1],VbaFormatPlan.Json(result),new UTF8Encoding(false));
   return ran?0:1;
  }
 }
}
