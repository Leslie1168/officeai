using System;using System.IO;using System.Drawing;using System.Reflection;using System.Windows.Forms;using System.Runtime.InteropServices;using System.Diagnostics;using System.Threading;using System.Text;using OfficeAi;
internal static class Program
{
 private delegate bool WindowProc(IntPtr hwnd,IntPtr param);
 [DllImport("user32.dll")]static extern bool EnumChildWindows(IntPtr parent,WindowProc callback,IntPtr param);
 [DllImport("user32.dll")]static extern bool EnumWindows(WindowProc callback,IntPtr param);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int GetClassName(IntPtr hwnd,StringBuilder value,int length);
 [DllImport("user32.dll")]static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern bool SetProp(IntPtr hwnd,string name,IntPtr value);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern IntPtr GetProp(IntPtr hwnd,string name);
 [DllImport("user32.dll")]static extern bool ShowWindow(IntPtr hwnd,int cmd);
 [DllImport("user32.dll")]static extern bool SetForegroundWindow(IntPtr hwnd);
 [DllImport("oleacc.dll")]static extern int AccessibleObjectFromWindow(IntPtr hwnd,uint id,ref Guid iid,[MarshalAs(UnmanagedType.IDispatch)]out object obj);
 [STAThread]static void Main(string[] args)
 {
  long window;int hostPid;
  if(args.Length!=2 || !long.TryParse(args[0],out window) || !int.TryParse(args[1],out hostPid))return;
  bool created;
  using(var instance=new Mutex(true,"Local\\OfficeAi.Client."+hostPid,out created))
  {
   if(!created){EnumWindows((hwnd,p)=>{if(GetProp(hwnd,"OfficeAI.HostPid").ToInt64()!=hostPid)return true;ShowWindow(hwnd,9);SetForegroundWindow(hwnd);return false;},IntPtr.Zero);return;}
   try
   {
    IntPtr parent=new IntPtr(window);uint pid;GetWindowThreadProcessId(parent,out pid);
    if(pid!=hostPid)throw new InvalidOperationException("原 Word 窗口已关闭，请从 Word 重新打开助手。");
    object native=null;
    EnumChildWindows(parent,(child,p)=>{var name=new StringBuilder(128);GetClassName(child,name,128);if(name.ToString()!="_WwG")return true;Guid iid=new Guid("00020400-0000-0000-C000-000000000046");object value;if(AccessibleObjectFromWindow(child,unchecked((uint)-16),ref iid,out value)==0 && value!=null){native=value;return false;}return true;},IntPtr.Zero);
    if(native==null)throw new InvalidOperationException("请先在 Word 中打开文档，再打开助手。");
    dynamic app=((dynamic)native).Application;
    var flags=BindingFlags.Static|BindingFlags.NonPublic;var type=typeof(Connect);
    type.GetField("CurrentApp",flags).SetValue(null,(object)app);type.GetField("Kind",flags).SetValue(null,Connect.HostKind.Word);type.GetField("CurrentHost",flags).SetValue(null,"Word");
    Application.EnableVisualStyles();Application.SetCompatibleTextRenderingDefault(false);
    var form=new Form {Text="OfficeAI · Word",Size=new Size(470,840),MinimumSize=new Size(360,540),BackColor=Color.FromArgb(17,21,25),StartPosition=FormStartPosition.Manual};
    var area=Screen.PrimaryScreen.WorkingArea;form.Location=new Point(Math.Max(area.Left,area.Right-form.Width-16),area.Top+20);form.Height=Math.Min(form.Height,area.Height-40);
    var host=new TaskPaneHost {Dock=DockStyle.Fill};form.Controls.Add(host);
    form.Shown+=(sender,e)=>SetProp(form.Handle,"OfficeAI.HostPid",new IntPtr(hostPid));
    var timer=new System.Windows.Forms.Timer {Interval=2000};
    timer.Tick+=(sender,e)=>{try{if(Process.GetProcessById(hostPid).HasExited || (int)app.Documents.Count==0)form.Close();}catch(ArgumentException){form.Close();}catch{}};
    timer.Start();form.FormClosed+=(sender,e)=>{timer.Stop();timer.Dispose();};
    Application.Run(form);
   }
   catch(Exception ex){MessageBox.Show(ex.Message,"OfficeAI",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
  }
 }
}
