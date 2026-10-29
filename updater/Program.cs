using System;using System.IO;using System.Drawing;using System.Reflection;using System.Windows.Forms;using System.Diagnostics;using System.Security.Principal;using System.Security.Cryptography;using System.Collections.Generic;using System.Web.Script.Serialization;using System.Threading.Tasks;
class UpdateForm:Form
{
 static readonly string[] Files={"OfficeAi.Addin.dll","OfficeAi.Client.exe","OfficeAi.VbaWorker.exe","ui\\app.js","ui\\ima-auto.js","ui\\index.html","ui\\style.css"};
 readonly Label status=new Label();readonly ProgressBar progress=new ProgressBar();readonly string resultPath;
 sealed class Change {internal string Target,Backup,Hash,Stage;internal bool Moved,Written;}
 UpdateForm(string output)
 {
  resultPath=output;Text="OfficeAI 0.5.2 更新";ClientSize=new Size(480,180);StartPosition=FormStartPosition.CenterScreen;FormBorderStyle=FormBorderStyle.FixedDialog;MaximizeBox=false;BackColor=Color.FromArgb(17,21,25);ForeColor=Color.White;Font=new Font("Microsoft YaHei UI",10);
  status.SetBounds(24,24,432,75);status.Text="正在更新实际加载的 OfficeAI…\n现有文件将保留备份，不会关闭你的 Office。";
  progress.SetBounds(24,120,432,20);Controls.Add(status);Controls.Add(progress);
  Shown+=async(sender,e)=>{object result=await Task.Run(()=>Install((text,percent)=>BeginInvoke(new Action(()=>{status.Text=text;progress.Value=percent;}))));
   Directory.CreateDirectory(Path.GetDirectoryName(resultPath));File.WriteAllText(resultPath,Json(result),new System.Text.UTF8Encoding(false));
   bool ok=(bool)result.GetType().GetProperty("ok").GetValue(result);status.Text=ok?"更新成功。\n保存文件并完全退出 Office 后，重新打开即可。":"更新未完成，请查看结果说明。\n原版文件已尽量恢复，请保存文件后重开更新器。";
   var close=new Button{Text="关闭",Left=370,Top=142,Width=85};close.Click+=(s,a)=>Close();Controls.Add(close); if(ok){var finishTimer=new System.Windows.Forms.Timer{Interval=2500};finishTimer.Tick+=(s,a)=>{finishTimer.Stop();Close();};finishTimer.Start();}
  };
 }
 static string Json(object obj)=>new JavaScriptSerializer().Serialize(obj);
 static byte[] Payload(string file)
 {using(var stream=Assembly.GetExecutingAssembly().GetManifestResourceStream("Payload."+file.Replace('\\','.')))using(var ms=new MemoryStream()){if(stream==null)throw new IOException("更新文件缺失："+file);stream.CopyTo(ms);return ms.ToArray();}}
 static string Hash(byte[] bytes){using(var hash=SHA256.Create())return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-","");}
 static string Under(string root,string relative)
 {string resolved=Path.GetFullPath(Path.Combine(root,relative));if(!resolved.StartsWith(Path.GetFullPath(root).TrimEnd('\\')+"\\",StringComparison.OrdinalIgnoreCase))throw new IOException("更新路径超出目标目录。");return resolved;}
 static object Install(Action<string,int> report)
 {
  var changed=new List<Change>();string root=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),"OfficeAi");
  string backup=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),"OfficeAi","backups",DateTime.Now.ToString("yyyyMMdd-HHmmss")+"-"+Guid.NewGuid().ToString("N").Substring(0,6));
  try
  {
   if(!Directory.Exists(root))throw new IOException("找不到原安装目录："+root);
   for(int i=0;i<Files.Length;i++)
   {
    string relative=Files[i];byte[] bytes=Payload(relative);string target=Under(root,relative),saved=Under(backup,relative);
    Directory.CreateDirectory(Path.GetDirectoryName(target));Directory.CreateDirectory(Path.GetDirectoryName(saved));
    var item=new Change{Target=target,Backup=saved,Hash=Hash(bytes),Stage=target+".stage-"+Guid.NewGuid().ToString("N")};changed.Add(item);
    report("正在更新 "+relative+"\n原文件保存在："+backup,(i*100)/Files.Length);
    File.WriteAllBytes(item.Stage,bytes);
    if(File.Exists(target)){File.Move(target,saved);item.Moved=true;}
    File.Move(item.Stage,target);item.Written=true;
    if(Hash(File.ReadAllBytes(target))!=item.Hash)throw new IOException("文件校验失败："+relative);
   }
   report("全部文件校验通过。",100);
   return new {ok=true,version="0.5.2",root,backup,files=Files,message="系统安装目录已更新。旧进程仍需保存文件后完全退出，再重新打开。"};
  }
  catch(Exception ex)
  {
   string rollback="";
   for(int i=changed.Count-1;i>=0;i--)
   {
    var item=changed[i];try{
     if(item.Written && File.Exists(item.Target) && Hash(File.ReadAllBytes(item.Target))==item.Hash)File.Delete(item.Target);
     if(item.Moved && File.Exists(item.Backup) && !File.Exists(item.Target))File.Move(item.Backup,item.Target);
    }catch(Exception error){rollback+=" 恢复失败："+error.Message;}
   }
   return new {ok=false,version="0.5.2",root,backup,message=ex.Message+rollback};
  }
  finally {foreach(var item in changed)try{if(File.Exists(item.Stage))File.Delete(item.Stage);}catch{}}
 }
 [STAThread]static void Main(string[] args)
 {
  string output=args.Length>1?Path.GetFullPath(args[1]):Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"OfficeAi","update-result.json");
  if(args.Length>0 && args[0]=="--verify")
  {
   var hashes=new Dictionary<string,string>();foreach(var file in Files)hashes[file]=Hash(Payload(file));Directory.CreateDirectory(Path.GetDirectoryName(output));File.WriteAllText(output,Json(new {ok=true,version="0.5.2",hashes}));return;
  }
  bool elevated=new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
  if(!elevated)
  {
   try {Process.Start(new ProcessStartInfo(Application.ExecutablePath,"--apply \""+output+"\""){UseShellExecute=true,Verb="runas",WindowStyle=ProcessWindowStyle.Hidden});}
   catch(Exception ex){Directory.CreateDirectory(Path.GetDirectoryName(output));File.WriteAllText(output,Json(new {ok=false,message="Windows 未允许管理员更新："+ex.Message}));}
   return;
  }
  bool created;
  using(var mutex=new System.Threading.Mutex(true,"Global\\OfficeAI_Update_052",out created))
  {
   if(!created){Directory.CreateDirectory(Path.GetDirectoryName(output));File.WriteAllText(output,Json(new{ok=false,message="另一个 OfficeAI 更新器正在运行，请等待它完成。"}));return;}
   Application.EnableVisualStyles();Application.SetCompatibleTextRenderingDefault(false);Application.Run(new UpdateForm(output));
  }
 }
}

