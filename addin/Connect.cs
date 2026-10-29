using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Office = Microsoft.Office.Core;

namespace OfficeAi
{
    /// <summary>
    /// 加载项入口。同时实现：
    ///   Extensibility.IDTExtensibility2   —— COM 加载项生命周期
    ///   Office.IRibbonExtensibility       —— 功能区
    ///   Office.ICustomTaskPaneConsumer    —— 拿到 ICTPFactory 以创建任务窗格
    /// ★ 类特性只能用 AutoDispatch，绝不能用 AutoDual（会让宿主 CLR 崩溃 0x80131506）。
    /// </summary>
    [ComVisible(true)]
    [Guid("7C3E9A21-4B6D-4F58-9E02-3A1B5C7D8E90")]
    [ProgId("OfficeAi.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, Office.IRibbonExtensibility,
                           Office.ICustomTaskPaneConsumer
    {
        private const string RibbonResourceName = "OfficeAi.Ribbon.xml";
        public const string Version = "0.5.2";

        public enum HostKind { Unknown = 0, Word = 1, Excel = 2, PowerPoint = 3 }

        internal static object CurrentApp;
        internal static HostKind Kind = HostKind.Unknown;
        internal static string CurrentHost = "?";

        private object _app;
        private Office.ICTPFactory _ctpFactory;
        private Office.CustomTaskPane _pane;
        private bool _paneVisible;
        private Process _clientProcess;

        /// <summary>
        /// 用「能力探测」认宿主：MS Office 与 WPS 的对象模型一致，
        /// 靠 Application.Name 不可靠（WPS 各组件名字各异），
        /// 而 Presentations / Workbooks / Documents 这三个集合在各自宿主里必定存在。
        /// 顺序很重要：先判 PowerPoint、再 Excel、最后 Word。
        /// </summary>
        internal static HostKind DetectKind(object app)
        {
            if (app == null) return HostKind.Unknown;
            dynamic a = app;
            if (Probe(() => a.Presentations)) return HostKind.PowerPoint;
            if (Probe(() => a.Workbooks)) return HostKind.Excel;
            if (Probe(() => a.Documents)) return HostKind.Word;
            return HostKind.Unknown;
        }

        private static bool Probe(Func<object> f)
        {
            try { return f() != null; } catch { return false; }
        }

        #region 日志

        /// <summary>
        /// 全项目统一日志入口（签名不变，Bridge.cs 等调用方零改动）。
        /// 实现转到 CrashLog.Log：改用**无缓冲 FileStream**（FILE_APPEND_DATA，一次 WriteFile 一行），
        /// 不再用 File.AppendAllText —— 后者每行都要 open/close，且数据先落在托管缓冲里，
        /// 进程被 AV 杀掉时最后一段会丢。格式保持不变：HH:mm:ss.fff [t线程号]  消息。
        /// </summary>
        internal static void Log(string msg)
        {
            CrashLog.Log(msg);
        }

        public Connect()
        {
            // ★ 取证装配点（T2）：VEH(first, 不可被覆盖) + 全内存 dump 线程 + 进程外看门狗 + WebView2 进程监视。
            // 必须最早安装：后面任何一行代码都可能触发 AV，先装才抓得到。
            if (Settings.GetBool("nativeDiagnostics", false)) CrashGuard.Install();
            Log("ctor v" + Version);
            LogStartupTimeline();
        }

        /// <summary>
        /// 把「这一次是哪个部署版本、进程是谁、取证开关是什么」写进崩溃时间线
        /// （addin.log + crash.log，后者 Flush(true) 落盘）。
        /// 为什么必须记：dump 拿回来第一件事就是归因——是哪个 DLL、哪个进程、什么配置下崩的。
        /// </summary>
        private static void LogStartupTimeline()
        {
            try
            {
                string dll = "?";
                string dllTime = "?";
                try
                {
                    dll = Assembly.GetExecutingAssembly().Location;
                    dllTime = File.GetLastWriteTime(dll).ToString("yyyy-MM-dd HH:mm:ss");
                }
                catch { }
                int pid = 0;
                try { pid = System.Diagnostics.Process.GetCurrentProcess().Id; } catch { }

                CrashLog.Timeline("Connect ctor v" + Version +
                                  " addinVer=" + Assembly.GetExecutingAssembly().GetName().Version +
                                  " pid=" + pid + " dll=" + dll + " dllWriteTime=" + dllTime);
                CrashLog.Timeline("settings: watchdog=" + Settings.GetBool("watchdog", true) +
                                  " crashDumpAllAv=" + Settings.GetBool("crashDumpAllAv", false) +
                                  " autoOpenPane=" + Settings.GetBool("autoOpenPane", true) +
                                  " theme=" + Settings.Get("theme", "light"));
            }
            catch (Exception e) { Log("LogStartupTimeline 失败: " + e.Message); }
        }

        #endregion

        #region IDTExtensibility2

        public void OnConnection(object Application, ext_ConnectMode ConnectMode,
                                 object AddInInst, ref Array custom)
        {
            _app = Application;
            CurrentApp = Application;
            try
            {
                string nm = "";
                try { dynamic d = Application; nm = (string)d.Name; } catch { }

                Kind = DetectKind(Application);
                CurrentHost = Kind == HostKind.Unknown
                    ? (string.IsNullOrEmpty(nm) ? "?" : nm)
                    : Kind.ToString();
                Log("OnConnection kind=" + Kind + " Name=" + nm);
                // 崩前时间线的锚点：加载项接入宿主的确切时刻与身份
                CrashLog.Timeline("OnConnection kind=" + Kind + " Name=" + nm +
                                  " host=" + CurrentHost + " connectMode=" + ConnectMode);

                // 第5条：打开文件时自动开启侧边栏。
                // 用定时器监测「活动文档从无到有」，比订阅各宿主 COM 事件更跨宿主、更稳。
                if (Settings.GetBool("autoOpenPane", true))
                    StartDocWatcher();
            }
            catch (Exception e)
            {
                Log("OnConnection 异常: " + e.Message);
            }
        }

        private System.Windows.Forms.Timer _docWatcher;
        private bool _hadDoc;

        private void StartDocWatcher()
        {
            if (_docWatcher != null) return;
            _docWatcher = new System.Windows.Forms.Timer { Interval = 1200 };
            _docWatcher.Tick += (s, e) =>
            {
                try
                {
                    if (_app == null) return;
                    // 面板已经显示出来，就不需要再周期性访问 Word 对象模型了。
                    // 定时器里反复碰 COM 会与 WebView2 的回调互相重入，把宿主搞崩
                    // （崩在 wwlib.dll / clr!GetCCWFromIUnknown）。面板可见时暂停轮询。
                    if (_paneVisible) { PauseDocWatcher(); return; }
                    bool has = HasActiveDocument();
                    if (has && !_hadDoc)
                    {
                        // 从「没有文档」变成「有文档」——正是"打开文件"的时机
                        _hadDoc = true;
                        if (_ctpFactory != null || Kind == HostKind.Word)
                        {
                            Log("检测到打开文件，自动开启侧边栏");
                            ShowPane(true);
                        }
                    }
                    else if (!has)
                    {
                        _hadDoc = false;
                    }
                }
                catch { }
            };
            _docWatcher.Start();
            // 初始视为“无文档”，这样即使宿主启动时就带着文档，
            // 首次 tick 检测到有文档也能触发自动开窗（覆盖“双击文件直接打开”场景）。
            _hadDoc = false;
            Log("DocWatcher 已启动");
            CrashLog.Timeline("DocWatcher 已启动（1200ms 轮询 ActiveDocument，面板可见时暂停）");
        }

        private void StopDocWatcher()
        {
            try { if (_docWatcher != null) { _docWatcher.Stop(); _docWatcher.Dispose(); } }
            catch { }
            _docWatcher = null;
            Log("DocWatcher 已停止（面板已开，不再轮询 COM）");
            CrashLog.Timeline("DocWatcher 已停止（面板已开，不再轮询 COM）");
        }

        /// <summary>面板可见时暂停轮询（保留定时器，面板收起后还能恢复）。</summary>
        private void PauseDocWatcher()
        {
            try { if (_docWatcher != null) _docWatcher.Stop(); } catch { }
            CrashLog.Timeline("DocWatcher 已暂停（面板可见，避免与 WebView2 回调互相重入）");
        }

        private bool HasActiveDocument()
        {
            try
            {
                dynamic a = _app;
                if (a == null) return false;
                switch (Kind)
                {
                    case HostKind.Word: return (bool)a.Visible && a.ActiveDocument != null;
                    case HostKind.Excel: return a.ActiveWorkbook != null;
                    case HostKind.PowerPoint: return a.ActivePresentation != null;
                }
            }
            catch { }
            return false;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Log("OnDisconnection " + RemoveMode);
            // 崩溃时间线 + WebView2 进程账本总结（谁先死：WebView2 子进程 还是 宿主）
            CrashLog.Timeline("OnDisconnection " + RemoveMode + " pane=" + (_pane != null) +
                              " paneVisible=" + _paneVisible + " docWatcher=" + (_docWatcher != null));
            CrashGuard.NotifyHostShutdown("OnDisconnection " + RemoveMode);
            try { if (_docWatcher != null) { _docWatcher.Stop(); _docWatcher.Dispose(); } }
            catch { }
            _docWatcher = null;
            // 不手动 Marshal.ReleaseComObject：把 RCW 的引用计数减穿会让
            // 宿主侧持有的接口指针变悬空，随机闪退。交给 GC/宿主自然释放。
            try { if (_pane != null) { _pane.Visible = false; } }
            catch { }
            if (_pane != null) CrashLog.Timeline("CTP 窗格已隐藏并释放（OnDisconnection）");
            _pane = null;
            _app = null;
            CurrentApp = null;
            Kind = HostKind.Unknown;
            CrashLog.Timeline("OnDisconnection 完成：_app/_pane 已置空");
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom)
        {
            Log("OnStartupComplete");
            CrashLog.Timeline("OnStartupComplete（宿主启动流程结束，加载项已就绪）");
        }
        public void OnBeginShutdown(ref Array custom)
        {
            Log("OnBeginShutdown");
            // ★ 关键时刻：宿主开始正常关闭。此时写一行「谁先死」的账本总结，
            // 用于和崩溃路径（宿主先死）区分开。
            CrashLog.Timeline("OnBeginShutdown（宿主开始正常关闭）");
            CrashGuard.NotifyHostShutdown("OnBeginShutdown");
        }

        #endregion

        #region Ribbon

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                Assembly asm = Assembly.GetExecutingAssembly();
                using (Stream s = asm.GetManifestResourceStream(RibbonResourceName))
                {
                    if (s != null)
                    {
                        using (var r = new StreamReader(s, new UTF8Encoding(false)))
                            return r.ReadToEnd();
                    }
                }
                Log("GetCustomUI: 资源缺失");
            }
            catch (Exception e)
            {
                Log("GetCustomUI 异常: " + e.Message);
            }
            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                   "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                   "<ribbon><tabs><tab id=\"OfficeAiTab\" label=\"AI 助手\">" +
                   "<group id=\"g1\" label=\"侧边栏\">" +
                   "<button id=\"btnTogglePane\" label=\"打开侧边栏\" size=\"large\" onAction=\"OnActionTogglePane\"/>" +
                   "</group></tab></tabs></ribbon></customUI>";
        }

        /// <summary>
        /// Ribbon 图标回调：按控件 Id 取同名嵌入 PNG。
        /// ★ 必须缓存，且 Bitmap / MemoryStream 绝不能 Dispose：
        ///   IPictureDisp 底层一直引用这个 Bitmap，Ribbon 每次重绘都会访问它。
        ///   早期版本用 using 把 Bitmap 提前 Dispose，导致 wwlib 重绘图标时
        ///   访问已释放的 GDI+ 对象（use-after-free），表现为 Word 随机闪退
        ///   （wwlib.dll 0xc0000005 固定偏移 0x183ba19，栈在 clr!GetCCWFromIUnknown），
        ///   在侧边栏大量渲染（抓图缩略图 / 主题切换全量重绘）后高发。
        /// </summary>
        private static readonly System.Collections.Generic.Dictionary<string, object> _iconCache =
            new System.Collections.Generic.Dictionary<string, object>();

        public object GetImage(object control)
        {
            try
            {
                string id = "";
                try { dynamic c = control; id = (string)c.Id; } catch { }
                if (string.IsNullOrEmpty(id)) return null;

                lock (_iconCache)
                {
                    object cached;
                    if (_iconCache.TryGetValue(id, out cached)) return cached;
                }

                string res = "OfficeAi.icons." + id + ".png";
                Assembly asm = Assembly.GetExecutingAssembly();
                using (Stream s = asm.GetManifestResourceStream(res))
                {
                    if (s == null) return null;
                    // 复制到常驻 MemoryStream：GDI+ 的 Bitmap(Stream) 惰性解码，
                    // 原始流必须在其生命周期内保持存活
                    var ms = new MemoryStream();
                    s.CopyTo(ms);
                    ms.Position = 0;
                    var bmp = new System.Drawing.Bitmap(ms);
                    object pic = PictureConverter.ToPictureDisp(bmp);
                    lock (_iconCache) { _iconCache[id] = pic; }
                    return pic;
                }
            }
            catch (Exception e)
            {
                Log("GetImage 失败: " + e.Message);
                return null;
            }
        }

        public void OnActionTogglePane(object control)
        {
            Log("OnActionTogglePane: 进入（CTPFactory=" + (_ctpFactory != null) + "）");
            CrashLog.Timeline("用户点击 Ribbon「打开侧边栏」（当前可见=" + _paneVisible + "）");
            try
            {
                if (Kind == HostKind.Word) { ShowPane(true); return; }
                if (_ctpFactory == null)
                {
                    MessageBox.Show(
                        "任务窗格工厂还没就绪。\n\n可能原因：宿主在加载项连接时未触发 CTPFactoryAvailable。\n" +
                        "请重启宿主程序再试；若一直如此，把 addin.log 发我。",
                        "Office AI 助手", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                ShowPane(!_paneVisible);
            }
            catch (Exception ex)
            {
                Log("OnActionTogglePane 异常: " + ex);
                MessageBox.Show("打开侧边栏失败：\n\n" + ex.Message, "Office AI 助手",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnActionAbout(object control)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Office AI 助手 · " + Version);
            sb.AppendLine();
            sb.AppendLine("宿主：" + CurrentHost + "（" + Kind + "）");
            sb.AppendLine("程序集版本：" + Assembly.GetExecutingAssembly().GetName().Version);
            sb.AppendLine("Office PIA：" + (typeof(Office.IRibbonExtensibility).Assembly.GetName().Version));
            sb.AppendLine("CTP 工厂：" + (_ctpFactory != null ? "已就绪" : "不可用"));
            sb.AppendLine("任务窗格：" + (_pane != null ? "已创建" : "未创建") +
                          (_pane != null ? "（" + (_paneVisible ? "显示" : "隐藏") + "）" : ""));
            sb.AppendLine("日志：" + Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OfficeAi", "addin.log"));
            MessageBox.Show(sb.ToString(), "关于 Office AI 助手",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion

        #region ICustomTaskPaneConsumer

        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        {
            _ctpFactory = CTPFactoryInst;
            Log("CTPFactoryAvailable: " + (CTPFactoryInst != null));
            CrashLog.Timeline("CTPFactoryAvailable=" + (CTPFactoryInst != null) + "（拿到任务窗格工厂）");
            // 拿到工厂后确保文档监测在跑：处理「双击文件直接打开宿主」以及
            // CTPFactoryAvailable 与 OnConnection 触发先后不定的情况。
            if (CTPFactoryInst != null) StartDocWatcher();
        }

        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hwnd,int cmd);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hwnd);
        private void OpenWordClient()
        {
            if (_clientProcess != null && !_clientProcess.HasExited)
            {
                _clientProcess.Refresh();IntPtr handle=_clientProcess.MainWindowHandle;
                if(handle!=IntPtr.Zero){ShowWindow(handle,9);SetForegroundWindow(handle);}return;
            }
            string root=Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string exe=Path.Combine(root,"OfficeAi.Client.exe");
            if(!File.Exists(exe))throw new FileNotFoundException("独立助手组件缺失，请更新 OfficeAI。",exe);
            dynamic a=_app;long handleValue=Convert.ToInt64(a.ActiveWindow.Hwnd);uint pid;GetWindowThreadProcessId(new IntPtr(handleValue),out pid);
            _clientProcess=Process.Start(new ProcessStartInfo(exe,handleValue+" "+pid){UseShellExecute=false,CreateNoWindow=true,WorkingDirectory=root});
            _paneVisible=true;PauseDocWatcher();
            Log("Word 独立助手已启动 clientPid="+_clientProcess.Id+" hostPid="+pid+"；未创建 Word 内嵌网页");
        }
        private void ShowPane(bool visible)
        {
            if(Kind==HostKind.Word)
            {
                if(visible)OpenWordClient();
                else if(_clientProcess!=null && !_clientProcess.HasExited)_clientProcess.CloseMainWindow();
                return;
            }
            if (_ctpFactory == null) return;
            bool created = false;
            if (_pane == null)
            {
                Log("CreateCTP: " + TaskPaneHost.ProgIdValue);
                CrashLog.Timeline("CreateCTP 开始 progId=" + TaskPaneHost.ProgIdValue);
                _pane = _ctpFactory.CreateCTP(TaskPaneHost.ProgIdValue, "AI 助手",
                                              Type.Missing);
                Log("CreateCTP 成功");
                created = true;
                try { _pane.Width = 420; } catch (Exception e) { Log("设宽度失败: " + e.Message); }
            }
            _paneVisible = visible;
            _pane.Visible = visible;
            Log("窗格可见性 -> " + visible);
            // 时间线里带上窗格状态：崩溃报告要能看出「崩时窗格是开着还是关着」，
            // 以及这次动作是「首次创建」还是「切换可见性」。
            string paneState = "?";
            try { paneState = "width=" + _pane.Width + " visible=" + _pane.Visible; } catch { }
            CrashLog.Timeline("CTP 窗格 " + (created ? "创建并" : "") + "设置可见性 -> " + visible +
                              "  [" + paneState + "]");
        }

        /// <summary>供前端「打开侧边栏」按钮调用</summary>
        internal void TogglePaneFromUi() { ShowPane(!_paneVisible); }

        #endregion
    }
}
