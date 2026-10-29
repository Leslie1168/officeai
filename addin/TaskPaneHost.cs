using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace OfficeAi
{
    /// <summary>
    /// 任务窗格的宿主控件。注册成 ActiveX 控件后由 Office 的 CreateCTP 创建，
    /// 内部用 WebView2 承载前端界面。
    /// </summary>
    [ComVisible(true)]
    [Guid("B4D8F1A2-9C3E-4A7B-8D15-6E2F0C9A3B47")]
    [ProgId(TaskPaneHost.ProgIdValue)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class TaskPaneHost : UserControl
    {
        public const string ProgIdValue = "OfficeAi.TaskPaneHost";
        private const string VirtualHost = "officeai.local";
        private const string ImgHost = "officeai-img.local";

        private WebView2 _web;
        private Bridge _bridge;
        private Label _fallback;
        private bool _initializing;
        private bool _closed;

        /// <summary>
        /// UI 线程锚点。WebView2 的 host object 回调不保证跑在宿主（Word）的 UI 线程上，
        /// 而 Office 对象模型是 STA 的，从别的线程直接访问会 AV（崩在 clr!GetCCWFromIUnknown）。
        /// 所有触碰 Office COM 的操作都必须经 UiInvoke 编组回这里所在的 UI 线程。
        /// </summary>
        internal static Control UiAnchor;

        internal static void UiInvoke(Action work)
        {
            var c = UiAnchor;
            if (c == null || c.IsDisposed || !c.IsHandleCreated) throw new InvalidOperationException("侧边栏尚未就绪或已关闭，请重新打开。");
            try
            {
                if (c.InvokeRequired) c.Invoke(work);
                else work();
            }
            catch (Exception e)
            {
                Connect.Log("UiInvoke 失败: " + e.Message);
            }
        }

        internal static T UiInvoke<T>(Func<T> work)
        {
            var c = UiAnchor;
            if (c == null || c.IsDisposed || !c.IsHandleCreated) throw new InvalidOperationException("侧边栏尚未就绪或已关闭，请重新打开。");
            if (c.InvokeRequired) return (T)c.Invoke(work);
            return work();
        }

        internal static void UiPost(Action work)
        {
            var c=UiAnchor;
            if(c==null || c.IsDisposed || !c.IsHandleCreated) throw new InvalidOperationException("侧边栏已关闭，请重新打开。");
            c.BeginInvoke(new Action(() => {if(!c.IsDisposed) work();}));
        }

        public TaskPaneHost()
        {
            Connect.Log("TaskPaneHost ctor");
            UiAnchor = this;
            BuildUi();
        }

        private void BuildUi()
        {
            BackColor = System.Drawing.Color.FromArgb(250, 250, 252);

            _fallback = new Label
            {
                Dock = DockStyle.Fill,
                Text = "正在初始化界面…",
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9.5F),
                ForeColor = System.Drawing.Color.FromArgb(120, 120, 130),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };
            Controls.Add(_fallback);

            _web = new WebView2 { Dock = DockStyle.Fill, Visible = false };
            Controls.Add(_web);
            _web.BringToFront();

            Load += OnLoaded;
        }

        private void OnLoaded(object sender, EventArgs e)
        {
            // ★ 关键：必须保证 AddHostObjectToScript 与 host object 的所有回调都跑在
            // 创建 WebView2 的同一 UI 线程上。WebView2 通过 COM/IDispatch 编组，
            // host object 在错线程上注册会让 JS 调过来的方法在 broker 线程上执行，
            // 碰 Office COM 时直接崩在 clr!GetCCWFromIUnknown。
            // 这里用 Task.Run 跑异步初始化，但用 Control.Invoke 把后续
            // AddHostObjectToScript / 设置虚拟主机等动作回到 UI 线程。
            if (_initializing || _closed) return;
            _initializing = true;
            _ = InitWebViewAsync();
        }

        private async System.Threading.Tasks.Task InitWebViewAsync()
        {
            try
            {
                // ★ 验收/取证用的 WebView2 远程调试通道（CDP）。
                //   默认完全关闭：只有显式设置了 OFFICEAI_WEBVIEW2_DEBUG_PORT 才生效，
                //   这样「正常使用」与「崩溃复现」两条路径的环境完全一致，不会污染取证。
                //   为什么不用环境变量 WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS：
                //   实测把它写进该环境变量后，WebView2 子进程命令行里并没有出现
                //   --remote-debugging-port、UDF 里也没生成 DevToolsActivePort（Word 未透传）。
                //   → 走 CoreWebView2Environment.CreateAsync(options) 才是可靠通道。
                //
                // ★ API 坑：EnsureCoreWebView2Async 的参数是 CoreWebView2Environment，
                //   不是 CoreWebView2EnvironmentOptions（写成 opts 直接编译不过）。
                //   而且 WebView2 一个进程只允许一个默认环境：只有开启调试时才自建环境，
                //   平时传 null 走默认环境，行为与改动前完全一致。
                var opts = new CoreWebView2EnvironmentOptions();
                int port;
                if (int.TryParse(Environment.GetEnvironmentVariable("OFFICEAI_WEBVIEW2_DEBUG_PORT"), out port) && port > 0 && port < 65536)
                    opts.AdditionalBrowserArguments = "--remote-debugging-port=" + port;
                string profile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OfficeAi", "WebView2", System.Diagnostics.Process.GetCurrentProcess().ProcessName);
                var env = await CoreWebView2Environment.CreateAsync(null, profile, opts);
                if (_closed || IsDisposed) return;
                await _web.EnsureCoreWebView2Async(env);
            }
            catch (Exception ex)
            {
                if (_closed || IsDisposed) return;
                Connect.Log("WebView2 初始化失败: " + ex);
                SafeSetFallback("WebView2 初始化失败：\n" + ex.Message +
                                "\n\n请确认已安装 WebView2 Runtime。");
                return;
            }

            if (_closed || IsDisposed) return;
            // 回到创建 WebView2 的 UI 线程注册 host object
            if (InvokeRequired)
            {
                try { BeginInvoke(new Action(SetupCoreWebViewOnUi)); }
                catch (Exception ex) { Connect.Log("回到 UI 线程失败: " + ex.Message); }
            }
            else
            {
                SetupCoreWebViewOnUi();
            }
        }

        private void SafeSetFallback(string text)
        {
            if (InvokeRequired) { try { BeginInvoke(new Action(() => { if (_fallback != null) _fallback.Text = text; })); } catch { } }
            else { if (_fallback != null) _fallback.Text = text; }
        }

        private void SetupCoreWebViewOnUi()
        {
            try
            {
                if (_closed || IsDisposed || _bridge != null) return;
                var s = _web.CoreWebView2.Settings;
                s.AreDefaultContextMenusEnabled = true;
                s.IsStatusBarEnabled = false;
                s.AreDevToolsEnabled = true;
                s.IsZoomControlEnabled = false;

                _bridge = new Bridge();
                // ★ 关键：AddHostObjectToScript 与 Bridge 类本身都必须在 UI 线程创建/注册，
                // 这样 JS 调 chrome.webview.hostObjects.bridge.xxx() 时，CLR 会在同一 STA 上
                // 执行，访问 Office COM 才安全。
                _web.CoreWebView2.AddHostObjectToScript("bridge", _bridge);
                _web.CoreWebView2.NavigationStarting += (sender, args) =>
                {
                    Uri uri;
                    if (!Uri.TryCreate(args.Uri, UriKind.Absolute, out uri) ||
                        uri.Scheme != "https" || uri.Host != VirtualHost) args.Cancel = true;
                };
                _web.CoreWebView2.NewWindowRequested += (sender, args) => { args.Handled = true; };

                _bridge.PushToJs = (json) =>
                {
                    try
                    {
                        if (InvokeRequired) BeginInvoke(new Action(() => Post(json)));
                        else Post(json);
                    }
                    catch (Exception ex) { Connect.Log("PushToJs 失败: " + ex.Message); }
                };

                _web.CoreWebView2.WebMessageReceived += (a, b) =>
                {
                    try { Connect.Log("前端消息: " + b.TryGetWebMessageAsString()); }
                    catch { }
                };

                string uiDir = Path.Combine(InstallRoot(), "ui");
                if (File.Exists(Path.Combine(uiDir, "index.html")))
                {
                    _web.CoreWebView2.SetVirtualHostNameToFolderMapping(
                        VirtualHost, uiDir, CoreWebView2HostResourceAccessKind.Allow);

                    string imgDir = ImageExtractor.CacheRoot;
                    try
                    {
                        Directory.CreateDirectory(imgDir);
                        _web.CoreWebView2.SetVirtualHostNameToFolderMapping(
                            ImgHost, imgDir, CoreWebView2HostResourceAccessKind.Allow);
                    }
                    catch (Exception ie) { Connect.Log("图片目录映射失败: " + ie.Message); }

                    _web.CoreWebView2.Navigate("https://" + VirtualHost + "/index.html");
                    Connect.Log("导航到 https://" + VirtualHost + "/index.html");
                }
                else
                {
                    Connect.Log("ui 目录缺失: " + uiDir);
                    _web.CoreWebView2.NavigateToString(
                        "<html><body style=\"font-family:'Microsoft YaHei UI';padding:20px\">" +
                        "<h3>界面文件缺失</h3><p>找不到 " + uiDir + "</p></body></html>");
                }

                _fallback.Visible = false;
                _web.Visible = true;
            }
            catch (Exception ex)
            {
                Connect.Log("SetupCoreWebView 异常: " + ex);
                SafeSetFallback("WebView2 初始化失败：\n" + ex.Message);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && !_closed)
            {
                _closed = true;
                if (_bridge != null) { _bridge.CancelAsk(); _bridge.PushToJs = null; }
                if (ReferenceEquals(UiAnchor, this)) UiAnchor = null;
            }
            base.Dispose(disposing);
        }

        internal static string InstallRoot()
        {
            try { return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); }
            catch { return ""; }
        }

        private void Post(string json)
        {
            try
            {
                if (_web != null && _web.CoreWebView2 != null)
                    _web.CoreWebView2.PostWebMessageAsJson(json);
            }
            catch (Exception e) { Connect.Log("Post 失败: " + e.Message); }
        }
    }
}
