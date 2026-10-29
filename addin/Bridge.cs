using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace OfficeAi
{
    /// <summary>
    /// 前端（WebView2 里的 JS）与后端（C#）之间的桥。
    /// JS 侧通过 chrome.webview.hostObjects.bridge 调用；返回值一律是 JSON 字符串。
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Bridge
    {
        private const int MaxChars = 200000;      // 总长度上限
        private const int MaxRowsPerSheet = 300;  // 每个工作表最多读多少行
        private const int MaxColsPerSheet = 40;   // 每个工作表最多读多少列

        #region 基础

        /// <summary>
        /// 把「访问 Office 对象模型」的工作编组到宿主的 UI 线程执行。
        /// Office 的 Application 对象是 STA 的；WebView2 的 host object 回调不保证
        /// 跑在 UI 线程上，跨线程直接访问会在 CLR 的 COM 编组层崩溃
        /// （clr!GetCCWFromIUnknown / wwlib.dll 0xc0000005）。所有碰 COM 的方法都要经这里。
        /// </summary>
        private static string OnUi(Func<string> work)
        {
            return TaskPaneHost.UiInvoke(work);
        }

        public string GetHostInfo()
        {
            string root = TaskPaneHost.InstallRoot().Replace("\\", "/");
            Connect.Log("Bridge.GetHostInfo 被调用");
            return "{\"host\":\"" + Esc(Connect.CurrentHost) + "\"," +
                   "\"version\":\"" + Esc(Connect.Version) + "\"," +
                   "\"installRoot\":\"" + Esc(root) + "\"," +
                   "\"kind\":\"" + Esc(Connect.Kind.ToString()) + "\"," +
                   "\"uiMode\":\"" + (Process.GetCurrentProcess().ProcessName=="OfficeAi.Client" ? "separate" : "embedded") + "\"," +
                   "\"uiProcessId\":" + Process.GetCurrentProcess().Id + "," +
                   "\"imgHost\":\"https://officeai-img.local\"," +
                   "\"webview\":\"ok\"}";
        }

        public void LogFromJs(string msg) { Connect.Log("[js] " + msg); }

        /// <summary>
        /// 【T2 新增，纯增量：不改任何已有方法的签名】进程侧存活心跳。
        /// 前端探针（渲染阶段/缩略图 load 计数）可以调它，让关键节点落进 crash.log 并 Flush(true)：
        /// 崩溃后把「JS 探针时间戳」与「crash.log 里的 t+毫秒」对齐，
        /// 就能判定崩溃发生在渲染中还是渲染完成后。频率建议 ≤ 每 3 秒一次。
        /// </summary>
        public void Heartbeat(string tag) { CrashLog.Heartbeat(tag); }

        public string Echo(string text) { Debug.WriteLine(text); return "[backend] " + text; }

        /// <summary>由 TaskPaneHost 注入：把一条 JSON 消息推给前端（已切到 UI 线程）</summary>
        internal Action<string> PushToJs;
        private readonly object _requestGate = new object();
        private CancellationTokenSource _request;
        public string CancelAsk()
        {
            lock (_requestGate) { if (_request != null) { _request.Cancel(); _request = null; } }
            return "{\"ok\":true}";
        }
        private static JavaScriptSerializer Serializer() => new JavaScriptSerializer { MaxJsonLength = 4 * 1024 * 1024 };


        #endregion

        #region 设置

        public string GetSettings()
        {
            return Settings.LoadJson();
        }

        public string SaveSettings(string json)
        {
            bool ok = Settings.SaveJson(json);
            return "{\"ok\":" + (ok ? "true" : "false") + "}";
        }

        #endregion

        #region 对话

        /// <summary>
        /// 发起一次对话。立即返回，内容通过 PushToJs 以
        /// {"type":"delta"|"done"|"error", ...} 的形式流式推给前端。
        /// </summary>
        public string Ask(string reqJson)
        {
            try
            {
                var ser = Serializer();
                var d = ser.DeserializeObject(reqJson) as Dictionary<string, object>;
                if (d == null) return "{\"ok\":false,\"message\":\"请求格式不对\"}";

                string prompt = d.ContainsKey("prompt") ? Convert.ToString(d["prompt"]) : "";
                string mode = d.ContainsKey("mode") ? Convert.ToString(d["mode"]) : "ask";
                bool withCtx = d.ContainsKey("withContext") && Convert.ToBoolean(d["withContext"]);
                string docText = d.ContainsKey("docText") ? Convert.ToString(d["docText"]) : "";

                string key = Settings.Get("apiKey", "");
                string baseUrl = Settings.Get("baseUrl", "https://api.deepseek.com");
                string model = Settings.Get("model", "deepseek-chat");

                string user = prompt;
                if (withCtx && !string.IsNullOrEmpty(docText))
                {
                    user = "【当前文档内容】\n" + docText +
                           "\n\n【我的要求】\n" + prompt;
                }

                if (string.IsNullOrWhiteSpace(prompt)) return "{\"ok\":false,\"message\":\"请输入要求\"}";
                string requestId = d.ContainsKey("requestId") ? Convert.ToString(d["requestId"]) : "";
                var history = new List<Dictionary<string, string>>();
                object[] turns = d.ContainsKey("history") ? d["history"] as object[] : null;
                if (turns != null)
                    foreach (object turn in turns)
                    {
                        var item = turn as Dictionary<string, object>;
                        if (item == null || !item.ContainsKey("role") || !item.ContainsKey("content")) continue;
                        string role = Convert.ToString(item["role"]), content = Convert.ToString(item["content"]);
                        if (role != "user" && role != "assistant") continue;
                        history.Add(new Dictionary<string, string> { { "role", role }, { "content", content.Length > 12000 ? content.Substring(0, 12000) : content } });
                    }
                if (history.Count > 16) history.RemoveRange(0, history.Count - 16);
                CancellationTokenSource request;
                lock (_requestGate)
                {
                    if (_request != null) _request.Cancel();
                    _request = request = new CancellationTokenSource();
                }
                Action<string> push = json =>
                {
                    lock (_requestGate)
                    {
                        if (!ReferenceEquals(_request, request) || request.IsCancellationRequested) return;
                        var message = Serializer().DeserializeObject(json) as Dictionary<string, object>;
                        message["requestId"] = requestId;
                        PushToJs?.Invoke(Serializer().Serialize(message));
                    }
                };
                Connect.Log("Ask: mode=" + mode + " withCtx=" + withCtx +
                            " docLen=" + (docText == null ? 0 : docText.Length));

                // 写作模式：让模型输出结构化编辑计划（一次性拿到完整 JSON），
                // 前端解析后渲染预览卡片，用户确认才真正落到文档。
                if (mode == "write")
                {
                    DeepSeek.ChatOnce(key, baseUrl, model, DeepSeek.WritePlanSystemPrompt, user,
                        content =>
                        {
                            Connect.Log("Ask(write) 计划返回 " +
                                        (content == null ? 0 : content.Length) + " 字");
                            if (push != null)
                                push("{\"type\":\"plan\",\"raw\":\"" + Esc(content) + "\"}");
                        },
                        err =>
                        {
                            Connect.Log("Ask(write) 失败: " + err.Replace("\n", " | "));
                            if (push != null) push("{\"type\":\"error\",\"message\":\"" + Esc(err) + "\"}");
                        }, request.Token, history);
                    return "{\"ok\":true}";
                }

                DeepSeek.StreamChat(key, baseUrl, model, DeepSeek.AskSystemPrompt, history, user,
                    delta => { if (push != null) push("{\"type\":\"delta\",\"text\":\"" + Esc(delta) + "\"}"); },
                    () => { Connect.Log("Ask: 流结束"); if (push != null) push("{\"type\":\"done\"}"); },
                    err =>
                    {
                        Connect.Log("Ask 失败: " + err.Replace("\n", " | "));
                        if (push != null) push("{\"type\":\"error\",\"message\":\"" + Esc(err) + "\"}");
                    }, request.Token);

                return "{\"ok\":true}";
            }
            catch (Exception e)
            {
                Connect.Log("Ask 异常: " + e);
                return "{\"ok\":false,\"message\":\"" + Esc(e.Message) + "\"}";
            }
        }

        #endregion

        #region 生成 PPT

        /// <summary>
        /// 由大纲生成 .pptx。立即返回，完成后通过 PushToJs 推
        /// {"type":"ppt","ok":bool,"path":"...","message":"..."}。
        /// 请求 JSON：{"outline":"...","theme":"tech-cyan","filename":"..."}
        /// </summary>
        public string GeneratePpt(string reqJson)
        {
            try
            {
                var ser = Serializer();
                var d = ser.DeserializeObject(reqJson) as Dictionary<string, object>;
                if (d == null) return "{\"ok\":false,\"message\":\"请求格式不对\"}";

                string outline = d.ContainsKey("outline") ? Convert.ToString(d["outline"]) : "";
                string theme = d.ContainsKey("theme") ? Convert.ToString(d["theme"]) : null;
                string fname = d.ContainsKey("filename") ? Convert.ToString(d["filename"]) : null;
                if (string.IsNullOrWhiteSpace(outline))
                    return "{\"ok\":false,\"message\":\"大纲是空的\"}";

                string script = Path.Combine(Path.Combine(TaskPaneHost.InstallRoot(), "engine"),
                                             "make_pptx.py");
                string python = Path.Combine(Path.Combine(
                    Path.Combine(TaskPaneHost.InstallRoot(), "engine"), "venv"),
                    @"Scripts\python.exe");
                if (!File.Exists(script) || !File.Exists(python))
                    return "{\"ok\":false,\"message\":\"PPT 引擎缺失：" + Esc(script) + "\"}";

                string outPath = BuildPptPath(fname);
                string outlineFile = Path.Combine(Path.GetTempPath(),
                    "officeai_outline_" + DateTime.Now.Ticks + ".txt");
                File.WriteAllText(outlineFile, outline, new UTF8Encoding(false));

                Action<string> push = PushToJs;
                Connect.Log("GeneratePpt: -> " + outPath);

                Task.Run(() =>
                {
                    string resp = "";
                    string err = "";
                    try
                    {
                        var psi = new ProcessStartInfo
                        {
                            FileName = python,
                            Arguments = "\"" + script + "\" -i \"" + outlineFile + "\" -o \"" +
                                        outPath + "\"" +
                                        (string.IsNullOrEmpty(theme) ? "" : " --theme " + theme),
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            RedirectStandardOutput = true,
                            RedirectStandardError = true,
                            WorkingDirectory = Path.GetDirectoryName(script)
                        };
                        using (var p = Process.Start(psi))
                        {
                            resp = p.StandardOutput.ReadToEnd();
                            err = p.StandardError.ReadToEnd();
                            p.WaitForExit(60000);
                        }
                    }
                    catch (Exception e) { err = e.Message; }
                    finally
                    {
                        try { File.Delete(outlineFile); } catch { }
                    }

                    bool ok = File.Exists(outPath) && new FileInfo(outPath).Length > 0;
                    Connect.Log("GeneratePpt 完成 ok=" + ok +
                                (string.IsNullOrEmpty(err) ? "" : " err=" + err.Replace("\n", " | ")));

                    if (push != null)
                    {
                        push("{\"type\":\"ppt\",\"ok\":" + (ok ? "true" : "false") +
                             ",\"path\":\"" + Esc(outPath) + "\"" +
                             ",\"message\":\"" + Esc(ok ? "已生成" : (err.Length > 0 ? err : resp)) +
                             "\"}");
                    }
                    if (ok) OpenFile(outPath);
                });

                return "{\"ok\":true}";
            }
            catch (Exception e)
            {
                Connect.Log("GeneratePpt 异常: " + e);
                return "{\"ok\":false,\"message\":\"" + Esc(e.Message) + "\"}";
            }
        }

        /// <summary>默认存到桌面；重名自动加序号</summary>
        private static string BuildPptPath(string fname)
        {
            string dir;
            try { dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory); }
            catch { dir = Path.GetTempPath(); }
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir)) dir = Path.GetTempPath();

            string name = string.IsNullOrWhiteSpace(fname) ? "演示文稿" : fname.Trim();
            foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
            if (!name.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase)) name += ".pptx";

            string full = Path.Combine(dir, name);
            int i = 2;
            while (File.Exists(full))
                full = Path.Combine(dir, Path.GetFileNameWithoutExtension(name) + "(" + (i++) + ").pptx");
            return full;
        }

        private static void OpenFile(string path)
        {
            try
            {
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
                Connect.Log("OpenFile: " + path);
            }
            catch (Exception e) { Connect.Log("OpenFile 失败: " + e.Message); }
        }

        #endregion

        #region 图片抓取（第7条）

        /// <summary>抓取当前文档的全部内嵌图片，落盘到缓存并返回清单（含缩略图相对路径）。</summary>
        public string ExtractImages()
        {
            try
            {
                Connect.Log("→ ExtractImages 进入");
                return OnUi(() => ImageExtractor.Extract(Connect.CurrentApp, Connect.Kind));
            }
            catch (Exception e)
            {
                Connect.Log("ExtractImages 异常: " + e);
                return "{\"ok\":false,\"error\":\"" + Esc(e.Message) + "\"}";
            }
        }

        /// <summary>把某张缓存图片另存为用户指定路径。</summary>
        public string SaveImageAs(string srcRel, string destPath)
        {
            try
            {
                string src = Path.Combine(ImageExtractor.CacheRoot, srcRel.Replace('/', '\\'));
                return ImageExtractor.SaveTo(src, destPath);
            }
            catch (Exception e)
            {
                return "{\"ok\":false,\"message\":\"" + Esc(e.Message) + "\"}";
            }
        }

        /// <summary>把某张图片当作“图片本身”插入到当前光标处。返回文字消息（ERR: 开头表示失败）。</summary>
        public string InsertImageAtCursor(string srcRel)
        {
            try
            {
                string src = Path.Combine(ImageExtractor.CacheRoot, srcRel.Replace('/', '\\'));
                if (!File.Exists(src)) return "{\"ok\":false,\"message\":\"图片不存在\"}";
                Connect.Log("InsertImageAtCursor: " + src);
                string msg = OnUi(() => ImageInsert.Insert(Connect.CurrentApp, Connect.Kind, src));
                bool ok = !msg.StartsWith("ERR:", StringComparison.Ordinal);
                return ok ? "{\"ok\":true,\"message\":\"" + Esc(msg) + "\"}"
                          : "{\"ok\":false,\"message\":\"" + Esc(msg.Substring(4)) + "\"}";
            }
            catch (Exception e)
            {
                Connect.Log("InsertImageAtCursor 异常: " + e);
                return "{\"ok\":false,\"message\":\"" + Esc(e.Message) + "\"}";
            }
        }

        /// <summary>把图片真身放进剪贴板（CF_DIB），这样到 Word/PPT 里 Ctrl+V 就是图片本身。</summary>
        public string CopyImage(string srcRel)
        {
            try
            {
                string src = Path.Combine(ImageExtractor.CacheRoot, srcRel.Replace('/', '\\'));
                if (!File.Exists(src)) return "{\"ok\":false,\"message\":\"图片不存在\"}";
                // 剪贴板数据要驻留在带消息泵的 STA 线程上才能长期有效，
                // 这里借 Bridge.OnUi（已是 Office 的 STA UI 线程）执行。
                bool ok = OnUi(() => ImageBoard.CopyToClipboard(src)) == "1";
                Connect.Log("CopyImage(" + srcRel + ") -> " + ok);
                return ok ? "{\"ok\":true,\"message\":\"图片已复制，在文档里 Ctrl+V 粘贴即是图片本身\"}"
                          : "{\"ok\":false,\"message\":\"复制到剪贴板失败\"}";
            }
            catch (Exception e)
            {
                Connect.Log("CopyImage 异常: " + e);
                return "{\"ok\":false,\"message\":\"" + Esc(e.Message) + "\"}";
            }
        }

        // Retained as inert compatibility endpoints for stale cached pages.
        public string PreviewFormat(string reqJson) => VbaFormatPlan.Json(new {ok=false,message="独立排版页已移除，请在写作框提出格式要求。"});
        public string ApplyFormat(string reqJson) => PreviewFormat(reqJson);

        /// <summary>用系统默认程序打开某张缓存图片（真实文件）。</summary>
        public string OpenImage(string srcRel)
        {
            try
            {
                Connect.Log("OpenImage 进入: " + srcRel);
                string src = Path.Combine(ImageExtractor.CacheRoot, srcRel.Replace('/', '\\'));
                if (!File.Exists(src)) return "{\"ok\":false,\"message\":\"图片不存在\"}";
                ShellStart(src, null);
                Connect.Log("OpenImage: " + src);
                return "{\"ok\":true}";
            }
            catch (Exception e)
            {
                return "{\"ok\":false,\"message\":\"" + Esc(e.Message) + "\"}";
            }
        }

        /// <summary>在资源管理器中定位某张缓存图片。</summary>
        public string RevealImage(string srcRel)
        {
            try
            {
                string src = Path.Combine(ImageExtractor.CacheRoot, srcRel.Replace('/', '\\'));
                if (!File.Exists(src)) return "{\"ok\":false,\"message\":\"图片不存在\"}";
                ShellStart("explorer.exe", "/select,\"" + src + "\"");
                return "{\"ok\":true}";
            }
            catch (Exception e)
            {
                return "{\"ok\":false,\"message\":\"" + Esc(e.Message) + "\"}";
            }
        }

        /// <summary>批量导出：把本次抓取的所有图片复制到一个目录。</summary>
        public string ExportAllImages(string dir)
        {
            try
            {
                if (string.IsNullOrEmpty(dir)) return "{\"ok\":false,\"message\":\"没给目录\"}";
                string docPath = null;
                try
                {
                    dynamic a = Connect.CurrentApp;
                    if (Connect.Kind == Connect.HostKind.Word) docPath = (string)a.ActiveDocument.FullName;
                    else if (Connect.Kind == Connect.HostKind.Excel) docPath = (string)a.ActiveWorkbook.FullName;
                    else if (Connect.Kind == Connect.HostKind.PowerPoint) docPath = (string)a.ActivePresentation.FullName;
                }
                catch { }
                if (docPath == null) return "{\"ok\":false,\"message\":\"没有打开的文档\"}";

                string stem = Path.GetFileNameWithoutExtension(docPath);
                int hash = Math.Abs(docPath.ToLowerInvariant().GetHashCode());
                string srcDir = Path.Combine(ImageExtractor.CacheRoot,
                    stem + "_" + hash.ToString("X8"), "orig");
                if (!Directory.Exists(srcDir))
                    return "{\"ok\":false,\"message\":\"请先在「图片」栏抓取一次\"}";

                string target = Path.Combine(dir, stem + "_图片");
                Directory.CreateDirectory(target);
                int n = 0;
                foreach (var f in Directory.GetFiles(srcDir))
                {
                    File.Copy(f, Path.Combine(target, Path.GetFileName(f)), true);
                    n++;
                }
                ShellStart(target, null);
                Connect.Log("ExportAllImages: " + n + " 张 -> " + target);
                return "{\"ok\":true,\"count\":" + n + ",\"path\":\"" + Esc(target) + "\"}";
            }
            catch (Exception e)
            {
                Connect.Log("ExportAllImages 异常: " + e);
                return "{\"ok\":false,\"message\":\"" + Esc(e.Message) + "\"}";
            }
        }

        /// <summary>
        /// 用系统外壳打开文件/程序。必须放到后台线程执行：
        /// WebView2 的 host object 回调跑在宿主（Word）的 UI 线程上，
        /// 直接在这里 ShellExecute 会泵动消息循环、让 Word 重入自身而崩溃（wwlib.dll 0xc0000005）。
        /// </summary>
        private static void ShellStart(string file, string args)
        {
            string f = file, a = args;
            var t = new Thread(() =>
            {
                try
                {
                    var psi = a == null
                        ? new ProcessStartInfo(f) { UseShellExecute = true }
                        : new ProcessStartInfo(f, a) { UseShellExecute = true };
                    Process.Start(psi);
                }
                catch (Exception e) { Connect.Log("ShellStart 失败: " + e.Message); }
            });
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        /// <summary>让前端选一个导出目录（返回绝对路径或空）。</summary>
        public string PickFolder()
        {
            // 同上：对话框必须在后台 STA 线程弹，不能占用宿主的 UI 线程
            return RunOnStaThread(() =>
            {
                using (var dlg = new FolderBrowserDialog())
                {
                    dlg.Description = "选择图片导出目录";
                    dlg.ShowNewFolderButton = true;
                    return dlg.ShowDialog() == DialogResult.OK ? dlg.SelectedPath : "";
                }
            });
        }

        /// <summary>让前端选一个保存文件的目标路径（另存为）。</summary>
        public string PickSavePath(string defaultName)
        {
            string path = RunOnStaThread(() =>
            {
                using (var dlg = new SaveFileDialog())
                {
                    dlg.Title = "另存为";
                    dlg.Filter = "所有文件|*.*";
                    dlg.FileName = string.IsNullOrEmpty(defaultName) ? "image" : defaultName;
                    dlg.OverwritePrompt = true;
                    return dlg.ShowDialog() == DialogResult.OK ? dlg.FileName : "";
                }
            });
            return "{\"path\":\"" + Esc(path.Replace("\\", "/")) + "\"}";
        }

        /// <summary>在独立 STA 线程上跑一段带对话框的代码，避免重入宿主 UI 线程。</summary>
        private static string RunOnStaThread(Func<string> work)
        {
            string result = "";
            var t = new Thread(() => { try { result = work(); } catch (Exception e) { Connect.Log("RunOnStaThread 异常: " + e.Message); } });
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
            return result;
        }

        #endregion

        #region 写入（必须由用户点「确认导入」触发）

        /// <summary>
        /// 把内容写进文档。method: "type"（模拟输入）| "vba"
        /// 返回 {"ok":bool,"message":"...","method":"...","vba":"..."}
        /// </summary>
        public string ApplyText(string text, string method)
        {
            try
            {
                Connect.Log("ApplyText: method=" + method + " len=" +
                            (text == null ? 0 : text.Length));
                return OnUi(() => Writer.Apply(Connect.CurrentApp, Connect.Kind, text, method));
            }
            catch (Exception e)
            {
                Connect.Log("ApplyText 异常: " + e);
                return "{\"ok\":false,\"message\":\"" + Esc(e.Message) + "\"}";
            }
        }

        #endregion

        #region 精准编辑（第1条）

        /// <summary>读取当前选区（含 Word 表格单元格、Excel 区域、PPT 文本框）。</summary>
        public string GetSelection()
        {
            return OnUi(() => Precise.GetSelection(Connect.CurrentApp, Connect.Kind));
        }

        /// <summary>把当前选区替换为 newText。</summary>
        public string ReplaceSelection(string newText)
        {
            return OnUi(() => Precise.ReplaceSelection(Connect.CurrentApp, Connect.Kind, newText));
        }

        /// <summary>全文查找替换。all=true 替换全部。</summary>
        public string FindReplace(string find, string replace, bool all)
        {
            return OnUi(() => Precise.FindReplace(Connect.CurrentApp, Connect.Kind, find, replace, all));
        }

        /// <summary>
        /// 执行一组编辑动作（写作模式的核心）。请求 JSON：
        /// {"ops":[{"op":"replace","find":"X","replace":"Y","all":true}, ...], "method":"type"}
        /// 返回 {"ok":bool,"message":"...","detail":"..."}。
        /// 必须由用户点「确认执行」触发。
        /// </summary>
        private sealed class EditSnapshot
        {
            public object Document, Range;
            public string Text;
            public bool Selected;
            public string VbaPreview;
        }
        private readonly Dictionary<string, EditSnapshot> _edits = new Dictionary<string, EditSnapshot>();
        public string CaptureEdit(string selectionJson)
        {
            try { return OnUi(() => {
                dynamic app = Connect.CurrentApp;
                var data = Serializer().DeserializeObject(ReadDocumentCore()) as Dictionary<string, object>;
                if (data == null || !Convert.ToBoolean(data["ok"])) return Serializer().Serialize(data);
                string token = Guid.NewGuid().ToString("N");
                if (_edits.Count >= 20) _edits.Clear();
                var snap = new EditSnapshot();
                if (Connect.Kind == Connect.HostKind.Word)
                {
                    snap.Document = app.ActiveDocument;
                    snap.Text = (string)((dynamic)snap.Document).Content.Text;
                    snap.Range = app.Selection.Range.Duplicate;
                    if (!string.IsNullOrWhiteSpace(selectionJson))
                    {
                        var selection = Serializer().DeserializeObject(selectionJson) as Dictionary<string, object>;
                        if (selection != null && selection.ContainsKey("start") && selection.ContainsKey("end"))
                        {
                            dynamic doc = snap.Document;
                            if (!selection.ContainsKey("document") || Convert.ToString(selection["document"]) != (string)doc.FullName)
                                return Fail("锁定选区属于另一份文档，请重新选取。");
                            snap.Range = doc.Range(Convert.ToInt32(selection["start"]), Convert.ToInt32(selection["end"]));
                            if ((string)((dynamic)snap.Range).Text != Convert.ToString(selection["text"]))
                                return Fail("锁定选区的原文已变化，请重新选取。");
                        }
                    }
                    snap.Selected = ((int)((dynamic)snap.Range).End > (int)((dynamic)snap.Range).Start);
                    if (snap.Selected)
                        data["selection"] = new Dictionary<string, object> {
                            { "text", (string)((dynamic)snap.Range).Text }, { "kind", "word" }
                        };
                }
                else
                {
                    snap.Document = Connect.Kind == Connect.HostKind.Excel ? (object)app.ActiveWorkbook : (object)app.ActivePresentation;
                    snap.Text = Convert.ToString(data["text"]);
                    if (Connect.Kind == Connect.HostKind.Excel)
                    {
                        snap.Range = app.Selection;
                        snap.Selected = snap.Range != null;
                    }
                    else
                    {
                        try { if ((int)app.ActiveWindow.Selection.Type == 3) { snap.Range = app.ActiveWindow.Selection.TextRange; snap.Selected = true; } } catch { }
                    }
                }
                _edits[token] = snap;
                data["token"] = token;
                return Serializer().Serialize(data);
            }); } catch (Exception ex) { return Fail(ex.Message); }
        }

        public string PreviewVbaPlan(string reqJson)
        {
            try {return OnUi(() => {
                if(Connect.Kind!=Connect.HostKind.Word) throw new InvalidOperationException("本轮 VBA 格式功能适用于 Word。");
                var request=Serializer().DeserializeObject(reqJson) as Dictionary<string,object>;
                string token=request!=null && request.ContainsKey("token") ? Convert.ToString(request["token"]) : "";
                EditSnapshot snap;
                if(!_edits.TryGetValue(token,out snap)) throw new InvalidOperationException("预览已失效，请重新发送要求。");
                dynamic app=Connect.CurrentApp;
                if(!Equals((object)app.ActiveDocument,snap.Document) || (string)app.ActiveDocument.Content.Text!=snap.Text)
                    throw new InvalidOperationException("原文或文档已变化，请重新发送要求。");
                var ops=request.ContainsKey("ops") ? request["ops"] as object[] : null;
                var plan=VbaFormatPlan.Build(ops,snap.Text,(int)((dynamic)snap.Range).Start,(int)((dynamic)snap.Range).End);
                snap.VbaPreview=Serializer().Serialize(ops);
                return VbaFormatPlan.Json(new {ok=true,code=plan.Code,summary=plan.Summary});
            });} catch(Exception ex) {return VbaFormatPlan.Json(new {ok=false,message=ex.Message});}
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
        private string BeginVbaPlan(Dictionary<string,object> request, object[] ops, string token, EditSnapshot snap)
        {
            if(Connect.Kind!=Connect.HostKind.Word) return VbaFormatPlan.Json(new {ok=false,message="VBA 格式功能目前适用于 Word。"});
            if(snap.VbaPreview!=Serializer().Serialize(ops)) return VbaFormatPlan.Json(new {ok=false,message="请先核对并生成 VBA 预览。"});
            string exe=Path.Combine(TaskPaneHost.InstallRoot(),"OfficeAi.VbaWorker.exe");
            if(!File.Exists(exe)) return VbaFormatPlan.Json(new {ok=false,message="VBA 执行组件缺失，请更新插件。"});
            dynamic app=Connect.CurrentApp;
            long hwnd=Convert.ToInt64(app.ActiveWindow.Hwnd);uint pid;GetWindowThreadProcessId(new IntPtr(hwnd),out pid);
            string requestId=request.ContainsKey("requestId") ? Convert.ToString(request["requestId"]) : Guid.NewGuid().ToString("N");
            string payload=VbaFormatPlan.Json(new {hwnd,pid,document=(string)((dynamic)snap.Document).FullName,hash=VbaFormatPlan.Hash(snap.Text),
                start=(int)((dynamic)snap.Range).Start,end=(int)((dynamic)snap.Range).End,ops});
            _edits.Remove(token);
            Action<string> push=PushToJs;
            Task.Run(() => {
                string folder=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"OfficeAi","vba-jobs");
                Directory.CreateDirectory(folder);string job=Guid.NewGuid().ToString("N");
                string input=Path.Combine(folder,job+".request.json"),output=Path.Combine(folder,job+".result.json");
                bool completed=false;
                try
                {
                    File.WriteAllText(input,payload,new UTF8Encoding(false));
                    using(var process=Process.Start(new ProcessStartInfo(exe,"\""+input+"\" \""+output+"\"") {UseShellExecute=false,CreateNoWindow=true,WorkingDirectory=Path.GetDirectoryName(exe)}))
                    {
                        completed=process.WaitForExit(90000);
                        if(!completed) throw new TimeoutException("VBA 尚未返回结果，请检查文档，勿重复执行。记录："+output);
                    }
                    var result=Serializer().DeserializeObject(File.ReadAllText(output,Encoding.UTF8)) as Dictionary<string,object>;
                    result["type"]="vbaResult";result["requestId"]=requestId;
                    Connect.Log("VBA 格式任务完成 ok="+result["ok"]);
                    push?.Invoke(Serializer().Serialize(result));
                }
                catch(Exception ex) {push?.Invoke(VbaFormatPlan.Json(new {type="vbaResult",requestId,ok=false,message=ex.Message}));}
                finally {if(completed) {try {File.Delete(input);} catch{} try {File.Delete(output);} catch{}}}
            });
            return VbaFormatPlan.Json(new {ok=true,queued=true,requestId});
        }

        public string ApplyPlan(string reqJson)
        {
            try
            {
                var d = Serializer().DeserializeObject(reqJson)
                        as Dictionary<string, object>;
                if (d == null) return "{\"ok\":false,\"message\":\"请求格式不对\"}";
                // JavaScriptSerializer 把 JSON 数组反序列化成 object[]，不是 List<object>，
                // 两种都兼容一下，否则会拿到 null 变成「0 个动作」。
                var ops = new List<object>();
                if (d.ContainsKey("ops") && d["ops"] != null)
                {
                    var arr = d["ops"] as object[];
                    if (arr != null) ops.AddRange(arr);
                    else
                    {
                        var lst = d["ops"] as System.Collections.IEnumerable;
                        if (lst != null)
                            foreach (var o in lst) ops.Add(o);
                    }
                }
                string method = d.ContainsKey("method") ? Convert.ToString(d["method"]) : "type";
                Connect.Log("ApplyPlan: " + ops.Count + " 个动作 method=" + method);
                string token = d.ContainsKey("token") ? Convert.ToString(d["token"]) : "";
                return OnUi(() => {
                    EditSnapshot snap;
                    if (!_edits.TryGetValue(token, out snap)) return Fail("修改预览已失效，请重新发送写作要求。");
                    dynamic app = Connect.CurrentApp;
                    object current = Connect.Kind == Connect.HostKind.Word ? (object)app.ActiveDocument :
                        Connect.Kind == Connect.HostKind.Excel ? (object)app.ActiveWorkbook : (object)app.ActivePresentation;
                    if (!Equals(current, snap.Document)) return Fail("当前文档已切换，请回到生成预览时的文档。");
                    string text = Connect.Kind == Connect.HostKind.Word ? (string)((dynamic)current).Content.Text :
                        Convert.ToString(((Dictionary<string, object>)Serializer().DeserializeObject(ReadDocumentCore()))["text"]);
                    if (text != snap.Text) return Fail("原文已变化，请重新生成预览，避免覆盖你的修改。");
                    if(VbaFormatPlan.HasFormat(ops.ToArray())) return BeginVbaPlan(d,ops.ToArray(),token,snap);
                    foreach (var item in ops)
                    {
                        var op = item as Dictionary<string, object>;
                        if (op != null && op.ContainsKey("op") && Convert.ToString(op["op"]) == "rewrite" && !snap.Selected)
                            return Fail("生成预览时没有锁定文字选区，请选中文字后重新发送。");
                    }
                    if (Connect.Kind == Connect.HostKind.Word)
                    {
                        bool rewriting = ops.Exists(item => { var op = item as Dictionary<string, object>; return op != null && op.ContainsKey("op") && Convert.ToString(op["op"]) == "rewrite"; });
                        if (!rewriting) ((dynamic)snap.Range).Collapse(0);
                        ((dynamic)snap.Range).Select();
                    }
                    if (Connect.Kind == Connect.HostKind.Excel && snap.Range != null)
                    {
                        ((dynamic)snap.Range).Worksheet.Activate();
                        ((dynamic)snap.Range).Select();
                    }
                    if (Connect.Kind == Connect.HostKind.PowerPoint && snap.Range != null) ((dynamic)snap.Range).Select();
                    _edits.Remove(token); // one-shot: a failed partial plan cannot be blindly replayed
                    return Editor.ExecutePlan(Connect.CurrentApp, Connect.Kind, ops, method);
                });
            }
            catch (Exception e)
            {
                Connect.Log("ApplyPlan 异常: " + e);
                return "{\"ok\":false,\"message\":\"" + Esc(e.Message) + "\"}";
            }
        }

        #endregion

        #region ima 资料库（第3条）

        /// <summary>测试 ima 凭据是否可用。</summary>
        public string ImaTest() { return ImaClient.TestConnection(); }

        /// <summary>列出知识库（query 空则列全部）。</summary>
        public string ImaListKb(string query) { return ImaClient.ListKnowledgeBases(query); }

        /// <summary>在指定知识库里搜索内容。</summary>
        public string ImaSearch(string kbId, string query) { return ImaClient.SearchKnowledge(kbId, query); }

        public string ImaPlan(string prompt, string libraries) { return ImaClient.PlanSearch(prompt, libraries); }

        /// <summary>按 media_id 取原文（URL 型直接抓回正文，笔记型返回 notebook_id）。</summary>
        public string ImaGetMedia(string mediaId) { return ImaClient.GetMedia(mediaId); }

        #endregion

        #region 读取文档

        /// <summary>
        /// 读取当前宿主里的文档内容。
        /// 返回 {"ok":bool,"length":int,"detail":"...","text":"...","error":"..."}
        /// </summary>
        public string ReadDocument()
        {
            try
            {
                return OnUi(ReadDocumentCore);
            }
            catch (Exception e)
            {
                Connect.Log("ReadDocument 异常: " + e);
                return Fail(e.Message);
            }
        }

        private string ReadDocumentCore()
        {
            try
            {
                object app = Connect.CurrentApp;
                if (app == null)
                    return Fail("还没取到宿主程序对象，请重启 Office 后再试。");

                string text;
                string detail;
                switch (Connect.Kind)
                {
                    case Connect.HostKind.Excel:
                        text = ReadExcel(app, out detail);
                        break;
                    case Connect.HostKind.PowerPoint:
                        text = ReadPpt(app, out detail);
                        break;
                    case Connect.HostKind.Word:
                        text = ReadWord(app, out detail);
                        break;
                    default:
                        return Fail("认不出当前宿主程序（既没有 ActiveDocument / ActiveWorkbook / ActivePresentation）。");
                }

                bool cut = false;
                if (text != null && text.Length > MaxChars)
                {
                    text = text.Substring(0, MaxChars);
                    cut = true;
                }
                if (cut) detail += "，内容过长已截断";

                Connect.Log("ReadDocument: " + detail + " 共 " + (text == null ? 0 : text.Length) + " 字");
                return "{\"ok\":true,\"length\":" + (text == null ? 0 : text.Length) +
                       ",\"detail\":\"" + Esc(detail) + "\",\"text\":\"" + Esc(text) + "\"}";
            }
            catch (Exception e)
            {
                Connect.Log("ReadDocument 异常: " + e);
                return Fail(e.Message);
            }
        }

        #region Word

        private static string ReadWord(object app, out string detail)
        {
            dynamic a = app;
            dynamic doc = a.ActiveDocument;
            if (doc == null)
                throw new InvalidOperationException("当前没有打开的文档。");

            string name = Try(() => (string)doc.Name, "(未命名)");
            string full = Try(() => (string)doc.FullName, "");

            int tables = Try(() => (int)doc.Tables.Count, 0);
            int paragraphs = Try(() => (int)doc.Paragraphs.Count, 0);

            string body = "";
            try { body = (string)doc.Content.Text; } catch { }
            if (string.IsNullOrEmpty(body))
            {
                // 有些宿主（如 WPS）对 Content.Text 支持不同，退化为逐段读取
                var sb = new StringBuilder();
                try
                {
                    dynamic ps = doc.Paragraphs;
                    int n = Math.Min((int)ps.Count, 2000);
                    for (int i = 1; i <= n; i++) sb.Append((string)ps[i].Range.Text);
                }
                catch { }
                body = sb.ToString();
            }

            // 选区单独给出，便于「针对选中的这段」
            string sel = Try(() => (string)a.Selection.Text, "");

            var outSb = new StringBuilder();
            outSb.Append("【文档】").Append(name).Append('\n');
            if (!string.IsNullOrEmpty(full)) outSb.Append("【路径】").Append(full).Append('\n');
            outSb.Append('\n').Append(Clean(body));

            detail = "Word · " + paragraphs + " 段 · " + tables + " 个表格";

            return outSb.ToString();
        }

        #endregion

        #region Excel

        private static string ReadExcel(object app, out string detail)
        {
            dynamic a = app;
            dynamic wb = a.ActiveWorkbook;
            if (wb == null)
                throw new InvalidOperationException("当前没有打开的工作簿。");

            var sb = new StringBuilder();
            string name = Try(() => (string)wb.Name, "(未命名)");
            sb.Append("【工作簿】").Append(name).Append('\n');

            int sheetCount = Try(() => (int)wb.Worksheets.Count, 0);
            sb.Append("【工作表数】").Append(sheetCount).Append('\n');
            detail = "Excel · " + sheetCount + " 个工作表";

            int totalCells = 0;
            dynamic sheets = wb.Worksheets;
            for (int i = 1; i <= sheetCount; i++)
            {
                dynamic ws = null;
                try { ws = sheets[i]; } catch { continue; }
                if (ws == null) continue;

                string sname = Try(() => (string)ws.Name, "Sheet" + i);
                sb.Append("\n===== 工作表 ").Append(i).Append("：").Append(sname)
                  .Append(" =====\n");

                object v = null;
                try { v = ws.UsedRange.Value2; } catch { }
                object[,] arr = v as object[,];
                if (arr == null)
                {
                    sb.Append("(空表)\n");
                    continue;
                }

                int rows = arr.GetLength(0);
                int cols = arr.GetLength(1);
                int rMax = Math.Min(rows, MaxRowsPerSheet);
                int cMax = Math.Min(cols, MaxColsPerSheet);

                for (int r = 1; r <= rMax; r++)
                {
                    var line = new StringBuilder();
                    for (int c = 1; c <= cMax; c++)
                    {
                        if (c > 1) line.Append('\t');
                        object cell = arr[r, c];
                        line.Append(cell == null ? "" : Convert.ToString(cell));
                        totalCells++;
                    }
                    sb.Append(line.ToString().TrimEnd('\t')).Append('\n');
                }
                if (rows > rMax || cols > cMax)
                {
                    sb.Append("…（本表共 ").Append(rows).Append(" 行 ")
                      .Append(cols).Append(" 列，已截断到 ")
                      .Append(rMax).Append(" 行 ").Append(cMax).Append(" 列）\n");
                }
            }

            detail += " · 约 " + totalCells + " 个单元格";
            return sb.ToString();
        }

        #endregion

        #region PowerPoint

        private static string ReadPpt(object app, out string detail)
        {
            dynamic a = app;
            dynamic pres = a.ActivePresentation;
            if (pres == null)
                throw new InvalidOperationException("当前没有打开的演示文稿。");

            var sb = new StringBuilder();
            string name = Try(() => (string)pres.Name, "(未命名)");
            sb.Append("【演示文稿】").Append(name).Append('\n');

            int slideCount = Try(() => (int)pres.Slides.Count, 0);
            detail = "PowerPoint · " + slideCount + " 张幻灯片";
            int shapeCount = 0;

            dynamic slides = pres.Slides;
            for (int i = 1; i <= slideCount; i++)
            {
                dynamic slide = null;
                try { slide = slides[i]; } catch { continue; }
                if (slide == null) continue;

                sb.Append("\n===== 第 ").Append(i).Append(" 张 =====\n");
                int n = Try(() => (int)slide.Shapes.Count, -1);
                if (n < 0) continue;
                for (int j = 1; j <= n; j++)
                {
                    dynamic sh = null;
                    try { sh = slide.Shapes[j]; }
                    catch (Exception ex)
                    { Connect.Log("    sh[" + j + "] 取形状失败: " + ex.Message); continue; }
                    if (sh == null) continue;

                    // 直接读文本内容，不用 HasText。
                    // 坑：PowerPoint 的 HasText 是 VARIANT_BOOL，经 dynamic 转成 bool 时
                    // 实测恒为 false，导致整个 PPT 读出来一篇空白。
                    string t;
                    try { t = Convert.ToString(sh.TextFrame.TextRange.Text); }
                    catch { continue; }
                    if (string.IsNullOrEmpty(t) || t.Trim().Length == 0) continue;
                    sb.Append(Clean(t)).Append('\n');
                    shapeCount++;
                }
            }

            detail += " · " + shapeCount + " 个文本框";
            return sb.ToString();
        }

        #endregion

        #endregion

        #region 辅助

        private static string Clean(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            // Word 的段落符 \r / \x07(单元格) 统一成换行
            return s.Replace("\r", "\n").Replace('\u0007', '\t').Replace("\u000b", "\n");
        }

        private static T Try<T>(Func<T> f, T fallback)
        {
            try { return f(); } catch { return fallback; }
        }

        private static string Fail(string msg)
        {
            return "{\"ok\":false,\"length\":0,\"detail\":\"\",\"error\":\"" + Esc(msg) + "\"}";
        }

        private static string Esc(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"")
                    .Replace("\r", "\\r").Replace("\n", "\\n")
                    .Replace("\t", "\\t");
        }

        #endregion
    }
}
