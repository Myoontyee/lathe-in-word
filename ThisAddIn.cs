using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using Extensibility;
using Microsoft.Web.WebView2.WinForms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace LatheAddIn
{
    // ── Word ↔ JS 桥接对象（暴露给 WebView2 JavaScript） ─────────────────
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class WordBridge
    {
        private Word.Application _app;
        public WordBridge(Word.Application app) { _app = app; }

        public string GetSelectedText()
        {
            try
            {
                string t = _app.Selection.Text ?? "";
                if (t.EndsWith("\r")) t = t.Substring(0, t.Length - 1);
                return t;
            }
            catch { return ""; }
        }

        public string GetDocumentText()
        {
            try { return _app.ActiveDocument.Content.Text ?? ""; }
            catch { return ""; }
        }

        public string GetDocumentTitle()
        {
            try { return _app.ActiveDocument.Name ?? ""; }
            catch { return ""; }
        }

        public double GetFontSize()
        {
            try { return _app.Selection.Font.Size; }
            catch { return 0; }
        }

        public void SetFontSize(double size)
        {
            try { _app.Selection.Font.Size = (float)size; }
            catch { }
        }

        public void SetBold(bool bold)
        {
            try { _app.Selection.Font.Bold = bold ? 1 : 0; }
            catch { }
        }

        public void SetItalic(bool italic)
        {
            try { _app.Selection.Font.Italic = italic ? 1 : 0; }
            catch { }
        }

        public void SetUnderline(bool underline)
        {
            try { _app.Selection.Font.Underline = underline ? Word.WdUnderline.wdUnderlineSingle : Word.WdUnderline.wdUnderlineNone; }
            catch { }
        }

        public int GetParagraphCount()
        {
            try { return _app.ActiveDocument.Paragraphs.Count; }
            catch { return 0; }
        }

        public string GetParagraphText(int index)
        {
            try
            {
                string t = _app.ActiveDocument.Paragraphs[index].Range.Text ?? "";
                return t.TrimEnd('\r');
            }
            catch { return ""; }
        }

        public void SelectParagraph(int index)
        {
            try { _app.ActiveDocument.Paragraphs[index].Range.Select(); }
            catch { }
        }

        public string GetAllParagraphs()
        {
            try
            {
                var doc = _app.ActiveDocument;
                var sb = new StringBuilder();
                int count = doc.Paragraphs.Count;
                for (int i = 1; i <= count; i++)
                {
                    string t = (doc.Paragraphs[i].Range.Text ?? "").TrimEnd('\r');
                    sb.Append(i).Append(": ").AppendLine(t);
                }
                return sb.ToString();
            }
            catch { return ""; }
        }

        public void InsertWithTrackChanges(string newText)
        {
            try
            {
                bool was = _app.ActiveDocument.TrackRevisions;
                _app.ActiveDocument.TrackRevisions = true;
                _app.Selection.TypeText(newText);
                if (!was) _app.ActiveDocument.TrackRevisions = false;
            }
            catch { }
        }

        public void ReplaceSelection(string newText)
        {
            try { _app.Selection.TypeText(newText); }
            catch { }
        }

        public void ReplaceParagraph(int index, string newText)
        {
            try
            {
                Word.Range r = _app.ActiveDocument.Paragraphs[index].Range;
                // don't include the paragraph mark
                r.End = r.End - 1;
                r.Text = newText;
            }
            catch { }
        }

        public void InsertNewParagraph(string text)
        {
            try
            {
                Word.Range r = _app.Selection.Range;
                r.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                r.InsertParagraphAfter();
                r.InsertAfter(text);
            }
            catch { }
        }

        public void InsertAsComment(string text)
        {
            try { _app.ActiveDocument.Comments.Add(_app.Selection.Range, text); }
            catch { }
        }

        public void SetFontColor(string hexColor)
        {
            try
            {
                hexColor = hexColor.TrimStart('#');
                int r = Convert.ToInt32(hexColor.Substring(0, 2), 16);
                int g = Convert.ToInt32(hexColor.Substring(2, 2), 16);
                int b = Convert.ToInt32(hexColor.Substring(4, 2), 16);
                _app.Selection.Font.Color = (Word.WdColor)(r | (g << 8) | (b << 16));
            }
            catch { }
        }
    }

    // ── 主加载项 ──────────────────────────────────────────────────────────
    [ComVisible(true)]
    [Guid("2D1F1F7D-3E2D-4B37-9A6E-7C8D5E9F1A11")]
    [ProgId("LatheAddIn.Connect")]
    public class ThisAddIn : IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer
    {
        public static Word.Application WordApp { get; private set; }
        private Office.CustomTaskPane _taskPane;
        private LathePanel _panel;

        public void OnConnection(object application, ext_ConnectMode connectMode,
                                 object addInInst, ref Array custom)
        {
            WordApp = (Word.Application)application;
        }

        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            WordApp = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string ribbonID)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='LatheTab' label='Lathe'>
        <group id='LatheGroup' label='Lathe'>
          <button id='OpenLatheButton' label='Lathe' size='large'
                  getImage='GetLatheImage' onAction='OnOpenLathe'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public stdole.IPictureDisp GetLatheImage(Office.IRibbonControl control)
        {
            return PictureConverter.BitmapToPicture(DrawLogo(32, 32));
        }

        public void OnOpenLathe(Office.IRibbonControl control)
        {
            if (_taskPane != null)
                _taskPane.Visible = !_taskPane.Visible;
        }

        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        {
            _taskPane = CTPFactoryInst.CreateCTP(
                "LatheAddIn.LathePanel", "Lathe", Type.Missing);
            _taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            _taskPane.Width = 440;
            _taskPane.Visible = true;
            _panel = _taskPane.ContentControl as LathePanel;
        }

        private static Bitmap DrawLogo(int w, int h)
        {
            Bitmap bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                System.Drawing.Drawing2D.GraphicsPath path = RoundRect(1, 1, w - 2, h - 2, 6);
                g.FillPath(new SolidBrush(Color.FromArgb(99, 102, 241)), path);
                using (Font f = new Font("Segoe UI", w * 0.42f, FontStyle.Bold, GraphicsUnit.Pixel))
                {
                    SizeF sz = g.MeasureString("L", f);
                    g.DrawString("L", f, Brushes.White,
                        (w - sz.Width) / 2f - 1, (h - sz.Height) / 2f);
                }
            }
            return bmp;
        }

        private static System.Drawing.Drawing2D.GraphicsPath RoundRect(int x, int y, int w, int h, int r)
        {
            System.Drawing.Drawing2D.GraphicsPath p = new System.Drawing.Drawing2D.GraphicsPath();
            p.AddArc(x, y, r * 2, r * 2, 180, 90);
            p.AddArc(x + w - r * 2, y, r * 2, r * 2, 270, 90);
            p.AddArc(x + w - r * 2, y + h - r * 2, r * 2, r * 2, 0, 90);
            p.AddArc(x, y + h - r * 2, r * 2, r * 2, 90, 90);
            p.CloseFigure();
            return p;
        }

        [ComRegisterFunction]
        public static void Register(Type type)
        {
            RegistryKey key = Registry.CurrentUser.CreateSubKey(
                @"Software\Microsoft\Office\Word\Addins\LatheAddIn.Connect");
            key.SetValue("Description", "Lathe in Word - AI 侧边栏");
            key.SetValue("FriendlyName", "Lathe in Word");
            key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
            key.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
            key.Close();
        }

        [ComUnregisterFunction]
        public static void Unregister(Type type)
        {
            try { Registry.CurrentUser.DeleteSubKeyTree(
                @"Software\Microsoft\Office\Word\Addins\LatheAddIn.Connect"); }
            catch { }
        }
    }

    // ── 侧边栏面板 ────────────────────────────────────────────────────────
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]
    [ProgId("LatheAddIn.LathePanel")]
    public class LathePanel : UserControl
    {
        private WebView2 _webView;
        private WordBridge _bridge;

        public LathePanel()
        {
            BackColor = Color.FromArgb(24, 24, 27);
            Dock = DockStyle.Fill;
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            _bridge = new WordBridge(ThisAddIn.WordApp);

            _webView = new WebView2();
            _webView.Dock = DockStyle.Fill;
            _webView.CoreWebView2InitializationCompleted += OnWebViewReady;
            Controls.Add(_webView);

            string udp = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "LatheInWord", "WebView2");
            Directory.CreateDirectory(udp);
            _webView.CreationProperties =
                new Microsoft.Web.WebView2.WinForms.CoreWebView2CreationProperties
                { UserDataFolder = udp };
            _webView.EnsureCoreWebView2Async();
        }

        private void OnWebViewReady(object sender,
            Microsoft.Web.WebView2.Core.CoreWebView2InitializationCompletedEventArgs e)
        {
            if (!e.IsSuccess) return;
            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = true;
            _webView.CoreWebView2.AddHostObjectToScript("word", _bridge);
            _webView.CoreWebView2.WebMessageReceived += OnWebMessage;

            string uiPath = WriteUiFile();
            _webView.CoreWebView2.Navigate("file:///" + uiPath.Replace('\\', '/'));
        }

        private void OnWebMessage(object sender,
            Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs e)
        {
            string msg = e.TryGetWebMessageAsString();
            if (msg == "reload")
            {
                string uiPath = WriteUiFile();
                _webView.CoreWebView2.Navigate("file:///" + uiPath.Replace('\\', '/'));
            }
        }

        private string WriteUiFile()
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "LatheInWord", "ui");
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, "index.html");
            File.WriteAllText(path, BuildHtml(), Encoding.UTF8);
            return path;
        }

        private string BuildHtml()
        {
            string apiKey = "";
            string baseUrl = "https://api.anthropic.com";
            string model = "claude-sonnet-4-6";

            try
            {
                string cfg = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                    ".claude", "settings.json");
                if (File.Exists(cfg))
                {
                    string j = File.ReadAllText(cfg, Encoding.UTF8);
                    apiKey  = JGet(j, "ANTHROPIC_AUTH_TOKEN") ?? JGet(j, "ANTHROPIC_API_KEY") ?? "";
                    baseUrl = JGet(j, "ANTHROPIC_BASE_URL") ?? baseUrl;
                    model   = JGet(j, "ANTHROPIC_DEFAULT_SONNET_MODEL") ?? JGet(j, "ANTHROPIC_MODEL") ?? model;
                }
            }
            catch { }

            baseUrl = baseUrl.TrimEnd('/');

            return HtmlTemplate
                .Replace("{{API_KEY}}", Esc(apiKey))
                .Replace("{{BASE_URL}}", Esc(baseUrl))
                .Replace("{{MODEL}}", Esc(model));
        }

        private static string Esc(string s)
        {
            return (s ?? "").Replace("\\", "\\\\").Replace("'", "\\'");
        }

        private static string JGet(string json, string key)
        {
            string s = "\"" + key + "\"";
            int i = json.IndexOf(s);
            if (i < 0) return null;
            int c = json.IndexOf(':', i + s.Length);
            if (c < 0) return null;
            int q1 = json.IndexOf('"', c + 1);
            if (q1 < 0) return null;
            int q2 = json.IndexOf('"', q1 + 1);
            if (q2 < 0) return null;
            return json.Substring(q1 + 1, q2 - q1 - 1);
        }

        [ComRegisterFunction]
        public static void Register(Type t) { }
        [ComUnregisterFunction]
        public static void Unregister(Type t) { }

        // ── 完整 UI HTML ──────────────────────────────────────────────────
        private const string HtmlTemplate = @"<!DOCTYPE html>
<html lang='zh'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width,initial-scale=1'>
<title>Lathe in Word</title>
<style>
:root{
  --bg:#18181b;--bg2:#1c1c1f;--bg3:#27272a;--border:#3f3f46;
  --text:#e4e4e7;--muted:#71717a;--accent:#6366f1;--accent2:#4f46e5;
  --user-bg:#6366f1;--ai-bg:#27272a;--green:#22c55e;--red:#ef4444;
  --tool-bg:#1a2332;--tool-border:#2d4a7a;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow:hidden}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text);
  display:flex;flex-direction:column;height:100vh}

/* ─ header ─ */
#hdr{display:flex;align-items:center;gap:8px;padding:8px 10px;
  border-bottom:1px solid var(--border);background:var(--bg2);flex-shrink:0}
.logo{width:24px;height:24px;background:var(--accent);border-radius:5px;
  display:flex;align-items:center;justify-content:center;
  font-weight:800;font-size:13px;color:#fff;flex-shrink:0}
#mdl-name{font-size:13px;font-weight:600;color:#fff;flex:1;
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.hbtn{background:var(--bg3);border:1px solid var(--border);border-radius:5px;
  color:var(--muted);font-size:11px;padding:3px 7px;cursor:pointer;white-space:nowrap}
.hbtn:hover{color:var(--text);background:var(--border)}

/* ─ sessions bar ─ */
#sess-bar{display:flex;align-items:center;gap:4px;padding:5px 8px;
  border-bottom:1px solid var(--border);background:var(--bg2);flex-shrink:0;overflow-x:auto}
.s-tab{background:transparent;border:1px solid transparent;border-radius:5px;
  color:var(--muted);font-size:12px;padding:3px 10px;cursor:pointer;white-space:nowrap;
  display:flex;align-items:center;gap:4px}
.s-tab:hover{background:var(--bg3);color:var(--text)}
.s-tab.active{background:var(--bg3);border-color:var(--border);color:var(--text)}
.s-tab .del{color:var(--muted);font-size:10px;padding:0 2px}
.s-tab .del:hover{color:var(--red)}
#new-sess{background:transparent;border:none;color:var(--muted);font-size:16px;
  cursor:pointer;padding:2px 6px;flex-shrink:0}
#new-sess:hover{color:var(--accent)}

/* ─ quick actions ─ */
#qa{display:flex;flex-wrap:wrap;gap:4px;padding:6px 8px;
  border-bottom:1px solid var(--border);flex-shrink:0}
.qa-btn{background:var(--bg3);border:1px solid var(--border);border-radius:5px;
  color:var(--text);font-size:11px;padding:4px 8px;cursor:pointer;white-space:nowrap}
.qa-btn:hover{background:var(--border);border-color:var(--accent)}
.qa-btn.accent{background:var(--accent);border-color:var(--accent);color:#fff}
.qa-btn.accent:hover{background:var(--accent2)}

/* ─ selection indicator ─ */
#sel-bar{display:none;align-items:flex-start;gap:6px;padding:5px 10px;
  background:#1a1a2e;border-bottom:1px solid #2d2d5e;flex-shrink:0}
#sel-bar.show{display:flex}
#sel-preview{font-size:11px;color:#a5b4fc;flex:1;
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:300px}
#sel-clear{background:transparent;border:none;color:var(--muted);font-size:12px;cursor:pointer}

/* ─ chat ─ */
#chat{flex:1;overflow-y:auto;padding:10px 10px 0;
  display:flex;flex-direction:column;gap:8px}
.msg{display:flex;flex-direction:column;gap:2px}
.msg.user{align-items:flex-end}
.msg.ai{align-items:flex-start}
.bubble{max-width:92%;padding:8px 12px;border-radius:12px;
  font-size:13.5px;line-height:1.65;white-space:pre-wrap;word-break:break-word}
.msg.user .bubble{background:var(--user-bg);color:#fff;border-bottom-right-radius:3px}
.msg.ai .bubble{background:var(--ai-bg);color:var(--text);border-bottom-left-radius:3px}
.msg-actions{display:flex;gap:4px;margin-top:2px;flex-wrap:wrap}
.mab{background:transparent;border:1px solid var(--border);border-radius:4px;
  color:var(--muted);font-size:10px;padding:2px 6px;cursor:pointer}
.mab:hover{color:var(--text);border-color:var(--accent)}
.thinking{color:var(--muted);font-size:12px;font-style:italic;padding:4px 0}

/* ─ tool call bubble ─ */
.tool-call{background:var(--tool-bg);border:1px solid var(--tool-border);border-radius:8px;
  padding:6px 10px;font-size:11px;color:#7dd3fc;max-width:92%;margin:2px 0}
.tool-call .tool-name{font-weight:700;color:#38bdf8}
.tool-call .tool-result{color:#86efac;margin-top:3px}

/* ─ token bar ─ */
#tok-bar{display:flex;gap:12px;align-items:center;padding:4px 10px;
  background:var(--bg2);border-top:1px solid var(--border);
  font-size:11px;color:var(--muted);flex-shrink:0}
#tok-bar b{color:var(--text)}
.tok-cloud{margin-left:auto}

/* ─ params panel ─ */
#params{padding:6px 8px;border-top:1px solid var(--border);
  background:var(--bg2);flex-shrink:0}
#params summary{font-size:11px;color:var(--muted);cursor:pointer;user-select:none;
  display:flex;align-items:center;gap:4px}
#params summary:hover{color:var(--text)}
.param-row{display:flex;align-items:center;gap:6px;margin-top:5px}
.param-row label{font-size:10px;color:var(--muted);width:60px;flex-shrink:0}
.param-row input[type=range]{flex:1;accent-color:var(--accent)}
.param-row input[type=number]{width:60px;background:var(--bg3);border:1px solid var(--border);
  border-radius:4px;color:var(--text);font-size:11px;padding:2px 4px;text-align:right}
.param-val{font-size:10px;color:var(--text);width:28px;text-align:right;flex-shrink:0}

/* ─ input ─ */
#inp-area{padding:8px 8px 10px;border-top:1px solid var(--border);
  display:flex;gap:6px;align-items:flex-end;flex-shrink:0}
textarea#inp{flex:1;background:var(--bg3);border:1px solid var(--border);border-radius:8px;
  color:var(--text);font-size:13px;font-family:inherit;padding:7px 10px;
  resize:none;outline:none;min-height:36px;max-height:100px;line-height:1.5}
textarea#inp:focus{border-color:var(--accent)}
#send{background:var(--accent);border:none;border-radius:8px;color:#fff;
  cursor:pointer;padding:0 12px;font-size:15px;flex-shrink:0;height:36px}
#send:hover{background:var(--accent2)}
#send:disabled{background:var(--bg3);cursor:default}

/* ─ modal ─ */
.modal{position:fixed;inset:0;background:rgba(0,0,0,.7);
  display:flex;align-items:center;justify-content:center;z-index:999}
.modal.hidden{display:none}
.modal-box{background:var(--bg2);border:1px solid var(--border);border-radius:10px;
  width:90%;max-height:80vh;display:flex;flex-direction:column}
.modal-hdr{display:flex;align-items:center;padding:10px 14px;
  border-bottom:1px solid var(--border)}
.modal-hdr h3{flex:1;font-size:14px}
.modal-hdr button{background:transparent;border:none;color:var(--muted);
  font-size:18px;cursor:pointer}
.modal-body{flex:1;overflow-y:auto;padding:10px 14px;
  font-size:12px;line-height:1.7;white-space:pre-wrap;color:var(--text)}
.modal-foot{padding:8px 14px;border-top:1px solid var(--border);display:flex;gap:6px}
.modal-foot button{background:var(--bg3);border:1px solid var(--border);border-radius:5px;
  color:var(--text);font-size:12px;padding:4px 12px;cursor:pointer}
.modal-foot button:hover{background:var(--border)}

/* ─ scrollbar ─ */
::-webkit-scrollbar{width:5px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}
</style>
</head>
<body>

<!-- Header -->
<div id='hdr'>
  <div class='logo'>L</div>
  <div id='mdl-name'>{{MODEL}}</div>
  <button class='hbtn' onclick='reloadConfig()' title='重新读取 cc-switch 配置'>↻ 刷新</button>
  <button class='hbtn' onclick='showCloud()'>☁ 已发</button>
  <button class='hbtn' onclick='exportMd()'>↓ MD</button>
</div>

<!-- Sessions -->
<div id='sess-bar'>
  <button id='new-sess' onclick='newSession()' title='新建会话'>＋</button>
</div>

<!-- Quick Actions -->
<div id='qa'>
  <button class='qa-btn accent' onclick='readSel()'>📖 读取选中</button>
  <button class='qa-btn' onclick='quickAct(""rewrite"")'>✏ 改写</button>
  <button class='qa-btn' onclick='quickAct(""translate"")'>🌐 翻译</button>
  <button class='qa-btn' onclick='quickAct(""summarize"")'>📋 总结</button>
  <button class='qa-btn' onclick='quickAct(""expand"")'>📝 扩写</button>
  <button class='qa-btn' onclick='quickAct(""polish"")'>✨ 润色</button>
  <button class='qa-btn' onclick='quickAct(""comment"")'>💬 批注</button>
  <button class='qa-btn' onclick='insertLast()'>⬇ 插入文档</button>
</div>

<!-- Selection indicator -->
<div id='sel-bar'>
  <span>📌 选中：</span>
  <span id='sel-preview'></span>
  <button id='sel-clear' onclick='clearSel()'>✕</button>
</div>

<!-- Chat -->
<div id='chat'></div>

<!-- Token bar -->
<div id='tok-bar'>
  <span>本地处理 <b id='t-local'>0</b> tokens</span>
  <span>已发送云端 <b id='t-sent'>0</b> tokens</span>
  <span>已接收 <b id='t-recv'>0</b> tokens</span>
  <span class='tok-cloud' id='doc-sent-indicator'></span>
</div>

<!-- Params -->
<details id='params'>
  <summary>⚙ 参数设置</summary>
  <div class='param-row'>
    <label>Temperature</label>
    <input type='range' id='p-temp' min='0' max='2' step='0.01' value='1'
      oninput=""document.getElementById('v-temp').textContent=this.value"">
    <span class='param-val' id='v-temp'>1</span>
  </div>
  <div class='param-row'>
    <label>Top P</label>
    <input type='range' id='p-topp' min='0' max='1' step='0.01' value='1'
      oninput=""document.getElementById('v-topp').textContent=this.value"">
    <span class='param-val' id='v-topp'>1</span>
  </div>
  <div class='param-row'>
    <label>Top K</label>
    <input type='range' id='p-topk' min='0' max='100' step='1' value='0'
      oninput=""document.getElementById('v-topk').textContent=this.value"">
    <span class='param-val' id='v-topk'>0</span>
  </div>
  <div class='param-row'>
    <label>Max Tokens</label>
    <input type='number' id='p-maxtok' min='100' max='8192' value='4096'>
    <span class='param-val'></span>
  </div>
</details>

<!-- Input -->
<div id='inp-area'>
  <textarea id='inp' rows='1' placeholder='输入消息… Enter发送 Shift+Enter换行'></textarea>
  <button id='send' onclick='sendMsg()'>↑</button>
</div>

<!-- Cloud modal -->
<div id='cloud-modal' class='modal hidden'>
  <div class='modal-box'>
    <div class='modal-hdr'>
      <h3>☁ 已发送到云端的内容</h3>
      <button onclick=""document.getElementById('cloud-modal').classList.add('hidden')"">✕</button>
    </div>
    <div class='modal-body' id='cloud-body'></div>
    <div class='modal-foot'>
      <button onclick=""document.getElementById('cloud-modal').classList.add('hidden')"">关闭</button>
    </div>
  </div>
</div>

<script>
// ─── Config ───────────────────────────────────────────────────────────────
var API_KEY  = '{{API_KEY}}';
var BASE_URL = '{{BASE_URL}}';
var MODEL    = '{{MODEL}}';

// ─── Word Tool Definitions (Claude Tool Use API) ───────────────────────
var WORD_TOOLS = [
  {
    name: 'get_selected_text',
    description: 'Get the currently selected text in the Word document. Use this to read what the user has highlighted.',
    input_schema: { type: 'object', properties: {}, required: [] }
  },
  {
    name: 'get_document_text',
    description: 'Get the full text content of the Word document.',
    input_schema: { type: 'object', properties: {}, required: [] }
  },
  {
    name: 'get_paragraph_count',
    description: 'Get the total number of paragraphs in the document.',
    input_schema: { type: 'object', properties: {}, required: [] }
  },
  {
    name: 'get_paragraph_text',
    description: 'Get the text of a specific paragraph by index (1-based).',
    input_schema: {
      type: 'object',
      properties: {
        index: { type: 'integer', description: 'Paragraph index, starting from 1' }
      },
      required: ['index']
    }
  },
  {
    name: 'get_all_paragraphs',
    description: 'Get all paragraphs with their index numbers. Use this to understand the document structure.',
    input_schema: { type: 'object', properties: {}, required: [] }
  },
  {
    name: 'select_paragraph',
    description: 'Select a specific paragraph by index (1-based). Must call this before modifying font/format of that paragraph.',
    input_schema: {
      type: 'object',
      properties: {
        index: { type: 'integer', description: 'Paragraph index, starting from 1' }
      },
      required: ['index']
    }
  },
  {
    name: 'get_font_size',
    description: 'Get the font size of the current selection in points.',
    input_schema: { type: 'object', properties: {}, required: [] }
  },
  {
    name: 'set_font_size',
    description: 'Set the font size of the current selection in points.',
    input_schema: {
      type: 'object',
      properties: {
        size: { type: 'number', description: 'Font size in points (e.g. 12, 14, 24)' }
      },
      required: ['size']
    }
  },
  {
    name: 'set_bold',
    description: 'Set bold formatting on the current selection.',
    input_schema: {
      type: 'object',
      properties: {
        bold: { type: 'boolean', description: 'true to bold, false to remove bold' }
      },
      required: ['bold']
    }
  },
  {
    name: 'set_italic',
    description: 'Set italic formatting on the current selection.',
    input_schema: {
      type: 'object',
      properties: {
        italic: { type: 'boolean', description: 'true for italic, false to remove' }
      },
      required: ['italic']
    }
  },
  {
    name: 'set_underline',
    description: 'Set underline formatting on the current selection.',
    input_schema: {
      type: 'object',
      properties: {
        underline: { type: 'boolean', description: 'true for underline, false to remove' }
      },
      required: ['underline']
    }
  },
  {
    name: 'replace_selection',
    description: 'Replace the current selection with new text.',
    input_schema: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'The new text to insert' }
      },
      required: ['text']
    }
  },
  {
    name: 'replace_paragraph',
    description: 'Replace the text of a specific paragraph by index (1-based).',
    input_schema: {
      type: 'object',
      properties: {
        index: { type: 'integer', description: 'Paragraph index, starting from 1' },
        text: { type: 'string', description: 'The new paragraph text' }
      },
      required: ['index', 'text']
    }
  },
  {
    name: 'insert_with_track_changes',
    description: 'Insert text at the current selection position with Track Changes enabled, so the user can review and accept/reject.',
    input_schema: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'Text to insert' }
      },
      required: ['text']
    }
  },
  {
    name: 'insert_new_paragraph',
    description: 'Insert a new paragraph after the current selection.',
    input_schema: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'Text for the new paragraph' }
      },
      required: ['text']
    }
  },
  {
    name: 'insert_as_comment',
    description: 'Add a comment annotation to the current selection in the Word document.',
    input_schema: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'Comment text' }
      },
      required: ['text']
    }
  },
  {
    name: 'set_font_color',
    description: 'Set the font color of the current selection.',
    input_schema: {
      type: 'object',
      properties: {
        color: { type: 'string', description: 'Hex color code, e.g. #FF0000 for red' }
      },
      required: ['color']
    }
  }
];

var SYSTEM_PROMPT = 'You are Lathe, an autonomous AI agent embedded in Microsoft Word. You have direct access to Word document manipulation tools.\n\nCRITICAL RULES:\n1. When asked to modify the document, USE YOUR TOOLS IMMEDIATELY. Do not ask the user for information you can retrieve yourself.\n2. If the user says ""first line"" or ""第一行"": call select_paragraph with index=1.\n3. If the user asks to change font size: first select_paragraph(index), then get_font_size(), then set_font_size(new_size).\n4. If the user asks to modify text content: use get_all_paragraphs() to find it, then replace_paragraph().\n5. NEVER ask ""what is the current font size?"" or ""which paragraph?"" - get the info yourself using tools.\n6. After completing the task, briefly report what you did (e.g. ""已将第一行字号从12pt扩大到24pt"").\n7. Chain multiple tool calls automatically to complete complex tasks.\n8. Always prefer using tools over asking the user for clarification.';

// ─── State ────────────────────────────────────────────────────────────────
var sessions   = [{ id: 1, name: 'Session 1', history: [] }];
var curSess    = 0;
var selText    = '';
var lastAiMsg  = '';
var cloudLog   = [];
var tokLocal   = 0;
var tokSent    = 0;
var tokRecv    = 0;

// ─── Init ─────────────────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', function() {
  renderSessions();
  renderChat();
  addMsg('ai', 'Lathe 已就绪 (Agent 模式)。\n\nAgent 可以直接操作 Word 文档——无需手动读取，直接下指令即可。\n\ncc-switch: ' + BASE_URL + '\n模型: ' + MODEL);
  document.getElementById('inp').addEventListener('input', autosize);
  document.getElementById('inp').addEventListener('keydown', function(e) {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendMsg(); }
  });
});

// ─── Sessions ─────────────────────────────────────────────────────────────
function newSession() {
  var id = Date.now();
  sessions.push({ id: id, name: 'Session ' + (sessions.length + 1), history: [] });
  curSess = sessions.length - 1;
  renderSessions();
  renderChat();
}

function switchSess(idx) {
  curSess = idx;
  renderSessions();
  renderChat();
}

function deleteSess(idx) {
  if (sessions.length === 1) return;
  sessions.splice(idx, 1);
  if (curSess >= sessions.length) curSess = sessions.length - 1;
  renderSessions();
  renderChat();
}

function renderSessions() {
  var bar = document.getElementById('sess-bar');
  var tabs = bar.querySelectorAll('.s-tab');
  for (var i = 0; i < tabs.length; i++) bar.removeChild(tabs[i]);
  var newBtn = document.getElementById('new-sess');
  for (var i = 0; i < sessions.length; i++) {
    (function(idx) {
      var tab = document.createElement('div');
      tab.className = 's-tab' + (idx === curSess ? ' active' : '');
      tab.textContent = sessions[idx].name;
      tab.onclick = function() { switchSess(idx); };
      if (sessions.length > 1) {
        var del = document.createElement('span');
        del.className = 'del';
        del.textContent = '×';
        del.onclick = function(e) { e.stopPropagation(); deleteSess(idx); };
        tab.appendChild(del);
      }
      bar.insertBefore(tab, newBtn);
    })(i);
  }
}

// ─── Chat rendering ───────────────────────────────────────────────────────
function renderChat() {
  var chat = document.getElementById('chat');
  chat.innerHTML = '';
  var hist = sessions[curSess].history;
  for (var i = 0; i < hist.length; i++) {
    appendMsgEl(hist[i].role, hist[i].content);
  }
  chat.scrollTop = chat.scrollHeight;
}

function addMsg(role, text) {
  sessions[curSess].history.push({ role: role, content: text });
  appendMsgEl(role, text);
  document.getElementById('chat').scrollTop = document.getElementById('chat').scrollHeight;
}

function appendMsgEl(role, text) {
  var chat = document.getElementById('chat');
  var div = document.createElement('div');
  div.className = 'msg ' + role;
  var bub = document.createElement('div');
  bub.className = 'bubble';
  bub.textContent = text;
  div.appendChild(bub);

  if (role === 'ai') {
    var acts = document.createElement('div');
    acts.className = 'msg-actions';
    var b1 = document.createElement('button');
    b1.className = 'mab'; b1.textContent = '⬇ Track Changes';
    (function(t){ b1.onclick = function() { insertToWord(t, true); }; })(text);
    var b2 = document.createElement('button');
    b2.className = 'mab'; b2.textContent = '⬇ 直接插入';
    (function(t){ b2.onclick = function() { insertToWord(t, false); }; })(text);
    var b3 = document.createElement('button');
    b3.className = 'mab'; b3.textContent = '💬 批注';
    (function(t){ b3.onclick = function() { insertComment(t); }; })(text);
    var b4 = document.createElement('button');
    b4.className = 'mab'; b4.textContent = '📋 复制';
    (function(t){ b4.onclick = function() { navigator.clipboard.writeText(t); }; })(text);
    acts.appendChild(b1);
    acts.appendChild(b2);
    acts.appendChild(b3);
    acts.appendChild(b4);
    div.appendChild(acts);
    lastAiMsg = text;
  }
  chat.appendChild(div);
}

function showToolCall(name, input, result) {
  var chat = document.getElementById('chat');
  var div = document.createElement('div');
  div.className = 'tool-call';
  var nameEl = document.createElement('div');
  nameEl.className = 'tool-name';
  nameEl.textContent = '🔧 ' + name + '(' + JSON.stringify(input) + ')';
  div.appendChild(nameEl);
  if (result !== undefined) {
    var resEl = document.createElement('div');
    resEl.className = 'tool-result';
    var rs = String(result);
    resEl.textContent = '→ ' + (rs.length > 120 ? rs.slice(0, 120) + '…' : rs);
    div.appendChild(resEl);
  }
  chat.appendChild(div);
  chat.scrollTop = chat.scrollHeight;
}

function updateTokBar() {
  document.getElementById('t-local').textContent = tokLocal;
  document.getElementById('t-sent').textContent  = tokSent;
  document.getElementById('t-recv').textContent  = tokRecv;
}

// ─── Tool Executor ────────────────────────────────────────────────────────
async function executeTool(name, input) {
  var w = window.chrome.webview.hostObjects.word;
  try {
    switch (name) {
      case 'get_selected_text':      return await w.GetSelectedText();
      case 'get_document_text':      return await w.GetDocumentText();
      case 'get_paragraph_count':    return String(await w.GetParagraphCount());
      case 'get_paragraph_text':     return await w.GetParagraphText(input.index);
      case 'get_all_paragraphs':     return await w.GetAllParagraphs();
      case 'select_paragraph':       await w.SelectParagraph(input.index); return 'selected paragraph ' + input.index;
      case 'get_font_size':          return String(await w.GetFontSize());
      case 'set_font_size':          await w.SetFontSize(input.size); return 'font size set to ' + input.size;
      case 'set_bold':               await w.SetBold(input.bold); return 'bold set to ' + input.bold;
      case 'set_italic':             await w.SetItalic(input.italic); return 'italic set to ' + input.italic;
      case 'set_underline':          await w.SetUnderline(input.underline); return 'underline set to ' + input.underline;
      case 'replace_selection':      await w.ReplaceSelection(input.text); return 'replaced selection';
      case 'replace_paragraph':      await w.ReplaceParagraph(input.index, input.text); return 'replaced paragraph ' + input.index;
      case 'insert_with_track_changes': await w.InsertWithTrackChanges(input.text); return 'inserted with track changes';
      case 'insert_new_paragraph':   await w.InsertNewParagraph(input.text); return 'inserted new paragraph';
      case 'insert_as_comment':      await w.InsertAsComment(input.text); return 'inserted comment';
      case 'set_font_color':         await w.SetFontColor(input.color); return 'color set to ' + input.color;
      default:                       return 'unknown tool: ' + name;
    }
  } catch(ex) {
    return 'error: ' + ex;
  }
}

// ─── Agent loop ───────────────────────────────────────────────────────────
async function callAPI(messages) {
  var params = {
    model: MODEL,
    max_tokens: parseInt(document.getElementById('p-maxtok').value) || 4096,
    system: SYSTEM_PROMPT,
    tools: WORD_TOOLS,
    messages: messages
  };
  var temp = parseFloat(document.getElementById('p-temp').value);
  if (temp !== 1) params.temperature = temp;
  var topp = parseFloat(document.getElementById('p-topp').value);
  if (topp !== 1) params.top_p = topp;
  var topk = parseInt(document.getElementById('p-topk').value);
  if (topk > 0) params.top_k = topk;

  return new Promise(function(resolve, reject) {
    var xhr = new XMLHttpRequest();
    xhr.open('POST', BASE_URL + '/v1/messages');
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.setRequestHeader('x-api-key', API_KEY);
    xhr.setRequestHeader('anthropic-version', '2023-06-01');
    xhr.onload = function() {
      try {
        var d = JSON.parse(xhr.responseText);
        if (d.error) reject(d.error.message || JSON.stringify(d.error));
        else resolve(d);
      } catch(ex) { reject(xhr.responseText.slice(0, 300)); }
    };
    xhr.onerror = function() { reject('网络错误'); };
    xhr.send(JSON.stringify(params));
  });
}

async function runAgent(userText) {
  var btn = document.getElementById('send');
  btn.disabled = true;

  var think = document.createElement('div');
  think.className = 'thinking';
  think.textContent = '🤖 Agent 正在思考…';
  document.getElementById('chat').appendChild(think);
  document.getElementById('chat').scrollTop = document.getElementById('chat').scrollHeight;

  // Build conversation from session history (exclude just-added user msg)
  var hist = sessions[curSess].history;
  var messages = [];
  for (var i = 0; i < hist.length - 1; i++) {
    var r = hist[i].role === 'user' ? 'user' : 'assistant';
    messages.push({ role: r, content: hist[i].content });
  }
  messages.push({ role: 'user', content: userText });

  var finalText = '';

  try {
    for (var iter = 0; iter < 15; iter++) {
      think.textContent = '🤖 Agent 正在思考… (round ' + (iter + 1) + ')';

      var resp = await callAPI(messages);

      // Read real token counts from API response
      if (resp.usage) {
        tokSent += resp.usage.input_tokens || 0;
        tokRecv += resp.usage.output_tokens || 0;
        tokLocal += (resp.usage.input_tokens || 0) + (resp.usage.output_tokens || 0);
        updateTokBar();
      }

      // Log to cloud tracker
      cloudLog.push({ role: 'api', content: JSON.stringify(resp.content).slice(0, 200), time: new Date().toLocaleTimeString() });

      var toolBlocks = [];
      var textParts = [];
      for (var ci = 0; ci < resp.content.length; ci++) {
        var block = resp.content[ci];
        if (block.type === 'tool_use') toolBlocks.push(block);
        if (block.type === 'text' && block.text) textParts.push(block.text);
      }

      // Show any text the model produces mid-stream
      if (textParts.length > 0) {
        var partialText = textParts.join('\n');
        if (resp.stop_reason === 'end_turn' || toolBlocks.length === 0) {
          finalText = partialText;
        }
      }

      // Done
      if (resp.stop_reason === 'end_turn' || toolBlocks.length === 0) {
        break;
      }

      // Execute tools
      messages.push({ role: 'assistant', content: resp.content });
      var toolResults = [];
      for (var ti = 0; ti < toolBlocks.length; ti++) {
        var tb = toolBlocks[ti];
        think.textContent = '🔧 执行: ' + tb.name + '…';
        var result = await executeTool(tb.name, tb.input || {});
        showToolCall(tb.name, tb.input || {}, result);
        toolResults.push({ type: 'tool_result', tool_use_id: tb.id, content: String(result) });
      }
      messages.push({ role: 'user', content: toolResults });
    }
  } catch(ex) {
    finalText = '错误: ' + ex;
  }

  think.remove();
  btn.disabled = false;

  if (finalText) {
    addMsg('ai', finalText);
  }
}

// ─── Word bridge helpers ───────────────────────────────────────────────────
async function readSel() {
  try {
    var text = await window.chrome.webview.hostObjects.word.GetSelectedText();
    text = (text || '').trim();
    if (!text) { addMsg('ai', '未检测到选中文字，请先在 Word 文档中选中文字。'); return; }
    selText = text;
    document.getElementById('sel-preview').textContent = text.slice(0, 80) + (text.length > 80 ? '…' : '');
    document.getElementById('sel-bar').classList.add('show');
    addMsg('ai', '已读取选中内容（' + text.length + ' 字）：\n\n' + text.slice(0, 300) + (text.length > 300 ? '\n…（已截断）' : ''));
  } catch(ex) { addMsg('ai', '读取失败: ' + ex); }
}

function clearSel() {
  selText = '';
  document.getElementById('sel-bar').classList.remove('show');
}

async function insertToWord(text, trackChanges) {
  try {
    if (trackChanges) {
      await window.chrome.webview.hostObjects.word.InsertWithTrackChanges(text);
      addMsg('ai', '✓ 已以 Track Changes 模式插入 Word。');
    } else {
      await window.chrome.webview.hostObjects.word.ReplaceSelection(text);
      addMsg('ai', '✓ 已直接插入 Word。');
    }
  } catch(ex) { addMsg('ai', '插入失败: ' + ex); }
}

async function insertComment(text) {
  try {
    await window.chrome.webview.hostObjects.word.InsertAsComment(text);
    addMsg('ai', '✓ 已以批注形式添加到 Word。');
  } catch(ex) { addMsg('ai', '批注失败: ' + ex); }
}

function insertLast() {
  if (lastAiMsg) insertToWord(lastAiMsg, true);
  else addMsg('ai', '暂无 AI 回复可插入。');
}

// ─── Quick actions ────────────────────────────────────────────────────────
var qaPrompts = {
  rewrite:   '请改写以下文字，保持原意但提升表达质量：\n\n',
  translate: '请将以下文字翻译成中文（如已是中文则翻译成英文）：\n\n',
  summarize: '请总结以下文字的核心要点：\n\n',
  expand:    '请扩写以下文字，丰富细节和论述：\n\n',
  polish:    '请润色以下文字，使其更专业流畅：\n\n',
  comment:   '请为以下文字提供专业批注和修改建议：\n\n'
};

function quickAct(type) {
  var base = qaPrompts[type] || '';
  if (selText) {
    document.getElementById('inp').value = base + selText;
  } else {
    document.getElementById('inp').value = base;
    addMsg('ai', '提示：先点击【📖 读取选中】将 Word 选中内容带入，再使用快捷操作效果更好。');
  }
  document.getElementById('inp').focus();
  autosize();
}

// ─── Send message ─────────────────────────────────────────────────────────
function sendMsg() {
  var text = document.getElementById('inp').value.trim();
  if (!text) return;
  if (!API_KEY) { addMsg('ai', '未找到 API Key，请检查 cc-switch 配置。'); return; }

  document.getElementById('inp').value = '';
  autosize();
  addMsg('user', text);

  cloudLog.push({ role: 'user', content: text, time: new Date().toLocaleTimeString() });

  runAgent(text);
}

// ─── Export Markdown ──────────────────────────────────────────────────────
function exportMd() {
  var sess = sessions[curSess];
  var lines = ['# ' + sess.name, '导出时间: ' + new Date().toLocaleString(), '---', ''];
  for (var i = 0; i < sess.history.length; i++) {
    var m = sess.history[i];
    if (m.role === 'user') {
      lines.push('**User:**');
      lines.push(m.content);
    } else {
      lines.push('**Lathe:**');
      lines.push(m.content);
    }
    lines.push('');
  }
  var md = lines.join('\n');
  var blob = new Blob([md], { type: 'text/markdown' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = 'lathe-' + sess.name.replace(/\s+/g, '-') + '-' + Date.now() + '.md';
  a.click();
  URL.revokeObjectURL(url);
}

// ─── Cloud content ────────────────────────────────────────────────────────
function showCloud() {
  var body = document.getElementById('cloud-body');
  if (!cloudLog.length) { body.textContent = '暂无发送记录。'; }
  else {
    var lines = [];
    for (var i = 0; i < cloudLog.length; i++) {
      var e = cloudLog[i];
      var label = e.role === 'user' ? '▶ 发出' : (e.role === 'api' ? '◀ API响应' : '◀ 收到');
      lines.push('[' + e.time + '] ' + label + '\n' + e.content + '\n');
    }
    body.textContent = lines.join('\n---\n');
  }
  document.getElementById('cloud-modal').classList.remove('hidden');
}

// ─── Misc ─────────────────────────────────────────────────────────────────
function reloadConfig() {
  window.chrome.webview.postMessage('reload');
}

function autosize() {
  var ta = document.getElementById('inp');
  ta.style.height = 'auto';
  ta.style.height = Math.min(ta.scrollHeight, 100) + 'px';
}
</script>
</body>
</html>";
    }

    // ── Bitmap → IPictureDisp ─────────────────────────────────────────────
    internal static class PictureConverter
    {
        [DllImport("oleaut32.dll")]
        private static extern int OleCreatePictureIndirect(
            ref PICTDESC pd, ref Guid riid, bool fOwn, out stdole.IPictureDisp ppv);

        [StructLayout(LayoutKind.Sequential)]
        private struct PICTDESC
        {
            public int cbSizeofstruct, picType;
            public IntPtr hbitmap, hpal;
            public short sReserved;
        }

        public static stdole.IPictureDisp BitmapToPicture(Bitmap bmp)
        {
            PICTDESC pd = new PICTDESC();
            pd.cbSizeofstruct = Marshal.SizeOf(pd);
            pd.picType = 1;
            pd.hbitmap = bmp.GetHbitmap();
            Guid iid = new Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB");
            stdole.IPictureDisp pic;
            OleCreatePictureIndirect(ref pd, ref iid, true, out pic);
            return pic;
        }
    }
}
