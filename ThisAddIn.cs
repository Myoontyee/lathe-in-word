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

namespace LatheAddIn
{
    [ComVisible(true)]
    [Guid("2D1F1F7D-3E2D-4B37-9A6E-7C8D5E9F1A11")]
    [ProgId("LatheAddIn.Connect")]
    public class ThisAddIn : IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer
    {
        private object _addInInstance;
        private Office.CustomTaskPane _taskPane;

        public void OnConnection(object application, ext_ConnectMode connectMode,
                                 object addInInst, ref Array custom)
        {
            _addInInstance = addInInst;
        }

        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            _addInInstance = null;
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
            return PictureConverter.BitmapToPicture(DrawLatheLogo(32, 32));
        }

        public void OnOpenLathe(Office.IRibbonControl control)
        {
            if (_taskPane != null)
                _taskPane.Visible = !_taskPane.Visible;
        }

        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        {
            _taskPane = CTPFactoryInst.CreateCTP(
                "LatheAddIn.LathePanel",
                "Lathe",
                Type.Missing);

            _taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            _taskPane.Width = 420;
            _taskPane.Visible = true;
        }

        private static Bitmap DrawLatheLogo(int w, int h)
        {
            Bitmap bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                using (System.Drawing.Drawing2D.GraphicsPath path = RoundRect(1, 1, w - 2, h - 2, 6))
                using (SolidBrush bg = new SolidBrush(Color.FromArgb(99, 102, 241)))
                    g.FillPath(bg, path);
                using (Font f = new Font("Segoe UI", w * 0.42f, FontStyle.Bold, GraphicsUnit.Pixel))
                using (SolidBrush wb = new SolidBrush(Color.White))
                {
                    SizeF sz = g.MeasureString("L", f);
                    g.DrawString("L", f, wb, (w - sz.Width) / 2f - 1, (h - sz.Height) / 2f);
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

        public LathePanel()
        {
            BackColor = Color.FromArgb(24, 24, 27);
            Dock = DockStyle.Fill;
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            _webView = new WebView2();
            _webView.Dock = DockStyle.Fill;
            _webView.CoreWebView2InitializationCompleted += OnWebViewReady;
            Controls.Add(_webView);

            string udp = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "LatheInWord", "WebView2");
            Directory.CreateDirectory(udp);

            _webView.CreationProperties = new Microsoft.Web.WebView2.WinForms.CoreWebView2CreationProperties
            {
                UserDataFolder = udp
            };
            _webView.EnsureCoreWebView2Async();
        }

        private void OnWebViewReady(object sender,
            Microsoft.Web.WebView2.Core.CoreWebView2InitializationCompletedEventArgs e)
        {
            if (!e.IsSuccess) return;
            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = false;
            _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
            _webView.CoreWebView2.NavigateToString(BuildHtml());
        }

        // 供刷新按钮调用（从 JS 通过 WebMessage 触发）
        private void Reload()
        {
            if (_webView != null && _webView.CoreWebView2 != null)
                _webView.CoreWebView2.NavigateToString(BuildHtml());
        }

        private string BuildHtml()
        {
            string apiKey = "";
            string baseUrl = "https://api.anthropic.com";
            string model = "claude-sonnet-4-6";

            try
            {
                string settingsPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                    ".claude", "settings.json");

                if (File.Exists(settingsPath))
                {
                    string json = File.ReadAllText(settingsPath, Encoding.UTF8);
                    apiKey  = JsonGet(json, "ANTHROPIC_AUTH_TOKEN")
                           ?? JsonGet(json, "ANTHROPIC_API_KEY") ?? "";
                    baseUrl = JsonGet(json, "ANTHROPIC_BASE_URL") ?? baseUrl;
                    model   = JsonGet(json, "ANTHROPIC_DEFAULT_SONNET_MODEL")
                           ?? JsonGet(json, "ANTHROPIC_MODEL") ?? model;
                }
            }
            catch { }

            baseUrl = baseUrl.TrimEnd('/');
            string statusMsg = apiKey.Length > 0
                ? "已从 cc-switch 自动读取配置，直接开始对话。"
                : "未找到配置，请检查 cc-switch 是否已启用。";

            StringBuilder sb = new StringBuilder();
            sb.Append("<!DOCTYPE html><html lang='zh'><head><meta charset='utf-8'>");
            sb.Append("<meta name='viewport' content='width=device-width,initial-scale=1'>");
            sb.Append("<title>Lathe</title><style>");
            sb.Append("*{box-sizing:border-box;margin:0;padding:0}");
            sb.Append("body{font-family:'Segoe UI',system-ui,sans-serif;background:#18181b;color:#e4e4e7;height:100vh;display:flex;flex-direction:column}");
            sb.Append("header{padding:10px 12px;border-bottom:1px solid #27272a;display:flex;align-items:center;gap:8px}");
            sb.Append(".logo{width:22px;height:22px;background:#6366f1;border-radius:5px;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;color:#fff;flex-shrink:0}");
            sb.Append("h1{font-size:14px;font-weight:600;color:#fff;flex:1}");
            sb.Append(".mdl{font-size:10px;color:#52525b;max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}");
            sb.Append("button.reload{background:#27272a;border:1px solid #3f3f46;border-radius:6px;color:#a1a1aa;font-size:12px;padding:3px 8px;cursor:pointer}");
            sb.Append("button.reload:hover{background:#3f3f46;color:#fff}");
            sb.Append("#chat{flex:1;overflow-y:auto;padding:10px 10px 0;display:flex;flex-direction:column;gap:8px}");
            sb.Append(".msg{display:flex}.msg.user{flex-direction:row-reverse}");
            sb.Append(".bubble{max-width:88%;padding:7px 11px;border-radius:12px;font-size:13px;line-height:1.6;white-space:pre-wrap;word-break:break-word}");
            sb.Append(".msg.user .bubble{background:#6366f1;color:#fff;border-bottom-right-radius:3px}");
            sb.Append(".msg.ai .bubble{background:#27272a;color:#e4e4e7;border-bottom-left-radius:3px}");
            sb.Append("#inp-area{padding:8px 8px 10px;border-top:1px solid #27272a;display:flex;gap:6px;align-items:flex-end}");
            sb.Append("textarea{flex:1;background:#27272a;border:1px solid #3f3f46;border-radius:8px;color:#e4e4e7;font-size:13px;font-family:inherit;padding:6px 10px;resize:none;outline:none;min-height:34px;max-height:96px;line-height:1.5}");
            sb.Append("textarea:focus{border-color:#6366f1}");
            sb.Append("#send{background:#6366f1;border:none;border-radius:8px;color:#fff;cursor:pointer;padding:0 11px;font-size:14px;flex-shrink:0;height:34px}");
            sb.Append("#send:hover{background:#4f46e5}#send:disabled{background:#3f3f46;cursor:default}");
            sb.Append(".thinking{color:#52525b;font-size:12px;padding:2px 0;font-style:italic}");
            sb.Append("</style></head><body>");
            sb.Append("<header><div class='logo'>L</div><h1>Lathe</h1>");
            sb.Append("<span class='mdl'>" + model + "</span>");
            sb.Append("<button class='reload' onclick='reload()' title='重新读取 cc-switch 配置'>↻</button>");
            sb.Append("</header>");
            sb.Append("<div id='chat'><div class='msg ai'><div class='bubble'>" + statusMsg + "</div></div></div>");
            sb.Append("<div id='inp-area'><textarea id='inp' rows='1' placeholder='输入消息… (Enter发送 Shift+Enter换行)'></textarea>");
            sb.Append("<button id='send'>↑</button></div>");
            sb.Append("<script>");
            sb.Append("var K='" + apiKey.Replace("'", "\\'") + "';");
            sb.Append("var U='" + baseUrl + "';");
            sb.Append("var M='" + model + "';");
            sb.Append("var inp=document.getElementById('inp');");
            sb.Append("var btn=document.getElementById('send');");
            sb.Append("var chat=document.getElementById('chat');");
            sb.Append("function reload(){window.chrome.webview.postMessage('reload');}");
            sb.Append("inp.addEventListener('input',function(){inp.style.height='auto';inp.style.height=Math.min(inp.scrollHeight,96)+'px';});");
            sb.Append("inp.addEventListener('keydown',function(e){if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();send();}});");
            sb.Append("btn.addEventListener('click',send);");
            sb.Append("function addMsg(r,t){var d=document.createElement('div');d.className='msg '+r;var b=document.createElement('div');b.className='bubble';b.textContent=t;d.appendChild(b);chat.appendChild(d);chat.scrollTop=chat.scrollHeight;return b;}");
            sb.Append("function send(){");
            sb.Append("  var t=inp.value.trim();if(!t)return;");
            sb.Append("  if(!K){addMsg('ai','未找到 API Key，请检查 cc-switch 配置。');return;}");
            sb.Append("  inp.value='';inp.style.height='auto';addMsg('user',t);btn.disabled=true;");
            sb.Append("  var th=document.createElement('div');th.className='thinking';th.textContent='正在思考…';chat.appendChild(th);chat.scrollTop=chat.scrollHeight;");
            sb.Append("  var x=new XMLHttpRequest();x.open('POST',U+'/v1/messages');");
            sb.Append("  x.setRequestHeader('Content-Type','application/json');");
            sb.Append("  x.setRequestHeader('x-api-key',K);");
            sb.Append("  x.setRequestHeader('anthropic-version','2023-06-01');");
            sb.Append("  x.onload=function(){th.remove();btn.disabled=false;");
            sb.Append("    try{var d=JSON.parse(x.responseText);addMsg('ai',d.content&&d.content[0]?d.content[0].text:x.responseText);}");
            sb.Append("    catch(ex){addMsg('ai','解析失败: '+x.responseText.slice(0,200));}");
            sb.Append("  };");
            sb.Append("  x.onerror=function(){th.remove();btn.disabled=false;addMsg('ai','网络错误。');};");
            sb.Append("  x.send(JSON.stringify({model:M,max_tokens:1024,messages:[{role:'user',content:t}]}));");
            sb.Append("}");
            sb.Append("</script></body></html>");
            return sb.ToString();
        }

        // WebMessage 接收：JS 发来 'reload' 时重新读取配置
        private void OnWebMessageReceived(object sender,
            Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs e)
        {
            if (e.TryGetWebMessageAsString() == "reload")
                _webView.CoreWebView2.NavigateToString(BuildHtml());
        }

        private static string JsonGet(string json, string key)
        {
            string search = "\"" + key + "\"";
            int i = json.IndexOf(search);
            if (i < 0) return null;
            int c = json.IndexOf(':', i + search.Length);
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
            public int cbSizeofstruct;
            public int picType;
            public IntPtr hbitmap;
            public IntPtr hpal;
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
