using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
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

        // ── IDTExtensibility2 ─────────────────────────────────────────────
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

        // ── IRibbonExtensibility ──────────────────────────────────────────
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

        // ── ICustomTaskPaneConsumer ───────────────────────────────────────
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

        // ── Logo 绘制（32×32 SVG 风格矢量字母 L） ─────────────────────────
        private static Bitmap DrawLatheLogo(int w, int h)
        {
            Bitmap bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                // 圆角背景
                using (System.Drawing.Drawing2D.GraphicsPath path = RoundRect(1, 1, w - 2, h - 2, 6))
                using (SolidBrush bg = new SolidBrush(Color.FromArgb(99, 102, 241)))
                    g.FillPath(bg, path);

                // 白色 "L" 字
                using (Font f = new Font("Segoe UI", w * 0.42f, FontStyle.Bold, GraphicsUnit.Pixel))
                using (SolidBrush wb = new SolidBrush(Color.White))
                {
                    SizeF sz = g.MeasureString("L", f);
                    g.DrawString("L", f, wb,
                        (w - sz.Width) / 2f - 1,
                        (h - sz.Height) / 2f);
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

        // ── COM 注册 / 注销 ───────────────────────────────────────────────
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
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(
                    @"Software\Microsoft\Office\Word\Addins\LatheAddIn.Connect");
            }
            catch { }
        }
    }

    // ── 侧边栏面板：WebView2 ──────────────────────────────────────────────
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

        // 必须等 HWND 创建后再初始化 WebView2，否则一片黑
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            _webView = new WebView2();
            _webView.Dock = DockStyle.Fill;
            _webView.CoreWebView2InitializationCompleted += OnWebViewReady;
            Controls.Add(_webView);

            string udp = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "LatheInWord", "WebView2");
            System.IO.Directory.CreateDirectory(udp);

            // 直接传路径字符串，避免 GetAwaiter().GetResult() 在 COM STA 线程死锁
            _webView.CreationProperties = new Microsoft.Web.WebView2.WinForms.CoreWebView2CreationProperties
            {
                UserDataFolder = udp
            };
            _webView.EnsureCoreWebView2Async();
        }

        private void OnWebViewReady(object sender, Microsoft.Web.WebView2.Core.CoreWebView2InitializationCompletedEventArgs e)
        {
            if (!e.IsSuccess) return;

            // 禁用右键菜单和开发者工具（生产环境）
            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = false;

            _webView.CoreWebView2.NavigateToString(GetSidebarHtml());
        }

        private string GetSidebarHtml()
        {
            // 读取 Claude Code CLI 设置的环境变量，自动填入 Key
            string apiKey = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY") ?? "";

            return @"<!DOCTYPE html>
<html lang='zh'>
<head>
<meta charset='utf-8'>
<meta name='viewport' content='width=device-width,initial-scale=1'>
<title>Lathe</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Segoe UI', system-ui, sans-serif;
    background: #18181b;
    color: #e4e4e7;
    height: 100vh;
    display: flex;
    flex-direction: column;
  }
  header {
    padding: 14px 16px 12px;
    border-bottom: 1px solid #27272a;
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .logo {
    width: 26px; height: 26px;
    background: #6366f1;
    border-radius: 6px;
    display: flex; align-items: center; justify-content: center;
    font-weight: 700; font-size: 15px; color: #fff;
    flex-shrink: 0;
  }
  header h1 { font-size: 16px; font-weight: 600; color: #fff; flex: 1; }
  #key-bar {
    padding: 8px 12px;
    background: #1c1c1f;
    border-bottom: 1px solid #27272a;
    display: flex;
    gap: 6px;
    align-items: center;
  }
  #key-bar input {
    flex: 1;
    background: #27272a;
    border: 1px solid #3f3f46;
    border-radius: 6px;
    color: #a1a1aa;
    font-size: 11px;
    padding: 5px 8px;
    outline: none;
    font-family: monospace;
  }
  #key-bar button {
    background: #3f3f46;
    border: none; border-radius: 6px;
    color: #d4d4d8; font-size: 11px;
    padding: 5px 10px; cursor: pointer;
  }
  #key-bar button:hover { background: #52525b; }
  #chat {
    flex: 1;
    overflow-y: auto;
    padding: 14px 14px 0;
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  .msg { display: flex; gap: 8px; }
  .msg.user { flex-direction: row-reverse; }
  .bubble {
    max-width: 84%;
    padding: 9px 13px;
    border-radius: 12px;
    font-size: 13.5px;
    line-height: 1.6;
    white-space: pre-wrap;
    word-break: break-word;
  }
  .msg.user .bubble { background: #6366f1; color: #fff; border-bottom-right-radius: 3px; }
  .msg.ai   .bubble { background: #27272a; color: #e4e4e7; border-bottom-left-radius: 3px; }
  #input-area {
    padding: 10px 10px 12px;
    border-top: 1px solid #27272a;
    display: flex;
    gap: 7px;
    align-items: flex-end;
  }
  textarea {
    flex: 1;
    background: #27272a;
    border: 1px solid #3f3f46;
    border-radius: 8px;
    color: #e4e4e7;
    font-size: 13.5px;
    font-family: inherit;
    padding: 8px 11px;
    resize: none;
    outline: none;
    min-height: 38px;
    max-height: 110px;
    line-height: 1.5;
  }
  textarea:focus { border-color: #6366f1; }
  #send {
    background: #6366f1;
    border: none; border-radius: 8px;
    color: #fff; cursor: pointer;
    padding: 0 13px; font-size: 16px;
    flex-shrink: 0; height: 38px;
  }
  #send:hover { background: #4f46e5; }
  #send:disabled { background: #3f3f46; cursor: default; }
  .thinking { color: #52525b; font-size: 12px; padding: 2px 0; font-style: italic; }
</style>
</head>
<body>
<header>
  <div class='logo'>L</div>
  <h1>Lathe</h1>
</header>
<div id='key-bar'>
  <input id='apikey' type='password' placeholder='Anthropic API Key (sk-ant-…)' value='" + apiKey + @"'>
  <button onclick='saveKey()'>保存</button>
</div>
<div id='chat'>
  <div class='msg ai'>
    <div class='bubble'>你好！我是 Lathe。" + (apiKey.Length > 0 ? "已从环境变量自动读取 API Key，可以直接开始对话。" : "请先在上方填入 Anthropic API Key。") + @"</div>
  </div>
</div>
<div id='input-area'>
  <textarea id='inp' rows='1' placeholder='输入消息… (Enter 发送, Shift+Enter 换行)'></textarea>
  <button id='send'>↑</button>
</div>
<script>
var apiKey = document.getElementById('apikey').value;

function saveKey() {
  apiKey = document.getElementById('apikey').value.trim();
  addMsg('ai', 'API Key 已更新。');
}

var inp = document.getElementById('inp');
var btn = document.getElementById('send');
var chat = document.getElementById('chat');

inp.addEventListener('input', function() {
  inp.style.height = 'auto';
  inp.style.height = Math.min(inp.scrollHeight, 110) + 'px';
});
inp.addEventListener('keydown', function(e) {
  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(); }
});
btn.addEventListener('click', send);

function addMsg(role, text) {
  var d = document.createElement('div');
  d.className = 'msg ' + role;
  var b = document.createElement('div');
  b.className = 'bubble';
  b.textContent = text;
  d.appendChild(b);
  chat.appendChild(d);
  chat.scrollTop = chat.scrollHeight;
  return b;
}

function send() {
  var text = inp.value.trim();
  if (!text) return;
  if (!apiKey) { addMsg('ai', '请先填入 API Key。'); return; }
  inp.value = '';
  inp.style.height = 'auto';
  addMsg('user', text);
  btn.disabled = true;

  var thinking = document.createElement('div');
  thinking.className = 'thinking';
  thinking.textContent = '正在思考…';
  chat.appendChild(thinking);
  chat.scrollTop = chat.scrollHeight;

  var xhr = new XMLHttpRequest();
  xhr.open('POST', 'https://api.anthropic.com/v1/messages');
  xhr.setRequestHeader('Content-Type', 'application/json');
  xhr.setRequestHeader('x-api-key', apiKey);
  xhr.setRequestHeader('anthropic-version', '2023-06-01');
  xhr.onload = function() {
    thinking.remove();
    btn.disabled = false;
    try {
      var data = JSON.parse(xhr.responseText);
      var reply = data.content && data.content[0] ? data.content[0].text : xhr.responseText;
      addMsg('ai', reply);
    } catch(ex) {
      addMsg('ai', '解析失败: ' + xhr.responseText.slice(0, 200));
    }
  };
  xhr.onerror = function() {
    thinking.remove();
    btn.disabled = false;
    addMsg('ai', '网络错误，请检查连接。');
  };
  xhr.send(JSON.stringify({
    model: 'claude-sonnet-4-6',
    max_tokens: 1024,
    messages: [{ role: 'user', content: text }]
  }));
}
</script>
</body>
</html>";
        }

        [ComRegisterFunction]
        public static void Register(Type t) { }

        [ComUnregisterFunction]
        public static void Unregister(Type t) { }
    }

    // ── Bitmap → IPictureDisp 转换工具 ────────────────────────────────────
    internal static class PictureConverter
    {
        [DllImport("oleaut32.dll")]
        private static extern int OleCreatePictureIndirect(
            ref PICTDESC pPictDesc, ref Guid riid, bool fOwn, out stdole.IPictureDisp ppvObj);

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
            pd.picType = 1; // PICTYPE_BITMAP
            pd.hbitmap = bmp.GetHbitmap();
            Guid iid = new Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB");
            stdole.IPictureDisp pic;
            OleCreatePictureIndirect(ref pd, ref iid, true, out pic);
            return pic;
        }
    }
}
