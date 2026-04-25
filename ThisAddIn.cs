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
            InitWebView();
        }

        private void InitWebView()
        {
            _webView = new WebView2();
            _webView.Dock = DockStyle.Fill;
            _webView.CoreWebView2InitializationCompleted += OnWebViewReady;
            Controls.Add(_webView);
            _webView.EnsureCoreWebView2Async();
        }

        private void OnWebViewReady(object sender, Microsoft.Web.WebView2.Core.CoreWebView2InitializationCompletedEventArgs e)
        {
            if (e.IsSuccess)
                _webView.CoreWebView2.NavigateToString(GetSidebarHtml());
        }

        private string GetSidebarHtml()
        {
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
    padding: 18px 20px 14px;
    border-bottom: 1px solid #27272a;
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .logo {
    width: 28px; height: 28px;
    background: #6366f1;
    border-radius: 6px;
    display: flex; align-items: center; justify-content: center;
    font-weight: 700; font-size: 16px; color: #fff;
    flex-shrink: 0;
  }
  header h1 { font-size: 17px; font-weight: 600; color: #fff; }
  #chat {
    flex: 1;
    overflow-y: auto;
    padding: 16px 16px 0;
    display: flex;
    flex-direction: column;
    gap: 12px;
  }
  .msg { display: flex; gap: 8px; }
  .msg.user { flex-direction: row-reverse; }
  .bubble {
    max-width: 82%;
    padding: 10px 14px;
    border-radius: 12px;
    font-size: 14px;
    line-height: 1.55;
    white-space: pre-wrap;
  }
  .msg.user .bubble { background: #6366f1; color: #fff; border-bottom-right-radius: 3px; }
  .msg.ai .bubble { background: #27272a; color: #e4e4e7; border-bottom-left-radius: 3px; }
  #input-area {
    padding: 12px 12px 14px;
    border-top: 1px solid #27272a;
    display: flex;
    gap: 8px;
    align-items: flex-end;
  }
  textarea {
    flex: 1;
    background: #27272a;
    border: 1px solid #3f3f46;
    border-radius: 8px;
    color: #e4e4e7;
    font-size: 14px;
    font-family: inherit;
    padding: 9px 12px;
    resize: none;
    outline: none;
    min-height: 40px;
    max-height: 120px;
    line-height: 1.5;
  }
  textarea:focus { border-color: #6366f1; }
  button#send {
    background: #6366f1;
    border: none;
    border-radius: 8px;
    color: #fff;
    cursor: pointer;
    padding: 9px 14px;
    font-size: 18px;
    line-height: 1;
    flex-shrink: 0;
    height: 40px;
  }
  button#send:hover { background: #4f46e5; }
  .thinking { color: #71717a; font-size: 13px; padding: 4px 0; }
</style>
</head>
<body>
<header>
  <div class='logo'>L</div>
  <h1>Lathe</h1>
</header>
<div id='chat'>
  <div class='msg ai'>
    <div class='bubble'>你好！我是 Lathe，你的 Word AI 写作助手。\n\n请输入你的问题或粘贴文档内容，我来帮你润色、翻译、总结或续写。</div>
  </div>
</div>
<div id='input-area'>
  <textarea id='inp' rows='1' placeholder='输入消息… (Enter 发送, Shift+Enter 换行)'></textarea>
  <button id='send'>&#9650;</button>
</div>
<script>
const chat = document.getElementById('chat');
const inp  = document.getElementById('inp');
const btn  = document.getElementById('send');

// API Key 暂时硬编码占位，后续换成设置界面
const API_KEY = 'YOUR_ANTHROPIC_API_KEY';

inp.addEventListener('input', () => {
  inp.style.height = 'auto';
  inp.style.height = Math.min(inp.scrollHeight, 120) + 'px';
});

inp.addEventListener('keydown', e => {
  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(); }
});
btn.addEventListener('click', send);

function addMsg(role, text) {
  const d = document.createElement('div');
  d.className = 'msg ' + role;
  const b = document.createElement('div');
  b.className = 'bubble';
  b.textContent = text;
  d.appendChild(b);
  chat.appendChild(d);
  chat.scrollTop = chat.scrollHeight;
  return b;
}

async function send() {
  const text = inp.value.trim();
  if (!text) return;
  inp.value = '';
  inp.style.height = 'auto';
  addMsg('user', text);

  const thinking = document.createElement('div');
  thinking.className = 'thinking';
  thinking.textContent = 'Lathe 正在思考…';
  chat.appendChild(thinking);
  chat.scrollTop = chat.scrollHeight;

  try {
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 1024,
        messages: [{ role: 'user', content: text }]
      })
    });
    const data = await res.json();
    thinking.remove();
    const reply = data.content && data.content[0] ? data.content[0].text : JSON.stringify(data);
    addMsg('ai', reply);
  } catch(err) {
    thinking.remove();
    addMsg('ai', '请求失败：' + err.message + '\n\n请在代码里填入你的 Anthropic API Key。');
  }
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
