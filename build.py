"""
Lathe in Word — 构建 + 注册脚本
用法: python build.py
"""
import subprocess, shutil, os, sys

BASE = os.path.dirname(os.path.abspath(__file__))
BIN  = os.path.join(BASE, 'bin', 'Release')
CSC  = r'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
REGASM = r'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'

REFS = [
    r'C:\Windows\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll',
    r'C:\Windows\assembly\GAC\stdole\7.0.3300.0__b03f5f7f11d50a3a\stdole.dll',
    r'C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Word.dll',
    r'C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL',
    os.path.join(BASE, r'webview2\lib\net462\Microsoft.Web.WebView2.Core.dll'),
    os.path.join(BASE, r'webview2\lib\net462\Microsoft.Web.WebView2.WinForms.dll'),
    r'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.dll',
    r'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Core.dll',
    r'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Windows.Forms.dll',
    r'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Drawing.dll',
]

WEBVIEW2_DLLS = [
    os.path.join(BASE, r'webview2\lib\net462\Microsoft.Web.WebView2.Core.dll'),
    os.path.join(BASE, r'webview2\lib\net462\Microsoft.Web.WebView2.WinForms.dll'),
    os.path.join(BASE, r'webview2\build\native\x64\WebView2Loader.dll'),
]

def run(cmd, desc):
    print(f'[{desc}]', ' '.join(os.path.basename(c) for c in cmd[:2]))
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        print('FAILED:\n', r.stdout[-2000:], r.stderr[-500:])
        sys.exit(1)
    return r

# 1. 关闭 Word
subprocess.run(['taskkill', '/f', '/im', 'WINWORD.EXE'], capture_output=True)

# 2. 编译
os.makedirs(BIN, exist_ok=True)
out_dll = os.path.join(BIN, 'LatheAddIn.dll')
ref_args = ['/r:' + r for r in REFS]
cmd = [CSC, '/target:library', '/optimize+', f'/out:{out_dll}', '/langversion:5'] + ref_args + [os.path.join(BASE, 'ThisAddIn.cs')]
run(cmd, 'compile')
print('  ->', out_dll)

# 3. 复制 WebView2 依赖
for src in WEBVIEW2_DLLS:
    dst = os.path.join(BIN, os.path.basename(src))
    shutil.copy2(src, dst)
    print('  copy', os.path.basename(src))

# 4. 注册
run([REGASM, out_dll, '/codebase', '/silent'], 'regasm')

print('\n✓ 构建并注册完成，请重新打开 Word。')
