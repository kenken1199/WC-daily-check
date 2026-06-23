"""
usb_watcher.py を Windows スタートアップに登録します。
このスクリプトを一度だけ実行してください。

削除したい場合は以下のファイルを消してください:
  %APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\Startup\\wc_usb_watcher.vbs
"""
import os
import sys

STARTUP_DIR = os.path.join(
    os.environ["APPDATA"],
    r"Microsoft\Windows\Start Menu\Programs\Startup"
)
VBS_PATH    = os.path.join(STARTUP_DIR, "wc_usb_watcher.vbs")
WATCHER_PY  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "usb_watcher.py")
PYTHONW     = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")

if not os.path.exists(PYTHONW):
    PYTHONW = sys.executable

vbs = f'Set sh = CreateObject("WScript.Shell")\n'
vbs += f'sh.Run """{PYTHONW}"" ""{WATCHER_PY}""", 0, False\n'

with open(VBS_PATH, "w", encoding="utf-8") as f:
    f.write(vbs)

print("=" * 50)
print("スタートアップ登録完了")
print(f"  登録ファイル : {VBS_PATH}")
print(f"  監視スクリプト: {WATCHER_PY}")
print()
print("次のステップ:")
print("  1. USB ドライブのボリュームラベルを確認する")
print("     （エクスプローラーでドライブを右クリック → プロパティ）")
print("  2. usb_watcher.py の USB_LABEL をそのラベルに変更する")
print("     現在の設定: USB_LABEL = 'WC_DATA'")
print("  3. PC を再起動すると自動で監視が始まります")
print("=" * 50)
