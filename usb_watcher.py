"""
USB 監視ツール
指定ラベルの USB が挿入されると RECORD フォルダを開いた状態で
analyze.py を起動し、ユーザーが日付フォルダを選択できるようにします。

使い方:
  1. USB_LABEL を USB ドライブのボリュームラベルに合わせる
  2. setup_startup.py を一度だけ実行してスタートアップに登録する
"""
import sys
import os
import time
import string
import datetime
import subprocess
import ctypes

# ============================================================
# ★ 設定 ★
USB_LABEL      = "緑3F瓶充填"   # USB のボリュームラベル（エクスプローラーで確認・変更可）
CHECK_INTERVAL = 3              # 監視間隔（秒）
# ============================================================

SCRIPT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "analyze.py")
LOG_PATH    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "usb_watcher.log")
_kernel32   = ctypes.windll.kernel32
DRIVE_REMOVABLE = 2

_LOCK_FILE = os.path.join(os.environ.get("TEMP", os.path.dirname(os.path.abspath(__file__))),
                          "wc_usb_watcher.pid")


def _log(msg):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}\n"
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line)
    print(line, end="")


def _single_instance_check():
    """同じスクリプトが既に動いていれば終了する（PIDファイルでロック）"""
    my_pid = os.getpid()
    if os.path.exists(_LOCK_FILE):
        try:
            old_pid = int(open(_LOCK_FILE).read().strip())
            h = _kernel32.OpenProcess(0x0400, False, old_pid)
            if h:
                _kernel32.CloseHandle(h)
                _log(f"既に PID={old_pid} が起動中のため終了します")
                sys.exit(0)
        except Exception:
            pass
    with open(_LOCK_FILE, "w") as f:
        f.write(str(my_pid))
    _log(f"起動 PID={my_pid}  SCRIPT_PATH={SCRIPT_PATH}")


def _get_removable_drives():
    result = {}
    bitmask = _kernel32.GetLogicalDrives()
    for letter in string.ascii_uppercase:
        if bitmask & 1:
            drive = f"{letter}:\\"
            if _kernel32.GetDriveTypeW(drive) == DRIVE_REMOVABLE:
                label_buf = ctypes.create_unicode_buffer(1024)
                _kernel32.GetVolumeInformationW(
                    drive, label_buf, 1024, None, None, None, None, 0
                )
                result[drive] = label_buf.value
        bitmask >>= 1
    return result


def _find_record_folder(drive):
    path = os.path.join(drive, "RECORD")
    return path if os.path.isdir(path) else drive


def main():
    _single_instance_check()

    pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if not os.path.exists(pythonw):
        pythonw = sys.executable

    seen = {}   # drive -> label（ラベル確定済みのドライブ）

    _log("監視開始")

    while True:
        time.sleep(CHECK_INTERVAL)
        current = _get_removable_drives()

        # 抜かれたドライブを解除（再挿入に対応）
        for drive in list(seen):
            if drive not in current:
                _log(f"ドライブ抜去: {drive}")
                del seen[drive]

        for drive, label in current.items():
            if drive in seen:
                continue
            if label == "":
                _log(f"ラベル未確定: {drive} → 次サイクルで再チェック")
                continue
            seen[drive] = label
            _log(f"新ドライブ確定: {drive}  label={repr(label)}  一致={label == USB_LABEL}")
            if label != USB_LABEL:
                continue
            folder = _find_record_folder(drive)
            _log(f"RECORDフォルダ: {folder}")
            cmd = [pythonw, SCRIPT_PATH, "--browse", folder]
            _log(f"起動コマンド: {cmd}")
            subprocess.Popen(cmd)


if __name__ == "__main__":
    main()
