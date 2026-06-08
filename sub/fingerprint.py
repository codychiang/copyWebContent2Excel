import subprocess, hashlib, sys

_EXPECTED = "fac0c302cd1aa5f610c0e3e384fc5635545a8d6d6a0234691431edb563e4cb61"

def _cpu_id():
    try:
        out = subprocess.check_output(
            'wmic cpu get ProcessorId /value', shell=True, stderr=subprocess.DEVNULL
        ).decode()
        for line in out.splitlines():
            if 'ProcessorId=' in line:
                return line.split('=', 1)[1].strip()
    except Exception:
        pass
    return ""

def verify():
    fp = hashlib.sha256(_cpu_id().encode()).hexdigest()
    if fp != _EXPECTED:
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("依賴錯誤", "此程式無法在此電腦上執行。")
            root.destroy()
        except Exception:
            pass
        sys.exit(1)
