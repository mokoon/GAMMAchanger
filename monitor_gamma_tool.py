import tkinter as tk
from tkinter import ttk, messagebox
import ctypes
from ctypes import wintypes
import numpy as np
import keyboard
import json
import os
import atexit  # 프로그램 종료 이벤트를 감지하기 위해 추가

gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32
user32.GetDC.argtypes = [wintypes.HWND]
user32.GetDC.restype = wintypes.HDC
gdi32.SetDeviceGammaRamp.argtypes = [wintypes.HDC, ctypes.c_void_p]
gdi32.SetDeviceGammaRamp.restype = wintypes.BOOL
gdi32.CreateDCW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_void_p]
gdi32.CreateDCW.restype = wintypes.HDC
gdi32.DeleteDC.argtypes = [wintypes.HDC]
gdi32.DeleteDC.restype = wintypes.BOOL

class GammaController:
    def __init__(self, root):
        self.root = root
        self.root.title("Monitor Gamma Tool - Final Safe Version")
        self.root.geometry("480x650")
        self.root.attributes('-topmost', True)

        self.brightness = tk.DoubleVar(value=1.0)
        self.contrast = tk.DoubleVar(value=1.0)
        self.gamma = tk.DoubleVar(value=1.0)
        self.new_preset_hk = tk.StringVar(value="f1")
        self.reset_hk_var = tk.StringVar(value="f9")
        
        self.config_file = "gamma_config.json"
        self.data = self.load_data()
        self.presets = self.data.get("presets", {})
        self.reset_hk_var.set(self.data.get("reset_hotkey", "f9"))

        self.setup_ui()
        self.refresh_all_hotkeys()
        self.apply_gamma()

        # 프로그램이 비정상 종료되어도 감마를 복구하도록 등록
        atexit.register(self.force_reset)

    def setup_ui(self):
        adj_frame = ttk.LabelFrame(self.root, text="실시간 조절")
        adj_frame.pack(padx=10, pady=5, fill="x")
        self.create_slider(adj_frame, "밝기", self.brightness, 0.5, 2.0)
        self.create_slider(adj_frame, "명비", self.contrast, 0.5, 2.0)
        self.create_slider(adj_frame, "감마", self.gamma, 0.1, 3.0)

        input_frame = ttk.LabelFrame(self.root, text="새 프리셋 추가")
        input_frame.pack(padx=10, pady=5, fill="x")
        ttk.Label(input_frame, text="이름:").grid(row=0, column=0, padx=5, pady=10, sticky="e")
        self.name_entry = ttk.Entry(input_frame, width=15)
        self.name_entry.grid(row=0, column=1, padx=5, pady=10, sticky="w")
        ttk.Button(input_frame, text="프리셋 저장", command=self.save_preset, width=12).grid(row=0, column=2, padx=5, pady=10)
        ttk.Label(input_frame, text="실행 키:").grid(row=1, column=0, padx=5, pady=10, sticky="e")
        ttk.Entry(input_frame, textvariable=self.new_preset_hk, state="readonly", width=15).grid(row=1, column=1, padx=5, pady=10, sticky="w")
        self.btn_set_new_hk = ttk.Button(input_frame, text="키 지정", command=lambda: self.listen_for_key(self.new_preset_hk, self.btn_set_new_hk))
        self.btn_set_new_hk.grid(row=1, column=2, padx=5, pady=10)

        list_frame = ttk.LabelFrame(self.root, text="저장된 프리셋 목록")
        list_frame.pack(padx=10, pady=5, fill="both", expand=True)
        self.tree = ttk.Treeview(list_frame, columns=("name", "hk"), show="headings", height=8)
        self.tree.heading("name", text="프리셋 이름")
        self.tree.heading("hk", text="단축키")
        self.tree.column("name", width=250)
        self.tree.column("hk", width=120)
        self.tree.pack(padx=5, pady=5, fill="both", expand=True)
        self.update_treeview()
        ttk.Button(list_frame, text="선택 항목 삭제", command=self.delete_preset).pack(pady=5)

        global_frame = ttk.LabelFrame(self.root, text="복구 설정")
        global_frame.pack(padx=10, pady=5, fill="x")
        ttk.Label(global_frame, text="복구 키:").pack(side="left", padx=5)
        ttk.Entry(global_frame, textvariable=self.reset_hk_var, state="readonly", width=10).pack(side="left", padx=5)
        self.btn_set_reset_hk = ttk.Button(global_frame, text="키 지정", command=lambda: self.listen_for_key(self.reset_hk_var, self.btn_set_reset_hk))
        self.btn_set_reset_hk.pack(side="left", padx=5)

    def create_slider(self, parent, label, var, start, end):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", padx=5, pady=2)
        ttk.Label(frame, text=label, width=8).pack(side="left")
        ttk.Scale(frame, from_=start, to=end, variable=var, orient="horizontal", command=lambda e: self.apply_gamma()).pack(side="left", fill="x", expand=True)
        label_val = ttk.Label(frame, text=f"{var.get():.2f}", width=5)
        label_val.pack(side="right")
        var.trace_add("write", lambda *args: label_val.config(text=f"{var.get():.2f}"))

    def listen_for_key(self, target_var, button_widget):
        button_widget.config(text="키 누르세요", state="disabled")
        def on_key_event(event):
            target_var.set(event.name)
            keyboard.unhook(hook_id)
            button_widget.config(text="키 지정", state="normal")
            self.refresh_all_hotkeys()
            return False
        hook_id = keyboard.on_press(on_key_event)

    def apply_gamma(self, *args):
        bv, cv, gv = self.brightness.get(), self.contrast.get(), self.gamma.get()
        ramp = np.zeros((3, 256), dtype=np.uint16)
        for i in range(256):
            v = pow(i / 255.0, 1/gv)
            res = ((v - 0.5) * cv + 0.5) * bv
            val = int(np.clip(res, 0, 1) * 65535)
            ramp[0][i] = ramp[1][i] = ramp[2][i] = val

        hdc = gdi32.CreateDCW("DISPLAY", "\\\\.\\DISPLAY1", None, None)
        if hdc:
            gdi32.SetDeviceGammaRamp(hdc, ramp.ctypes.data)
            gdi32.DeleteDC(hdc)

    # UI 없이 API만 직접 호출하는 강제 복구 함수
    def force_reset(self):
        ramp = np.zeros((3, 256), dtype=np.uint16)
        for i in range(256):
            val = int(i / 255.0 * 65535)
            ramp[0][i] = ramp[1][i] = ramp[2][i] = val
        hdc = gdi32.CreateDCW("DISPLAY", "\\\\.\\DISPLAY1", None, None)
        if hdc:
            gdi32.SetDeviceGammaRamp(hdc, ramp.ctypes.data)
            gdi32.DeleteDC(hdc)

    def make_handler(self, b, c, g):
        def handler():
            self.root.after(0, lambda: [self.brightness.set(b), self.contrast.set(c), self.gamma.set(g), self.apply_gamma()])
        return handler

    def refresh_all_hotkeys(self):
        try:
            keyboard.unhook_all()
            for name, p in self.presets.items():
                hk = p.get("hotkey")
                if hk:
                    keyboard.add_hotkey(hk, self.make_handler(p['b'], p['c'], p['g']), suppress=False)
            reset_hk = self.reset_hk_var.get()
            if reset_hk:
                keyboard.add_hotkey(reset_hk, self.make_handler(1.0, 1.0, 1.0), suppress=False)
            self.save_data()
        except Exception as e:
            print(f"Hotkey Error: {e}")

    def save_preset(self):
        name = self.name_entry.get().strip()
        if not name: return
        self.presets[name] = {
            "b": round(self.brightness.get(), 2), 
            "c": round(self.contrast.get(), 2),
            "g": round(self.gamma.get(), 2), 
            "hotkey": self.new_preset_hk.get()
        }
        self.update_treeview()
        self.refresh_all_hotkeys()
        self.name_entry.delete(0, tk.END)

    def delete_preset(self):
        selected_item = self.tree.selection()
        if not selected_item: return
        item_values = self.tree.item(selected_item)['values']
        if item_values and item_values[0] in self.presets:
            del self.presets[item_values[0]]
            self.update_treeview()
            self.refresh_all_hotkeys()

    def load_data(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f: return json.load(f)
            except: pass
        return {"presets": {}, "reset_hotkey": "f9"}

    def save_data(self):
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump({"presets": self.presets, "reset_hotkey": self.reset_hk_var.get()}, f, indent=4, ensure_ascii=False)

    def update_treeview(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for name, p in self.presets.items():
            self.tree.insert("", "end", values=(name, p.get("hotkey", "")))

if __name__ == "__main__":
    root = tk.Tk()
    app = GammaController(root)
    
    # 창의 X 버튼을 눌러 종료할 때의 처리
    def on_closing():
        app.force_reset() # 화면 복구
        root.destroy()    # 창 닫기
        
    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.mainloop()
