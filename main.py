import tkinter as tk
from tkinter import ttk
import threading
import pyperclip
import keyboard
import pyautogui
import pystray
from pystray import MenuItem, Menu as TrayMenu
from PIL import Image, ImageDraw
import json
import os
import time
from datetime import datetime

MAX_HISTORY = 50
POLL_INTERVAL = 0.5
HISTORY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "history.json")

pyautogui.PAUSE = 0

# ── Tray icon image ──────────────────────────────────────────────────────────

def _make_tray_icon() -> Image.Image:
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([8, 14, 56, 60], radius=5, fill="#89b4fa")
    d.rectangle([22, 8, 42, 22], fill="#313244", outline="#cba6f7", width=1)
    d.rectangle([26, 11, 38, 19], fill="#1e1e2e")
    d.rectangle([14, 30, 50, 34], fill="#1e1e2e")
    d.rectangle([14, 40, 46, 44], fill="#1e1e2e")
    d.rectangle([14, 50, 38, 54], fill="#1e1e2e")
    return img

# ── Table detection ──────────────────────────────────────────────────────────

def parse_table(text: str):
    lines = text.splitlines()
    if len(lines) < 2:
        return None
    rows = [line.split("\t") for line in lines if line.strip()]
    col_counts = [len(r) for r in rows]
    if max(col_counts) < 2:
        return None
    dominant = max(set(col_counts), key=col_counts.count)
    if col_counts.count(dominant) / len(col_counts) < 0.5:
        return None
    return rows


def make_entry(text, rows=None, col_states=None, row_states=None,
               header_state=True, ts=None):
    return {
        "text": text,
        "rows": rows,
        "col_states": col_states,
        "row_states": row_states,
        "header_state": header_state,
        "time": ts or datetime.now().strftime("%H:%M"),
    }

# ── History persistence ──────────────────────────────────────────────────────

def save_history(history: list, lock: threading.Lock):
    try:
        with lock:
            data = [
                {
                    "text": e["text"],
                    "rows": e.get("rows"),
                    "col_states": e.get("col_states"),
                    "row_states": e.get("row_states"),
                    "header_state": e.get("header_state", True),
                    "time": e.get("time", ""),
                }
                for e in history
            ]
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def load_history() -> list:
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        entries = []
        for item in data:
            e = make_entry(
                item["text"],
                rows=item.get("rows"),
                col_states=item.get("col_states"),
                row_states=item.get("row_states"),
                header_state=item.get("header_state", True),
                ts=item.get("time", ""),
            )
            entries.append(e)
        return entries
    except Exception:
        return []


# ── Table filter popup ───────────────────────────────────────────────────────

class TableFilterPopup:
    POPUP_W = 860
    POPUP_H = 620

    def __init__(self, parent, rows, on_autpaste, on_state_save, original_text,
                 init_col_states=None, init_row_states=None, init_header_state=True):
        self.parent = parent
        self.rows = rows
        self.on_autpaste = on_autpaste
        self.on_state_save = on_state_save
        self.original_text = original_text
        self.headers = rows[0]
        self.data_rows = rows[1:]

        self.col_vars = [
            tk.BooleanVar(value=(init_col_states[i]
                                 if init_col_states and i < len(init_col_states) else True))
            for i in range(len(self.headers))
        ]
        self.row_vars = [
            tk.BooleanVar(value=(init_row_states[i]
                                 if init_row_states and i < len(init_row_states) else True))
            for i in range(len(self.data_rows))
        ]
        self.header_var = tk.BooleanVar(value=init_header_state)
        self._build()

    def _build(self):
        self.win = tk.Toplevel(self.parent)
        self.win.title("표 필터 — 열/행 선택")
        self.win.configure(bg="#1e1e2e")
        self.win.resizable(True, True)
        self.win.grab_set()
        self.win.attributes("-topmost", True)
        self.win.lift()

        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        x = (sw - self.POPUP_W) // 2
        y = (sh - self.POPUP_H) // 2
        self.win.geometry(f"{self.POPUP_W}x{self.POPUP_H}+{x}+{y}")

        self._build_buttons(self.win)
        self._build_preview(self.win)
        tk.Frame(self.win, bg="#313244", height=1).pack(fill=tk.X, padx=10, side=tk.BOTTOM)

        top_frame = tk.Frame(self.win, bg="#1e1e2e")
        top_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 4))
        self._build_col_panel(top_frame)
        self._build_row_panel(top_frame)
        self._update_preview()

    def _build_col_panel(self, parent):
        frame = tk.LabelFrame(parent, text="  열 선택  ", bg="#1e1e2e", fg="#89b4fa",
                               font=("Segoe UI", 9, "bold"), bd=1, relief=tk.GROOVE, labelanchor="n")
        frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        btn_row = tk.Frame(frame, bg="#1e1e2e")
        btn_row.pack(fill=tk.X, pady=(6, 2), padx=6)
        self._small_btn(btn_row, "전체 선택",
                        lambda: self._set_all(self.col_vars, True)).pack(side=tk.LEFT, padx=(0, 4))
        self._small_btn(btn_row, "전체 해제",
                        lambda: self._set_all(self.col_vars, False)).pack(side=tk.LEFT)

        sf, inner = self._scrollable_frame(frame)
        sf.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 6))
        for i, (name, var) in enumerate(zip(self.headers, self.col_vars)):
            label = name.strip() or f"열{i+1}"
            tk.Checkbutton(inner, text=label, variable=var,
                           bg="#181825", fg="#cdd6f4", selectcolor="#313244",
                           activebackground="#181825", activeforeground="#cdd6f4",
                           anchor=tk.W, font=("Segoe UI", 9),
                           command=self._update_preview).pack(fill=tk.X, padx=4, pady=1)

    def _build_row_panel(self, parent):
        frame = tk.LabelFrame(parent, text="  행 선택  ", bg="#1e1e2e", fg="#a6e3a1",
                               font=("Segoe UI", 9, "bold"), bd=1, relief=tk.GROOVE, labelanchor="n")
        frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))

        tk.Checkbutton(frame, text="헤더(1행) 포함", variable=self.header_var,
                       bg="#1e1e2e", fg="#f9e2af", selectcolor="#313244",
                       activebackground="#1e1e2e", activeforeground="#f9e2af",
                       font=("Segoe UI", 9, "bold"), anchor=tk.W,
                       command=self._update_preview).pack(fill=tk.X, padx=8, pady=(6, 2))
        tk.Frame(frame, bg="#313244", height=1).pack(fill=tk.X, padx=6, pady=2)

        btn_row = tk.Frame(frame, bg="#1e1e2e")
        btn_row.pack(fill=tk.X, pady=(2, 2), padx=6)
        self._small_btn(btn_row, "전체 선택",
                        lambda: self._set_all(self.row_vars, True)).pack(side=tk.LEFT, padx=(0, 4))
        self._small_btn(btn_row, "전체 해제",
                        lambda: self._set_all(self.row_vars, False)).pack(side=tk.LEFT)

        sf, inner = self._scrollable_frame(frame)
        sf.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 6))
        for i, (row, var) in enumerate(zip(self.data_rows, self.row_vars)):
            raw = "\t".join(row)
            preview = raw[:20].replace("\n", " ")
            label = f"{i+2}행: {preview}{'…' if len(raw) > 20 else ''}"
            tk.Checkbutton(inner, text=label, variable=var,
                           bg="#181825", fg="#cdd6f4", selectcolor="#313244",
                           activebackground="#181825", activeforeground="#cdd6f4",
                           anchor=tk.W, font=("Segoe UI", 9),
                           command=self._update_preview).pack(fill=tk.X, padx=4, pady=1)

    def _build_preview(self, parent):
        pf = tk.LabelFrame(parent, text="  미리보기  ", bg="#1e1e2e", fg="#cba6f7",
                            font=("Segoe UI", 9, "bold"), bd=1, relief=tk.GROOVE, labelanchor="n")
        pf.pack(fill=tk.X, padx=10, pady=6, side=tk.BOTTOM)
        pf.configure(height=170)
        pf.pack_propagate(False)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Preview.Treeview", background="#181825", foreground="#cdd6f4",
                        fieldbackground="#181825", rowheight=22, font=("Segoe UI", 9))
        style.configure("Preview.Treeview.Heading", background="#313244", foreground="#89b4fa",
                        font=("Segoe UI", 9, "bold"), relief="flat")
        style.map("Preview.Treeview", background=[("selected", "#89b4fa")],
                  foreground=[("selected", "#1e1e2e")])

        tc = tk.Frame(pf, bg="#181825")
        tc.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        sy = tk.Scrollbar(tc, orient=tk.VERTICAL)
        sy.pack(side=tk.RIGHT, fill=tk.Y)
        sx = tk.Scrollbar(tc, orient=tk.HORIZONTAL)
        sx.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree = ttk.Treeview(tc, style="Preview.Treeview", show="headings",
                                  yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.pack(fill=tk.BOTH, expand=True)
        sy.config(command=self.tree.yview)
        sx.config(command=self.tree.xview)

    def _update_preview(self):
        sel_col_idx = [i for i, v in enumerate(self.col_vars) if v.get()]
        sel_rows = [r for r, v in zip(self.data_rows, self.row_vars) if v.get()]
        col_names = [self.headers[i].strip() or f"열{i+1}" for i in sel_col_idx]
        self.tree["columns"] = col_names
        for name in col_names:
            self.tree.heading(name, text=name)
            self.tree.column(name, width=max(80, len(name) * 9), minwidth=50, stretch=True)
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in sel_rows:
            vals = [row[i] if i < len(row) else "" for i in sel_col_idx]
            self.tree.insert("", tk.END, values=vals)

    def _build_buttons(self, parent):
        bf = tk.Frame(parent, bg="#1e1e2e")
        bf.pack(fill=tk.X, padx=10, pady=(0, 10), side=tk.BOTTOM)
        self._action_btn(bf, "붙여넣기",    "#89b4fa", "#1e1e2e",
                         self._paste_filtered).pack(side=tk.LEFT, padx=(0, 6))
        self._action_btn(bf, "전체 붙여넣기", "#a6e3a1", "#1e1e2e",
                         self._paste_all).pack(side=tk.LEFT, padx=(0, 6))
        self._action_btn(bf, "취소",        "#f38ba8", "#1e1e2e",
                         self.win.destroy).pack(side=tk.LEFT)

    def _get_filtered_text(self):
        sel_col_idx = [i for i, v in enumerate(self.col_vars) if v.get()]
        sel_rows = [r for r, v in zip(self.data_rows, self.row_vars) if v.get()]
        result = []
        if self.header_var.get():
            result.append([self.headers[i] for i in sel_col_idx])
        for row in sel_rows:
            result.append([row[i] if i < len(row) else "" for i in sel_col_idx])
        return "\n".join("\t".join(r) for r in result)

    def _flush_state(self):
        self.on_state_save([v.get() for v in self.col_vars],
                           [v.get() for v in self.row_vars],
                           self.header_var.get())

    def _paste_filtered(self):
        text = self._get_filtered_text()
        self._flush_state()
        self.win.destroy()
        self.on_autpaste(text)

    def _paste_all(self):
        self._flush_state()
        self.win.destroy()
        self.on_autpaste(self.original_text)

    def _set_all(self, var_list, value):
        for v in var_list:
            v.set(value)
        self._update_preview()

    @staticmethod
    def _scrollable_frame(parent):
        outer = tk.Frame(parent, bg="#181825", relief=tk.FLAT)
        canvas = tk.Canvas(outer, bg="#181825", highlightthickness=0, bd=0)
        sb = tk.Scrollbar(outer, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        inner = tk.Frame(canvas, bg="#181825")
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(win_id, width=e.width))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-1 * e.delta / 120), "units"))
        return outer, inner

    @staticmethod
    def _small_btn(parent, text, cmd):
        return tk.Button(parent, text=text, command=cmd,
                         bg="#313244", fg="#cdd6f4", relief=tk.FLAT,
                         font=("Segoe UI", 8), cursor="hand2", padx=6, pady=2,
                         activebackground="#45475a", activeforeground="#cdd6f4")

    @staticmethod
    def _action_btn(parent, text, bg, fg, cmd):
        return tk.Button(parent, text=text, command=cmd,
                         bg=bg, fg=fg, relief=tk.FLAT,
                         font=("Segoe UI", 10, "bold"), cursor="hand2",
                         padx=14, pady=5,
                         activebackground=bg, activeforeground=fg)


# ── Main clipboard manager ───────────────────────────────────────────────────

class ClipboardManager:
    def __init__(self):
        self.history: list[dict] = load_history()
        self.filtered: list[dict] = []
        self.lock = threading.Lock()
        self.last_clip = ""
        self.window_visible = False
        self.popup_open = False
        self.tray_icon = None

        self._build_ui()
        self._apply_filter()          # populate list from loaded history
        self._start_monitor()
        self._register_hotkey()
        self._build_tray()

        self.root.protocol("WM_DELETE_WINDOW", self._hide_window)
        self.root.withdraw()
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            pass
        finally:
            keyboard.unhook_all()

    # ── Tray ────────────────────────────────────────────────────────────────

    def _build_tray(self):
        icon_img = _make_tray_icon()
        menu = TrayMenu(
            MenuItem("열기",  lambda icon, item: self.root.after(0, self._show_window)),
            MenuItem("종료",  lambda icon, item: self.root.after(0, self._quit)),
        )
        self.tray_icon = pystray.Icon(
            "smart_clipboard", icon_img, "Smart Clipboard Manager", menu
        )
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def _quit(self):
        save_history(self.history, self.lock)
        if self.tray_icon:
            self.tray_icon.stop()
        keyboard.unhook_all()
        self.root.destroy()

    # ── UI ──────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.root = tk.Tk()
        self.root.geometry("500x580")
        self.root.resizable(True, True)
        self.root.configure(bg="#1e1e2e")
        self.root.attributes("-topmost", True)
        self._update_title()

        # ── Toolbar: search + clear-all button
        toolbar = tk.Frame(self.root, bg="#1e1e2e", pady=8, padx=10)
        toolbar.pack(fill=tk.X)

        tk.Label(toolbar, text="검색", bg="#1e1e2e", fg="#cdd6f4",
                 font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(0, 6))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._apply_filter())
        tk.Entry(toolbar, textvariable=self.search_var,
                 bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                 relief=tk.FLAT, font=("Segoe UI", 10), bd=4
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Button(toolbar, text="✕", command=lambda: self.search_var.set(""),
                  bg="#313244", fg="#cdd6f4", relief=tk.FLAT,
                  font=("Segoe UI", 9), cursor="hand2", padx=4
                  ).pack(side=tk.LEFT, padx=(4, 0))

        tk.Button(toolbar, text="전체 삭제", command=self._clear_all,
                  bg="#f38ba8", fg="#1e1e2e", relief=tk.FLAT,
                  font=("Segoe UI", 8, "bold"), cursor="hand2", padx=8, pady=3,
                  activebackground="#eba0ac", activeforeground="#1e1e2e"
                  ).pack(side=tk.LEFT, padx=(8, 0))

        # ── Scrollable item list
        list_outer = tk.Frame(self.root, bg="#1e1e2e", padx=10)
        list_outer.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.list_canvas = tk.Canvas(list_outer, bg="#181825",
                                     highlightthickness=0, bd=0)
        list_sb = tk.Scrollbar(list_outer, orient=tk.VERTICAL,
                                command=self.list_canvas.yview)
        self.list_canvas.configure(yscrollcommand=list_sb.set)
        list_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.list_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.items_frame = tk.Frame(self.list_canvas, bg="#181825")
        self._items_win = self.list_canvas.create_window(
            (0, 0), window=self.items_frame, anchor="nw")
        self.items_frame.bind("<Configure>",
            lambda e: self.list_canvas.configure(
                scrollregion=self.list_canvas.bbox("all")))
        self.list_canvas.bind("<Configure>",
            lambda e: self.list_canvas.itemconfig(self._items_win, width=e.width))
        self.list_canvas.bind("<MouseWheel>", self._on_list_scroll)

        # ── Context menu
        self._ctx_entry = None
        self.context_menu = tk.Menu(
            self.root, tearoff=0, bg="#313244", fg="#cdd6f4",
            activebackground="#89b4fa", activeforeground="#1e1e2e", relief=tk.FLAT)
        self.context_menu.add_command(label="삭제", command=self._delete_ctx_item)
        self.context_menu.add_command(label="전체 삭제", command=self._clear_all)

        # ── Status bar
        self.status_var = tk.StringVar(value="클립보드 모니터링 중...")
        tk.Label(self.root, textvariable=self.status_var,
                 bg="#181825", fg="#6c7086",
                 font=("Segoe UI", 8), anchor=tk.W, padx=10, pady=4
                 ).pack(fill=tk.X, side=tk.BOTTOM)

    def _update_title(self):
        count = len(self.history)
        self.root.title(f"Smart Clipboard Manager  [{count}개]")

    def _on_list_scroll(self, e):
        self.list_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

    # ── Item row rendering ───────────────────────────────────────────────────

    def _render_item_row(self, parent, entry: dict, idx: int):
        is_table = entry["rows"] is not None
        bg  = "#1e1e2e" if idx % 2 == 0 else "#24273a"
        hbg = "#313244"

        row_frame = tk.Frame(parent, bg=bg, pady=3, padx=6)
        row_frame.pack(fill=tk.X)

        # ── [표] badge
        if is_table:
            tk.Label(row_frame, text="[표]",
                     bg="#313244", fg="#f9e2af",
                     font=("Segoe UI", 7, "bold"), padx=3, pady=1
                     ).pack(side=tk.LEFT, padx=(0, 5))

        # ── Main label
        if is_table:
            rows = entry["rows"]
            n_data = len(rows) - 1
            n_cols = len(rows[0])
            hdr = ", ".join(h.strip() for h in rows[0][:3])
            if n_cols > 3:
                hdr += "…"
            display  = f"{n_data}행 × {n_cols}열  |  {hdr}"
            label_fg = "#f9e2af"
        else:
            raw = entry["text"]
            display  = raw[:60].replace("\n", " ").replace("\r", "")
            if len(raw) > 60:
                display += "…"
            label_fg = "#cdd6f4"

        lbl = tk.Label(row_frame, text=display, bg=bg, fg=label_fg,
                       font=("Segoe UI", 9), anchor=tk.W, cursor="hand2")
        lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # ── Right side: time + buttons
        right = tk.Frame(row_frame, bg=bg)
        right.pack(side=tk.RIGHT)

        ts = entry.get("time", "")
        if ts:
            tk.Label(right, text=ts, bg=bg, fg="#6c7086",
                     font=("Segoe UI", 7)).pack(side=tk.LEFT, padx=(0, 6))

        if is_table:
            tk.Button(right, text="편집",
                      bg="#2a2a3e", fg="#cba6f7", relief=tk.FLAT,
                      font=("Segoe UI", 8), cursor="hand2", width=5, pady=1,
                      command=lambda e=entry: self._open_edit_popup(e),
                      activebackground="#45475a", activeforeground="#cba6f7"
                      ).pack(side=tk.LEFT, padx=(0, 3))

        tk.Button(right, text="붙여넣기",
                  bg="#2a2a3e", fg="#89b4fa", relief=tk.FLAT,
                  font=("Segoe UI", 8), cursor="hand2", width=7, pady=1,
                  command=lambda e=entry: self._do_copy_and_paste(e["text"]),
                  activebackground="#45475a", activeforeground="#89b4fa"
                  ).pack(side=tk.LEFT)

        # ── Hover + events
        def _enter(_):
            row_frame.configure(bg=hbg); lbl.configure(bg=hbg); right.configure(bg=hbg)
        def _leave(_):
            row_frame.configure(bg=bg); lbl.configure(bg=bg); right.configure(bg=bg)

        for w in [row_frame, lbl]:
            w.bind("<Enter>", _enter)
            w.bind("<Leave>", _leave)
            w.bind("<Double-Button-1>",
                   lambda _, e=entry: self._do_copy_and_paste(e["text"]))
            w.bind("<Button-3>",
                   lambda ev, e=entry: self._on_right_click(ev, e))
            w.bind("<MouseWheel>", self._on_list_scroll)

    # ── Clipboard monitor ────────────────────────────────────────────────────

    def _start_monitor(self):
        threading.Thread(target=self._monitor_loop, daemon=True).start()

    def _monitor_loop(self):
        while True:
            try:
                clip = pyperclip.paste()
                if clip and clip != self.last_clip:
                    self.last_clip = clip
                    self.root.after(0, lambda c=clip: self._handle_new_clip(c))
            except Exception:
                pass
            time.sleep(POLL_INTERVAL)

    def _handle_new_clip(self, text: str):
        rows = parse_table(text)
        if rows and not self.popup_open:
            self.popup_open = True
            entry = make_entry(text, rows=rows,
                               col_states=[True] * len(rows[0]),
                               row_states=[True] * (len(rows) - 1),
                               header_state=True)
            self._add_entry(entry)
            self.status_var.set(
                f"표 감지됨: {len(rows)}행 × {len(rows[0])}열 — 필터 팝업 열림")
            self._open_table_popup(entry)
        else:
            self._add_entry(make_entry(text))

    # ── Table popup ──────────────────────────────────────────────────────────

    def _open_table_popup(self, entry: dict):
        def on_autpaste(filtered_text):
            self._apply_filter()
            self._do_copy_and_paste(filtered_text)

        def on_state_save(col_states, row_states, header_state):
            entry["col_states"] = col_states
            entry["row_states"] = row_states
            entry["header_state"] = header_state

        TableFilterPopup(
            self.root, entry["rows"],
            on_autpaste=on_autpaste,
            on_state_save=on_state_save,
            original_text=entry["text"],
            init_col_states=entry.get("col_states"),
            init_row_states=entry.get("row_states"),
            init_header_state=entry.get("header_state", True),
        )
        self.root.after(100, self._check_popup_closed)

    def _check_popup_closed(self):
        toplevels = [w for w in self.root.winfo_children()
                     if isinstance(w, tk.Toplevel) and w.winfo_exists()]
        if toplevels:
            self.root.after(200, self._check_popup_closed)
        else:
            self.popup_open = False

    def _open_edit_popup(self, entry: dict):
        if not self.popup_open:
            self.popup_open = True
            self._open_table_popup(entry)

    # ── History management ───────────────────────────────────────────────────

    def _add_entry(self, entry: dict):
        with self.lock:
            self.history = [e for e in self.history if e["text"] != entry["text"]]
            self.history.insert(0, entry)
            if len(self.history) > MAX_HISTORY:
                self.history = self.history[:MAX_HISTORY]
        self._apply_filter()
        self._update_title()
        save_history(self.history, self.lock)
        preview = entry["text"][:40].replace("\n", " ").replace("\t", " | ")
        self.status_var.set(f"저장됨: {preview}{'...' if len(entry['text']) > 40 else ''}")

    # ── Paste action ─────────────────────────────────────────────────────────

    def _do_copy_and_paste(self, text: str):
        self.last_clip = text
        pyperclip.copy(text)

        was_visible = self.window_visible
        self.root.withdraw()
        self.window_visible = False

        def _paste():
            try:
                pyautogui.hotkey("ctrl", "v")
            except Exception:
                pass
            if was_visible:
                self.root.after(150, self._restore_window)

        self.root.after(200, _paste)
        preview = text[:40].replace("\n", " ").replace("\t", " | ")
        self.status_var.set(f"붙여넣기: {preview}{'...' if len(text) > 40 else ''}")

    def _restore_window(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        self.window_visible = True

    # ── Filter / list update ─────────────────────────────────────────────────

    def _apply_filter(self):
        query = self.search_var.get().lower()
        with self.lock:
            self.filtered = (
                [e for e in self.history if query in e["text"].lower()]
                if query else list(self.history)
            )
        for w in self.items_frame.winfo_children():
            w.destroy()
        for i, entry in enumerate(self.filtered):
            self._render_item_row(self.items_frame, entry, i)
        self.items_frame.update_idletasks()
        self.list_canvas.configure(scrollregion=self.list_canvas.bbox("all"))

    # ── Events ───────────────────────────────────────────────────────────────

    def _on_right_click(self, event, entry: dict):
        self._ctx_entry = entry
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def _delete_ctx_item(self):
        if self._ctx_entry is None:
            return
        with self.lock:
            self.history = [e for e in self.history if e is not self._ctx_entry]
        self._ctx_entry = None
        self._apply_filter()
        self._update_title()
        save_history(self.history, self.lock)

    def _clear_all(self):
        with self.lock:
            self.history.clear()
        self._apply_filter()
        self._update_title()
        save_history(self.history, self.lock)
        self.status_var.set("전체 삭제 완료")

    # ── Hotkey ───────────────────────────────────────────────────────────────

    def _register_hotkey(self):
        try:
            keyboard.add_hotkey("win+shift+v", self._toggle_window, suppress=False)
        except Exception as e:
            print(f"단축키 등록 실패: {e}")

    def _toggle_window(self):
        self.root.after(0, self._do_toggle)

    def _do_toggle(self):
        if self.window_visible:
            self._hide_window()
        else:
            self._show_window()

    def _show_window(self):
        self._apply_filter()
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        self.window_visible = True

    def _hide_window(self):
        self.root.withdraw()
        self.window_visible = False


if __name__ == "__main__":
    ClipboardManager()
