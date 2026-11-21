import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

# Hỗ trợ kéo thả file/thư mục nếu có tkinterdnd2
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

# Thư viện điều khiển Excel & Máy in Windows
try:
    import win32com.client
    import win32print
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

SUPPORTED_EXTS = (".xlsx", ".xlsm", ".xls", ".xlsb")

PAPER_SIZES = {
    "Giữ nguyên theo file gốc": None,
    "A3 (297 x 420 mm)": 8,       # xlPaperA3
    "A4 (210 x 297 mm)": 9,       # xlPaperA4
    "A5 (148 x 210 mm)": 11,      # xlPaperA5
    "Letter (8.5 x 11 in)": 1,    # xlPaperLetter
}

ORIENTATIONS = {
    "Giữ nguyên theo file gốc": None,
    "Dọc (Portrait)": 1,          # xlPortrait
    "Ngang (Landscape)": 2,       # xlLandscape
}

MISSING_SHEET_ACTIONS = {
    "Bỏ qua file và ghi vào nhật ký lỗi": "skip",
    "In sheet đầu tiên của file": "first",
}


class InHangLoatExcelPro:
    def __init__(self, root):
        self.root = root
        self.root.title("IN HÀNG LOẠT EXCEL PRO")
        self.root.geometry("820x740")
        self.root.minsize(760, 680)
        self.root.configure(bg="#0F1115")

        self.file_list = []
        self.is_printing = False
        self.stop_requested = False

        self._setup_styles()
        self._build_ui()
        self._load_printers()

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        bg_main = "#0F1115"
        bg_card = "#161A22"
        fg_text = "#F3F4F6"
        fg_muted = "#9CA3AF"
        accent = "#1D74F5"

        style.configure("Main.TFrame", background=bg_main)
        style.configure("Card.TFrame", background=bg_card, relief="flat")
        style.configure("Title.TLabel", background=bg_card, foreground=fg_text, font=("Segoe UI", 14, "bold"))
        style.configure("Section.TLabel", background=bg_card, foreground=fg_text, font=("Segoe UI", 11, "bold"))
        style.configure("Body.TLabel", background=bg_card, foreground=fg_text, font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=bg_card, foreground=fg_muted, font=("Segoe UI", 9))
        style.configure("Badge.TLabel", background="#0E2E24", foreground="#34D399", font=("Segoe UI", 9, "bold"), padding=(8, 3))

        style.configure("TCombobox", fieldbackground="#1F2430", background="#1F2430", foreground=fg_text, arrowcolor=fg_text)
        style.map("TCombobox", fieldbackground=[("readonly", "#1F2430")], selectbackground=[("readonly", "#1F2430")], selectforeground=[("readonly", fg_text)])

        style.configure("TCheckbutton", background=bg_card, foreground=fg_text, font=("Segoe UI", 10))
        style.map("TCheckbutton", background=[("active", bg_card)])

        style.configure("Green.Horizontal.TProgressbar", troughcolor="#1F2430", background="#10B981", thickness=14)

    def _build_ui(self):
        container = ttk.Frame(self.root, style="Main.TFrame", padding=16)
        container.pack(fill="both", expand=True)

        # Header
        header_card = ttk.Frame(container, style="Card.TFrame", padding=(16, 12))
        header_card.pack(fill="x", pady=(0, 10))

        ttk.Label(header_card, text="🖨  IN HÀNG LOẠT EXCEL PRO", style="Title.TLabel").pack(side="left")
        ttk.Label(header_card, text="Bản độc lập • Không cần Python", style="Badge.TLabel").pack(side="right")

        # 1. Chọn dữ liệu
        sec1 = ttk.Frame(container, style="Card.TFrame", padding=14)
        sec1.pack(fill="both", expand=True, pady=(0, 10))

        top_sec1 = ttk.Frame(sec1, style="Card.TFrame")
        top_sec1.pack(fill="x", pady=(0, 8))
        ttk.Label(top_sec1, text="1. Chọn dữ liệu Excel", style="Section.TLabel").pack(side="left")

        btn_bar = ttk.Frame(top_sec1, style="Card.TFrame")
        btn_bar.pack(side="right")

        self._make_btn(btn_bar, "+ Chọn nhiều file", self.select_files, "#2563EB").pack(side="left", padx=4)
        self._make_btn(btn_bar, "📁 Chọn cả thư mục", self.select_folder, "#374151").pack(side="left", padx=4)
        self._make_btn(btn_bar, "🗑 Xóa danh sách", self.clear_files, "#7F1D1D").pack(side="left", padx=(4, 0))

        # Vùng kéo thả + danh sách file
        drop_frame = tk.Frame(sec1, bg="#12161F", highlightbackground="#2E3646", highlightthickness=1, bd=0)
        drop_frame.pack(fill="both", expand=True)

        hint_text = "Kéo thả nhiều file hoặc thư mục Excel vào đây (Hỗ trợ XLSX, XLSM, XLS, XLSB)"
        if not HAS_DND:
            hint_text = "Danh sách file Excel sẽ in (Hỗ trợ XLSX, XLSM, XLS, XLSB)"

        self.lbl_drop = tk.Label(drop_frame, text=hint_text, bg="#12161F", fg="#9CA3AF", font=("Segoe UI", 9, "italic"), pady=6)
        self.lbl_drop.pack(fill="x")

        list_container = tk.Frame(drop_frame, bg="#12161F")
        list_container.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        scrollbar = ttk.Scrollbar(list_container)
        scrollbar.pack(side="right", fill="y")

        self.listbox = tk.Listbox(
            list_container,
            bg="#12161F",
            fg="#E5E7EB",
            selectbackground="#2563EB",
            selectforeground="#FFFFFF",
            font=("Consolas", 9),
            bd=0,
            highlightthickness=0,
            yscrollcommand=scrollbar.set,
            height=7,
        )
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)

        if HAS_DND:
            for widget in (drop_frame, self.lbl_drop, self.listbox):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)

        self.lbl_count = ttk.Label(sec1, text="Tổng cộng: 0 file", style="Muted.TLabel")
        self.lbl_count.pack(anchor="w", pady=(6, 0))

        # 2. Chọn Sheet
        sec2 = ttk.Frame(container, style="Card.TFrame", padding=14)
        sec2.pack(fill="x", pady=(0, 10))

        row_s2 = ttk.Frame(sec2, style="Card.TFrame")
        row_s2.pack(fill="x")
        ttk.Label(row_s2, text="2. Chọn Sheet cần in", style="Section.TLabel").pack(side="left")
        ttk.Label(row_s2, text="Mặc định: WL", style="Badge.TLabel").pack(side="right")

        grid_s2 = ttk.Frame(sec2, style="Card.TFrame")
        grid_s2.pack(fill="x", pady=(8, 0))

        ttk.Label(grid_s2, text="Tên Sheet mục tiêu:", style="Body.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.var_sheet_name = tk.StringVar(value="WL")
        entry_sheet = tk.Entry(grid_s2, textvariable=self.var_sheet_name, bg="#1F2430", fg="#FFFFFF", insertbackground="#FFFFFF", relief="flat", font=("Segoe UI", 10, "bold"), width=18)
        entry_sheet.grid(row=0, column=1, sticky="w", padx=(0, 20), ipady=4)

        ttk.Label(grid_s2, text="Nếu không tìm thấy Sheet:", style="Body.TLabel").grid(row=0, column=2, sticky="w", padx=(0, 8))
        self.var_missing_action = tk.StringVar(value=list(MISSING_SHEET_ACTIONS.keys())[0])
        cb_missing = ttk.Combobox(grid_s2, textvariable=self.var_missing_action, values=list(MISSING_SHEET_ACTIONS.keys()), state="readonly", width=34)
        cb_missing.grid(row=0, column=3, sticky="w")

        # 3. Thiết lập in
        sec3 = ttk.Frame(container, style="Card.TFrame", padding=14)
        sec3.pack(fill="x", pady=(0, 10))

        ttk.Label(sec3, text="3. Thiết lập in ấn", style="Section.TLabel").pack(anchor="w", pady=(0, 8))

        grid_s3 = ttk.Frame(sec3, style="Card.TFrame")
        grid_s3.pack(fill="x")

        # Hàng 1: Máy in & Số bản
        ttk.Label(grid_s3, text="Chọn máy in:", style="Body.TLabel").grid(row=0, column=0, sticky="w", pady=4, padx=(0, 8))
        self.var_printer = tk.StringVar()
        self.cb_printer = ttk.Combobox(grid_s3, textvariable=self.var_printer, state="readonly", width=44)
        self.cb_printer.grid(row=0, column=1, sticky="w", pady=4, padx=(0, 20))

        ttk.Label(grid_s3, text="Số bản in:", style="Body.TLabel").grid(row=0, column=2, sticky="w", pady=4, padx=(0, 8))
        self.var_copies = tk.IntVar(value=1)
        spin_copies = tk.Spinbox(grid_s3, from_=1, to=999, textvariable=self.var_copies, bg="#1F2430", fg="#FFFFFF", buttonbackground="#374151", relief="flat", font=("Segoe UI", 10), width=8)
        spin_copies.grid(row=0, column=3, sticky="w", pady=4, ipady=3)

        # Hàng 2: Khổ giấy & Hướng giấy
        ttk.Label(grid_s3, text="Khổ giấy:", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=4, padx=(0, 8))
        self.var_paper = tk.StringVar(value=list(PAPER_SIZES.keys())[0])
        cb_paper = ttk.Combobox(grid_s3, textvariable=self.var_paper, values=list(PAPER_SIZES.keys()), state="readonly", width=44)
        cb_paper.grid(row=1, column=1, sticky="w", pady=4, padx=(0, 20))

        ttk.Label(grid_s3, text="Hướng giấy:", style="Body.TLabel").grid(row=1, column=2, sticky="w", pady=4, padx=(0, 8))
        self.var_orient = tk.StringVar(value=list(ORIENTATIONS.keys())[0])
        cb_orient = ttk.Combobox(grid_s3, textvariable=self.var_orient, values=list(ORIENTATIONS.keys()), state="readonly", width=24)
        cb_orient.grid(row=1, column=3, sticky="w", pady=4)

        # Hàng 3: Giữ nguyên vùng in và tỷ lệ gốc
        self.var_keep_setup = tk.BooleanVar(value=True)
        chk_keep = ttk.Checkbutton(
            sec3,
            text="Giữ nguyên thiết lập vùng in (Print Area) và tỷ lệ thu phóng (Zoom / Fit-to-Page) của file gốc",
            variable=self.var_keep_setup,
        )
        chk_keep.pack(anchor="w", pady=(8, 0))

        # Nút Bắt đầu in & Dừng
        action_bar = ttk.Frame(container, style="Main.TFrame")
        action_bar.pack(fill="x", pady=(0, 10))

        self.btn_start = tk.Button(
            action_bar,
            text="🖨  BẮT ĐẦU IN HÀNG LOẠT",
            bg="#1D74F5",
            fg="#FFFFFF",
            activebackground="#1558C0",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 12, "bold"),
            relief="flat",
            cursor="hand2",
            pady=10,
            command=self.start_printing,
        )
        self.btn_start.pack(side="left", fill="x", expand=True)

        self.btn_stop = tk.Button(
            action_bar,
            text="⏹ DỪNG",
            bg="#DC2626",
            fg="#FFFFFF",
            activebackground="#991B1B",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            cursor="hand2",
            pady=10,
            padx=18,
            state="disabled",
            command=self.request_stop,
        )
        self.btn_stop.pack(side="right", padx=(10, 0))

        # Thanh tiến trình & Nhật ký lỗi
        sec4 = ttk.Frame(container, style="Card.TFrame", padding=12)
        sec4.pack(fill="both", expand=True)

        status_top = ttk.Frame(sec4, style="Card.TFrame")
        status_top.pack(fill="x", pady=(0, 6))

        self.lbl_status = ttk.Label(status_top, text="Sẵn sàng...", foreground="#34D399", background="#161A22", font=("Segoe UI", 9, "bold"))
        self.lbl_status.pack(side="left")

        self.lbl_progress_pct = ttk.Label(status_top, text="0 / 0 (0%)", style="Muted.TLabel")
        self.lbl_progress_pct.pack(side="right")

        self.progress = ttk.Progressbar(sec4, style="Green.Horizontal.TProgressbar", mode="determinate")
        self.progress.pack(fill="x", pady=(0, 8))

        log_header = ttk.Frame(sec4, style="Card.TFrame")
        log_header.pack(fill="x", pady=(0, 4))
        ttk.Label(log_header, text="Nhật ký tiến trình & lỗi:", style="Muted.TLabel").pack(side="left")
        self._make_btn(log_header, "Xuất nhật ký (.txt)", self.export_log, "#374151").pack(side="right")

        self.txt_log = tk.Text(
            sec4,
            bg="#0B0D11",
            fg="#D1D5DB",
            font=("Consolas", 9),
            height=6,
            bd=0,
            highlightthickness=1,
            highlightbackground="#2E3646",
            state="disabled",
        )
        self.txt_log.pack(fill="both", expand=True)
        self.txt_log.tag_config("ERROR", foreground="#F87171")
        self.txt_log.tag_config("SUCCESS", foreground="#34D399")
        self.txt_log.tag_config("WARN", foreground="#FBBF24")

    def _make_btn(self, parent, text, cmd, bg):
        return tk.Button(
            parent,
            text=text,
            command=cmd,
            bg=bg,
            fg="#FFFFFF",
            activebackground="#4B5563",
            activeforeground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=4,
        )

    def _load_printers(self):
        if not HAS_WIN32:
            self.cb_printer["values"] = ["(Chưa cài pywin32 - Chỉ chạy trên Windows)"]
            self.cb_printer.current(0)
            return
        try:
            printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
            default_printer = win32print.GetDefaultPrinter()
            self.cb_printer["values"] = printers
            if default_printer in printers:
                self.cb_printer.set(default_printer)
            elif printers:
                self.cb_printer.current(0)
        except Exception as e:
            self._log(f"Không thể lấy danh sách máy in: {e}", "ERROR")

    def _add_paths(self, paths):
        added = 0
        for p in paths:
            p = os.path.normpath(p.strip('"'))
            if os.path.isdir(p):
                for root_dir, _, files in os.walk(p):
                    for f in sorted(files):
                        if f.lower().endswith(SUPPORTED_EXTS) and not f.startswith("~$"):
                            full_p = os.path.join(root_dir, f)
                            if full_p not in self.file_list:
                                self.file_list.append(full_p)
                                self.listbox.insert("end", full_p)
                                added += 1
            elif os.path.isfile(p):
                if os.path.basename(p).lower().endswith(SUPPORTED_EXTS) and not os.path.basename(p).startswith("~$"):
                    if p not in self.file_list:
                        self.file_list.append(p)
                        self.listbox.insert("end", p)
                        added += 1
        self.lbl_count.config(text=f"Tổng cộng: {len(self.file_list)} file")
        if added:
            self._log(f"Đã thêm {added} file Excel vào danh sách.")

    def _on_drop(self, event):
        raw = event.data
        paths = self.root.tk.splitlist(raw)
        self._add_paths(paths)

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Chọn các file Excel cần in",
            filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls *.xlsb"), ("All Files", "*.*")],
        )
        if files:
            self._add_paths(files)

    def select_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục chứa file Excel")
        if folder:
            self._add_paths([folder])

    def clear_files(self):
        if self.is_printing:
            return
        self.file_list.clear()
        self.listbox.delete(0, "end")
        self.lbl_count.config(text="Tổng cộng: 0 file")
        self.progress["value"] = 0
        self.lbl_progress_pct.config(text="0 / 0 (0%)")
        self.lbl_status.config(text="Sẵn sàng...")

    def _log(self, msg, tag=None):
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {msg}\n"

        def append():
            self.txt_log.config(state="normal")
            if tag:
                self.txt_log.insert("end", line, tag)
            else:
                self.txt_log.insert("end", line)
            self.txt_log.see("end")
            self.txt_log.config(state="disabled")

        self.root.after(0, append)

    def export_log(self):
        content = self.txt_log.get("1.0", "end").strip()
        if not content:
            messagebox.showinfo("Thông báo", "Nhật ký hiện đang trống.")
            return
        save_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialfile=f"NhatKyInExcel_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            filetypes=[("Text Files", "*.txt")],
        )
        if save_path:
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(content)
            messagebox.showinfo("Thành công", "Đã xuất file nhật ký!")

    def request_stop(self):
        if self.is_printing:
            self.stop_requested = True
            self._log("Đang yêu cầu dừng sau khi hoàn tất file hiện tại...", "WARN")

    def start_printing(self):
        if not HAS_WIN32:
            messagebox.showerror("Lỗi môi trường", "Ứng dụng cần chạy trên Windows có cài Microsoft Excel.")
            return
        if not self.file_list:
            messagebox.showwarning("Chưa có dữ liệu", "Vui lòng chọn hoặc kéo thả ít nhất 1 file Excel.")
            return
        printer = self.var_printer.get()
        if not printer:
            messagebox.showwarning("Chưa chọn máy in", "Vui lòng chọn máy in trước khi bắt đầu.")
            return

        self.is_printing = True
        self.stop_requested = False
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.progress["maximum"] = len(self.file_list)
        self.progress["value"] = 0

        thread = threading.Thread(target=self._print_worker, daemon=True)
        thread.start()

    def _print_worker(self):
        pythoncom.CoInitialize()
        excel = None
        total = len(self.file_list)
        success_count = 0
        error_count = 0

        target_sheet = self.var_sheet_name.get().strip() or "WL"
        missing_action = MISSING_SHEET_ACTIONS[self.var_missing_action.get()]
        printer_name = self.var_printer.get()
        copies = max(1, int(self.var_copies.get()))
        paper_code = PAPER_SIZES[self.var_paper.get()]
        orient_code = ORIENTATIONS[self.var_orient.get()]
        keep_setup = self.var_keep_setup.get()

        self._log(f"=== BẮT ĐẦU IN {total} FILE | Máy in: {printer_name} | Sheet: {target_sheet} | Số bản: {copies} ===")

        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.EnableEvents = False

            for idx, file_path in enumerate(self.file_list, start=1):
                if self.stop_requested:
                    self._log("Đã dừng tiến trình theo yêu cầu người dùng.", "WARN")
                    break

                fname = os.path.basename(file_path)
                self.root.after(0, lambda f=fname, i=idx: self.lbl_status.config(text=f"Đang in ({i}/{total}): {f}"))

                wb = None
                try:
                    wb = excel.Workbooks.Open(file_path, ReadOnly=True, UpdateLinks=0)
                    sheet_names = [s.Name for s in wb.Sheets]

                    # Tìm sheet khớp không phân biệt hoa thường
                    matched_name = None
                    for sname in sheet_names:
                        if sname.strip().lower() == target_sheet.lower():
                            matched_name = sname
                            break

                    if matched_name:
                        ws = wb.Sheets(matched_name)
                    elif missing_action == "first":
                        ws = wb.Sheets(1)
                        self._log(f"[Cảnh báo] {fname}: Không có sheet '{target_sheet}', chuyển sang in sheet '{ws.Name}'.", "WARN")
                    else:
                        raise ValueError(f"Không tìm thấy sheet '{target_sheet}' (Các sheet hiện có: {', '.join(sheet_names)})")

                    # Thiết lập PageSetup nếu người dùng có chọn thay đổi
                    ps = ws.PageSetup
                    if paper_code is not None:
                        ps.PaperSize = paper_code
                    if orient_code is not None:
                        ps.Orientation = orient_code

                    # Nếu không tick giữ nguyên thiết lập gốc thì mới reset FitToPages, ngược lại giữ nguyên PrintArea & Zoom/FitToPages của file
                    if not keep_setup:
                        ps.PrintArea = ""
                        ps.Zoom = False
                        ps.FitToPagesWide = 1
                        ps.FitToPagesTall = False

                    ws.PrintOut(Copies=copies, ActivePrinter=printer_name, Collate=True)
                    success_count += 1
                    self._log(f"[OK] {fname} -> Sheet '{ws.Name}' ({copies} bản)", "SUCCESS")

                except Exception as file_err:
                    error_count += 1
                    self._log(f"[LỖI] {fname}: {file_err}", "ERROR")
                finally:
                    if wb is not None:
                        try:
                            wb.Close(SaveChanges=False)
                        except Exception:
                            pass

                pct = int((idx / total) * 100)
                self.root.after(0, lambda i=idx, p=pct: (
                    self.progress.configure(value=i),
                    self.lbl_progress_pct.config(text=f"{i} / {total} ({p}%)")
                ))

        except Exception as app_err:
            self._log(f"[LỖI NGHIÊM TRỌNG] Không khởi động được Microsoft Excel: {app_err}", "ERROR")
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

            self.is_printing = False
            self.root.after(0, lambda: (
                self.btn_start.config(state="normal"),
                self.btn_stop.config(state="disabled"),
                self.lbl_status.config(text=f"Hoàn tất! Thành công: {success_count} | Lỗi: {error_count}")
            ))
            self._log(f"=== KẾT THÚC: Thành công {success_count}/{total} file | Lỗi: {error_count} ===")


def main():
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = InHangLoatExcelPro(root)
    root.mainloop()


if __name__ == "__main__":
    main()
