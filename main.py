from PIL import ImageGrab, ImageTk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue
import sys
import subprocess
import os

# Local imports
from utils import get_current_directory, get_capture_directory, create_capture_directory
from capture import capture_macro
from pdf_utils import create_pdf_from_images, create_searchable_pdf

class CaptureProgressWindow(tk.Toplevel):
    def __init__(self, parent, cancel_event):
        super().__init__(parent)
        self.title("캡처 진행 중")
        self.geometry("300x100")
        self.cancel_event = cancel_event

        ttk.Style(self).configure("TLabel", font=("Helvetica", 10))
        ttk.Style(self).configure("TButton", font=("Helvetica", 10))

        self.progress_label = ttk.Label(self, text="캡처를 시작합니다...")
        self.progress_label.pack(pady=10, expand=True)

        cancel_button = ttk.Button(self, text="캡처 중단", command=self.cancel_capture)
        cancel_button.pack(pady=5, expand=True)

        self.protocol("WM_DELETE_WINDOW", self.cancel_capture)

    def cancel_capture(self):
        self.cancel_event.set()
        self.destroy()


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("자동 캡처")
        self.geometry("450x350")
        self.minsize(400, 300)

        self.capture_area = None

        # Style
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TLabel", font=("Helvetica", 12))
        style.configure("TButton", font=("Helvetica", 12), padding=5)
        style.configure("TEntry", font=("Helvetica", 12), padding=5)
        style.configure("TNotebook.Tab", font=("Helvetica", 12, "bold"))

        self.current_directory = get_current_directory()
        self.capture_directory = tk.StringVar()
        self.capture_directory.set(get_capture_directory(self.current_directory, "output"))
        create_capture_directory(self.capture_directory.get())

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # Notebook (Tabs)
        notebook = ttk.Notebook(self, padding="10 10 10 10")
        notebook.grid(row=0, column=0, sticky="nsew")

        capture_tab = ttk.Frame(notebook)
        pdf_tab = ttk.Frame(notebook)

        notebook.add(capture_tab, text="캡처")
        notebook.add(pdf_tab, text="PDF 변환")

        # --- Capture Tab ---
        capture_tab.columnconfigure(1, weight=1)

        ttk.Label(capture_tab, text="저장 폴더:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.capture_dir_entry = ttk.Entry(capture_tab, textvariable=self.capture_directory)
        self.capture_dir_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        browse_capture_button = ttk.Button(capture_tab, text="폴더 선택", command=self.browse_capture_folder)
        browse_capture_button.grid(row=0, column=2, padx=5)

        ttk.Label(capture_tab, text="캡처 주기 (초):").grid(
            row=1, column=0, sticky="w", padx=5, pady=5
        )
        self.capture_interval_entry = ttk.Entry(capture_tab)
        self.capture_interval_entry.grid(row=1, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.capture_interval_entry.insert(0, "1.5")

        ttk.Label(capture_tab, text="시작 페이지:").grid(
            row=2, column=0, sticky="w", padx=5, pady=5
        )
        self.start_page_entry = ttk.Entry(capture_tab)
        self.start_page_entry.grid(row=2, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.start_page_entry.insert(0, "1")

        ttk.Label(capture_tab, text="종료 페이지:").grid(
            row=3, column=0, sticky="w", padx=5, pady=5
        )
        self.end_page_entry = ttk.Entry(capture_tab)
        self.end_page_entry.grid(row=3, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        self.end_page_entry.insert(0, "1")

        capture_buttons_frame = ttk.Frame(capture_tab)
        capture_buttons_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=10)
        capture_buttons_frame.columnconfigure(0, weight=1)
        capture_buttons_frame.columnconfigure(1, weight=1)

        select_area_button = ttk.Button(
            capture_buttons_frame, text="캡처 영역 선택", command=self.select_area
        )
        select_area_button.grid(row=0, column=0, sticky="ew", padx=5)

        capture_button = ttk.Button(
            capture_buttons_frame, text="캡처 시작", command=self.start_capture
        )
        capture_button.grid(row=0, column=1, sticky="ew", padx=5)

        open_folder_button = ttk.Button(
            capture_tab, text="캡처 폴더 열기", command=self.open_capture_folder
        )
        open_folder_button.grid(row=5, column=0, columnspan=3, sticky="ew", padx=5, pady=5)

        # --- PDF Tab ---
        pdf_tab.columnconfigure(1, weight=1)

        ttk.Label(pdf_tab, text="이미지 폴더:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.pdf_source_entry = ttk.Entry(pdf_tab)
        self.pdf_source_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        self.pdf_source_entry.insert(0, self.capture_directory.get())

        browse_button = ttk.Button(pdf_tab, text="폴더 선택", command=self.browse_pdf_source)
        browse_button.grid(row=0, column=2, padx=5)

        self.ocr_var = tk.BooleanVar()
        ocr_checkbutton = ttk.Checkbutton(pdf_tab, text="OCR 텍스트 인식 (느림)", variable=self.ocr_var)
        ocr_checkbutton.grid(row=1, column=0, columnspan=3, sticky="w", padx=5, pady=5)

        pdf_button = ttk.Button(pdf_tab, text="PDF로 변환", command=self.convert_to_pdf)
        pdf_button.grid(row=2, column=0, columnspan=3, sticky="ew", padx=5, pady=5)

    def browse_capture_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.capture_directory.set(folder_path)

    def open_capture_folder(self):
        folder_path = self.capture_directory.get()
        if sys.platform == "win32":
            os.startfile(folder_path)
        elif sys.platform == "darwin":
            subprocess.run(["open", folder_path])
        else:
            subprocess.run(["xdg-open", folder_path])

    def select_area(self):
        self.withdraw()
        screenshot = ImageGrab.grab()
        self.selection_window = tk.Toplevel(self)
        self.selection_window.attributes("-fullscreen", True)

        self.canvas = tk.Canvas(self.selection_window, cursor="cross")
        self.canvas.pack(fill="both", expand=True)

        self.tk_img = ImageTk.PhotoImage(screenshot)
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)

        self.selection_window.bind("<ButtonPress-1>", self.on_mouse_press)
        self.selection_window.bind("<B1-Motion>", self.on_mouse_drag)
        self.selection_window.bind("<ButtonRelease-1>", self.on_mouse_release)
        self.selection_window.bind("<Escape>", self.cancel_selection)

    def cancel_selection(self, event=None):
        self.selection_window.destroy()
        self.deiconify()

    def on_mouse_press(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y, outline="red", width=2
        )

    def on_mouse_drag(self, event):
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_mouse_release(self, event):
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)

        x1 = min(self.start_x, end_x) + 2
        y1 = min(self.start_y, end_y) + 2
        x2 = max(self.start_x, end_x) - 2
        y2 = max(self.start_y, end_y) - 2

        self.capture_area = (x1, y1, x2, y2)
        self.selection_window.destroy()
        self.deiconify()
        messagebox.showinfo("성공", f"캡처 영역이 선택되었습니다: {self.capture_area}")

    def start_capture(self):
        if not self.capture_area:
            messagebox.showerror("오류", "캡처 영역을 먼저 선택해주세요.")
            return

        try:
            capture_dir = self.capture_directory.get()
            create_capture_directory(capture_dir)

            capture_interval = float(self.capture_interval_entry.get())
            start_page = int(self.start_page_entry.get())
            end_page = int(self.end_page_entry.get())

            self.cancel_event = threading.Event()
            self.progress_queue = queue.Queue()

            self.progress_window = CaptureProgressWindow(self, self.cancel_event)

            self.capture_thread = threading.Thread(
                target=capture_macro,
                args=(
                    capture_dir,
                    start_page,
                    end_page,
                    capture_interval,
                    self.capture_area,
                    self.cancel_event,
                    self.progress_queue,
                ),
            )
            self.capture_thread.start()
            self.after(100, self.check_progress)

        except ValueError:
            messagebox.showerror("오류", "숫자 값을 올바르게 입력해주세요.")

    def check_progress(self):
        try:
            msg = self.progress_queue.get_nowait()
            if msg == "done":
                self.progress_window.destroy()
                messagebox.showinfo("성공", "캡처가 완료되었습니다.")
            elif msg == "cancelled":
                messagebox.showinfo("취소", "캡처가 중단되었습니다.")
            elif msg.startswith("Error:"):
                self.progress_window.destroy()
                messagebox.showerror("오류", msg)
            else:
                self.progress_window.progress_label.config(text=msg)
                self.after(100, self.check_progress)
        except queue.Empty:
            self.after(100, self.check_progress)

    def browse_pdf_source(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.pdf_source_entry.delete(0, tk.END)
            self.pdf_source_entry.insert(0, folder_path)

    def convert_to_pdf(self):
        source_dir = self.pdf_source_entry.get()
        if not source_dir or not os.path.isdir(source_dir):
            messagebox.showerror("오류", "유효한 이미지 폴더를 선택해주세요.")
            return

        pdf_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Documents", "*.pdf"), ("All Files", "*.*")]
        )
        if not pdf_path:
            return

        if self.ocr_var.get():
            create_searchable_pdf(source_dir, pdf_path)
        else:
            create_pdf_from_images(source_dir, pdf_path)


if __name__ == "__main__":
    app = Application()
    app.mainloop()
