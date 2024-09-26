# ui/main_window.py

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.font as tkFont
import pandas as pd
import logging
import json
import os
import threading
import queue
from data.processing import process_data
from pdf.generator import generate_pdf
from utils.resources import resource_path, sanitize_font_name
from utils.fonts import FONT_PATH
from ui.log_handler import TextHandler
import fitz  # PyMuPDF

class MainWindow:
    """工作證生成器主應用程式"""
    def __init__(self, root):
        self.root = root
        self.root.title("SBR工作證生成器")

        # 設定全局字體大小
        default_font = tkFont.nametofont("TkDefaultFont")
        default_font.configure(size=12)

        # 設置應用程式圖標
        try:
            self.root.iconbitmap(resource_path("app_icon.ico"))  # 替換為您的 ICO 文件名
        except Exception as e:
            logging.error(f"設置應用程式圖標時出錯: {e}")

        # 初始化變數
        self.excel_file = tk.StringVar()
        self.image_folder = tk.StringVar()
        self.pdf_filename = tk.StringVar(value="workpasses_double_sided.pdf")
        self.data = pd.DataFrame()

        # 手動偏移量變數
        self.offset_x = tk.DoubleVar()
        self.offset_y = tk.DoubleVar()

        # 加載設定
        self.load_settings()

        # 設置進度隊列
        self.queue = queue.Queue()
        self.progress_var = tk.DoubleVar()
        self.is_generating = False  # 用於判斷是否正在生成

        # 設置 UI 元素
        self.setup_ui()

        # 設置日誌處理器
        self.setup_logging()
        
        # 視窗置中
        self.center_window()

    def center_window(self):
        """將視窗置中"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def setup_ui(self):
        """設置應用程式的 UI 元素"""
        # 定義字體
        entry_font = tkFont.Font(size=12)
        text_font = tkFont.Font(size=12)
        tree_font = tkFont.Font(size=12)
        heading_font = tkFont.Font(size=12, weight='bold')

        # Excel 選擇
        frame_excel = ttk.Frame(self.root, padding="10")
        frame_excel.grid(row=0, column=0, sticky="W", padx=5, pady=5)

        ttk.Label(frame_excel, text="Excel 文件:", font=entry_font).grid(row=0, column=0, sticky="W")
        ttk.Entry(frame_excel, textvariable=self.excel_file, width=50, font=entry_font).grid(row=0, column=1, padx=5)
        ttk.Button(frame_excel, text="選擇", command=self.select_excel).grid(row=0, column=2)

        # 圖片資料夾選擇
        frame_image = ttk.Frame(self.root, padding="10")
        frame_image.grid(row=1, column=0, sticky="W", padx=5, pady=5)

        ttk.Label(frame_image, text="圖片資料夾:", font=entry_font).grid(row=0, column=0, sticky="W")
        ttk.Entry(frame_image, textvariable=self.image_folder, width=50, font=entry_font).grid(row=0, column=1, padx=5)
        ttk.Button(frame_image, text="選擇", command=self.select_image_folder).grid(row=0, column=2)

        # 偏移量設定
        frame_offset = ttk.Frame(self.root, padding="10")
        frame_offset.grid(row=2, column=0, sticky="W", padx=5, pady=5)

        ttk.Label(frame_offset, text="水平偏移量 (點):", font=entry_font).grid(row=0, column=0, sticky="W")
        ttk.Entry(frame_offset, textvariable=self.offset_x, width=10, font=entry_font).grid(row=0, column=1, padx=5)

        ttk.Label(frame_offset, text="垂直偏移量 (點):", font=entry_font).grid(row=0, column=2, sticky="W")
        ttk.Entry(frame_offset, textvariable=self.offset_y, width=10, font=entry_font).grid(row=0, column=3, padx=5)

        # 預覽區域
        frame_preview = ttk.Frame(self.root, padding="10")
        frame_preview.grid(row=3, column=0, sticky="NSEW", padx=5, pady=5)

        ttk.Label(frame_preview, text="預覽:", font=entry_font).grid(row=0, column=0, sticky="W")

        # 增加水平和垂直滾動條
        tree_scroll_y = ttk.Scrollbar(frame_preview, orient="vertical")
        tree_scroll_x = ttk.Scrollbar(frame_preview, orient="horizontal")

        self.tree = ttk.Treeview(
            frame_preview,
            columns=("公司名稱", "姓名", "工作證號碼", "有效期限", "圖片路徑"),
            show='headings',
            height=10,
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set
        )
        self.tree.grid(row=1, column=0, sticky="NSEW")

        # 設置 Treeview 的字體
        style = ttk.Style()
        style.configure("Treeview", font=tree_font)  # 設定內容字體
        style.configure("Treeview.Heading", font=heading_font)  # 設定標題字體

        # 配置滾動條
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_y.grid(row=1, column=1, sticky='NS')

        tree_scroll_x.config(command=self.tree.xview)
        tree_scroll_x.grid(row=2, column=0, sticky='EW')

        # 設定欄位標題和寬度
        self.tree.heading("公司名稱", text="公司名稱")
        self.tree.column("公司名稱", width=150, anchor='center')

        self.tree.heading("姓名", text="姓名")
        self.tree.column("姓名", width=100, anchor='center')

        self.tree.heading("工作證號碼", text="工作證號碼")
        self.tree.column("工作證號碼", width=120, anchor='center')

        self.tree.heading("有效期限", text="有效期限")
        self.tree.column("有效期限", width=100, anchor='center')

        self.tree.heading("圖片路徑", text="圖片路徑")
        self.tree.column("圖片路徑", width=200, anchor='center')  # 增加寬度以顯示完整路徑

        # PDF 檔名
        frame_pdf = ttk.Frame(self.root, padding="10")
        frame_pdf.grid(row=4, column=0, sticky="W", padx=5, pady=5)

        ttk.Label(frame_pdf, text="PDF 檔名:", font=entry_font).grid(row=0, column=0, sticky="W")
        ttk.Entry(frame_pdf, textvariable=self.pdf_filename, width=50, font=entry_font).grid(row=0, column=1, padx=5)
        ttk.Button(frame_pdf, text="選擇保存位置", command=self.select_pdf_filename).grid(row=0, column=2)

        # 進度條
        frame_progress = ttk.Frame(self.root, padding="10")
        frame_progress.grid(row=5, column=0, sticky="W", padx=5, pady=5)

        self.progress_bar = ttk.Progressbar(frame_progress, variable=self.progress_var, maximum=100, length=400)
        self.progress_bar.grid(row=0, column=0, pady=5)

        # 進度百分比標籤
        self.progress_label = ttk.Label(frame_progress, text="0%", font=entry_font)
        self.progress_label.grid(row=1, column=0, pady=5)

        # 日誌顯示區域
        frame_log = ttk.Frame(self.root, padding="10")
        frame_log.grid(row=6, column=0, sticky="NSEW", padx=5, pady=5)

        ttk.Label(frame_log, text="日誌:", font=entry_font).grid(row=0, column=0, sticky="W")
        self.log_text = tk.Text(frame_log, height=10, width=80, state='disabled', wrap='word', font=text_font)
        self.log_text.grid(row=1, column=0, sticky="NSEW")

        # 增加滾動條到日誌
        log_scroll_y = ttk.Scrollbar(frame_log, orient="vertical", command=self.log_text.yview)
        log_scroll_y.grid(row=1, column=1, sticky='NS')
        self.log_text.configure(yscrollcommand=log_scroll_y.set)

        # 生成和取消按鈕
        frame_buttons = ttk.Frame(self.root, padding="10")
        frame_buttons.grid(row=7, column=0, sticky="W", padx=5, pady=5)

        self.generate_button = ttk.Button(frame_buttons, text="生成 PDF", command=self.start_generate_pdf)
        self.generate_button.grid(row=0, column=0, padx=5)

        self.cancel_button = ttk.Button(frame_buttons, text="取消", command=self.cancel_generate_pdf, state='disabled')
        self.cancel_button.grid(row=0, column=1, padx=5)

        # 設定 Grid 權重，使 UI 元素隨視窗調整大小
        self.root.grid_rowconfigure(3, weight=1)  # 預覽區域
        self.root.grid_rowconfigure(6, weight=1)  # 日誌區域
        self.root.grid_columnconfigure(0, weight=1)

    def setup_logging(self):
        """設置日誌處理器，將日誌寫入 Text 控件"""
        text_handler = TextHandler(self.log_text)
        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        text_handler.setFormatter(formatter)
        logging.getLogger().addHandler(text_handler)

    def select_excel(self):
        """選擇 Excel 文件"""
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_file.set(file_path)
            self.load_data()

    def select_image_folder(self):
        """選擇圖片資料夾"""
        folder_path = filedialog.askdirectory(title="選擇圖片資料夾")
        if folder_path:
            self.image_folder.set(folder_path)
            self.load_data()

    def select_pdf_filename(self):
        """選擇保存 PDF 文件的位置"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="保存 PDF 文件"
        )
        if file_path:
            self.pdf_filename.set(file_path)

    def load_data(self):
        """加載並處理數據"""
        excel_path = self.excel_file.get()
        image_folder = self.image_folder.get()

        if not excel_path or not image_folder:
            return

        try:
            df = pd.read_excel(excel_path, engine='openpyxl')
            required_columns = ['公司名稱', '姓名', '工作證號碼', '訓練日期']
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("錯誤", f"Excel 文件缺少必要的欄位: {required_columns}")
                return

            self.data = process_data(df, image_folder)

            # 更新預覽
            for item in self.tree.get_children():
                self.tree.delete(item)
            for _, row in self.data.iterrows():
                self.tree.insert("", "end", values=tuple(row))

            logging.info("數據加載並預處理完成")
        except Exception as e:
            logging.error(f"加載數據時出錯: {e}")
            messagebox.showerror("錯誤", f"加載數據時出錯: {e}")

    def start_generate_pdf(self):
        """開始生成 PDF"""
        if self.data.empty:
            messagebox.showwarning("警告", "沒有可生成的數據。請確認已選擇 Excel 文件和圖片資料夾。")
            return

        if self.is_generating:
            return  # 已經在生成中，防止重複點擊

        pdf_filename = self.pdf_filename.get()
        if not pdf_filename.endswith(".pdf"):
            pdf_filename += ".pdf"

        # 檢查模板文件是否存在
        template_pdf_front = resource_path("templates/工作證模板(正).pdf")  # 正面模板
        template_pdf_back = resource_path("templates/工作證模板(背).pdf")  # 背面模板

        if not os.path.exists(template_pdf_front) or not os.path.exists(template_pdf_back):
            messagebox.showerror("錯誤", "模板 PDF 文件不存在。請確認 '工作證模板(正).pdf' 和 '工作證模板(背).pdf' 在 templates 目錄中。")
            return

        # 載入字體
        try:
            font_name = sanitize_font_name("kaiu")  # 假設字體名稱為 'kaiu'
            logging.info(f"成功載入字體: kaiu.ttf, 字體名稱: {font_name}")
        except Exception as e:
            logging.error(f"載入字體檔案時出錯: {e}")
            messagebox.showerror("錯誤", f"載入字體檔案時出錯: {e}")
            return

        # 創建 PDF 文檔
        self.doc = fitz.open()

        # 設置進度條
        total_steps = len(self.data) * 2  # 正面和背面
        self.progress_var.set(0)
        self.progress_bar['maximum'] = total_steps

        # 獲取偏移量
        offset_x = self.offset_x.get()
        offset_y = self.offset_y.get()

        # 定義進度更新回調
        def progress_callback():
            self.queue.put(1)

        # 開始 PDF 生成的線程
        self.is_generating = True
        self.generate_button.config(state='disabled')
        self.cancel_button.config(state='normal')
        self.thread = threading.Thread(target=self.generate_pdf_thread, args=(
            self.doc, self.data, template_pdf_front, template_pdf_back, self.image_folder.get(), font_name, pdf_filename, progress_callback, offset_x, offset_y))
        self.thread.start()

        # 開始檢查進度
        self.root.after(100, self.process_queue)

    def cancel_generate_pdf(self):
        """取消生成 PDF"""
        if self.is_generating:
            if messagebox.askyesno("確認", "確定要取消生成 PDF 嗎？"):
                self.is_generating = False
                self.generate_button.config(state='normal')
                self.cancel_button.config(state='disabled')
                self.progress_var.set(0)
                self.progress_label.config(text="0%")
                logging.info("已取消生成 PDF")
                # 關閉文檔
                if hasattr(self, 'doc'):
                    self.doc.close()

    def generate_pdf_thread(self, doc, data, template_pdf_front, template_pdf_back, image_folder, font_name, pdf_filename, progress_callback, offset_x, offset_y):
        """PDF 生成線程函數"""
        try:
            generate_pdf(doc, data, template_pdf_front, template_pdf_back, image_folder, font_name, progress_callback, offset_x, offset_y, self)
            # 將檔名設置到 PDF metadata
            doc.set_metadata({"title": os.path.basename(pdf_filename)})
            if self.is_generating:
                doc.save(pdf_filename)
                logging.info(f"成功保存 PDF 工作證文件: {pdf_filename}")
                self.queue.put("done")
            doc.close()
            # 保存設定
            self.save_settings()
        except Exception as e:
            logging.error(f"保存 PDF 文件時出錯: {e}")
            self.queue.put(f"error:{e}")
            if hasattr(self, 'doc'):
                doc.close()

    def process_queue(self):
        """處理進度隊列，用於更新進度條和處理完成訊息"""
        try:
            while True:
                msg = self.queue.get_nowait()
                if isinstance(msg, int):
                    current_progress = self.progress_var.get() + msg
                    self.progress_var.set(current_progress)
                    percentage = int((current_progress / self.progress_bar['maximum']) * 100)
                    self.progress_label.config(text=f"{percentage}%")
                elif isinstance(msg, str):
                    if msg == "done":
                        messagebox.showinfo("成功", f"成功生成 PDF 工作證文件: {self.pdf_filename.get()}")
                        self.progress_var.set(0)
                        self.progress_label.config(text="0%")
                    elif msg.startswith("error"):
                        error_msg = msg.split(":", 1)[1]
                        messagebox.showerror("錯誤", f"保存 PDF 文件時出錯: {error_msg}")
                        self.progress_var.set(0)
                        self.progress_label.config(text="0%")
                    self.is_generating = False
                    self.generate_button.config(state='normal')
                    self.cancel_button.config(state='disabled')
        except queue.Empty:
            pass
        finally:
            if self.is_generating:
                self.root.after(100, self.process_queue)

    def load_settings(self):
        """加載設定"""
        if os.path.exists('config/config.json'):
            with open('config/config.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
                self.offset_x.set(config.get('offset_x', -1.8))
                self.offset_y.set(config.get('offset_y', -1.6))
        else:
            self.offset_x.set(-1.8)
            self.offset_y.set(-1.6)

    def save_settings(self):
        """保存設定"""
        if not os.path.exists('config'):
            os.makedirs('config')
        config = {
            'offset_x': self.offset_x.get(),
            'offset_y': self.offset_y.get()
        }
        with open('config/config.json', 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
