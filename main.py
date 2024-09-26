# main.py

import logging
import sys
import tkinter as tk
from ui.main_window import MainWindow
from utils.resources import set_dpi_awareness, get_system_dpi

def main():
    set_dpi_awareness()  # 設定 DPI 感知

    logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")

    root = tk.Tk()

    # 獲取系統 DPI 以調整 Tkinter 縮放比例
    try:
        if sys.platform == "win32":
            dpi = get_system_dpi()
            if dpi == 0:
                raise ValueError("獲取 DPI 失敗")
            scaling_factor = dpi / 96
            root.tk.call('tk', 'scaling', scaling_factor)
            logging.info(f"系統 DPI: {dpi}, 縮放因子: {scaling_factor}")
    except Exception as e:
        logging.warning(f"獲取系統 DPI 時出錯: {e}")
        # 設定預設縮放因子
        root.tk.call('tk', 'scaling', 1.0)

    app = MainWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main()
