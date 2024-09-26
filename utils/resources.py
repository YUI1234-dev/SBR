# utils/resources.py

import sys
import os
import ctypes
import logging
from datetime import datetime, timedelta
import math
from PIL import ImageFont
import fitz  # PyMuPDF

def set_dpi_awareness():
    """
    設定應用程式的 DPI 感知，以確保在高 DPI 顯示器上有更好的顯示效果。
    """
    try:
        if sys.platform == "win32":
            ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
    except Exception as e:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
            logging.warning(f"SetProcessDpiAwareness 失敗，回退到 SetProcessDPIAware: {e}")
        except Exception as ex:
            logging.warning(f"設定 DPI 感知時出錯: {e}, {ex}")

def get_system_dpi():
    """
    獲取系統的 DPI 設定，用於調整應用程式的縮放比例。
    """
    try:
        user32 = ctypes.windll.user32
        hdc = user32.GetDC(0)
        LOGPIXELSX = 88
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, LOGPIXELSX)
        user32.ReleaseDC(0, hdc)
        return dpi
    except Exception as e:
        logging.warning(f"獲取系統 DPI 時出錯: {e}")
        return 96  # 預設 DPI

def resource_path(relative_path):
    """
    獲取資源文件的絕對路徑，支援打包後的 EXE。
    """
    try:
        # PyInstaller 會將資源文件解壓縮到 _MEIPASS 目錄
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def sanitize_font_name(font_name):
    """清理字體名稱，去除不允許的字符"""
    return ''.join(c for c in font_name if c.isalnum())

def generate_offsets(n, radius):
    """生成 n 個方向的偏移量，形成一個圓形"""
    angles = [2 * math.pi * i / n for i in range(n)]
    return [(radius * math.cos(angle), radius * math.sin(angle)) for angle in angles]

def fit_text_in_box(page, text, rect, max_fontsize, min_fontsize, font_name):
    """
    縮小字體直到文字適合矩形框，並使文字水平及垂直居中。
    """
    from .fonts import FONT_PATH  # 確保從 fonts 模組匯入 FONT_PATH
    fontsize = max_fontsize

    while fontsize >= min_fontsize:
        try:
            # 使用 Pillow 計算文字寬高
            pillow_font = ImageFont.truetype(FONT_PATH, int(fontsize))
        except IOError:
            logging.error(f"無法載入字體檔案: {FONT_PATH}")
            return min_fontsize

        bbox = pillow_font.getbbox(text)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]

        logging.debug(f"文字 '{text}' 字體大小 {fontsize} 寬度: {text_width}, 高度: {text_height}, 矩形框寬度: {rect.width}, 高度: {rect.height}")

        if text_width <= rect.width and text_height <= rect.height:
            # 計算水平和垂直居中位置
            x_start = rect.x0 + (rect.width - text_width) / 2
            y_start = rect.y0 + (rect.height - text_height) / 2 + text_height

            # 插入文字時強制指定字體名稱
            try:
                # 生成 144 個方向的偏移，半徑為 0.18 點
                offsets = generate_offsets(144, 0.18)
                for dx, dy in offsets:
                    page.insert_text(
                        fitz.Point(x_start + dx, y_start + dy),
                        text,
                        fontsize=fontsize,
                        fontname=font_name,  # 使用清理過的字體名稱
                        fontfile=FONT_PATH
                    )
                # 最後在原位置正常繪製一次
                page.insert_text(
                    fitz.Point(x_start, y_start),
                    text,
                    fontsize=fontsize,
                    fontname=font_name,
                    fontfile=FONT_PATH
                )
                logging.info(f"成功插入文字: {text}，字體大小：{fontsize}")
            except Exception as e:
                logging.error(f"插入文字 '{text}' 時出錯: {e}")
            return fontsize
        else:
            logging.debug(f"嘗試字體大小 {fontsize} 但文字 '{text}' 仍然超出矩形框範圍")
            fontsize -= 1

    logging.warning(f"文字 '{text}' 未能適合矩形框，使用最小字體大小 {min_fontsize}")
    return min_fontsize

def convert_to_minguo_date(minguo_date_str):
    """將民國日期格式 (YYY.MM.DD) 轉換為西元日期，計算新日期為原日期 +3年 -1天，並轉回民國格式"""
    try:
        parts = minguo_date_str.split('.')
        if len(parts) != 3:
            raise ValueError("日期格式不正確")
        minguo_year, month, day = map(int, parts)
        gregorian_year = minguo_year + 1911
        original_date = datetime(gregorian_year, month, day)
        # 計算新日期：原日期 +3年 -1天
        try:
            new_date = original_date.replace(year=original_date.year + 3)
        except ValueError:
            # 處理閏年問題
            new_date = original_date + (datetime(original_date.year + 3, 3, 1) - datetime(original_date.year, 3, 1))
        new_date = new_date - timedelta(days=1)
        new_minguo_year = new_date.year - 1911
        return f"{new_minguo_year}.{new_date.month:02d}.{new_date.day:02d}"
    except Exception as e:
        logging.error(f"日期格式錯誤: {minguo_date_str}, 錯誤: {e}")
        return minguo_date_str  # 返回原始格式

def find_image_path(folder, name):
    """根據姓名在資料夾中尋找對應的圖片，不考慮副檔名"""
    for file in os.listdir(folder):
        file_name, file_ext = os.path.splitext(file)
        if file_name.lower() == name.lower():
            return os.path.join(folder, file)
    return None
