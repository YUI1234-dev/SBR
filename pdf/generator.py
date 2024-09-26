# pdf/generator.py

import fitz  # PyMuPDF
import logging
import os
from utils.resources import fit_text_in_box, generate_offsets

def generate_pdf(doc, data, template_pdf_front, template_pdf_back, image_folder, font_name, progress_callback, offset_x, offset_y, app=None):
    """生成 PDF 的主要函數"""
    logging.info(f"開始生成工作證, 共 {len(data)} 張")

    # 定義頁面尺寸，假設 A4 頁面 (595 x 842 點)
    page_width = 595
    page_height = 842

    # 載入模板 PDF
    try:
        template_doc_front = fitz.open(template_pdf_front)
        front_page = template_doc_front.load_page(0)
        front_width, front_height = front_page.rect.width, front_page.rect.height
        logging.info(f"正面模板尺寸: {front_width:.2f} 點 × {front_height:.2f} 點")
    except Exception as e:
        logging.error(f"載入正面模板 PDF 時發生錯誤: {e}")
        return

    try:
        template_doc_back = fitz.open(template_pdf_back)
        back_page = template_doc_back.load_page(0)
        back_width, back_height = back_page.rect.width, back_page.rect.height
        logging.info(f"背面模板尺寸: {back_width:.2f} 點 × {back_height:.2f} 點")
    except Exception as e:
        logging.error(f"載入背面模板 PDF 時發生錯誤: {e}")
        return

    # 使用模板尺寸作為工作證尺寸
    card_width = front_width
    card_height = front_height

    # 計算水平居中的 x_offset
    total_width = (card_width * 2) + 10  # 兩張工作證的總寬度 + 中間的間隔
    x_offset_start = (page_width - total_width) / 2  # 水平居中起始位置

    max_per_page = 8  # 每頁最多 8 張工作證

    # 獲取模板內容在頁面中的偏移量
    front_content_rect = front_page.bound()
    back_content_rect = back_page.bound()

    front_offset_x = front_content_rect.x0
    front_offset_y = front_content_rect.y0

    back_offset_x = back_content_rect.x0
    back_offset_y = back_content_rect.y0

    # 計算正反面內容的偏移差異
    offset_diff_x = back_offset_x - front_offset_x
    offset_diff_y = back_offset_y - front_offset_y

    logging.info(f"正面模板偏移量: x={front_offset_x}, y={front_offset_y}")
    logging.info(f"背面模板偏移量: x={back_offset_x}, y={back_offset_y}")
    logging.info(f"偏移差異: offset_diff_x={offset_diff_x}, offset_diff_y={offset_diff_y}")

    # 使用用戶提供的偏移量
    MANUAL_OFFSET_X = offset_x  # 水平微調，單位：點
    MANUAL_OFFSET_Y = offset_y  # 垂直微調，單位：點

    logging.info(f"手動調整偏移量: MANUAL_OFFSET_X={MANUAL_OFFSET_X}, MANUAL_OFFSET_Y={MANUAL_OFFSET_Y}")

    total_cards = len(data)
    total_pages = (total_cards + max_per_page - 1) // max_per_page  # 計算總頁數

    for i in range(0, total_cards, max_per_page):
        if app and not app.is_generating:
            logging.info("生成過程被取消")
            return

        current_batch = data.iloc[i:i+max_per_page]

        # 正面頁面生成，重置 y_offset_start
        y_offset_start = 20
        page_front = doc.new_page(width=page_width, height=page_height)  # 正面頁面
        logging.info(f"生成第 {i+1} 到 {i+len(current_batch)} 張工作證的正面")
        for j, row in enumerate(current_batch.itertuples(index=False), 0):
            if app and not app.is_generating:
                logging.info("生成過程被取消")
                return

            x_offset = x_offset_start + (j % 2) * (card_width + 10)
            y_offset = y_offset_start + (j // 2) * (card_height + 10)

            # 檢查是否超出頁面高度
            if y_offset + card_height > page_height - 20:
                logging.warning(f"工作證位置超出頁面範圍，跳過: {row.姓名}")
                continue

            # 定義插入矩形框
            insertion_rect = fitz.Rect(x_offset, y_offset, x_offset + card_width, y_offset + card_height)

            # 插入正面模板
            try:
                page_front.show_pdf_page(insertion_rect, template_doc_front, 0)
                logging.info(f"成功插入正面模板到位置: {insertion_rect}")
            except Exception as e:
                logging.error(f"插入正面模板時出錯: {e}")
                continue

            # 插入圖片和文字
            # 定義圖片和文字的矩形框，相對於插入矩形框
            image_rect = fitz.Rect(x_offset + 146, y_offset + 52.5, x_offset + 252.5, y_offset + 141)
            company_rect = fitz.Rect(x_offset + 81.5, y_offset + 53.5, x_offset + 144.5, y_offset + 73)
            name_rect = fitz.Rect(x_offset + 81.5, y_offset + 75, x_offset + 144.5, y_offset + 94.5)
            id_rect = fitz.Rect(x_offset + 81.5, y_offset + 96.5, x_offset + 144.5, y_offset + 115.7)
            valid_rect = fitz.Rect(x_offset + 81.5, y_offset + 118, x_offset + 144.5, y_offset + 139.7)
     
            # 插入圖片
            try:
                if row.圖片路徑:
                    page_front.insert_image(image_rect, filename=row.圖片路徑, keep_proportion=False)
                    logging.info(f"成功插入圖片，位置: {image_rect}")
                else:
                    logging.warning(f"沒有提供圖片路徑，跳過插入圖片: {row.姓名}")
            except Exception as e:
                logging.error(f"插入圖片時出錯: {e}")

            # 插入文字
            fit_text_in_box(page_front, f"{row.公司名稱}", company_rect, max_fontsize=10, min_fontsize=5, font_name=font_name)
            fit_text_in_box(page_front, f"{row.姓名}", name_rect, max_fontsize=10, min_fontsize=5, font_name=font_name)
            fit_text_in_box(page_front, f"{row.工作證號碼}", id_rect, max_fontsize=10, min_fontsize=5, font_name=font_name)
            fit_text_in_box(page_front, f"{row.有效期限}", valid_rect, max_fontsize=10, min_fontsize=5, font_name=font_name)

            logging.info(f"工作證正面生成完成: {row.姓名}")

            if progress_callback:
                progress_callback()

        # 背面頁面生成，重置 y_offset_start
        y_offset_start = 20
        page_back = doc.new_page(width=page_width, height=page_height)  # 背面頁面
        logging.info(f"生成第 {i+1} 到 {i+len(current_batch)} 張工作證的背面")

        # 反轉工作證的順序
        reversed_batch = current_batch.iloc[::-1].reset_index(drop=True)

        for j, row in enumerate(reversed_batch.itertuples(index=False), 0):
            if app and not app.is_generating:
                logging.info("生成過程被取消")
                return

            # 調整背面模板插入位置
            x_offset = x_offset_start + (1 - (j % 2)) * (card_width + 10) - offset_diff_x + offset_x
            y_offset = y_offset_start + (j // 2) * (card_height + 10) - offset_diff_y + offset_y

            # 檢查是否超出頁面高度
            if y_offset + card_height > page_height - 20:
                logging.warning(f"工作證位置超出頁面範圍，跳過背面模板插入")
                continue

            # 定義插入矩形框
            insertion_rect = fitz.Rect(x_offset, y_offset, x_offset + card_width, y_offset + card_height)

            # 插入背面模板
            try:
                page_back.show_pdf_page(insertion_rect, template_doc_back, 0)
                logging.info(f"成功插入背面模板到位置: {insertion_rect}")
            except Exception as e:
                logging.error(f"插入背面模板時出錯: {e}")
                continue

            logging.info(f"工作證背面生成完成")

            if progress_callback:
                progress_callback()

    # 如果總頁數為奇數，添加一個空白頁，以確保雙面列印時頁面數量為偶數
    if total_pages % 2 != 0:
        page_back = doc.new_page(width=page_width, height=page_height)
        logging.info("添加一個空白頁，以確保雙面列印時頁面數量為偶數")
