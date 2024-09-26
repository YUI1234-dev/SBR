# data/processing.py

import logging
import pandas as pd
from datetime import datetime, timedelta
import os
from utils.resources import convert_to_minguo_date, find_image_path

def process_data(df, image_folder):
    """處理數據，包括計算訓練日期和匹配圖片路徑"""
    processed_rows = []
    for index, row in df.iterrows():
        # 計算訓練日期：原始日期 + 3 年 - 1 天，格式轉換為民國 YYY.MM.DD
        try:
            original_minguo_date = str(row['訓練日期'])
            minguo_date_str = convert_to_minguo_date(original_minguo_date)
        except Exception as e:
            logging.error(f"處理訓練日期錯誤，使用原始值: {row['訓練日期']}, 錯誤: {e}")
            minguo_date_str = row['訓練日期']

        # 匹配圖片路徑
        name = str(row['姓名']).strip()
        image_path = find_image_path(image_folder, name)
        if not image_path:
            logging.warning(f"找不到圖片: {name}")

        processed_rows.append({
            '公司名稱': row['公司名稱'],
            '姓名': name,
            '工作證號碼': row['工作證號碼'],
            '有效期限': minguo_date_str,
            '圖片路徑': image_path if image_path else ""
        })

    processed_df = pd.DataFrame(processed_rows)
    return processed_df
