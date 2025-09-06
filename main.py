import pandas as pd
import os
import re
from tqdm import tqdm

# --- تنظیمات ---
excel_file = 'files/qavanin.xlsx'          # فایل ورودی
output_file = 'files/qavanin_output.xlsx'  # فایل خروجی
base_folder = r'D:\pdf advar'

col_shomare_koli = 'شماره کلی'
col_shomare_parvande = 'شماره پرونده و ردیف '
col_tarikh = 'تاريخ‌تصويب 1'
col_doreh = 'دوره قانونگذاری'
col_lib_folder = 'شماره فولدر کتابخانه'

# --- مپ تبدیل دوره فارسی به نام پوشه ---
doreh_to_folder = {
    'اول': 'd1-ok', 'دوم': 'd2-ok', 'سوم': 'd3-ok', 'چهارم': 'd4-ok', 'پنجم': 'd5-ok',
    'ششم': 'd6-ok', 'هفتم': 'd7-ok', 'هشتم': 'd8-ok', 'نهم': 'd9-ok', 'دهم': 'd10-ok',
    'یازدهم': 'd11-ok', 'دوازدهم': 'd12-ok', 'سیزدهم': 'd13-ok', 'چهاردهم': 'd14-ok',
    'پانزدهم': 'd15-ok', 'شانزدهم': 'd16-ok', 'هفدهم': 'd17-ok', 'هجدهم': 'd18-ok',
    'نوزدهم': 'd19-ok', 'بیستم': 'd20-ok', 'بیست و یکم': 'd21-ok',
    'بیست و دوم': 'd22-ok', 'بیست و سوم': 'd23-ok', 'بیست و چهارم': 'd24-ok'
}

# --- آستانه تاریخی انقلاب (شمسی) ---
after_threshold = "1357/11/29"

# --- خواندن اکسل ---
df = pd.read_excel(excel_file)

# --- ستون‌های جدید ---
df["وضعیت وجود پوشه"] = ""
df["تطابق با الگو"] = ""
df["بررسی زیر پوشه"] = ""
df["وجود فایل"] = ""
df["مدیریت خطاها"] = ""

total_rows = len(df)

for idx, row in tqdm(df.iterrows(), total=total_rows, desc="در حال پردازش"):
    try:
        shomare_parvande_val = str(row[col_shomare_parvande]).strip()
        doreh_val = str(row[col_doreh]).strip()
        tarikh_val = str(row[col_tarikh]).strip()
        lib_folder_val = str(row[col_lib_folder]).strip()

        # --- مسیر دوره ---
        period_folder = doreh_to_folder.get(doreh_val)
        if not period_folder:
            df.at[idx, "وضعیت وجود پوشه"] = "ندارد"
            df.at[idx, "تطابق با الگو"] = "غیر مطابق"
            df.at[idx, "مدیریت خطاها"] = "پوشه دوره پیدا نشد"
            continue

        # --- مسیر کامل ---
        if tarikh_val and tarikh_val >= after_threshold:
            full_path = os.path.join(base_folder, "baed az enghelab eslami", period_folder)
        else:
            full_path = os.path.join(base_folder, "ghabl az enghelab", "majlis shorayeh melli", period_folder)

        # --- پیدا کردن پوشه ---
        target_folder_name_clean = shomare_parvande_val.replace(" ", "").strip()
        found_folders = []
        if os.path.exists(full_path):
            for root, dirs, files in os.walk(full_path):
                for d in dirs:
                    d_clean = d.replace(" ", "").strip()
                    if d_clean == target_folder_name_clean or d_clean.startswith(target_folder_name_clean):
                        found_folders.append(os.path.join(root, d))

            if not found_folders:
                df.at[idx, "وضعیت وجود پوشه"] = "ندارد"
                df.at[idx, "تطابق با الگو"] = "غیر مطابق"
                df.at[idx, "مدیریت خطاها"] = "پوشه پیدا نشد"
                continue
        else:
            df.at[idx, "وضعیت وجود پوشه"] = "ندارد"
            df.at[idx, "تطابق با الگو"] = "غیر مطابق"
            df.at[idx, "مدیریت خطاها"] = "مسیر دوره وجود ندارد"
            continue

        # ✅ از اینجا به بعد یعنی پوشه وجود دارد
        df.at[idx, "وضعیت وجود پوشه"] = "دارد"

        # --- بررسی تطابق الگو ---
        pattern = r"^\d+$|^\d+\s*مکرر\d*$"
        if re.match(pattern, shomare_parvande_val.replace(" ", "")):
            df.at[idx, "تطابق با الگو"] = "مطابق"
        else:
            df.at[idx, "تطابق با الگو"] = "غیر مطابق"

        # --- بررسی زیرپوشه ---
        has_subfolder = any(
            [os.path.isdir(os.path.join(folder, d)) for folder in found_folders for d in os.listdir(folder)]
        )
        if has_subfolder:
            df.at[idx, "بررسی زیر پوشه"] = "دارد"
            df.at[idx, "مدیریت خطاها"] = "خطای زیر پوشه"
        else:
            df.at[idx, "بررسی زیر پوشه"] = "ندارد"

        # --- بررسی فایل‌ها ---
        has_file = False
        for folder in found_folders:
            files = os.listdir(folder)
            pdfs = [f for f in files if f.lower().endswith('.pdf')]
            images = [
                f for f in files
                if lib_folder_val and lib_folder_val.lower() != "nan"
                and os.path.splitext(f)[0].replace(" ", "").strip().startswith(lib_folder_val.replace(" ", "").strip())
                and f.lower().endswith(('.jpg', '.jpeg', '.png', '.tif', '.bmp'))
            ]
            if pdfs or images:
                has_file = True
                break

        if has_file:
            df.at[idx, "وجود فایل"] = "دارد"
        else:
            df.at[idx, "وجود فایل"] = "ندارد"
            if not df.at[idx, "مدیریت خطاها"]:
                df.at[idx, "مدیریت خطاها"] = "نبود فایل"

        # --- اگر هیچ خطا ثبت نشده بود
        if not df.at[idx, "مدیریت خطاها"]:
            df.at[idx, "مدیریت خطاها"] = "بدون خطا"

    except Exception as e:
        df.at[idx, "وضعیت وجود پوشه"] = "نامشخص"
        df.at[idx, "تطابق با الگو"] = "نامشخص"
        df.at[idx, "مدیریت خطاها"] = str(e)

# --- ذخیره در اکسل ---
df.to_excel(output_file, index=False)
print(f"✅ خروجی ذخیره شد در: {output_file}")
