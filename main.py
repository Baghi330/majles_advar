import pandas as pd
import os
import re
from tqdm import tqdm

# --- تنظیمات ---
excel_file = 'files/qavanin.xlsx'          # فایل ورودی
output_file = 'files/qavanin_output8.xlsx'  # فایل خروجی
base_folder = r'D:\pdf advar'

# ستون‌ها
col_shomare_koli = 'شماره کلی'
col_shomare_parvande = 'شماره پرونده و ردیف '
col_tarikh = 'تاريخ‌تصويب 1'
col_doreh = 'دوره قانونگذاری'
col_lib_folder = 'شماره فولدر کتابخانه'

# --- مپ تبدیل دوره فارسی به نام پوشه (قبل انقلاب) ---
doreh_to_folder_before = {
    'اول': 'd1-ok', 'دوم': 'd2-ok', 'سوم': 'd3-ok', 'چهارم': 'd4-ok', 'پنجم': 'd5-ok',
    'ششم': 'd6-ok', 'هفتم': 'd7-ok', 'هشتم': 'd8-ok', 'نهم': 'd9-ok', 'دهم': 'd10-ok',
    'یازدهم': 'd11-ok', 'دوازدهم': 'd12-ok', 'سیزدهم': 'd13-ok', 'چهاردهم': 'd14-ok',
    'پانزدهم': 'd15-ok', 'شانزدهم': 'd16-ok', 'هفدهم': 'd17-ok', 'هجدهم': 'd18-ok',
    'نوزدهم': 'd19-ok', 'بیستم': 'd20-ok', 'بیست و یکم': 'd21-ok',
    'بیست و دوم': 'd22-ok', 'بیست و سوم': 'd23-ok', 'بیست و چهارم': 'd24-ok'
}

# --- مپ تبدیل دوره برای بعد انقلاب ---
doreh_to_folder_after = {
    str(i): f"d.{i}" for i in range(1, 13)   # d.1 تا d.12
}

# --- تابع پردازش یک دیتافریم ---
def process_dataframe(df, period_map, base_subfolder):
    df = df.copy()
    df["وضعیت وجود پوشه"] = ""
    df["تطابق با الگو"] = ""
    df["بررسی زیر پوشه"] = ""
    df["وجود فایل"] = ""
    df["مدیریت خطاها"] = ""

    for idx, row in tqdm(df.iterrows(), total=len(df), desc=f"پردازش {base_subfolder}"):
        try:
            shomare_parvande_val = str(row[col_shomare_parvande]).strip()
            doreh_val = str(row[col_doreh]).strip()
            lib_folder_val = str(row[col_lib_folder]).strip()

            # --- مسیر دوره ---
            period_folder = period_map.get(doreh_val)
            if not period_folder:
                df.at[idx, "وضعیت وجود پوشه"] = "ندارد"
                df.at[idx, "تطابق با الگو"] = "غیر مطابق"
                df.at[idx, "مدیریت خطاها"] = "پوشه دوره پیدا نشد"
                continue

            # --- مسیر کامل ---
            full_path = os.path.join(base_folder, base_subfolder, period_folder)

            # --- پیدا کردن پوشه فقط در سطح مستقیم ---
            target_folder_name_clean = shomare_parvande_val.replace(" ", "").strip()
            found_folders = []
            if os.path.exists(full_path):
                for d in os.listdir(full_path):  # فقط پوشه‌های مستقیم
                    folder_path = os.path.join(full_path, d)
                    if os.path.isdir(folder_path):
                        d_clean = d.replace(" ", "").strip()
                        if d_clean == target_folder_name_clean:
                          found_folders.append(folder_path)

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

            # ✅ پوشه پیدا شد
            df.at[idx, "وضعیت وجود پوشه"] = "دارد"

            # --- تطابق الگو ---
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
                # خطای زیر پوشه را موقتا قرار میدهیم و بعد فایل‌ها را هم بررسی می‌کنیم
                subfolder_error = "خطای زیر پوشه"
            else:
                df.at[idx, "بررسی زیر پوشه"] = "ندارد"
                subfolder_error = ""

            # --- بررسی فایل‌ها (با پشتیبانی از زیرپوشه‌ها) ---
            has_file = False
            matched_file = False
            for folder in found_folders:
                for root, _, files in os.walk(folder):  # بررسی داخل زیرپوشه‌ها
                    if files:
                        has_file = True
                    for f in files:
                        f_lower = f.lower()
                        if f_lower.endswith(('.pdf', '.jpg', '.jpeg', '.png', '.tif', '.bmp')):
                            # بررسی تطابق با شماره فولدر کتابخانه
                            if (
                                lib_folder_val
                                and lib_folder_val.lower() != "nan"
                                and lib_folder_val.replace(" ", "").strip() in os.path.splitext(f)[0].replace(" ", "").strip()
                            ):
                                matched_file = True
                                break
                    if matched_file:
                        break
                if matched_file:
                    break

            # --- بروزرسانی ستون‌ها ---
            if not has_file:
                df.at[idx, "وجود فایل"] = "ندارد"
                df.at[idx, "مدیریت خطاها"] = "نبود فایل"
            else:
                df.at[idx, "وجود فایل"] = "دارد"
                if matched_file:
                    # فایل بود و تطابق داشت
                    if subfolder_error:
                        df.at[idx, "مدیریت خطاها"] = subfolder_error
                    else:
                        df.at[idx, "مدیریت خطاها"] = "بدون خطا"
                else:
                    # فایل بود ولی تطابق نداشت
                    if subfolder_error:
                        df.at[idx, "مدیریت خطاها"] = subfolder_error + " / عدم تطابق با شماره فولدر کتابخانه"
                    else:
                        df.at[idx, "مدیریت خطاها"] = "عدم تطابق با شماره فولدر کتابخانه"

        except Exception as e:
            df.at[idx, "وضعیت وجود پوشه"] = "نامشخص"
            df.at[idx, "تطابق با الگو"] = "نامشخص"
            df.at[idx, "مدیریت خطاها"] = str(e)

    return df


# --- خواندن هر دو شیت از اکسل ---
df_before = pd.read_excel(excel_file, sheet_name=0)  # قبل انقلاب
df_after = pd.read_excel(excel_file, sheet_name=1)   # بعد انقلاب

# --- پردازش ---
result_before = process_dataframe(df_before, doreh_to_folder_before, os.path.join("ghabl az enghelab", "majlis shorayeh melli"))
result_after = process_dataframe(df_after, doreh_to_folder_after, "baed az enghelab eslami")

# --- شمارش خطاها ---
error_counts = pd.concat([result_before["مدیریت خطاها"], result_after["مدیریت خطاها"]])
error_counts = error_counts.value_counts().reset_index()
error_counts.columns = ["نوع خطا", "تعداد"]

# --- ذخیره در خروجی ---
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    result_before.to_excel(writer, index=False, sheet_name="قبل انقلاب")
    result_after.to_excel(writer, index=False, sheet_name="بعد انقلاب")
    error_counts.to_excel(writer, index=False, sheet_name="گزارش خطاها")

print(f"✅ خروجی ذخیره شد در: {output_file}")
