import os
import pandas as pd
from openpyxl.styles import Font, PatternFill

# تنظیمات مسیرها
input_folder = ".."
output_file = "duplicated.xlsx"
excel_files = ["1403.xlsx", "Digikala.xlsx", "others.xlsx"]


def extract_month_day(date_str):
    """استخراج ماه و روز از رشته تاریخ به صورت دستی"""
    if pd.isna(date_str):
        return None

    date_str = str(date_str).strip()

    # فرمت MM/YY یا MM/DD
    if '/' in date_str:
        parts = date_str.split('/')
        if len(parts) == 2:
            month = parts[0].zfill(2)
            day = '01'  # روز پیش‌فرض
            return f"{month}/{day}"
        elif len(parts) == 3:
            month = parts[0].zfill(2)
            day = parts[1].zfill(2)
            return f"{month}/{day}"

    # فرمت YYYY-MM-DD
    elif '-' in date_str and len(date_str.split('-')[0]) == 4:
        parts = date_str.split('-')
        month = parts[1].zfill(2)
        day = parts[2].zfill(2) if len(parts) > 2 else '01'
        return f"{month}/{day}"

    # فرمت‌های دیگر
    elif len(date_str) == 4:  # مثل 1223
        month = date_str[:2].zfill(2)
        day = date_str[2:].zfill(2) if len(date_str) > 2 else '01'
        return f"{month}/{day}"

    return None


def clean_numeric_value(value):
    """پاک‌سازی مقادیر عددی"""
    if pd.isna(value):
        return 0.0

    value_str = str(value).strip()

    # حذف پرانتز و تبدیل به منفی
    if '(' in value_str and ')' in value_str:
        value_str = '-' + value_str.replace('(', '').replace(')', '')

    # حذف کاما و سایر نویسه‌های غیرعددی
    value_str = ''.join(c for c in value_str if c.isdigit() or c in ['-', '.'])

    try:
        return float(value_str)
    except (ValueError, TypeError):
        return 0.0


def process_sheet_data(df, file_name, sheet_name):
    """پردازش داده‌های یک شیت"""
    results = []

    # پیدا کردن ستون تاریخ
    date_col = next((col for col in df.columns if 'تاریخ' in str(col)), None)
    if date_col is None:
        return results

    # ستون‌های عددی هدف
    numeric_cols = [col for col in df.columns if
                    any(x in str(col) for x in ['طاهر', 'مجتبی', 'محمد', 'تنخواه', 'مابه', 'درستی', 'طرف حساب'])]

    for idx, row in df.iterrows():
        # استخراج ماه و روز از تاریخ
        month_day = extract_month_day(row[date_col])
        if not month_day:
            continue

        # پردازش مقادیر عددی
        values = {}
        for col in numeric_cols:
            if col in row and pd.notna(row[col]):
                values[col] = clean_numeric_value(row[col])

        # فقط ردیف‌هایی که حداقل یک مقدار عددی دارند
        if values:
            results.append({
                'file': file_name,
                'sheet': sheet_name,
                'month_day': month_day,
                'values': values,
                'row': idx + 2,
                'description': str(row.get('توضیحات', ''))
            })

            # دیباگ: نمایش داده‌های پردازش شده
            print(f"\nپردازش ردیف {idx + 2}:")
            print(f"تاریخ خام: {row[date_col]} -> ماه/روز: {month_day}")
            print("مقادیر عددی پردازش شده:")
            for col, val in values.items():
                print(f"  {col}: {row[col]} -> {val}")

    return results


def find_duplicates():
    """پیدا کردن مقادیر تکراری"""
    all_data = []

    for file in excel_files:
        file_path = os.path.join(input_folder, file)
        if not os.path.exists(file_path):
            continue

        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                all_data.extend(process_sheet_data(df, file, sheet_name))
        except Exception as e:
            print(f"خطا در پردازش فایل {file}: {str(e)}")

    # گروه‌بندی داده‌ها بر اساس ماه و روز
    date_groups = {}
    for item in all_data:
        if item['month_day'] not in date_groups:
            date_groups[item['month_day']] = []
        date_groups[item['month_day']].append(item)

    # پیدا کردن تکراری‌ها
    duplicates = []
    for month_day, items in date_groups.items():
        if len(items) < 2:
            continue

        # مقایسه تمام جفت‌های ممکن در یک تاریخ
        for i in range(len(items)):
            for j in range(i + 1, len(items)):
                item1 = items[i]
                item2 = items[j]

                # پیدا کردن ستون‌های مشترک
                common_cols = set(item1['values'].keys()) & set(item2['values'].keys())
                similar_values = {}

                for col in common_cols:
                    val1 = item1['values'][col]
                    val2 = item2['values'][col]

                    # مقایسه مقادیر با تحمل خطای کم
                    if abs(val1 - val2) < 0.01:
                        similar_values[col] = val1

                # اگر حداقل دو ستون مشابه دارند
                if len(similar_values) >= 2:
                    duplicates.append({
                        'date': month_day,
                        'file1': item1['file'],
                        'sheet1': item1['sheet'],
                        'row1': item1['row'],
                        'file2': item2['file'],
                        'sheet2': item2['sheet'],
                        'row2': item2['row'],
                        'similar_values': similar_values,
                        'description1': item1['description'],
                        'description2': item2['description']
                    })

    return duplicates


def save_duplicates_to_excel(duplicates):
    """ذخیره نتایج در فایل اکسل"""
    if not duplicates:
        print("⚠️ هیچ مورد تکراری یافت نشد ⚠️")
        return

    # ایجاد دیتافریم
    df = pd.DataFrame(duplicates)

    # ذخیره فایل
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='تکراری‌ها', index=False)

            # استایل‌دهی
            workbook = writer.book
            worksheet = writer.sheets['تکراری‌ها']

            # استایل هدر
            header_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
            header_font = Font(color='FFFFFF', bold=True)

            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font

            # تنظیم عرض ستون‌ها
            for col in worksheet.columns:
                max_length = max(len(str(cell.value)) for cell in col)
                worksheet.column_dimensions[col[0].column_letter].width = max_length + 2

        print(f"✅ نتایج در فایل {output_file} ذخیره شد.")
    except Exception as e:
        print(f"❌ خطا در ذخیره فایل: {str(e)}")


if __name__ == "__main__":
    print("در حال جستجوی مقادیر تکراری...")
    duplicates = find_duplicates()
    save_duplicates_to_excel(duplicates)

    if duplicates:
        print("\nهشدار: مقادیر تکراری/مشابه یافت شد:")
        for dup in duplicates:
            print(f"\n- تاریخ: {dup['date']}")
            print(f"  - فایل اول: {dup['file1']} (شیت: {dup['sheet1']}, ردیف: {dup['row1']})")
            print(f"    توضیحات: {dup['description1']}")
            print(f"  - فایل دوم: {dup['file2']} (شیت: {dup['sheet2']}, ردیف: {dup['row2']})")
            print(f"    توضیحات: {dup['description2']}")
            print("  - مقادیر مشابه:")
            for col, val in dup['similar_values'].items():
                print(f"    * {col}: {val}")