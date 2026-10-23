# ============================================================
# ОТЧЁТ ПО ЗАПРАВКАМ С ДИНАМИКОЙ
# Читает все файлы .xls и .xlsx из папки «Заправки», объединяет,
# строит сводку по водителям и динамику по дням
#
# ИСПРАВЛЕНИЯ В ЭТОЙ ВЕРСИИ:
#   1. Скрипт смотрит в папку «Заправки»
#   2. Папку можно передать аргументом
#   3. Печатает абсолютный путь папки
#   4. Понятная ошибка, если папки нет
#   5. reset_index после фильтра
#   6. Очистка чисел: "50,5" → 50.5
#   7. Чтение .xls через python-calamine (BIFF5/7/8 — все форматы)
#   8. enumerate вместо row_idx+2 в циклах Excel
# ============================================================
import os
import sys
import glob
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# НАСТРОЙКИ
# ============================================================
ПАПКА_ПО_УМОЛЧАНИЮ = 'Заправки'

МАСКИ_ФАЙЛОВ = [
    'transactional_report_*.xls',
    'transactional_report_*.xlsx',
    'Отчет_*.xls',
    'Отчет_*.xlsx',
]


# ============================================================
# СТИЛИ ДЛЯ EXCEL
# ============================================================
def style_header(ws, row, columns):
    header_fill = PatternFill(start_color='667EEA', end_color='667EEA', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', size=11)
    border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC'),
    )
    for col_idx, col_name in enumerate(columns, 1):
        cell = ws.cell(row=row, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border


def style_cell(cell, align='left'):
    border = Border(
        left=Side(style='thin', color='EEEEEE'),
        right=Side(style='thin', color='EEEEEE'),
        top=Side(style='thin', color='EEEEEE'),
        bottom=Side(style='thin', color='EEEEEE'),
    )
    cell.border = border
    cell.alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)


def auto_width(ws, min_width=10, max_width=50):
    for col in ws.columns:
        max_length = 0
        col_letter = None
        for cell in col:
            if not hasattr(cell, 'column_letter'):
                continue
            if col_letter is None:
                col_letter = cell.column_letter
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        if col_letter:
            adjusted = min(max(max_length + 2, min_width), max_width)
            ws.column_dimensions[col_letter].width = adjusted


# ============================================================
# ЧТЕНИЕ ФАЙЛОВ
# ============================================================
def read_any_excel(filepath):
    """
    Читает .xls или .xlsx.

    Для .xls используем python-calamine — она читает BIFF4/5/7/8
    (включая старые форматы Excel 5.0/95, в которых приходят выгрузки
    топливных карт).

    Резервные варианты — на случай, если calamine не установлена:
    xlrd (для BIFF8), затем HTML.
    """
    ext = os.path.splitext(filepath)[1].lower()

    if ext == '.xls':
        # --- Попытка 1: python-calamine ---
        try:
            return pd.read_excel(filepath, engine='calamine', header=0)
        except ImportError:
            print("      ⚠️ python-calamine не установлена, пробую xlrd…")
        except Exception as e:
            print(f"      ⚠️ calamine не справился ({e.__class__.__name__}), пробую xlrd…")

        # --- Попытка 2: xlrd (для BIFF8) ---
        try:
            return pd.read_excel(filepath, engine='xlrd', header=0)
        except Exception as e:
            print(f"      ⚠️ xlrd не справился ({e.__class__.__name__}), пробую HTML…")

        # --- Попытка 3: HTML (бывает, что .xls — это HTML) ---
        try:
            tables = pd.read_html(filepath, header=0)
            if not tables:
                raise ValueError("в файле нет HTML-таблиц")
            return tables[0]
        except Exception as e2:
            raise RuntimeError(f"все способы провалились: {e2}")
    else:
        return pd.read_excel(filepath, engine='openpyxl', header=0)


# ============================================================
# ОЧИСТКА ЧИСЛОВЫХ КОЛОНОК
# ============================================================
def clean_numeric(series):
    """
    Превращает «грязную» колонку в числа.
      - неразрывный пробел (\xa0) — из 1С
      - обычные пробелы как разделитель тысяч: "1 234,5"
      - запятая как десятичный разделитель: "50,5" → "50.5"
    """
    s = series.astype(str)
    s = s.str.replace('\xa0', '', regex=False)
    s = s.str.replace(' ', '', regex=False)
    s = s.str.replace(',', '.', regex=False)
    return pd.to_numeric(s, errors='coerce')


# ============================================================
# ГЛАВНАЯ
# ============================================================
def main():
    print("=" * 70)
    print("📊 ОТЧЁТ ПО ЗАПРАВКАМ С ДИНАМИКОЙ")
    print("=" * 70)

    # ---------- 0. ПАПКА С ФАЙЛАМИ ----------
    if len(sys.argv) > 1:
        папка = sys.argv[1]
    else:
        папка = ПАПКА_ПО_УМОЛЧАНИЮ

    абс_папка = os.path.abspath(папка)
    print(f"\n📁 Папка с файлами: {абс_папка}")

    if not os.path.isdir(абс_папка):
        print(f"\n❌ Папка не существует: {абс_папка}")
        print(f"   Создай её или передай другую папку аргументом:")
        print(f"       python zapravki_dinamika.py ИмяПапки")
        return

    # ---------- 1. ИЩЕМ ФАЙЛЫ ----------
    files = []
    for mask in МАСКИ_ФАЙЛОВ:
        pattern = os.path.join(абс_папка, mask)
        files.extend(glob.glob(pattern))

    files = sorted(set(files))

    print(f"\n1. Найдено файлов: {len(files)}")
    for f in files:
        print(f"   • {os.path.basename(f)}")

    if not files:
        print("\n❌ Не найдено ни одного файла для обработки.")
        print(f"   Проверь имена файлов в папке: {абс_папка}")
        print(f"   Ожидаются маски: {МАСКИ_ФАЙЛОВ}")
        return

    # ---------- 2. ЧИТАЕМ И ОБЪЕДИНЯЕМ ----------
    print("\n2. Читаю файлы…")
    all_dfs = []
    files_ok = 0
    files_fail = 0

    for f in files:
        try:
            df = read_any_excel(f)
            all_dfs.append(df)
            print(f"   ✅ {os.path.basename(f)} — {len(df)} строк")
            files_ok += 1
        except Exception as e:
            print(f"   ❌ {os.path.basename(f)} — ошибка: {e}")
            files_fail += 1

    if not all_dfs:
        print("\n❌ Ни один файл не удалось прочитать.")
        return

    df_all = pd.concat(all_dfs, ignore_index=True)
    print(f"\n   Всего строк: {len(df_all)}")
    print(f"   Успешно: {files_ok}, с ошибками: {files_fail}")

    # ---------- 3. КОЛОНКИ К ЕДИНОМУ ВИДУ ----------
    print("\n3. Привожу колонки к единому виду…")

    переименование = {
        'Закреплена за': 'Водитель',
        'Кол-во': 'Литры',
        'Сумма со скидкой (в валюте клиента)': 'Сумма',
    }
    df_all = df_all.rename(columns=переименование)

    обязательные = ['Дата и время', 'Водитель', 'Литры', 'Вид операции']
    отсутствуют = [c for c in обязательные if c not in df_all.columns]
    if отсутствуют:
        print(f"   ❌ Нет колонок: {отсутствуют}")
        print(f"   Доступные колонки: {list(df_all.columns)}")
        return

    # ---------- 4. ФИЛЬТР "ПРЯМАЯ" ----------
    print("\n4. Фильтрую реальные заправки (Вид операции = Прямая)…")
    df_all = df_all[df_all['Вид операции'].astype(str).str.strip() == 'Прямая'].copy()
    df_all = df_all.reset_index(drop=True)
    print(f"   После фильтра: {len(df_all)} строк")

    # ---------- 5. ЛИТРЫ → ЧИСЛО ----------
    print("\n5. Привожу литры к числам…")
    df_all['Литры'] = clean_numeric(df_all['Литры'])

    bad_litry = df_all['Литры'].isna().sum()
    if bad_litry:
        print(f"   ⚠️ {bad_litry} строк с нечисловыми литрами — удаляю")
    df_all = df_all.dropna(subset=['Литры'])
    print(f"   Осталось строк: {len(df_all)}")

    if len(df_all) == 0:
        print("\n❌ После очистки литров не осталось данных.")
        return

    # ---------- 6. СУММА → ЧИСЛО ----------
    if 'Сумма' in df_all.columns:
        print("\n6. Привожу суммы к числам…")
        df_all['Сумма'] = clean_numeric(df_all['Сумма'])
        bad_summy = df_all['Сумма'].isna().sum()
        if bad_summy:
            print(f"   ⚠️ {bad_summy} строк с нечисловой суммой — заменяю на 0")
        df_all['Сумма'] = df_all['Сумма'].fillna(0)

    # ---------- 7. ДАТУ → DATETIME ----------
    print("\n7. Обрабатываю даты…")
    df_all['Дата и время'] = pd.to_datetime(
        df_all['Дата и время'],
        dayfirst=True,
        errors='coerce'
    )
    df_all = df_all.dropna(subset=['Дата и время'])

    if len(df_all) > 0:
        print(f"   Период: с {df_all['Дата и время'].min().date()} "
              f"по {df_all['Дата и время'].max().date()}")

    df_all['Дата'] = df_all['Дата и время'].dt.date
    df_all['День недели'] = df_all['Дата и время'].dt.day_name()

    # ---------- 8. СВОДКА ПО ВОДИТЕЛЯМ ----------
    print("\n8. Сводка по водителям…")

    svodka_voditeli = df_all.groupby('Водитель').agg(
        Всего_литров=('Литры', 'sum'),
        Количество_заправок=('Литры', 'count'),
        Средний_объём=('Литры', 'mean'),
    ).reset_index()

    if 'Сумма' in df_all.columns:
        summy = df_all.groupby('Водитель')['Сумма'].sum().reset_index()
        summy.columns = ['Водитель', 'Всего_сумма']
        svodka_voditeli = svodka_voditeli.merge(summy, on='Водитель', how='left')

    svodka_voditeli = svodka_voditeli.sort_values('Всего_литров', ascending=False)

    print("\n" + "-" * 70)
    print(f"{'Водитель':<30} {'Литров':>10} {'Заправок':>10} {'Средний':>10}")
    print("-" * 70)
    for _, row in svodka_voditeli.iterrows():
        print(f"{str(row['Водитель'])[:28]:<30} {row['Всего_литров']:>10.1f} "
              f"{row['Количество_заправок']:>10} {row['Средний_объём']:>10.1f}")
    print("-" * 70)
    print(f"{'ИТОГО':<30} {svodka_voditeli['Всего_литров'].sum():>10.1f} "
          f"{svodka_voditeli['Количество_заправок'].sum():>10}")

    # ---------- 9. ДИНАМИКА ПО ДНЯМ ----------
    print("\n9. Динамика по дням…")

    dinamika_dni = df_all.groupby('Дата').agg(
        Всего_литров=('Литры', 'sum'),
        Количество_заправок=('Литры', 'count'),
    ).reset_index()
    dinamika_dni['Дата'] = pd.to_datetime(dinamika_dni['Дата'])
    dinamika_dni = dinamika_dni.sort_values('Дата')

    # ---------- 10. EXCEL ----------
    print("\n10. Создаю Excel-файл…")

    wb = Workbook()

    # ЛИСТ 1: СВОДКА ПО ВОДИТЕЛЯМ
    ws1 = wb.active
    ws1.title = 'Сводка по водителям'

    columns1 = ['Водитель', 'Всего литров', 'Количество заправок', 'Средний объём']
    if 'Всего_сумма' in svodka_voditeli.columns:
        columns1.append('Всего сумма (₽)')
    style_header(ws1, 1, columns1)

    for i, (_, row) in enumerate(svodka_voditeli.iterrows(), start=2):
        ws1.cell(row=i, column=1, value=row['Водитель'])
        ws1.cell(row=i, column=2, value=round(row['Всего_литров'], 1))
        ws1.cell(row=i, column=3, value=int(row['Количество_заправок']))
        ws1.cell(row=i, column=4, value=round(row['Средний_объём'], 1))
        if 'Всего_сумма' in svodka_voditeli.columns:
            ws1.cell(row=i, column=5, value=round(row['Всего_сумма'], 0))
        for col_idx in range(1, len(columns1) + 1):
            style_cell(ws1.cell(row=i, column=col_idx),
                       'center' if col_idx > 1 else 'left')

    auto_width(ws1)
    ws1.freeze_panes = 'A2'

    # ЛИСТ 2: ДИНАМИКА ПО ДНЯМ
    ws2 = wb.create_sheet('Динамика по дням')
    columns2 = ['Дата', 'Всего литров', 'Количество заправок']
    style_header(ws2, 1, columns2)

    for i, (_, row) in enumerate(dinamika_dni.iterrows(), start=2):
        ws2.cell(row=i, column=1, value=row['Дата'].strftime('%d.%m.%Y'))
        ws2.cell(row=i, column=2, value=round(row['Всего_литров'], 1))
        ws2.cell(row=i, column=3, value=int(row['Количество_заправок']))
        for col_idx in range(1, 4):
            style_cell(ws2.cell(row=i, column=col_idx),
                       'center' if col_idx > 1 else 'left')

    auto_width(ws2)
    ws2.freeze_panes = 'A2'

    # ЛИСТ 3: ВСЕ ЗАПРАВКИ
    ws3 = wb.create_sheet('Все заправки')
    columns3 = ['Дата', 'Водитель', 'Литры']
    if 'Сумма' in df_all.columns:
        columns3.append('Сумма (₽)')
    columns3.append('Местоположение')
    style_header(ws3, 1, columns3)

    for i, (_, row) in enumerate(df_all.iterrows(), start=2):
        ws3.cell(row=i, column=1, value=row['Дата и время'].strftime('%d.%m.%Y %H:%M'))
        ws3.cell(row=i, column=2, value=row['Водитель'])
        ws3.cell(row=i, column=3, value=row['Литры'])
        col = 4
        if 'Сумма' in df_all.columns:
            ws3.cell(row=i, column=col, value=row['Сумма'])
            col += 1
        ws3.cell(row=i, column=col, value=row.get('Местоположение', ''))
        for col_idx in range(1, len(columns3) + 1):
            style_cell(ws3.cell(row=i, column=col_idx),
                       'center' if col_idx in (1, 3) else 'left')

    auto_width(ws3)
    ws3.freeze_panes = 'A2'

    # СОХРАНЯЕМ
    filename = f'Отчет_заправки_{datetime.now().strftime("%d.%m.%Y_%H-%M")}.xlsx'
    wb.save(filename)

    print("\n" + "=" * 70)
    print(f"🎉 ГОТОВО! Файл: {filename}")
    print("=" * 70)
    print(f"\nВсего заправок: {len(df_all)}")
    print(f"Всего литров: {df_all['Литры'].sum():.1f}")
    print(f"Водителей: {len(svodka_voditeli)}")
    print(f"Дней: {len(dinamika_dni)}")


if __name__ == '__main__':
    main()