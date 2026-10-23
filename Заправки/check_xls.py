# check_xls.py — смотрит, что на самом деле внутри .xls
# ЗАПУСК: python check_xls.py  (из корня проекта, где лежит папка «Заправки»)
import glob
import os

# Ищем файлы в подпапке «Заправки»
files = sorted(glob.glob(os.path.join('Заправки', 'transactional_report_*.xls')))

print(f"Найдено файлов: {len(files)}")
print()

if not files:
    print("❌ Файлы не найдены. Проверь, что папка «Заправки» есть рядом со скриптом.")
else:
    for f in files[:3]:  # возьмём первые 3 файла для примера
        print("=" * 70)
        print(f"Файл: {os.path.basename(f)}")
        print(f"Размер: {os.path.getsize(f)} байт")
        print("-" * 70)

        with open(f, 'rb') as fh:
            head = fh.read(300)

        # Первые 16 байт в hex — сигнатура формата
        print("Hex первых 16 байт:", head[:16].hex(' '))
        print()

        print("Первые 300 байт как текст (что видно):")
        print(repr(head[:300]))
        print()