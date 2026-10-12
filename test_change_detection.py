"""Проверка без сети и без Telegram: ловит ли бот изменения в файле с тем же именем.

Запуск (из папки с schedule_bot.py):   python test_change_detection.py

Скрипт сам создаёт два xlsx, подменяет «Облако» заглушкой и прогоняет
sync_refresh() несколько раз. Меняется ТОЛЬКО последний урок — именно такую
мелкую правку раньше бот не замечал.
"""
import io
import sys

import openpyxl

import schedule_bot as sb


def make_xlsx(last_subject):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "1 смена"
    ws.append(["№ урока", "класс", "5а", "5б"])
    ws.append([1, "8.30-9.10", "русс 24", "мат 12"])
    ws.append([2, "9.20-10.00", "лит 24", "мат 12"])
    ws.append([3, "10.10-10.50", "русс 24", "ин 5"])
    ws.append([4, "11.00-11.40", None, "физ 7"])          # у 5а тут пропуск
    ws.append([5, "11.50-12.30", last_subject, "ин 5"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


MON = "Расписание 12 октября 2026.xlsx"   # понедельник
TUE = "Расписание 13 октября 2026.xlsx"   # вторник

cloud = {MON: make_xlsx("физ 7"), TUE: make_xlsx("физ 7")}
sb.list_folder = lambda: [(n, [n]) for n in cloud]     # подмена Облака
sb.download = lambda links: cloud[links[0]]

ok = True


def check(title, cond):
    global ok
    ok = ok and cond
    print(("✅" if cond else "❌"), title)


check("1. первый запуск: оба файла новые", sorted(sb.sync_refresh()) == sorted([MON, TUE]))
check("2. повторный запуск без правок: ничего не менялось", sb.sync_refresh() == [])

cloud[MON] = make_xlsx("хим 9")      # заменили файл с тем же именем
check("3. поменяли один урок в файле с тем же именем: бот заметил", sb.sync_refresh() == [MON])
check("4. сразу после этого снова тишина", sb.sync_refresh() == [])

mon_text = sb.screen_schedule("2026-10-12", "5", "а")[0]
tue_text = sb.screen_schedule("2026-10-13", "5", "а")[0]
check("5. в расписании понедельника новый урок (хим), а не старый (физ)",
      "хим 9" in mon_text and "физ 7" not in mon_text)
check("6. вторник не затронут", "физ 7" in tue_text)
check("7. в понедельник первой строкой «Разговоры о важном» 8:00 - 8:30",
      mon_text.splitlines()[1] == "0. 8:00 - 8:30 Разговоры о важном")
check("8. во вторник «Разговоров о важном» нет", "Разговоры о важном" not in tue_text)
check("9. пропуск показан как «4. -»", "4. -" in mon_text.splitlines())

print("\n--- понедельник, 5а ---\n" + mon_text)
sys.exit(0 if ok else 1)
