import telebot
import os
import openpyxl
import sqlite3


token = "5115667742:AAGZ8QiWAuRDMXLI2lwvnsTgXAruGYo-Wqg"

bot = telebot.TeleBot(token)


@bot.message_handler(content_types=['document'])
def test(message):

    """Saving file"""
    save_dir = "."
    file_name = message.document.file_name
    file_id_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_id_info.file_path)
    src = file_name
    with open(save_dir + "/" + src, 'wb') as new_file:
        new_file.write(downloaded_file)

    """Reading file"""
    wb = openpyxl.load_workbook(filename=file_name)
    sheet = wb[wb.get_sheet_names()[0]]
    data = {}
    increment = 0
    db = sqlite3.connect("database.db")
    cur = db.cursor()
    for row in range(1, sheet.max_row + 1):
        tmp = sheet.cell(row=row, column=1).value
        if tmp is None or not ("1" <= str(tmp)[0] <= "9"):
            continue
        increment += 1
        data[increment] = {}
        data[increment]["group"] = str(sheet.cell(row=row, column=6).value)
        data[increment]["practice_name"] = str(sheet.cell(row=row, column=7).value)
        data[increment]["begin"] = str(sheet.cell(row=row, column=8).value.date().strftime("%d.%m.%Y"))
        data[increment]["end"] = str(sheet.cell(row=row, column=9).value.strftime("%d.%m.%Y"))
        cur.execute(
            """select * from orders where \"group\"=:group and practice_name=:practice_name and begin=:begin and end=:end""",
            data[increment])
        if not cur.fetchall():
            cur.execute("""insert into orders(\"group\", practice_name, begin, end, status) values (?, ?, ?, ?, ?)""",
                        (data[increment]["group"],
                         data[increment]["practice_name"],
                         data[increment]["begin"],
                         data[increment]["end"],
                         0))
            db.commit()

    db.close()
    path = os.path.join(os.path.abspath(os.path.dirname(__file__)), file_name)
    os.remove(path)


bot.infinity_polling()
