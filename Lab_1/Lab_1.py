import win32com.client
import sys
import time
import random

def search_replace_all(word_file, find_str, replace_str):
    ''' replace all occurrences of `find_str` w/ `replace_str` in `word_file` '''
    wdFindContinue = 1
    wdReplaceAll = 2

    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = 0
    app.DisplayAlerts = 0
    app.Documents.Open(word_file)

    app.Selection.Find.Execute(find_str, False, False, False, False, False, \
        True, wdFindContinue, False, replace_str, wdReplaceAll)
    app.ActiveDocument.Close(SaveChanges=True)
    app.Quit()

f = 'C:\\Users\\olyfs\\github\\interfaces\\Lab_1\\Исходник.docx'
search_replace_all(f, '#Family', input("Введите фамилию: "))
search_replace_all(f, '#Name', input("Введите имя: "))
search_replace_all(f, '#Patronymic', input("Введите отчество: "))
search_replace_all(f, '#Date', input("Введите дату рождения: "))

print("Готово. Данные успешно изменены.")
