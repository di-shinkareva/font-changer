
#Функция для изменения параметров текста в документе 
def modify_docx(file_path):
     try :
        #Окрываем .docx файл
        doc=Document(file_path) # type: ignore
        # Проходим по каждому абзацу документа
        for paragraph in doc.paragraphs:
            # Настраиваем текст внутри абзаца
            for run in paragraph.runs:
                run.font.name = 'Times New Roman'  # Устанавливаем шрифт
                run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman') # type: ignore
                # Сохраняем изменения в новый файл
        output_path = f"modified_{os.path.basename(file_path)}" # type: ignore
        doc.save(output_path)
        print(f"Документ сохранён: {output_path}")

     except Exception as e:
        # Если возникает ошибка, выводим сообщение
        print(f"Ошибка при обработке файла {file_path}: {e}") # type: ignore
        # Ищем все файлы .docx в папке
        docx_files = [f for f in os.listdir(folder_path) if f.endswith('.docx')] # type: ignore

    # Если файлы .docx не найдены, выводим сообщение
        if not docx_files:
         print("В указанной папке нет файлов .docx.")
        return
     # Обрабатываем каждый файл
     for file_name in docx_files:
        file_path = os.path.join(folder_path, file_name)  # type: ignore # Получаем полный путь к файлу
        print(f"Обрабатывается файл: {file_name}")
        modify_docx(file_path)  # Вызываем функцию для обработки файла



=======
import os # Для работы с файловой системой
from docx import Document # type: ignore # Для работы с
документами .docx # type: ignore
from docx.shared import pt # type: ignore # Для задания
размера шрифта # type: ignore
from docx.oxml.ns import qn # type: ignore #Для работы с XML-
структурой документа # type: ignore
from docx.oxml import parse_xml # type: ignore # Для обработки
межстрочного интервала # type: ignore
from docx.oxml.ns import nsdecls # type: ignore # Для работы с
пространством имен XML # type: ignore # type: ignore

# Устанавливаем шрифт
run._element.rPr.rFonts.set(qn('W:eastAsia'), # type: ignore
'Times New Roman')
# Устанавливаем шрифт для языков East Asia
run.font.size = Pt(14) # type: ignore
#
Устанавливаем размер шрифта # type: ignore
# Настраиваем межстрочный интервал
p_pr = paragraph._element.get_or_add_pPr () # type: ignore
spacing = parse_xml(r'<w:spacing %s w:line="360" w:lineRule="auto"/>' % nsdecls ('w'))
p_pr. append (spacing)

# Основная функция, которая обрабатывает все файлы в указанной папке
def main():
# Просим пользователя ввести путь к папке folder_path = input( "Введите путь к папке с файлами .docx: ").strip()
# Преобразуем путь в абсолютный (для работыс относительными путями)
folder_path = os. path. abspath(folder_path) # type: ignore
# Проверяем, существует ли указанная папка
if not os. path.isdir(folder_path):
    print(f"Ошибка: Папка не найдена. Убедитесь, что путь указан правильно: {folder_path}")
return
# Вызываем функцию для обработки файла
# Точка входа в программу
if__name__== " __main__" 
main()
