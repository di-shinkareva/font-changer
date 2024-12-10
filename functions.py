import os  # Для работы с файловой системой
from docx import Document  # Импортируем Document для работы с .docx файлами
from docx.oxml import parse_xml  # Для обработки межстрочного интервала
from docx.oxml.ns import (
    nsdecls,  # Для работы с пространством имен XML
    qn,  # Для работы с XML-структурой документа
)
from docx.shared import Pt  # Для задания размера шрифта

# Функция для изменения параметров текста в документе
def modify_docx(file_path):
    try:
        # Открываем .docx файл
        doc = Document(file_path)
        # Проходим по каждому абзацу документа
        for paragraph in doc.paragraphs:
            # Настраиваем текст внутри абзаца
            for run in paragraph.runs:
                run.font.name = 'Times New Roman'  # Устанавливаем шрифт
                run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')  # Устанавливаем шрифт для языков East Asia
                run.font.size = Pt(14)  # Устанавливаем размер шрифта
            # Настраиваем межстрочный интервал
            p_pr = paragraph._element.get_or_add_pPr()
            spacing = parse_xml(r'<w:spacing %s w:line="360" w:lineRule="auto"/>' % nsdecls('w'))
            p_pr.append(spacing)
        # Сохраняем изменения в новый файл
        output_path = f"modified_{os.path.basename(file_path)}"
        doc.save(output_path)  # Сохраняем в output_path
        print(f"Документ сохранён: {output_path}")
    except Exception as e:
        # Если возникает ошибка, выводим сообщение
        print(f"Ошибка при обработке файла {file_path}: {e}")

# Основная функция, которая обрабатывает все файлы в указанной папке
def main():
    # Просим пользователя ввести путь к папке
    folder_path = input("Введите путь к папке с файлами .docx: ").strip()
    # Преобразуем путь в абсолютный (для работы с относительными путями)
    folder_path = os.path.abspath(folder_path)
    # Проверяем, существует ли указанная папка
    if not os.path.isdir(folder_path):
        print(f"Ошибка: Папка не найдена. Убедитесь, что путь указан правильно: {folder_path}")
        return
    # Ищем все файлы .docx в папке
    docx_files = [f for f in os.listdir(folder_path) if f.endswith('.docx')]
    # Если файлы .docx не найдены, выводим сообщение
    if not docx_files:
        print("В указанной папке нет файлов .docx.")
        return
    # Обрабатываем каждый файл
    for file_name in docx_files:
        file_path = os.path.join(folder_path, file_name)  # Получаем полный путь к файлу
        print(f"Обрабатывается файл: {file_name}")
        modify_docx(file_path)  # Вызываем функцию для обработки файла

# Точка входа в программу
if __name__ == "__main__":
    main()
