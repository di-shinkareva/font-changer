## functions модуль для импорта функций для работы

# TODO - import <зависимость из задания>


def change_font_and_spacing(doc_path):
    """Изменяет шрифт, размер шрифта и межстрочный интервал в документе."""
    doc = Document (doc_path)
    #размер шрифта в пунктах 
for paragraph in doc.paragraphs:  
    for run in doc.paragraphs:
        run.font.name = 'Times New Roman'
        run.font.size = 14
    pass


def func2(n):
    """Функция 2"""
    pass
