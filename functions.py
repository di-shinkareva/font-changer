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




