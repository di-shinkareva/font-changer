#Функция для изменения параметров текста в документе 
def modify_docx(file_path):
    try:
        #Окрываем .docx файл
        doc=Document(file_path) # type: ignore

