import os
from payconpy.fpdf.focr.orc import *
from payconpy.fselenium.fselenium import *

def ocr(file_path, path_tesseract):
    # file_path = os.path.abspath("Arquivo_Usuario/base/guia.pdf")
    # path_tesseract = os.path.abspath("Arquivo_Robo/bin/Tesseract-OCR/tesseract.exe")
    # path_tesseract = os.path.abspath(path_tesseract)
    if not os.path.exists(path_tesseract):
        while not os.path.exists(path_tesseract):
            faz_log('*** COLOQUE OS BINÁRIOS DO TESSERACT NA PASTA BIN (O NOME DA PASTA DOS BINÁRIOS DEVE SER "Tesseract-OCR") ***')
            sleep(10)
        else:
            pass
    text = ocr_tesseract_v2(
        file_path,
        limit_pages=5,
        path_tesseract="Arquivo_Robo/bin/Tesseract-OCR/tesseract.exe",
        path_pages="Arquivo_Robo/pages",
        tempdir="Arquivo_Robo/tempdir",
    )
    return text