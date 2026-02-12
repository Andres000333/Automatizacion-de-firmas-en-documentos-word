import os
import tempfile
from pdf2docx import Converter
from docx2pdf import convert
from firmador import firmar_documento_tablas

def procesar_pdf(ruta_pdf, firmas, ruta_salida, fecha_aprobacion=None):
    """
    Flujo CORRECTO:
    PDF → Word → FIRMAR → PDF
    """

    carpeta_base = os.path.dirname(ruta_salida)

    carpeta_word = os.path.join(carpeta_base, "WORD")
    carpeta_pdf = os.path.join(carpeta_base, "PDF")

    os.makedirs(carpeta_word, exist_ok=True)
    os.makedirs(carpeta_pdf, exist_ok=True)

    nombre_base = os.path.splitext(os.path.basename(ruta_pdf))[0]

    # ===== RUTAS =====
    word_convertido = os.path.join(carpeta_word, f"{nombre_base}_ORIGINAL.docx")
    word_firmado = os.path.join(carpeta_word, f"{nombre_base}.docx")
    pdf_final = os.path.join(carpeta_pdf, f"{nombre_base}.pdf")

    # ===== 1) PDF → WORD =====
    cv = Converter(ruta_pdf)
    cv.convert(word_convertido)
    cv.close()

    # ===== 2) FIRMAR WORD =====
    firmar_documento_tablas(
        word_convertido,
        firmas,
        word_firmado,
        fecha_aprobacion=fecha_aprobacion
    )

    # ===== 3) WORD → PDF =====
    convert(word_firmado, pdf_final)
