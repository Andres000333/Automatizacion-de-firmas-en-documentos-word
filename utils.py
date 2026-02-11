import unicodedata
import re

def normalizar(texto):
    if not texto:
        return ""

    texto = texto.lower()

    texto = ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

    texto = texto.replace("\n", " ").replace("\t", " ")
    texto = re.sub(r'\s+', ' ', texto)

    return texto.strip()


def contiene_nombre(nombre, texto):
    nombre_n = normalizar(nombre)
    texto_n = normalizar(texto)
    return nombre_n in texto_n
