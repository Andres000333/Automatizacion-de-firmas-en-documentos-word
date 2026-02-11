import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
from firmador import firmar_documento_tablas, generar_preview

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Firmador automático Word")
app.geometry("620x700")


firmas = [
    {"nombre": "", "imagen": "", "tamano": 3.0},
    {"nombre": "", "imagen": "", "tamano": 3.0},
    {"nombre": "", "imagen": "", "tamano": 3.0},
    {"nombre": "", "imagen": "", "tamano": 3.0},
]

docs_seleccionados = []
carpeta_destino = ""


# ================= IMAGENES =================
def seleccionar_imagen(idx):
    ruta = filedialog.askopenfilename(
        filetypes=[("Imágenes", "*.png *.jpg *.jpeg")]
    )
    if ruta:
        firmas[idx]["imagen"] = ruta
        botones_img[idx].configure(text="Imagen ✔")


# ================= SELECCIONAR DOCUMENTOS =================
def seleccionar_docs():
    global docs_seleccionados
    docs_seleccionados = filedialog.askopenfilenames(
        filetypes=[("Word", "*.docx")]
    )
    lbl_docs.configure(text=f"{len(docs_seleccionados)} documentos seleccionados")


# ================= CARPETA =================
def seleccionar_carpeta():
    global carpeta_destino
    carpeta_destino = filedialog.askdirectory()
    lbl_carpeta.configure(text=carpeta_destino)


# ================= PREVIEW =================
def previsualizar():
    if not docs_seleccionados:
        messagebox.showerror("Error", "Selecciona documentos primero")
        return

    ok = generar_preview(docs_seleccionados[0], firmas)
    if not ok:
        messagebox.showerror("Error", "No se encontró ningún nombre.")


# ================= GENERAR MASIVO =================
def generar_todos():
    if not docs_seleccionados:
        messagebox.showerror("Error", "Selecciona documentos")
        return

    if not carpeta_destino:
        messagebox.showerror("Error", "Selecciona carpeta destino")
        return

    nombre_carpeta = entrada_carpeta.get()
    if not nombre_carpeta:
        messagebox.showerror("Error", "Escribe nombre de carpeta")
        return

    ruta_final = os.path.join(carpeta_destino, nombre_carpeta)
    os.makedirs(ruta_final, exist_ok=True)

    for ruta_doc in docs_seleccionados:
        nombre_archivo = os.path.basename(ruta_doc)
        ruta_salida = os.path.join(ruta_final, nombre_archivo)
        firmar_documento_tablas(ruta_doc, firmas, ruta_salida)

    messagebox.showinfo("Listo", "Todos los documentos fueron firmados ✔")


# ================= HEADER =================
header = ctk.CTkFrame(app, height=70, fg_color="#1f6aa5")
header.pack(fill="x")

ctk.CTkLabel(
    header,
    text="FIRMADOR AUTOMÁTICO WORD",
    font=("Arial", 20, "bold"),
    text_color="white"
).pack(pady=18)


# ================= SCROLL PRINCIPAL =================
scroll = ctk.CTkScrollableFrame(app)
scroll.pack(fill="both", expand=True, padx=15, pady=15)


# ================= FIRMANTES =================
ctk.CTkLabel(scroll, text="Firmantes", font=("Arial", 17, "bold")).pack(anchor="w", pady=(5,10))

botones_img = []

for i in range(4):
    fila = ctk.CTkFrame(scroll)
    fila.pack(fill="x", pady=8)

    entrada = ctk.CTkEntry(fila, placeholder_text=f"Nombre firmante {i+1}", height=34)
    entrada.pack(side="left", fill="x", expand=True, padx=6)

    entrada.bind(
        "<KeyRelease>",
        lambda e, idx=i, ent=entrada:
        firmas.__setitem__(idx, {**firmas[idx], "nombre": ent.get()})
    )

    btn = ctk.CTkButton(
        fila, text="Seleccionar firma", width=140,
        command=lambda idx=i: seleccionar_imagen(idx)
    )
    btn.pack(side="right", padx=5)

    botones_img.append(btn)

    # ===== SLIDER DE TAMAÑO =====
    slider_frame = ctk.CTkFrame(scroll)
    slider_frame.pack(fill="x", padx=10)

    ctk.CTkLabel(slider_frame, text="Tamaño firma (cm):").pack(side="left")

    slider = ctk.CTkSlider(
        slider_frame,
        from_=1.5,
        to=6.0,
        number_of_steps=45,
        width=200,
        command=lambda value, idx=i:
        firmas.__setitem__(
            idx,
            {**firmas[idx], "tamano": round(float(value), 2)}
        )
    )
    slider.set(3.0)
    slider.pack(side="left", padx=10)


# ================= DOCUMENTOS =================
ctk.CTkLabel(scroll, text="Documentos", font=("Arial", 17, "bold")).pack(anchor="w", pady=(20,10))

ctk.CTkButton(scroll, text="Seleccionar documentos Word", command=seleccionar_docs).pack(fill="x")
lbl_docs = ctk.CTkLabel(scroll, text="0 documentos seleccionados")
lbl_docs.pack(pady=(4,10))


# ================= DESTINO =================
ctk.CTkLabel(scroll, text="Carpeta destino", font=("Arial", 17, "bold")).pack(anchor="w", pady=(20,10))

ctk.CTkButton(scroll, text="Seleccionar carpeta destino", command=seleccionar_carpeta).pack(fill="x")
lbl_carpeta = ctk.CTkLabel(scroll, text="(carpeta no seleccionada)")
lbl_carpeta.pack(pady=(4,10))

entrada_carpeta = ctk.CTkEntry(scroll, placeholder_text="Nombre de carpeta de salida")
entrada_carpeta.pack(fill="x", pady=5)


# ================= BOTONES FINALES =================
ctk.CTkButton(
    scroll,
    text="PREVISUALIZAR (con firmas)",
    height=45,
    font=("Arial", 14, "bold"),
    command=previsualizar
).pack(fill="x", pady=(25,8))

ctk.CTkButton(
    scroll,
    text="FIRMAR TODOS LOS DOCUMENTOS",
    height=50,
    font=("Arial", 15, "bold"),
    fg_color="#2e8b57",
    hover_color="#256f46",
    command=generar_todos
).pack(fill="x", pady=(5,20))


app.mainloop()
