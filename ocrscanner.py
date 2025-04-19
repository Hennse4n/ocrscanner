import re
import xlsxwriter 
import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox as mb
from PIL import Image, ImageTk
import pytesseract 

def cargar_imagen():
    ruta_imagen = filedialog.askopenfilename()
    title = "Selecciona una imagen",
    filetypes=[("Archivos de imagen", "*.jpg *.jpeg *.png *.bmp *.gif")]

    if ruta_imagen:
        imagen = Image.open(ruta_imagen)

        max_width = 600
        max_height = 400
        if imagen.width > max_width or imagen.height > max_height:
            imagen.thumbnail((max_width, max_height))

        imagen_tk = ImageTk.PhotoImage(imagen)

        etiqueta_imagen.config(image=imagen_tk)
        etiqueta_imagen.image = imagen_tk

        try:
            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            imagen  = ruta_imagen
            image = Image.open(imagen)
            preprocesar_imagen(ruta_imagen)
            texto_extraido = pytesseract.image_to_string(image, lang="spa")
            
            if texto_extraido=="":
                mb.showinfo("Atención", "La imagen no contiene texto")
            else:
                cuadro_texto.delete(1.0, tk.END)
                cuadro_texto.insert(tk.END, texto_extraido)

        except Exception as e: 
            mb.showerror("Error", "Error al escanear la imagen")

def crear_excel(empresa, fecha_emision, serie, numero, nit, total):
    try:
        # Abrir diálogo para guardar archivo
        ruta_archivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Guardar archivo Excel como"
        )

        if not ruta_archivo:
            mb.showinfo("Cancelado", "No se guardó el archivo.")
            return

        # Crear el archivo Excel
        workbook = xlsxwriter.Workbook(ruta_archivo)
        worksheet = workbook.add_worksheet("Datos Factura")

        headers = ["Empresa", "Fecha de emisión", "Serie", "Número", "NIT del comprador", "Total pagado"]
        datos = [ empresa, fecha_emision, serie, numero, nit, total]

        # Escribir encabezados
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        # Escribir datos
        for col, dato in enumerate(datos):
            worksheet.write(1, col, dato)

        workbook.close()
        mb.showinfo("Éxito", f"Archivo guardado en:\n{ruta_archivo}")

    except Exception as e:
        mb.showerror("Error", f"No se pudo guardar en Excel: {e}")

def extraer_datos(texto):
    texto = texto.strip()

    texto = texto.strip()

    # Buscar el nombre de la empresa (línea después de "NIT:")
    empresa = re.search(r'(?i)NIT[:\-]?\s*\d+\s*\n(.+)', texto)

    # Fecha de emisión
    fecha = re.search(r'(?i)fecha\s*de\s*emisi[oó]n[:\-]?\s*(\d{2}/\d{2}/\d{4})', texto)

    # Serie
    serie = re.search(r'(?i)serie\s*[:\-]?\s*([A-Z0-9]+)', texto)

    # Número
    numero = re.search(r'(?i)n[uú]mero\s*[:\-]?\s*(\d+)', texto)

    # NIT del comprador (el segundo NIT que aparece)
    nits = re.findall(r'(?i)NIT[:\-]?\s*([A-Z0-9]+)', texto)
    nit_comprador = nits[1] if len(nits) > 1 else "No encontrado"

    # Total pagado
    total = re.search(r'(?i)TOTAL\s+(\d+\.\d{2})', texto)

    return {
        "empresa": empresa.group(1).strip() if empresa else "No encontrado",
        "fecha": fecha.group(1).strip() if fecha else "No encontrada",
        "serie": serie.group(1).strip() if serie else "No encontrada",
        "numero": numero.group(1).strip() if numero else "No encontrado",
        "nit": nit_comprador,
        "total": total.group(1).strip() if total else "No encontrado"
    }

def extraer_desde_cuadro_texto():
    texto = cuadro_texto.get("1.0", tk.END)
    datos = extraer_datos(texto)

    mensaje = (
        f"Empresa: {datos['empresa']}\n"
        f"Fecha de emisión: {datos['fecha']}\n"
        f"Serie: {datos['serie']}\n"
        f"Número: {datos['numero']}\n"
        f"NIT del comprador: {datos['nit']}\n"
        f"Total pagado: Q{datos['total']}"
    )
    if mensaje=="":
        mb.showinfo("Aviso", "no hay datos")
    else:
        mb.showinfo("Datos detectados", mensaje)

        crear_excel(
            empresa=datos["empresa"],
            fecha_emision=datos["fecha"],
            serie=datos["serie"],
            numero=datos["numero"],
            nit=datos["nit"],
            total=datos["total"]
        )

def preprocesar_imagen(ruta):
    # Cargar imagen en escala de grises
     imagen = cv2.imread(ruta, cv2.IMREAD_GRAYSCALE)
    # Aplicar umbral para convertir en blanco y negro
     _, imagen_bn = cv2.threshold(imagen, 150, 255, cv2.THRESH_BINARY)
     return Image.fromarray(imagen_bn)

wroot = tk.Tk()
wroot.geometry("1200x1300+0+20")
wroot.title("OCR Scanner")

texto = tk.Label(
    wroot, 
    text="Pulsa en <<Extraer texto>> para escanear tu imagen",
    padx=100,
    font=("Arial", 18, "bold"),
    anchor="center",
    fg="blue"
)

texto2 = tk.Label(
    wroot,
    text="Pulsa en <<generar excel>> para guardar el texto",
    padx=100,
    font=("Arial", 18, "bold"),
    anchor="center",
    fg="red"
)

texto3 = tk.Label(
    wroot,
    text="Bienvenido a OCR Scanner!",
    padx=100,
    font=("Century Gothic", 24, "bold"),
    anchor="center",
)

texto3.pack()
texto.pack() 
texto2.pack()

boton = tk.Button(wroot, text="Extraer texto", command=cargar_imagen)
boton.pack(pady=10, padx=100, anchor="nw")

boton2 = tk.Button(wroot, text="Crear Excel", command=extraer_desde_cuadro_texto)
boton2.pack(pady=10, padx=100, anchor="ne")

cuadro_texto = tk.Text(wroot, wrap=tk.WORD, height=10, width=40)
cuadro_texto.pack(pady=10, padx=10, anchor="center")
cuadro_texto.insert(tk.END, "Aquí aparecerá tu texto")

etiqueta_imagen = tk.Label(wroot, text="Aqui veras tu imagen", font=("Arial", 11, "bold"))
etiqueta_imagen.pack(padx=100, anchor="nw")

wroot.mainloop()
    

