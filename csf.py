import fitz
import re
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def getDatosFiscales(pdfPath): 
    pdf = fitz.open(pdfPath)
    text = ""
    
    for page in pdf:
        text += page.get_text()
        
    # Persona moral: busca razón social
    rzs = re.search(r'Denominación/Razón Social:\s*(.+)', text)
    if rzs:
        rzs = rzs.group(1).strip()
    else:
        nombres = re.search(r'Nombre \(s\):\s*(.+)', text)
        apellido1 = re.search(r'Primer Apellido:\s*(.+)', text)
        apellido2 = re.search(r'Segundo Apellido:\s*(.+)', text)
        if nombres and apellido1 and apellido2:
            rzs = f"{nombres.group(1).strip()} {apellido1.group(1).strip()} {apellido2.group(1).strip()}"
        else:
            rzs = "No se encontró la razón social"

    # RFC       
    rfc = re.search(r'RFC:\s*(.+)', text)
    rfc = rfc.group(1).strip() if rfc else "No se encontró el RFC"
    
    # Código Postal
    cpm = re.search(r'Código Postal:(\d{5})', text)
    cpm = cpm.group(1).strip() if cpm else "No se encontró el código postal"
    
    # Régimen Fiscal
    regimen = re.findall(r'Regímenes:\s*[\s\S]+?Régimen\s+(.+?)\s+\d{2}/\d{2}/\d{4}', text)
    regimen = regimen[-1].strip() if regimen else "No se encontró el régimen fiscal"

    return {
        "file": os.path.basename(pdfPath),
        "razon_social": rzs,
        "RFC": rfc,
        "codigo_postal": cpm,
        "regimen": regimen,
        "items": "",           
        "tickets": "",         
        "codigo_sucursal": ""  
    }

def getFolder():
    folder = filedialog.askdirectory()
    if not folder: 
        return
    datos = []
    for file in os.listdir(folder):
        if file.lower().endswith(".pdf"):
            full_path = os.path.join(folder, file)
            try:
                datos.append(getDatosFiscales(full_path))
            except Exception as e:
                messagebox.showerror("Error al procesar", f"{file}: {e}")
    if datos:
        df = pd.DataFrame(datos)
        output_file = os.path.join(folder, "datos_fiscales.xlsx")
        df.to_excel(output_file, index=False)
        messagebox.showinfo("Proceso completado", f"Archivo guardado en:\n{output_file}")
    else:
        messagebox.showwarning("Sin datos", "No se encontraron PDFs válidos.")

# UI con Tkinter
root = tk.Tk()
root.title("Extractor de Datos Fiscales SAT")
root.geometry("400x200")

label = tk.Label(root, text="Selecciona la carpeta con los PDFs", font=("Arial", 12))
label.pack(pady=20)

btn = tk.Button(root, text="Seleccionar Carpeta", command=getFolder, font=("Arial", 12))
btn.pack()

root.mainloop()
