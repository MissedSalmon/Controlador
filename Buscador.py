import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# Variable global para almacenar el DataFrame
df = None

def cargar_excel():
    global df
    archivo_path = filedialog.askopenfilename(
        title="Selecciona un archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if archivo_path:
        try:
            df = pd.read_excel(archivo_path)
            # Reemplazar NaN con una cadena vacía y luego convertir a cadena
            df['Nro Doc'] = df['Nro Doc'].fillna('').apply(lambda x: str(int(x)) if isinstance(x, (float, int)) else str(x))
            print("Columnas del archivo:", df.columns)
            print(df.head(10))  # Imprime las primeras 10 filas para ver el contenido
            messagebox.showinfo("Éxito", "Archivo cargado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo: {e}")

def buscar():
    global df
    if df is None:
        messagebox.showerror("Error", "Primero debes cargar un archivo Excel.")
        return
    
    nro_doc = entrada_nro_doc.get()
    if not nro_doc:
        messagebox.showwarning("Advertencia", "Por favor, ingresa un número de documento.")
        return
    
    # Filtrar las filas que coincidan con el número de documento
    resultados = df[df['Nro Doc'] == nro_doc][['Nro Recibo', 'Nombre Originante del Ingreso', 'Descripción Recibo', 'Fecha DGA', 'Monto']]
    
    # Limpiar el área de resultados antes de mostrar los nuevos resultados
    area_resultados.delete(1.0, tk.END)
    
    if not resultados.empty:
        for _, fila in resultados.iterrows():
            resultado = f"Nro Recibo: {fila['Nro Recibo']}\n"
            resultado += f"Nombre: {fila['Nombre Originante del Ingreso']}\n"
            resultado += f"Descripción: {fila['Descripción Recibo']}\n"
            resultado += f"Fecha DGA: {fila['Fecha DGA']}\n"
            resultado += f"Monto: {fila['Monto']}\n"
            resultado += "-" * 40 + "\n"
            area_resultados.insert(tk.END, resultado)
    else:
        area_resultados.insert(tk.END, "No se encontraron coincidencias.")

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Buscador de Recibos en Excel")

# Botón para cargar el archivo Excel
boton_cargar = tk.Button(ventana, text="Cargar Excel", command=cargar_excel)
boton_cargar.pack(pady=10)

# Entrada para el número de documento
label_nro_doc = tk.Label(ventana, text="Número de Documento:")
label_nro_doc.pack()
entrada_nro_doc = tk.Entry(ventana)
entrada_nro_doc.pack(pady=5)

# Botón para buscar
boton_buscar = tk.Button(ventana, text="Buscar", command=buscar)
boton_buscar.pack(pady=10)

# Área de resultados
area_resultados = tk.Text(ventana, width=80, height=20)
area_resultados.pack(pady=10)

# Ejecutar la aplicación
ventana.mainloop()
