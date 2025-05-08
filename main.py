from tkinter import *
from tkinter import filedialog, messagebox
import pandas as pd
from pandas import ExcelWriter
import sys
import os

# Configuración de paths
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

icon_path = os.path.join(base_path, "hoja-de-excel.ico")

# Clase para tooltips
class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.mostrar)
        self.widget.bind("<Leave>", self.ocultar)

    def mostrar(self, event=None):
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + 25
        
        self.tooltip_window = Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        
        label = Label(self.tooltip_window, text=self.text, background="#ffffe0", relief="solid", borderwidth=1, padx=5, pady=2)
        label.pack()

    def ocultar(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

ventana = Tk()
ventana.title("Excel Merger")
ventana.geometry("600x300")
ventana.resizable(False, False)
ventana.wm_iconbitmap(icon_path)

archivo1_path = StringVar()
archivo2_path = StringVar()
df_final = None

def cargar_archivo1():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if filename:
        archivo1_path.set(filename)

def cargar_archivo2():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if filename:
        archivo2_path.set(filename)

def eliminar_archivo1():
    archivo1_path.set("")

def eliminar_archivo2():
    archivo2_path.set("")

def mezclar_datos():
    global df_final
    try:
        if not archivo1_path.get() or not archivo2_path.get():
            messagebox.showerror("Error", "Debes cargar ambos archivos primero")
            return
        
        df_factura = pd.read_excel(
            archivo1_path.get(),
            dtype={'Fecha': str}
        )
        
        if df_factura['Fecha'].dtype == 'datetime64[ns]':
            df_factura['Fecha'] = df_factura['Fecha'].dt.strftime('%Y-%m-%d %H:%M:%S')
        
        df_datos = pd.read_excel(archivo2_path.get())
        
        if 'Nombre' not in df_datos.columns or 'CUIT' not in df_datos.columns:
            messagebox.showerror("Error", "El archivo de datos debe contener columnas 'Nombre' y 'CUIT'")
            return

        cuil_dict = df_datos.set_index('Nombre')['CUIT'].astype(str).to_dict()
        df_factura['CUIT'] = df_factura['Cliente'].map(cuil_dict).fillna('N/A')
        
        df_final = df_factura
        messagebox.showinfo("Éxito", "Datos mezclados correctamente!\nAhora puedes guardar el archivo")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")

def guardar_archivo():
    if df_final is None:
        messagebox.showerror("Error", "Primero debes mezclar los datos")
        return

    filename = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("Todos los archivos", "*.*")],
        title="Guardar archivo mezclado"
    )
    
    if filename:
        try:
            with ExcelWriter(filename, engine='openpyxl', datetime_format='yyyy-mm-dd hh:mm:ss') as writer:
                df_final.to_excel(writer, index=False)
            
            messagebox.showinfo("Éxito", f"Archivo guardado en:\n{filename}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo:\n{str(e)}")

# Interfaz gráfica
texto = Label(ventana, text="Seleccione los archivos a procesar:")
texto.config(pady=10)
texto.pack()

# Marco para archivo factura
frame_archivo1 = Frame(ventana)
frame_archivo1.pack(pady=5)

btn_archivo1 = Button(frame_archivo1, text="Cargar el archivo factura", command=cargar_archivo1)
btn_archivo1.config(pady=5, padx=10)
btn_archivo1.pack(side=LEFT, padx=5)

btn_eliminar1 = Button(frame_archivo1, text="X", command=eliminar_archivo1, bg="#ff4444", fg="white", width=3)
btn_eliminar1.pack(side=LEFT)
Tooltip(btn_eliminar1, "Quitar archivo")

lbl_archivo1 = Label(ventana, textvariable=archivo1_path)
lbl_archivo1.pack()

# Marco para archivo datos
frame_archivo2 = Frame(ventana)
frame_archivo2.pack(pady=5)

btn_archivo2 = Button(frame_archivo2, text="Cargar el archivo de datos", command=cargar_archivo2)
btn_archivo2.config(pady=5, padx=10)
btn_archivo2.pack(side=LEFT, padx=5)

btn_eliminar2 = Button(frame_archivo2, text="X", command=eliminar_archivo2, bg="#ff4444", fg="white", width=3)
btn_eliminar2.pack(side=LEFT)
Tooltip(btn_eliminar2, "Quitar archivo")

lbl_archivo2 = Label(ventana, textvariable=archivo2_path)
lbl_archivo2.pack()

# Botones principales
frame_botones = Frame(ventana)
frame_botones.pack(pady=10)

btn_merger = Button(frame_botones, text="Mezclar datos", command=mezclar_datos)
btn_merger.config(pady=5, padx=10, bg="#4CAF50", fg="white")
btn_merger.pack(side=LEFT, padx=5)

btn_guardar = Button(frame_botones, text="Guardar archivo", command=guardar_archivo)
btn_guardar.config(pady=5, padx=10, bg="#2196F3", fg="white")
btn_guardar.pack(side=LEFT, padx=5)

# Footer
footer_frame = Frame(ventana, bg="#f0f0f0")
footer_frame.pack(side=BOTTOM, fill=X, pady=10)

texto_developer = Label(footer_frame, text="Developed by Tobias Gallo", bg="#f0f0f0", fg="#666666")
texto_developer.pack(side=LEFT, padx=10)

texto_version = Label(footer_frame, text="v.0.2.0", bg="#f0f0f0", fg="#666666")
texto_version.pack(side=RIGHT, padx=10)

ventana.mainloop()
