import tkinter as tk
from tkinter import filedialog, messagebox
from main import iniciar_proceso

def seleccionar_entrada():
    ruta = filedialog.askopenfilename(
        title="Seleccionar archivo de empresas",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )
    if ruta:
        entrada_var.set(ruta)

def seleccionar_salida():
    ruta = filedialog.asksaveasfilename(
        title="Guardar resultados",
        defaultextension=".xlsx",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )
    if ruta:
        salida_var.set(ruta)

def ejecutar():
    entrada = entrada_var.get()
    salida = salida_var.get()

    if not entrada:
        messagebox.showwarning("Falta input", "Debes seleccionar un archivo de entrada.")
        return
    if not salida:
        messagebox.showwarning("Falta output", "Debes seleccionar una ruta para guardar.")
        return

    try:
        iniciar_proceso(entrada, salida)
        messagebox.showinfo("Completado", "Proceso terminado correctamente.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("Scraper SUNAT")
root.geometry("500x300")

entrada_var = tk.StringVar()
salida_var = tk.StringVar()

tk.Label(root, text="Archivo de entrada:").pack(pady=5)
tk.Entry(root, textvariable=entrada_var, width=50).pack()
tk.Button(root, text="Seleccionar archivo", command=seleccionar_entrada).pack()

tk.Label(root, text="Archivo de salida:").pack(pady=10)
tk.Entry(root, textvariable=salida_var, width=50).pack()
tk.Button(root, text="Guardar como", command=seleccionar_salida).pack()

tk.Button(root, text="Iniciar Proceso", command=ejecutar, bg="green", fg="white").pack(pady=20)

root.mainloop()
