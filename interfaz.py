import tkinter as tk
from tkinter import messagebox
from generar_informe5 import generar_informe
import os

os.chdir(os.path.dirname(os.path.abspath(__file__)))

def ejecutar_informe():

    turno = turno_var.get()

    if turno == "":
        messagebox.showwarning("Aviso","Seleccione un turno")
        return

    try:

        generar_informe(turno)

        messagebox.showinfo(
            "Proceso terminado",
            "Informe generado correctamente"
        )

    except Exception as e:

        messagebox.showerror("Error",str(e))


ventana = tk.Tk()
ventana.title("Automatizador de Informe SOC")
ventana.geometry("400x250")

titulo = tk.Label(
    ventana,
    text="AUTOMATIZADOR DE INFORME",
    font=("Arial",14,"bold")
)
titulo.pack(pady=10)

turno_var = tk.StringVar()

tk.Label(ventana,text="Seleccione turno").pack()

tk.Radiobutton(
    ventana,
    text="Mañana",
    variable=turno_var,
    value="mañana"
).pack()

tk.Radiobutton(
    ventana,
    text="Tarde",
    variable=turno_var,
    value="tarde"
).pack()

tk.Radiobutton(
    ventana,
    text="Noche",
    variable=turno_var,
    value="noche"
).pack()

boton = tk.Button(
    ventana,
    text="Generar Informe",
    command=ejecutar_informe,
    width=20,
    height=2
)

boton.pack(pady=20)

ventana.mainloop()