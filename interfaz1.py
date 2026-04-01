import tkinter as tk
from generar_informe5 import generar_informe
import os

os.chdir(os.path.dirname(os.path.abspath(__file__)))

def mostrar_mensaje_auto(texto, duracion=2000):
    popup = tk.Toplevel()
    popup.title("Proceso terminado")
    popup.geometry("300x100")
    popup.configure(bg="#2c3e50")

    label = tk.Label(
        popup,
        text=texto,
        bg="#2c3e50",
        fg="white",
        font=("Arial",10)
    )
    label.pack(expand=True)

    popup.update_idletasks()
    x = (popup.winfo_screenwidth() // 2) - 150
    y = (popup.winfo_screenheight() // 2) - 50
    popup.geometry(f"+{x}+{y}")

    popup.after(duracion, popup.destroy)

def ejecutar_informe():

    turno = turno_var.get()

    if turno == "":
        mostrar_mensaje_auto("Seleccione un turno", 2000)
        return

    try:
        generar_informe(turno)
        mostrar_mensaje_auto("Informe generado correctamente", 2500)

    except Exception as e:
        mostrar_mensaje_auto(f"Error: {str(e)}", 3000)


ventana = tk.Tk()
ventana.title("Automatizador de Informe SOC")
ventana.geometry("400x250")
ventana.configure(bg="#2c3e50")

titulo = tk.Label(
    ventana,
    text="AUTOMATIZADOR DE INFORME",
    font=("Arial",14,"bold"),
    bg="#2c3e50",
    fg="white"
)
titulo.pack(pady=10)

turno_var = tk.StringVar()

tk.Label(
    ventana,
    text="Seleccione turno",
    bg="#2c3e50",
    fg="white"
).pack()

for texto, valor in [("Mañana","mañana"),("Tarde","tarde"),("Noche","noche")]:
    tk.Radiobutton(
        ventana,
        text=texto,
        variable=turno_var,
        value=valor,
        bg="#2c3e50",
        fg="white",
        selectcolor="#34495e"
    ).pack()

boton = tk.Button(
    ventana,
    text="Generar Informe",
    command=ejecutar_informe,
    width=20,
    height=2,
    bg="#27ae60",
    fg="white"
)
boton.pack(pady=20)

ventana.mainloop()