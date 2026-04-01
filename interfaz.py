import tkinter as tk
from generar_informe5 import generar_informe
import os

os.chdir(os.path.dirname(os.path.abspath(__file__)))

# 🎨 COLORES CORPORATIVOS
AZUL = "#125BC0"
AZUL_OSCURO = "#0C4292"
BLANCO = "#FFFFFF"
NARANJA = "#FF6A00"
GRIS = "#BDC3C7"

def mostrar_mensaje_auto(texto, duracion=2000):
    popup = tk.Toplevel()
    popup.title("Proceso")
    popup.geometry("320x120")
    popup.configure(bg=AZUL_OSCURO)

    label = tk.Label(
        popup,
        text=texto,
        bg=AZUL_OSCURO,
        fg=BLANCO,
        font=("Segoe UI", 10)
    )
    label.pack(expand=True)

    popup.update_idletasks()
    x = (popup.winfo_screenwidth() // 2) - 160
    y = (popup.winfo_screenheight() // 2) - 60
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


# 🪟 VENTANA PRINCIPAL
ventana = tk.Tk()
ventana.title("Automatizador de Informe SOC")
ventana.geometry("420x270")
ventana.configure(bg=AZUL_OSCURO)

# 🏷️ TÍTULO
titulo = tk.Label(
    ventana,
    text="AUTOMATIZADOR DE INFORME",
    font=("Segoe UI", 14, "bold"),
    bg=AZUL_OSCURO,
    fg=BLANCO
)
titulo.pack(pady=15)

# 📌 SUBTÍTULO
subtitulo = tk.Label(
    ventana,
    text="Seleccione el turno",
    font=("Segoe UI", 10),
    bg=AZUL_OSCURO,
    fg=GRIS
)
subtitulo.pack()

turno_var = tk.StringVar()

# 🔘 RADIOBUTTONS
for texto, valor in [("Mañana", "mañana"), ("Tarde", "tarde"), ("Noche", "noche")]:
    tk.Radiobutton(
        ventana,
        text=texto,
        variable=turno_var,
        value=valor,
        bg=AZUL_OSCURO,
        fg=BLANCO,
        activebackground=AZUL_OSCURO,
        activeforeground=BLANCO,
        selectcolor=AZUL,
        font=("Segoe UI", 10)
    ).pack(anchor="w", padx=120, pady=2)

# 🎯 BOTÓN PRINCIPAL
def on_enter(e):
    boton.config(bg="#D68910")  # naranja más oscuro

def on_leave(e):
    boton.config(bg=NARANJA)

boton = tk.Button(
    ventana,
    text="Generar Informe",
    command=ejecutar_informe,
    width=22,
    height=2,
    bg=NARANJA,
    fg=BLANCO,
    activebackground="#D68910",
    activeforeground=BLANCO,
    font=("Segoe UI", 10, "bold"),
    bd=0,
    cursor="hand2"
)
boton.pack(pady=25)

boton.bind("<Enter>", on_enter)
boton.bind("<Leave>", on_leave)

ventana.mainloop()