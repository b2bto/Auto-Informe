import sys
import ctypes
import os
import tkinter as tk
import atexit

# =========================
# PALETA DE COLORES (CORPORATIVO)
# =========================
COLOR_FONDO = "#365E93"
COLOR_PANEL = "#34495E"
COLOR_TEXTO = "#FFFFFF"
COLOR_PRINCIPAL = "#2C3E50"
COLOR_NARANJA = "#e67e22"
COLOR_HOVER = "#fa9e4e"

# =========================
# EVITAR DOBLE INSTANCIA
# =========================
mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "NUSE-123-UNIQUE-MUTEX")
last_error = ctypes.windll.kernel32.GetLastError()

if last_error == 183:
    # =========================
    # TRAER VENTANA AL FRENTE
    # =========================
    user32 = ctypes.windll.user32

    # Buscar ventana por título
    hwnd = user32.FindWindowW(None, "Sistema de Reportes SOC")

    if hwnd:
        user32.ShowWindow(hwnd, 9)   # 9 = restaurar
        user32.SetForegroundWindow(hwnd)

    sys.exit()

def cerrar_mutex():
    try:
        ctypes.windll.kernel32.CloseHandle(mutex)
    except:
        pass

atexit.register(cerrar_mutex)

# =========================
# RUTA BASE (PYINSTALLER)
# =========================
base_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(base_dir)

from generar_informe5 import generar_informe

# =========================
# MENSAJE EMERGENTE
# =========================
def mostrar_mensaje_auto(texto, duracion=2000):
    popup = tk.Toplevel()
    popup.title("Proceso terminado")
    popup.geometry("320x110")
    popup.configure(bg=COLOR_PANEL)

    label = tk.Label(
        popup,
        text=texto,
        bg=COLOR_PANEL,
        fg=COLOR_TEXTO,
        font=("Segoe UI", 10)
    )
    label.pack(expand=True)

    popup.update_idletasks()
    x = (popup.winfo_screenwidth() // 2) - 160
    y = (popup.winfo_screenheight() // 2) - 55
    popup.geometry(f"+{x}+{y}")

    popup.after(duracion, popup.destroy)

# =========================
# LÓGICA
# =========================
def ejecutar_informe():

    turno = turno_var.get()

    if turno == "":
        mostrar_mensaje_auto("Seleccione un turno")
        return

    try:
        generar_informe(turno)
        mostrar_mensaje_auto("Informe generado correctamente")

    except Exception as e:
        mostrar_mensaje_auto(f"Error: {str(e)}")

# =========================
# VENTANA PRINCIPAL
# =========================
ventana = tk.Tk()
ventana.title("Sistema de Reportes SOC")
ventana.geometry("420x270")
ventana.configure(bg=COLOR_FONDO)

titulo = tk.Label(
    ventana,
    text="SISTEMA AUTOMATIZADO DE INFORMES",
    font=("Segoe UI", 12, "bold"),
    bg=COLOR_FONDO,
    fg=COLOR_TEXTO
)
titulo.pack(pady=12)

turno_var = tk.StringVar()

tk.Label(
    ventana,
    text="Seleccione turno",
    bg=COLOR_FONDO,
    fg=COLOR_TEXTO,
    font=("Segoe UI", 9)
).pack()

for texto, valor in [("Mañana","mañana"),("Tarde","tarde"),("Noche","noche")]:
    tk.Radiobutton(
        ventana,
        text=texto,
        variable=turno_var,
        value=valor,
        bg=COLOR_FONDO,
        fg=COLOR_TEXTO,
        selectcolor=COLOR_PANEL,
        activebackground=COLOR_FONDO,
        font=("Segoe UI", 9)
    ).pack(anchor="w", padx=40)

# =========================
# BOTÓN CORPORATIVO CON HOVER
# =========================
def on_enter(e):
    boton.config(bg=COLOR_HOVER)

def on_leave(e):
    boton.config(bg=COLOR_NARANJA)

boton = tk.Button(
    ventana,
    text="GENERAR INFORME",
    command=ejecutar_informe,
    width=22,
    height=2,
    bg=COLOR_NARANJA,
    fg="white",
    font=("Segoe UI", 10, "bold"),
    bd=0,
    relief="flat",
    activebackground=COLOR_HOVER,
    activeforeground="white",
    cursor="hand2"
)

boton.pack(pady=18)

boton.bind("<Enter>", on_enter)
boton.bind("<Leave>", on_leave)

ventana.mainloop()