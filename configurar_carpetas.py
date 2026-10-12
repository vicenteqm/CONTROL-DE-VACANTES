import json
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

ROOT = Path(__file__).resolve().parent

def main():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo("Configurar INVEX", "Primero selecciona la carpeta donde guardas el Excel BASE DE TIENDAS INVEX.", parent=root)
    base = filedialog.askdirectory(title="Selecciona la carpeta BASE DE TIENDAS")
    if not base:
        messagebox.showwarning("Configuración cancelada", "No se seleccionó la carpeta de la base.", parent=root)
        root.destroy(); return
    attendance = filedialog.askdirectory(title="Selecciona la carpeta ASISTENCIA")
    if not attendance:
        messagebox.showwarning("Configuración cancelada", "No se seleccionó la carpeta de asistencia.", parent=root)
        root.destroy(); return
    config = {
        "base_folder": base,
        "base_pattern": "*BASE*TIENDAS*.xlsx",
        "attendance_folder": attendance,
        "attendance_pattern": "*ASISTENCIA*.xlsx",
        "auto_push": True
    }
    (ROOT / "config.json").write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    messagebox.showinfo("Configuración guardada", "Las carpetas quedaron configuradas.\n\nCuando quieras probar la actualización, ejecuta actualizar.bat.\n\nProcura que cada carpeta tenga el Excel vigente y que esté guardado antes de ejecutar.", parent=root)
    root.destroy()

if __name__ == "__main__":
    main()
