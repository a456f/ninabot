# estado_global.py
import json
import os

archivo_estado = "estado_excel.json"

def guardar_estado(estado_excel: str, ruta: str):
    with open(archivo_estado, "w") as f:
        json.dump({
            "estado_excel": estado_excel,
            "ruta": ruta
        }, f, ensure_ascii=False)

def cargar_estado():
    if os.path.exists(archivo_estado):
        with open(archivo_estado, "r") as f:
            data = json.load(f)
            return data.get("estado_excel", "📊 Archivo Excel: No cargado ❌"), data.get("ruta", "")
    return "📊 Archivo Excel: No cargado ❌", ""
