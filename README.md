# 🧾 Gestor de Transferencias Bancarias (TBs)

Este script automatiza el emparejamiento de comprobantes de transferencias bancarias (TBs) con sus respectivos soportes, validando su correspondencia y estado ("ABONADO" o rechazado) a partir de un archivo Excel de control.

## 🚀 Funcionalidades

- Empareja soportes con los TBs correspondientes, validando por nombre y valor.
- Lee un archivo Excel para identificar la relación entre soportes y TBs.
- Clasifica el estado de cada soporte: `ABONADO` o `RECHAZADO`.
- Puede trabajar con soportes en formato PDF, PNG o JPG.
- Exporta los resultados a un archivo final (por ejemplo, CSV o Excel).

## 📁 Estructura de Archivos

- `/soportes/`: Carpeta con los archivos de soporte (comprobantes de abono).
- `/tbs/`: Carpeta con los archivos de TBs (transferencias).
- `datos.xlsx`: Archivo de control con columnas como:
  - `Nombre Propietario`
  - `Valor`
  - `Número de comprobante`
  - `Estado` (ABONADO / otro)

## 🛠️ Requisitos

- Python 3.x
- Bibliotecas:
  - `pandas`
  - `openpyxl`
  - `Pillow`
  - `pdf2image`
  - `pytesseract`
  - `opencv-python`

Puedes instalar los requerimientos con:

```bash
pip install -r requirements.txt
