# Procesador de TBs y Soportes de Pago

Aplicación de escritorio en Python con PyQt5 que permite combinar archivos PDF de **Transferencias Bancarias (TBs)** con sus respectivos **soportes de pago**, utilizando información de un archivo Excel.

## ✨ Funcionalidades

- Interfaz gráfica amigable con PyQt5.
- Carga de archivos:
  - PDF de TBs
  - PDF con soportes de pago
  - Archivo Excel con datos de egresos
- División automática de soportes en tres regiones por página.
- Verificación del contenido con OCR (`get_text`) para asegurarse que los soportes contengan la palabra "ABONADO".
- Generación de PDFs individuales combinando TB + soporte, nombrados por número de egreso y beneficiario.

## 🛠️ Tecnologías usadas

- Python 3.x
- PyQt5
- fitz / PyMuPDF
- Pandas

## 📦 Instalación

1. Clona el repositorio:

```bash
git clone https://github.com/tuusuario/tu-repo.git
