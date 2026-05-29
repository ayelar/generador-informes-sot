# Generador de Informes de Laboratorio

Aplicación desarrollada en Python y Streamlit para generar informes automáticos en Word a partir de datos cargados en Excel.

## Funcionalidades

- Carga de archivos Excel (.xlsx)
- Carga de plantilla Word (.docx)
- Reemplazo automático de variables
- Limpieza automática de campos vacíos
- Generación automática del informe final

## Tecnologías utilizadas

- Python
- Streamlit
- openpyxl
- python-docx

## Estructura del proyecto
```text
Generador_informes/
│
├── app.py
├── excel_utils.py
├── diccionario.py
├── word_utils.py
├── requirements.txt
└── README.md
```

## Cómo ejecutar

Instalar dependencias:

```bash
pip install -r requirements.txt

streamlit run app.py 
