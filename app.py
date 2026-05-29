import streamlit as st
from excel_utils import cargar_excel
from diccionario import construir_diccionario
from word_utils import generar_word

st.title("Generador de Informes")

excel_file = st.file_uploader(
    "Subir Excel",
    type=["xlsx"]
)

word_file = st.file_uploader(
    "Subir plantilla Word",
    type=["docx"]
)

if st.button("Generar informe"):

    if excel_file is not None and word_file is not None:

        # ---------------------------------
        # GUARDAR ARCHIVOS TEMPORALES
        # ---------------------------------
        excel_path = "excel_temp.xlsx"
        word_path = "plantilla_temp.docx"

# Guardar Excel
        with open(excel_path, "wb") as f:
            f.write(excel_file.getbuffer())
            
# Guardar Word
        with open(word_path, "wb") as f:
            f.write(word_file.getbuffer())
        
        # ======================================
        # SECCIÓN 1 - Cargar el excel
        # ======================================
        wb, ws_dic = cargar_excel(excel_path)

        st.success("Archivos cargados correctamente")

        # ======================================
        # SECCIÓN 2 - Construcción del diccionario
        # ======================================

        diccionario = construir_diccionario(wb, ws_dic)

        # ======================================
        # SECCIÓN 3 - Word: verificación cruzada y reemplazo (HÍBRIDA)
        # - Body + tablas: método convencional (sin runs)
        # - Headers/footers: método robusto (runs)
        # - Insensible a tildes y mayúsculas
        # - Postproceso: "- ± -" --> "-"
        # ======================================
        
        output_path = generar_word(
            word_path,
            diccionario
        )
            
        st.success("Informe generado correctamente")

        # ======================================
        # DESCARGA
        # ======================================

        with open(output_path, "rb") as f:

            st.download_button(
                label="Descargar informe",
                data=f,
                file_name="Informe_generado.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    else:

            st.error("Tenés que subir ambos archivos")
                


