#from openpyxl import load_workbook

# ======================================
# SECCIÓN 2 - Construcción del diccionario
# ======================================

from datetime import datetime

def construir_diccionario(wb, ws_dic):

    diccionario = {}

    # Asumimos encabezados en fila 1
    for row in ws_dic.iter_rows(min_row=2):
        clave = row[0].value          # Columna A
        ref_texto = row[2].value     # Columna C ("Hoja-Celda")

        # Salteos defensivos
        if clave is None or ref_texto is None:
            continue

        # Parseo de "Hoja-Celda"
        try:
            hoja_nombre, celda_ref = ref_texto.split("-")
        except ValueError:
            raise ValueError(f"Referencia mal formada: '{ref_texto}'")

        # Verificación de hoja
        if hoja_nombre not in wb.sheetnames:
            raise ValueError(f"La hoja '{hoja_nombre}' no existe (clave {clave})")

        ws_origen = wb[hoja_nombre]
        valor = ws_origen[celda_ref].value

        # Normalización del valor
        # --- Normalización ---
        if valor is None or str(valor).strip() == "":
            diccionario[clave] = "__BORRAR__"

        # Fechas
        elif isinstance(valor, datetime):
            diccionario[clave] = valor.strftime("%d/%m/%Y")

        # Números float → 1 decimal
        elif isinstance(valor, float):
            diccionario[clave] = f"{valor:.1f}"

        # Otros valores
        else:
            diccionario[clave] = str(valor).strip()


    print("\n--- Diccionario generado ---")
    for k, v in diccionario.items():

        print(f"{k}  -->  {v}")
    print("--- Fin del diccionario ---\n")

    elementos = []

    for i in range(1, 7):

        elem = str(
            diccionario.get(f"Elemento{i}", "")
        ).strip()

        if elem and elem != "__BORRAR__":
            elementos.append(elem)

    # Formato lindo
    if len(elementos) > 1:

        texto_elementos = (
            ", ".join(elementos[:-1])
            + " y "
            + elementos[-1]
        )

    elif len(elementos) == 1:

        texto_elementos = elementos[0]

    else:
        texto_elementos = ""

    # Nueva variable para Word
    diccionario["LISTA_ELEMENTOS"] = texto_elementos

    print("\nLISTA_ELEMENTOS:")
    print(diccionario["LISTA_ELEMENTOS"])

    return diccionario