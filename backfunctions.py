import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re 
import unicodedata
import zipfile
from openpyxl.utils.exceptions import InvalidFileException
import calendar

def UploaderAxaDependents(uploaded_file):
    dependientes_AXA, base_manual_AXA = None, None  
    
    if uploaded_file is not None:
        try:
            xls = pd.ExcelFile(uploaded_file, engine="openpyxl")

            # Validamos existencia de hojas
            if "BASE CENTRAL SSFF" in xls.sheet_names:
                dependientes_AXA = pd.read_excel(xls, sheet_name="BASE CENTRAL SSFF")
            else:
                st.warning("No se encontró la hoja 'BASE CENTRAL SSFF'.")

            if "BASES MANUALES" in xls.sheet_names:
                base_manual_AXA = pd.read_excel(xls, sheet_name="BASES MANUALES")
            else:
                st.warning("No se encontró la hoja 'BASES MANUALES'.")

        except (zipfile.BadZipFile, InvalidFileException):
            st.error("❌ El archivo no es un Excel válido, por favor sube un archivo tipo .xlsx o .xls.")
        except Exception as e:
            st.error(f"❌ Error leyendo AXA: {e}")

    return dependientes_AXA, base_manual_AXA
        

def UploaderHCDependents(uploaded_file):
    month_from_filename = 'Archivo'
    if uploaded_file is not None:
        try:
            dependientes_HC = pd.read_excel(uploaded_file, engine="openpyxl")

            file_name = uploaded_file.name.upper().strip()
            file_name = file_name.replace(".XLSX", "").replace(".XLS", "")
            name_parts = file_name.split()

            spanish_to_english = {
                "ENERO": "January", "FEBRERO": "February", "MARZO": "March",
                "ABRIL": "April", "MAYO": "May", "JUNIO": "June",
                "JULIO": "July", "AGOSTO": "August", "SEPTIEMBRE": "September",
                "OCTUBRE": "October", "NOVIEMBRE": "November", "DICIEMBRE": "December"
            }

            for part in name_parts:
                if part in spanish_to_english:
                    month_from_filename = spanish_to_english[part]
                    break
                else:
                    found_month = False
                    for month in calendar.month_name[1:]:
                        if month.upper() in part:
                            month_from_filename = month
                            found_month = True
                            break
                    if found_month:
                        break

            return dependientes_HC, month_from_filename

        except (zipfile.BadZipFile, InvalidFileException):
            st.error("❌ El archivo no es un Excel válido, por favor sube un archivo tipo .xlsx o .xls.")
            return None, month_from_filename

        except Exception as e:
            st.error(f"❌ Error leyendo HC: {e}")
            return None, month_from_filename

    return None, month_from_filename
       

def normalize_nombre(nombre):
    nombre = nombre.replace("?", "N").replace("Ñ", "N")
    nombre = unicodedata.normalize("NFD", nombre)
    nombre = nombre.encode("ascii", "ignore").decode("utf-8")
    nombre = re.sub(r"[.-]", " ", nombre)
    nombre = re.sub(r"\s+", " ", nombre).strip()
    return nombre

def ProcessDependents_Generate_excel(dependientes_AXA, dependientes_HC, base_manual):
    # 1. Limpiar nombres de columnas
    dependientes_AXA.columns = dependientes_AXA.columns.str.strip().str.upper()
    dependientes_HC.columns = dependientes_HC.columns.str.strip().str.upper()

    dependientes_AXA_columns_text = ["NOMBRE DEL ASEGURADO", "APELLIDO PATERNO DEL ASEGURADO", "APELLIDO MATERNO DEL ASEGURADO"]
    dependientes_HC_columns_text = ["NOMBRE", "AP_PATERNO", "AP_MATERNO"]

    for col in dependientes_AXA_columns_text:
        dependientes_AXA[col] = dependientes_AXA[col].fillna("").astype(str).str.strip().str.upper()

    for col in dependientes_HC_columns_text:
        dependientes_HC[col] = dependientes_HC[col].fillna("").astype(str).str.strip().str.upper()

    dependientes_AXA["NUMERO DE CERTIFICADO"] = dependientes_AXA["NUMERO DE CERTIFICADO"].astype(str).str.strip().str.upper()
    dependientes_HC["NOEMPLEADO"] = dependientes_HC["NOEMPLEADO"].astype(str).str.strip().str.upper()

    # 2. Crear columna NOMBRE COMPLETO
    dependientes_AXA["NOMBRE_COMPLETO"] = (
        dependientes_AXA["NOMBRE DEL ASEGURADO"] + " " +
        dependientes_AXA["APELLIDO PATERNO DEL ASEGURADO"] + " " +
        dependientes_AXA["APELLIDO MATERNO DEL ASEGURADO"]
    ).apply(normalize_nombre)

    dependientes_HC["NOMBRE_COMPLETO"] = (
        dependientes_HC["NOMBRE"] + " " +
        dependientes_HC["AP_PATERNO"] + " " +
        dependientes_HC["AP_MATERNO"]
    ).apply(normalize_nombre)

    # 3. Reordenar columnas
    axa_columns_order = ["NUMERO DE POLIZA","NOMBRE DEL ASEGURADO","APELLIDO PATERNO DEL ASEGURADO","APELLIDO MATERNO DEL ASEGURADO","NOMBRE_COMPLETO", 
                         "EDAD","FECHA DE ALTA","FECHA DE BAJA","PARENTESCO","ESTATUS DEL ASEGURADO",
                         "NUMERO DEL SUBGRUPO", "NUMERO DE CERTIFICADO","FECHA DE ANTIGUEDAD","FECHA DE NACIMIENTO","SEXO", "ALIAS"]

    hc_columns_order = ["EMPRESA","NOEMPLEADO","IGPAREN","IGSEXO","IGFALT","CALCULA EDAD",
                        "NOMBRE","AP_PATERNO", "AP_MATERNO","NOMBRE_COMPLETO",
                        "RFC_CLI","DIRECCION_CLI", "COLONIA_CLI","CP_CLI","ESTADO_CLI","DELMUN_CLI",
                        "CIUDAD_CLI","EMAIL_CLI", "TELEMP1_CLI","ESTUDIANTE","DEP_ECONOMICO","COHABITAEMP",
                        "DIVISION", "DESCRIPCION DIVISION","SUB-DIVISION","DESCRIPCION SUBDIV.",
                        "TIPO POLIZA","AREA DE NOMINA","DESCRIPCION AREA NOM."]

    dependientes_AXA = dependientes_AXA[axa_columns_order]
    dependientes_HC = dependientes_HC[hc_columns_order]

    # 4. Comparaciones y dismatchs
    HC_ids = set(dependientes_HC["NOEMPLEADO"])
    HC_names = set(dependientes_HC["NOMBRE_COMPLETO"])
    filtro_id_axa = ~dependientes_AXA["NUMERO DE CERTIFICADO"].isin(HC_ids)
    filtro_name_axa = ~dependientes_AXA["NOMBRE_COMPLETO"].isin(HC_names)
    candidatos_dismatch_axa = dependientes_AXA[filtro_id_axa | filtro_name_axa].copy()

    def clasificar_fino_axa_t(row):
        id_match = row["NUMERO DE CERTIFICADO"] in HC_ids
        name_match = row["NOMBRE_COMPLETO"] in HC_names
        if not id_match and not name_match:
            return "Revisar ID y el nombre en el HC"
        elif not id_match:
            return "Revisar el ID en el HC"
        elif not name_match:
            return "Revisar el nombre en el HC"
        else:
            return None

    candidatos_dismatch_axa["Tipo_Disparidad"] = candidatos_dismatch_axa.apply(clasificar_fino_axa_t, axis=1)
    dismatch_AXA = candidatos_dismatch_axa[candidatos_dismatch_axa["Tipo_Disparidad"].notna()].copy()

    AXA_ids = set(dependientes_AXA["NUMERO DE CERTIFICADO"])
    AXA_names = set(dependientes_AXA["NOMBRE_COMPLETO"])
    filtro_id_hc = ~dependientes_HC["NOEMPLEADO"].isin(AXA_ids)
    filtro_name_hc = ~dependientes_HC["NOMBRE_COMPLETO"].isin(AXA_names)
    candidatos_dismatch_hc = dependientes_HC[filtro_id_hc | filtro_name_hc].copy()

    ids_CIGNA = {"1747946", "1748045", "1748045"}

    def clasificar_fino_hc(row):
        id_match = row["NOEMPLEADO"] in AXA_ids
        name_match = row["NOMBRE_COMPLETO"] in AXA_names
        if row["NOEMPLEADO"] in ids_CIGNA:
            return "Dados de alta en CIGNA"
        if not id_match and not name_match:
            return "Revisar el ID y el nombre en AXA"
        elif not id_match:
            return "Revisar el ID en el AXA"
        elif not name_match:
            return "Revisar el nombre en AXA"
        else:
            return None

    candidatos_dismatch_hc["Tipo_Disparidad"] = candidatos_dismatch_hc.apply(clasificar_fino_hc, axis=1)
    dismatch_HC = candidatos_dismatch_hc[candidatos_dismatch_hc["Tipo_Disparidad"].notna()].copy()

    # 5. Crear archivo Excel en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        base_manual.to_excel(writer,sheet_name = "Base_Manual", index = False)
        dismatch_HC.to_excel(writer, sheet_name="Diferencias_en_HC", index=False)
        dismatch_AXA.to_excel(writer, sheet_name="Diferencias_en_AXA", index=False)
        # dependientes_AXA.to_excel(writer, sheet_name="Base_Original_AXA", index=False)
        # dependientes_HC.to_excel(writer, sheet_name="Base_Original_HC", index=False)
    output.seek(0)

    return output