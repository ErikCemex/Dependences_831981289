import pandas as pd 
import numpy as np 
import unicodedata
import re

# base_axa = input("Escribe el nombre del documento de la base de datos de AXA")
# base_hc = input("Escribe el nombre del documento de la base de datos de HC")

# dependientes_AXA = pd.read_excel(base_axa + ".xlsx", engine= "openpyxl")
# dependientes_HC = pd.read_excel(base_hc + ".xlsx", engine= "openpyxl")

dependientes_AXA = pd.read_excel("BASE DE ASEGURADOS ACTIVOS AXA - POLIZAS CENTRAL_Junio.xlsx", engine= "openpyxl")
dependientes_HC = pd.read_excel("HC DEPENDIENTES JUNIO.xlsx", engine= "openpyxl")

dependientes_AXA.columns = dependientes_AXA.columns.str.strip().str.upper()
dependientes_HC.columns = dependientes_HC.columns.str.strip().str.upper()

dependientes_AXA_columns_text = ["NOMBRE", "APELLIDO PATERNO", "APELLIDO MATERNO"]
dependientes_HC_columns_text = ["NOMBRE", "AP_PATERNO", "AP_MATERNO"]


for col in dependientes_AXA_columns_text:
    dependientes_AXA[col] = dependientes_AXA[col].fillna("").astype(str).str.strip().str.upper()

for col in dependientes_HC_columns_text:
    dependientes_HC[col] = dependientes_HC[col].fillna("").astype(str).str.strip().str.upper()


dependientes_AXA["CERTIFICADO"] = dependientes_AXA["CERTIFICADO"].astype(str).str.strip().str.upper()
dependientes_HC["NOEMPLEADO"] = dependientes_HC["NOEMPLEADO"].astype(str).str.strip().str.upper()



dependientes_AXA["NOMBRE_COMPLETO"] = (
    dependientes_AXA["NOMBRE"].fillna("").str.strip() + " " +
    dependientes_AXA["APELLIDO PATERNO"].fillna("").str.strip() + " " +
    dependientes_AXA["APELLIDO MATERNO"].fillna("").str.strip()
)


dependientes_HC["NOMBRE_COMPLETO"] = (
    dependientes_HC["NOMBRE"].fillna("").str.strip() + " " +
    dependientes_HC["AP_PATERNO"].fillna("").str.strip() + " " +
    dependientes_HC["AP_MATERNO"].fillna("").str.strip()
)


def normalize_nombre(nombre):
    # Replace Ñ (manually before normalization)
    nombre = nombre.replace("?", "N").replace("Ñ", "N")
    
    # Remove accents (e.g., Ü -> U, Á -> A)
    nombre = unicodedata.normalize("NFD", nombre)
    nombre = nombre.encode("ascii", "ignore").decode("utf-8")
    
    # Replace . and - with space
    nombre = re.sub(r"[.-]", " ", nombre)
    
    # Replace multiple spaces with single space
    nombre = re.sub(r"\s+", " ", nombre).strip()
    
    return nombre

# Apply to NOMBRE COMPLETO
dependientes_AXA["NOMBRE_COMPLETO"] = dependientes_AXA["NOMBRE_COMPLETO"].apply(normalize_nombre)
dependientes_HC["NOMBRE_COMPLETO"]  = dependientes_HC["NOMBRE_COMPLETO"].apply(normalize_nombre)

axa_columns_order = ["NO.POLIZA","NOMBRE","APELLIDO PATERNO","APELLIDO MATERNO","NOMBRE_COMPLETO", 
                     "EDAD","FECHA DE ALTA","FECHA DE BAJA","PARENTESCO","ESTATUS",
                     "SUBGRUPO", "CERTIFICADO","FECHA DE ANTIGÃŒEDAD","FECHA DE NACIMIENTO","SEXO"]

hc_columns_order = ["EMPRESA","NOEMPLEADO","IGPAREN","IGSEXO","IGFALT","CALCULA EDAD",
                     "NOMBRE","AP_PATERNO", "AP_MATERNO","NOMBRE_COMPLETO",
                     "RFC_CLI","DIRECCION_CLI", "COLONIA_CLI","CP_CLI","ESTADO_CLI","DELMUN_CLI",
                     "CIUDAD_CLI","EMAIL_CLI", "TELEMP1_CLI","ESTUDIANTE","DEP_ECONOMICO","COHABITAEMP",
                     "DIVISION", "DESCRIPCION DIVISION","SUB-DIVISION","DESCRIPCION SUBDIV.",
                     "TIPO POLIZA","AREA DE NOMINA","DESCRIPCION AREA NOM."]

dependientes_AXA = dependientes_AXA[axa_columns_order]
dependientes_HC = dependientes_HC[hc_columns_order]


HC_ids = set(dependientes_HC["NOEMPLEADO"])
HC_names = set(dependientes_HC["NOMBRE_COMPLETO"])
filtro_id_axa = ~dependientes_AXA["CERTIFICADO"].isin(HC_ids)
filtro_name_axa = ~dependientes_AXA["NOMBRE_COMPLETO"].isin(HC_names)
candidatos_dismatch_axa = dependientes_AXA[filtro_id_axa | filtro_name_axa].copy()

# 2. Recalcular la condición exacta
def clasificar_fino_axa_t(row):
    id_match = row["CERTIFICADO"] in HC_ids
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
candidatos_dismatch_axa = candidatos_dismatch_axa[candidatos_dismatch_axa["Tipo_Disparidad"].notna()].copy()


dismatch_AXA = candidatos_dismatch_axa.copy()


# ===================
# HC: misma lógica
# ===================


AXA_ids = set(dependientes_AXA["CERTIFICADO"])
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
candidatos_dismatch_hc = candidatos_dismatch_hc[candidatos_dismatch_hc["Tipo_Disparidad"].notna()].copy()

dismatch_HC = candidatos_dismatch_hc.copy()

# 3. Nos quedamos con los dismatchs reales

with pd.ExcelWriter("Dependents_June.xlsx", engine="openpyxl") as writer:
    dismatch_HC.to_excel(writer, sheet_name="Diferencias_en_HC", index=False) # Employees in HC but not in AXA
    dismatch_AXA.to_excel(writer, sheet_name="Diferencias_en_AXA", index=False) #Employees in AXA but not in HC
    dependientes_AXA.to_excel(writer, sheet_name = "Base_Original_AXA", index= False)
    dependientes_HC.to_excel(writer,sheet_name="Base_Original_HC",index= False)
