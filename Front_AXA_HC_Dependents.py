import streamlit as st
from backfunctions import UploaderAxaDependents, UploaderHCDependents, ProcessDependents_Generate_excel

dependientes_AXA = None
dependientes_HC = None
month_from_filename = None

st.set_page_config(page_title="Inconsistencias Dependientes", layout= 'wide')
st.image("Logo_8192884199.png", width=250)
st.markdown("<h1 style='text-align: center;'>Revisión de inconsistencias entre los asegurados por AXA y por los dependientes de Central</h1>", unsafe_allow_html=True)
st.write('-----------------------------------------------------------------------------------------------------------------------------')

col1, col2 = st.columns(2)

with col1:
    Upload_AXA_dependents = st.file_uploader(
        label='Introduce el archivo de los dependientes de AXA en formato excel', 
        key="uploader1")
    dependientes_AXA, base_manual_axa = UploaderAxaDependents(Upload_AXA_dependents)

with col2:
    Upload_HC_dependents = st.file_uploader(
        label='Introduce el archivo de los dependientes de Central (HC) en formato excel', 
        key="uploader2")
    dependientes_HC, month_from_filename = UploaderHCDependents(Upload_HC_dependents)



if dependientes_AXA is not None and dependientes_HC is not None:
    output = ProcessDependents_Generate_excel(dependientes_AXA, dependientes_HC, base_manual_axa)
    st.write('------------------------------------------------------------------------------------------------------------------------------')
    spacer1, center_col, spacer2 = st.columns([1, 2, 1])
    with center_col:
        try: 
            st.download_button(
                label= '📥 Descargar archivo Excel con diferencias entre las dos bases de datos',
                data= output,
                file_name= f'Dependents_{month_from_filename}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            st.error(f"❌ Ocurrió un error al procesar los archivos: {e}")
else:
    st.write('------------------------------------------------------------------------------------------------------------------------------')
    st.markdown(
    """
    <div style='text-align: center; background-color: #e1f5fe; padding: 10px; border-radius: 5px; color: #31708f; border: 1px solid #bce8f1;'>
        🔄 Por favor, sube ambos archivos para ver los resultados y generar la descarga.
    </div>
    """,
    unsafe_allow_html=True
    )