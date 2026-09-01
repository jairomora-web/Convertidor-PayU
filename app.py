import pandas as pd
import re
from datetime import datetime
import streamlit as st
import io

# Estilos visuales estilo BOLD
st.markdown("""
    <style>
    /* Fondo general */
    .stApp {
        background-color: #FAFAFA;
    }
    /* Título principal */
    h1 {
        color: #0F172A !important;
    }
    /* Botón de descarga */
    div.stDownloadButton > button {
        background-color: #FF0051 !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: bold !important;
        padding: 12px 24px !important;
        font-size: 16px !important;
    }
    /* Efecto al pasar el cursor sobre el botón */
    div.stDownloadButton > button:hover {
        background-color: #D90043 !important;
        color: #FFFFFF !important;
    }
    /* Caja donde se arrastran los archivos */
    [data-testid="stFileUploader"] {
        background-color: #FFFFFF;
        border: 2px dashed #FF0051;
        border-radius: 10px;
        padding: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# 1. Interfaz de Usuario
st.set_page_config(page_title="Convertidor PayU", page_icon="📄")
st.title("📄 Convertidor de Archivos PAYU")
st.write("Sube tus reportes de PayU para generar el archivo consolidado de transacciones y transferencias.")

uploaded_files = st.file_uploader("📤 Seleccione los archivos de PayU a procesar:", accept_multiple_files=True, type=['xlsx', 'xls'])

# 2. Función para procesar cada archivo
@st.cache_data
def procesar_archivo(file_name, file_content):
    df = pd.read_excel(io.BytesIO(file_content), header=4).dropna(how='all')

    df['ID_PAGO'] = df['DESCRIPCION'].apply(
        lambda x: re.search(r'\[(.*?)\]', str(x)).group(1) if pd.notna(x) and re.search(r'\[(.*?)\]', str(x)) else None
    )
    df['CONCEPTO'] = df['DESCRIPCION'].apply(
        lambda x: str(x).split(' [')[0].strip() if pd.notna(x) else None
    )

    df = df[df['ID_PAGO'].notna()].copy()
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    return df

# Solo se ejecuta si el usuario sube archivos
if uploaded_files:
    with st.spinner("🔍 Procesando archivos..."):
        # 3. Procesar y combinar archivos
        dfs = []
        for f in uploaded_files:
            # st.file_uploader devuelve un objeto tipo BytesIO que se puede leer directamente
            f.seek(0)
            dfs.append(procesar_archivo(f.name, f.read()))
            
        df_combinado = pd.concat(dfs, ignore_index=True)

        # 4. Separar transacciones y transferencias
        transferencias_mask = df_combinado['ID_PAGO'].str.contains('PAYMENT_ORDER', na=False)
        transferencias = df_combinado[transferencias_mask].copy()
        transacciones = df_combinado[~transferencias_mask].copy()

        # 5. Procesar transacciones (VENTAS)
        if not transacciones.empty:
            pivot = transacciones.pivot_table(
                index='ID_PAGO', columns='CONCEPTO', values=['CREDITOS', 'DEBITOS'],
                aggfunc='sum', fill_value=0
            )
            pivot.columns = [f'{col[1]}_{col[0]}' for col in pivot.columns]
            pivot = pivot.reset_index()

            columnas_necesarias = {
                'SALES_CREDITOS': 'VENTA', 'POL_COMMISSION_DEBITOS': 'COMISION',
                'IVA_POL_COMMISSION_DEBITOS': 'IVA_COMISION', 'RENTA_RETENTION_DEBITOS': 'RETEFUENTE',
                'ICA_RETENTION_DEBITOS': 'RETEICA', 'IVA_RETENTION_DEBITOS': 'RETEIVA'
            }

            for col in columnas_necesarias.keys():
                if col not in pivot.columns:
                    pivot[col] = 0.0

            reporte_transacciones = pivot.assign(
                VENTA = lambda x: x['SALES_CREDITOS'],
                COMISION = lambda x: x['POL_COMMISSION_DEBITOS'].abs(),
                IVA_COMISION = lambda x: x['IVA_POL_COMMISSION_DEBITOS'].abs(),
                RETEFUENTE = lambda x: x['RENTA_RETENTION_DEBITOS'].abs(),
                RETEICA = lambda x: x['ICA_RETENTION_DEBITOS'].abs(),
                RETEIVA = lambda x: x['IVA_RETENTION_DEBITOS'].abs(),
                NETO = lambda x: x['VENTA'] - x['COMISION'] - x['IVA_COMISION'] - x['RETEFUENTE'] - x['RETEICA'] - x['RETEIVA']
            )

            info_extra = transacciones.groupby('ID_PAGO').agg({
                'FECHA': 'first', 'DOCUMENTO': 'first', 'NUEVO SALDO': 'last'
            }).reset_index()

            reporte_final = pd.DataFrame({
                'FECHA.TRANSACCION': info_extra['FECHA'], 'TRANSACTION.ID': info_extra['ID_PAGO'],
                'CANAL RECAUDO': 'PAYU', 'PROCESADOR.DE.TRANSACCION': 'BANCOLOMBIA',
                'MEDIO.DE.PAGO': '', 'VTA.TOTAL': reporte_transacciones['VENTA'],
                'VLR.COMPRA': 0, 'PROPINA': 0, 'IVA': reporte_transacciones['IVA_COMISION'],
                'IAC': 0, 'COSTO.PROCESAMIENTO': reporte_transacciones['COMISION'],
                'RTE_FUENTE': reporte_transacciones['RETEFUENTE'], 'RTE_ICA': reporte_transacciones['RETEICA'],
                'RTE_IVA': reporte_transacciones['RETEIVA'], 'NETO.BANCO': reporte_transacciones['NETO'],
                'COMISION.BOLD': 0
            }).sort_values('FECHA.TRANSACCION')

            reporte_final['FECHA.TRANSACCION'] = reporte_final['FECHA.TRANSACCION'].dt.strftime('%d/%m/%Y')
        else:
            reporte_final = pd.DataFrame(columns=[
                'FECHA.TRANSACCION', 'TRANSACTION.ID', 'CANAL RECAUDO', 'PROCESADOR.DE.TRANSACCION',
                'MEDIO.DE.PAGO', 'VTA.TOTAL', 'VLR.COMPRA', 'PROPINA', 'IVA', 'IAC',
                'COSTO.PROCESAMIENTO', 'RTE_FUENTE', 'RTE_ICA', 'RTE_IVA', 'NETO.BANCO', 'COMISION.BOLD'
            ])

        # 6. Procesar transferencias
        if not transferencias.empty:
            reporte_transferencias = transferencias.groupby(['DOCUMENTO', 'FECHA']).agg({
                'DEBITOS': 'sum', 'CREDITOS': 'sum', 'CONCEPTO': lambda x: ', '.join(x)
            }).reset_index()

            monto_principal = transferencias[transferencias['CONCEPTO'] == 'PAYMENT_ORDER'].groupby(['DOCUMENTO', 'FECHA'])['DEBITOS'].sum().abs().reset_index(drop=True)
            comision = transferencias[transferencias['CONCEPTO'].str.contains('POL_COMMISION', na=False)].groupby(['DOCUMENTO', 'FECHA'])['DEBITOS'].sum().abs().reset_index(drop=True)
            iva_comision = transferencias[transferencias['CONCEPTO'].str.contains('IVA_PAYMENT', na=False)].groupby(['DOCUMENTO', 'FECHA'])['DEBITOS'].sum().abs().reset_index(drop=True)

            reporte_transferencias['MONTO_TRANSFERENCIA'] = monto_principal
            reporte_transferencias['COMISION'] = comision
            reporte_transferencias['IVA_COMISION'] = iva_comision
            reporte_transferencias['MONTO_NETO'] = reporte_transferencias['MONTO_TRANSFERENCIA'] - reporte_transferencias['COMISION'].fillna(0)
            reporte_transferencias['TOTAL'] = reporte_transferencias['MONTO_TRANSFERENCIA'] + reporte_transferencias['COMISION'].fillna(0) + reporte_transferencias['IVA_COMISION'].fillna(0)
            
            reporte_transferencias = reporte_transferencias.sort_values('FECHA').rename(columns={'FECHA': 'FECHA.TRANSFERENCIA'})
            reporte_transferencias['FECHA.TRANSFERENCIA'] = reporte_transferencias['FECHA.TRANSFERENCIA'].dt.strftime('%d/%m/%Y')

            columnas_finales = ['FECHA.TRANSFERENCIA', 'DOCUMENTO', 'MONTO_TRANSFERENCIA', 'MONTO_NETO', 'COMISION', 'IVA_COMISION', 'TOTAL']
            reporte_transferencias = reporte_transferencias[columnas_finales].fillna(0)
        else:
            columnas_finales = ['FECHA.TRANSFERENCIA', 'DOCUMENTO', 'MONTO_TRANSFERENCIA', 'MONTO_NETO', 'COMISION', 'IVA_COMISION', 'TOTAL']
            reporte_transferencias = pd.DataFrame(columns=columnas_finales)

        # 7. Exportar a Excel en memoria (BytesIO)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='DD/MM/YYYY') as writer:
            column_order = [
                'FECHA.TRANSACCION', 'TRANSACTION.ID', 'CANAL RECAUDO', 'PROCESADOR.DE.TRANSACCION',
                'MEDIO.DE.PAGO', 'VTA.TOTAL', 'VLR.COMPRA', 'PROPINA', 'IVA', 'IAC',
                'COSTO.PROCESAMIENTO', 'RTE_FUENTE', 'RTE_ICA', 'RTE_IVA', 'NETO.BANCO', 'COMISION.BOLD'
            ]
            reporte_final[column_order].to_excel(writer, sheet_name='Transacciones', index=False, startrow=1, header=False)
            reporte_transferencias.to_excel(writer, sheet_name='Transferencias', index=False, startrow=1, header=False)

            workbook = writer.book
            header_format = workbook.add_format({'bold': True, 'fg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center'})
            num_format = workbook.add_format({'num_format': '#,##0.00', 'align': 'right'})
            date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'align': 'center'})
            text_format = workbook.add_format({'align': 'center'})

            ws_trans = writer.sheets['Transacciones']
            for col_num, value in enumerate(column_order):
                ws_trans.write(0, col_num, value, header_format)
            ws_trans.add_table(0, 0, max(len(reporte_final), 1), len(column_order)-1, {'columns': [{'header': col} for col in column_order], 'style': 'Table Style Medium 9'})
            
            ws_trans.set_column('A:A', 12, date_format)
            ws_trans.set_column('B:B', 25)
            ws_trans.set_column('C:C', 12, text_format)
            ws_trans.set_column('D:D', 20, text_format)
            ws_trans.set_column('E:E', 12, text_format)
            for col in ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P']:
                ws_trans.set_column(f'{col}:{col}', 14, num_format)

            ws_transf = writer.sheets['Transferencias']
            for col_num, value in enumerate(reporte_transferencias.columns):
                ws_transf.write(0, col_num, value, header_format)
            ws_transf.add_table(0, 0, max(len(reporte_transferencias), 1), max(len(reporte_transferencias.columns)-1, 1), {'columns': [{'header': col} for col in reporte_transferencias.columns], 'style': 'Table Style Medium 10'})
            
            ws_transf.set_column('A:A', 12, date_format)
            ws_transf.set_column('B:B', 15)
            for col in ['C', 'D', 'E', 'F']:
                ws_transf.set_column(f'{col}:{col}', 15, num_format)

            ws_trans.freeze_panes(1, 0)
            ws_transf.freeze_panes(1, 0)
            
        st.success("✅ ¡Archivo procesado con éxito!")
        
        # Botón de descarga
        fecha_actual = datetime.now().strftime("%Y-%m-%d")
        st.download_button(
            label="⬇️ Descargar Reporte Final de PayU",
            data=output.getvalue(),
            file_name=f'Payu_{fecha_actual}.xlsx',
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
