from flask import Flask, request, jsonify, send_file
import pandas as pd
import numpy as np
import io
import unicodedata
import calendar

app = Flask(__name__)

def normalizar_columna(col):
    col = str(col)
    col = unicodedata.normalize('NFKD', col).encode('ASCII', 'ignore').decode('utf-8')
    return col.lower().strip()

def generar_mes(fecha):
    try:
        fecha = pd.to_datetime(fecha)
        nombre_mes = calendar.month_name[fecha.month].capitalize()
        nombre_mes = unicodedata.normalize('NFKD', nombre_mes).encode('ASCII', 'ignore').decode('utf-8')
        return f"{nombre_mes}_{str(fecha.year)[-2:]}"
    except:
        return np.nan

@app.route('/homologar', methods=['POST'])
def homologar():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    base = request.files['base']
    nombre_archivo = file.filename.lower()
    es_colpatria = "colpatria" in nombre_archivo

    xls = pd.ExcelFile(file)
    hoja = next((h for h in xls.sheet_names if "detalle" in h.lower() or "extracto" in h.lower()), xls.sheet_names[0])
    skiprows = 0 if es_colpatria else 4
    df = pd.read_excel(xls, sheet_name=hoja, skiprows=skiprows)

    df.columns = [normalizar_columna(c) for c in df.columns]
    df = df.loc[:, ~pd.Index(df.columns).duplicated(keep='first')]

    mapeo_flexible = {
        'sucursal': 'SUCURSALES',
        'fecha recaudo': 'MES',
        'fecha de aplicacion': 'MES',
        'fecha pago': 'MES',
        'ramo': 'RAMO',
        'ramo 1': 'RAMO',
        'poliza': 'N. PÓLIZA',
        'póliza': 'N. PÓLIZA',
        'certificado': 'CERTIFICADO',
        'endoso': 'CERTIFICADO',
        'doc. tomador': 'NIT',
        'nit': 'NIT',
        'nombre tomador': 'RESP. DE PAGO',
        'tomador': 'RESP. DE PAGO',
        'valor comision': 'COMISIÓN',
        'valor iva comision': 'IVA',
        'rte iva': 'RTE. IVA',
        'rte fte': 'RTE. FTE',
        'rte ica': 'ICA',
        'total comision': 'TOTAL PAGADO'
    }

    df.rename(columns={col: mapeo_flexible[col] for col in df.columns if col in mapeo_flexible}, inplace=True)
    if 'MES' in df.columns:
        df['MES'] = df['MES'].apply(generar_mes)

    df_base = pd.read_excel(base)
    estructura_final = df_base.columns.tolist()

    col_aseg = "NOMBRE CIA . ASEG"
    nombre_aseg = "COLPATRIA" if es_colpatria else "MUNDIAL SEGUROS GENERALES"
    if col_aseg not in df.columns:
        idx = estructura_final.index(col_aseg)
        df.insert(loc=idx, column=col_aseg, value=nombre_aseg)
    else:
        df[col_aseg] = nombre_aseg

    for col in estructura_final:
        if col not in df.columns:
            df[col] = np.nan
    df = df[estructura_final]
    df['Campos_Faltantes'] = df.isna().sum(axis=1)

    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(output, download_name='archivo_homologado.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run()
