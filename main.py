import os
import io
import pandas as pd
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.requests import Request

app = FastAPI(title="Conciliador XLSX")

# Configurar plantillas
templates = Jinja2Templates(directory="templates")

@app.get("/")
async def read_index(request: Request):
    return templates.TemplateResponse(request=request, name="index.html")

def clean_currency(value):
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    
    # Convertir a cadena y limpiar espacios y símbolos
    s_value = str(value).replace('$', '').strip()
    if not s_value or s_value == '.00':
        return 0.0

    # Determinar si el separador decimal es coma o punto
    # Si hay ambos, el que esté más a la derecha suele ser el decimal
    last_dot = s_value.rfind('.')
    last_comma = s_value.rfind(',')
    
    if last_dot > last_comma:
        # Formato US: 1,234.56 -> eliminar comas, mantener punto
        s_value = s_value.replace(',', '')
    elif last_comma > last_dot:
        # Formato EU: 1.234,56 -> eliminar puntos, cambiar coma por punto
        s_value = s_value.replace('.', '').replace(',', '.')
    else:
        # Solo hay uno o ninguno
        s_value = s_value.replace(',', '.')

    try:
        return float(s_value)
    except ValueError:
        return 0.0

@app.post("/api/process")
async def process_file(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Formato de archivo no soportado")

    try:
        # Leer el archivo excel completo inicialmente sin cabecera para encontrarla
        contents = await file.read()
        df_raw = pd.read_excel(io.BytesIO(contents), header=None)
        
        # Buscar la fila de cabecera
        header_row_index = 0
        target_keywords = ['fecha', 'descripcion', 'valor', 'saldo']
        for i, row in df_raw.iterrows():
            row_values = [str(x).lower().strip() for x in row.values if not pd.isna(x)]
            matches = sum(1 for kw in target_keywords if any(kw in rv for rv in row_values))
            if matches >= 2:
                header_row_index = i
                break
        
        # Volver a leer desde la fila encontrada
        df = pd.read_excel(io.BytesIO(contents), header=header_row_index)

        # Normalizar nombres de columnas
        df.columns = [str(c).strip() for c in df.columns]
        
        # Identificar columnas de Valor y Saldo (pueden venir como 'VALOR', 'Suma de VALOR', etc)
        valor_cols = [c for c in df.columns if 'valor' in c.lower()]
        saldo_cols = [c for c in df.columns if 'saldo' in c.lower()]
        
        cols_to_fix = valor_cols + saldo_cols
        
        for col in cols_to_fix:
            df[col] = df[col].apply(clean_currency)

        # Exportar a Excel con formato de tabla
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Data Procesada')
            
            workbook = writer.book
            worksheet = writer.sheets['Data Procesada']
            
            num_rows, num_cols = df.shape
            from openpyxl.utils import get_column_letter
            last_col_letter = get_column_letter(num_cols)
            # El rango incluye la cabecera (fila 1) hasta num_rows + 1
            table_range = f"A1:{last_col_letter}{num_rows + 1}"
            
            from openpyxl.worksheet.table import Table, TableStyleInfo
            tab = Table(displayName="TablaProcesada", ref=table_range)
            
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                   showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            tab.tableStyleInfo = style
            worksheet.add_table(tab)
            
            # Aplicar formato numérico explícito a las columnas de dinero
            # Esto soluciona el problema de que Excel haga "recuento" (count) en lugar de suma
            from openpyxl.styles import NamedStyle
            # Crear un estilo de número si no existe
            number_format = '#,##0.00'
            
            for i, col_name in enumerate(df.columns):
                col_letter = get_column_letter(i + 1)
                
                # Si es columna de dinero, aplicar formato
                is_money = any(kw in col_name.lower() for kw in ['valor', 'saldo'])
                
                if is_money:
                    for cell in worksheet[col_letter][1:]: # Saltamos la cabecera
                        cell.number_format = number_format
                
                # Ajustar ancho de columnas
                data_max = 0
                for val in df[col_name]:
                    val_str = str(val) if not pd.isna(val) else ""
                    if len(val_str) > data_max:
                        data_max = len(val_str)
                
                head_len = len(str(col_name))
                column_length = max(data_max, head_len) + 4 # Un poco más de margen
                worksheet.column_dimensions[col_letter].width = column_length

        output.seek(0)
        
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=procesado_{file.filename}"}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al procesar el archivo: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    import os
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
