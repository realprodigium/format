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
    
    # Eliminar símbolos de moneda, espacios y comas (separadores de miles)
    # Suponiendo formato estándar: $ 1.234,56 o 1,234.56
    s_value = str(value).replace('$', '').replace(' ', '').strip()
    
    # Manejar formato europeo (puntos para miles, comas para decimales)
    # Si hay puntos y comas, asumimos punto=mil, coma=decimal
    if '.' in s_value and ',' in s_value:
        s_value = s_value.replace('.', '').replace(',', '.')
    # Si solo hay coma y está al final (ej 1234,56), asumimos decimal
    elif ',' in s_value:
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
        # Leer el archivo excel
        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents))

        # Columnas a procesar
        target_columns = ['suma de valor', 'suma de saldo']
        
        # Normalizar nombres de columnas a minúsculas para búsqueda flexible
        df.columns = [str(c).lower().strip() for c in df.columns]
        
        for col in target_columns:
            if col in df.columns:
                df[col] = df[col].apply(clean_currency)
            else:
                # Si no encuentra la columna exacta, buscar parecidas
                found = False
                for actual_col in df.columns:
                    if col in actual_col:
                        df[actual_col] = df[actual_col].apply(clean_currency)
                        found = True
                        break

        # Exportar a Excel con formato de tabla
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Data Procesada')
            
            # Obtener el workbook y el worksheet para aplicar formato de tabla
            workbook = writer.book
            worksheet = writer.sheets['Data Procesada']
            
            # Definir el rango de la tabla
            num_rows, num_cols = df.shape
            # Convertir column index a letra (0=A, 1=B, etc)
            from openpyxl.utils import get_column_letter
            last_col_letter = get_column_letter(num_cols)
            table_range = f"A1:{last_col_letter}{num_rows + 1}"
            
            # Crear el objeto tabla
            from openpyxl.worksheet.table import Table, TableStyleInfo
            tab = Table(displayName="TablaProcesada", ref=table_range)
            
            # Estilo de tabla (claro, con filas alternas)
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                   showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            tab.tableStyleInfo = style
            worksheet.add_table(tab)
            
            # Ajustar ancho de columnas automáticamente
            for i, column in enumerate(df.columns):
                # Calcular el largo máximo de los datos en la columna de forma segura
                data_max = 0
                for val in df[column]:
                    val_str = str(val) if not pd.isna(val) else ""
                    if len(val_str) > data_max:
                        data_max = len(val_str)
                
                head_len = len(str(column))
                column_length = max(data_max, head_len) + 2
                worksheet.column_dimensions[get_column_letter(i + 1)].width = column_length

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
