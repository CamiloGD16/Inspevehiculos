from docx import Document
import json
import sys

# Recibir datos desde Power Automate
datos_json = sys.argv[1] if len(sys.argv) > 1 else '{}'
datos = json.loads(datos_json)

# Crear documento
doc = Document()
doc.add_heading('Reporte de Inspección', 0)

# Tabla
table = doc.add_table(rows=1, cols=3)
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Item'
hdr_cells[1].text = 'Cumple'
hdr_cells[2].text = 'Observación'

# Agregar filas desde los datos
for dato in datos.get('items', []):
    row_cells = table.add_row().cells
    row_cells[0].text = dato.get('Item', '')
    row_cells[1].text = dato.get('Cumple', '')
    row_cells[2].text = dato.get('Observacion', '')

doc.save('reporte.docx')
print("Word creado!")
