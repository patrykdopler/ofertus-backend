
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from docx import Document
from docx.shared import Inches
from openpyxl import load_workbook
from io import BytesIO
import zipfile

app = FastAPI()


def extract_positions(xlsx_bytes: bytes):
    # Odczytuje pozycje z arkusza 'Arkusz1'.
    # Szuka w kolumnie 7 (G) tekstów zaczynających się od 'Poz.'
    # oraz w kolumnie 6 (F) etykiet: 'Nazwa:', 'Opis:', 'Wypełnienia:', 'Ilość:'.
    wb = load_workbook(BytesIO(xlsx_bytes), data_only=True)
    ws = wb['Arkusz1']
    max_row = ws.max_row

    poz_rows = []
    for r in range(1, max_row + 1):
        v = ws.cell(row=r, column=7).value
        if isinstance(v, str) and v.strip().startswith('Poz.'):
            poz_rows.append(r)

    positions = []
    for idx, base_row in enumerate(poz_rows, start=1):
        block_start = base_row
        block_end = poz_rows[idx] - 1 if idx < len(poz_rows) else min(max_row, base_row + 40)

        info = {
            'lp': idx,
            'nazwa': ws.cell(row=base_row, column=7).value or '',
            'ilosc': '',
            'opis': '',
            'image': None,
        }

        for r in range(block_start, block_end + 1):
            label = ws.cell(row=r, column=6).value
            if isinstance(label, str):
                t = label.strip()
                if t == 'Ilość:':
                    val = ws.cell(row=r, column=7).value
                    if val is not None:
                        info['ilosc'] = str(val)
                elif t in ('Wypełnienia:', 'Opis:'):
                    val = ws.cell(row=r, column=7).value
                    if val:
                        text = str(val).replace('\n', ' ').replace('_x000D_', ' ')
                        if info['opis']:
                            info['opis'] += ' '
                        info['opis'] += text

        positions.append(info)

    return positions


def extract_images(xlsx_bytes: bytes):
    # Czyta pliki graficzne z xl/media w XLSX.
    images = []
    with zipfile.ZipFile(BytesIO(xlsx_bytes)) as z:
        media_files = sorted(
            name
            for name in z.namelist()
            if name.startswith('xl/media/')
            and (name.lower().endswith('.png') or name.lower().endswith('.jpg') or name.lower().endswith('.jpeg'))
        )
        for name in media_files:
            images.append(z.read(name))
    return images


def add_image_below_text(cell, image_bytes: bytes, target_px_width: int = 500):
    # Dodaje obraz pod istniejącym tekstem w komórce.
    if not image_bytes:
        return
    width_inch = target_px_width / 96.0  # 96 dpi ~ 96 px/cal
    paragraph = cell.add_paragraph()
    run = paragraph.add_run()
    run.add_picture(BytesIO(image_bytes), width=Inches(width_inch))


@app.post('/generate-offer')
async def generate_offer(xlsx: UploadFile = File(...)):
    # Główna funkcja API – przyjmuje XLSX i zwraca DOCX.
    xlsx_bytes = await xlsx.read()

    positions = extract_positions(xlsx_bytes)
    images = extract_images(xlsx_bytes)

    for idx, pos in enumerate(positions):
        if idx < len(images):
            pos['image'] = images[idx]

    doc = Document('template.docx')
    if not doc.tables:
        table = doc.add_table(rows=2, cols=4)
        hdr = table.rows[0].cells
        hdr[0].text = 'L.p.'
        hdr[1].text = 'Rysunek (Widok od zewnątrz), wymiary'
        hdr[2].text = 'Ilość sztuk'
        hdr[3].text = 'OPIS'
        template_row = table.rows[1]
    else:
        table = doc.tables[0]
        if len(table.rows) < 2:
            table.add_row()
        template_row = table.rows[1]

    # wyczyść wiersz wzorcowy
    for cell in template_row.cells:
        cell.text = ''

    # usuń nadmiarowe wiersze
    while len(table.rows) > 2:
        table._tbl.remove(table.rows[-1]._tr)

    # wypełnij pozycjami
    for idx, pos in enumerate(positions):
        if idx == 0:
            row_cells = template_row.cells
        else:
            new_row = table.add_row()
            row_cells = new_row.cells

        row_cells[0].text = str(pos['lp'])
        name_cell = row_cells[1]
        name_cell.text = pos['nazwa']
        if pos.get('image'):
            add_image_below_text(name_cell, pos['image'], target_px_width=500)
        row_cells[2].text = str(pos['ilosc'])
        row_cells[3].text = pos['opis']

    out = BytesIO()
    doc.save(out)
    out.seek(0)

    return StreamingResponse(
        out,
        media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        headers={'Content-Disposition': 'attachment; filename="oferta.docx"'}
    )
