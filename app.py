from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl
from reportlab.lib.pagesizes import A4, A3, letter, landscape, portrait
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer
from reportlab.lib.units import cm
import io
import os

app = Flask(__name__)
CORS(app)

def hex_to_color(hex_color):
    if not hex_color:
        return None
    h = str(hex_color).replace('#', '').strip()
    if len(h) == 8:
        h = h[2:]  # strip alpha (ARGB)
    if len(h) != 6:
        return None
    if h.upper() in ('FFFFFF', '000000', 'FFFFFFF'):
        return None
    try:
        r = int(h[0:2], 16) / 255.0
        g = int(h[2:4], 16) / 255.0
        b = int(h[4:6], 16) / 255.0
        return colors.Color(r, g, b)
    except:
        return None

def get_cell_bg_color(cell):
    try:
        fill = cell.fill
        if fill and fill.fill_type and fill.fill_type != 'none':
            fg = fill.fgColor
            if fg and fg.type == 'rgb':
                return hex_to_color(fg.rgb)
    except:
        pass
    return None

def get_cell_font_color(cell):
    try:
        font = cell.font
        if font and font.color and font.color.type == 'rgb':
            return hex_to_color(font.color.rgb)
    except:
        pass
    return None

def get_cell_bold(cell):
    try:
        return bool(cell.font and cell.font.bold)
    except:
        return False

def get_cell_align(cell):
    try:
        if cell.alignment and cell.alignment.horizontal:
            return cell.alignment.horizontal
        if cell.data_type == 'n':
            return 'right'
    except:
        pass
    return 'left'

def format_cell_value(cell):
    val = cell.value
    if val is None:
        return ''
    if isinstance(val, float):
        if val == int(val):
            return f'{int(val):,}'
        return f'{val:,.2f}'
    if isinstance(val, int):
        return f'{val:,}'
    return str(val)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'Excel to PDF API'})

@app.route('/convert/excel-to-pdf', methods=['POST', 'OPTIONS'])
def excel_to_pdf():
    if request.method == 'OPTIONS':
        response = jsonify({'status': 'ok'})
        response.headers['Access-Control-Allow-Origin'] = '*'
        response.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return response

    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    paper = request.form.get('paper', 'A4')
    orientation = request.form.get('orientation', 'auto')
    font_size = int(request.form.get('font_size', 9))

    try:
        file_bytes = io.BytesIO(file.read())
        wb = openpyxl.load_workbook(file_bytes, data_only=True)

        paper_map = {'A4': A4, 'A3': A3, 'Letter': letter}
        base_size = paper_map.get(paper, A4)

        pdf_buffer = io.BytesIO()

        all_elements = []

        for sheet_idx, sheet_name in enumerate(wb.sheetnames):
            ws = wb[sheet_name]
            max_row = ws.max_row
            max_col = ws.max_column
            if not max_row or not max_col:
                continue

            # Read data and styles
            table_data = []
            style_grid = []

            for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
                row_vals = []
                row_styles = []
                for cell in row:
                    row_vals.append(format_cell_value(cell))
                    row_styles.append({
                        'bg': get_cell_bg_color(cell),
                        'fg': get_cell_font_color(cell),
                        'bold': get_cell_bold(cell),
                        'align': get_cell_align(cell),
                    })
                table_data.append(row_vals)
                style_grid.append(row_styles)

            # Column widths
            col_widths = []
            for col_idx in range(1, max_col + 1):
                letter_col = openpyxl.utils.get_column_letter(col_idx)
                dim = ws.column_dimensions.get(letter_col)
                if dim and dim.width and dim.width > 1:
                    col_widths.append(float(dim.width) * 5.5)
                else:
                    col_widths.append(55.0)

            # Row heights
            row_heights = []
            for row_idx in range(1, max_row + 1):
                dim = ws.row_dimensions.get(row_idx)
                if dim and dim.height and dim.height > 1:
                    row_heights.append(float(dim.height) * 0.75)
                else:
                    row_heights.append(float(font_size) + 4.0)

            total_w = sum(col_widths)
            margin = 36

            # Determine orientation
            if orientation == 'landscape':
                page_size = landscape(base_size)
            elif orientation == 'portrait':
                page_size = portrait(base_size)
            else:
                usable_p = base_size[0] - margin * 2
                usable_l = base_size[1] - margin * 2
                if total_w > usable_p and total_w <= usable_l:
                    page_size = landscape(base_size)
                else:
                    page_size = portrait(base_size)

            usable_w = page_size[0] - margin * 2

            # Scale if needed
            if total_w > usable_w:
                scale = usable_w / total_w
                col_widths = [w * scale for w in col_widths]
                row_heights = [max(h * scale, font_size + 2) for h in row_heights]
                actual_font = max(6, int(font_size * scale))
            else:
                actual_font = font_size

            # Build table style
            style_cmds = [
                ('FONTSIZE', (0, 0), (-1, -1), actual_font),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('GRID', (0, 0), (-1, -1), 0.4, colors.Color(0.8, 0.8, 0.8)),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                ('BACKGROUND', (0, 0), (-1, -1), colors.white),
            ]

            for r_idx, row_styles in enumerate(style_grid):
                for c_idx, st in enumerate(row_styles):
                    coord = (c_idx, r_idx)
                    if st['bg']:
                        style_cmds.append(('BACKGROUND', coord, coord, st['bg']))
                    if st['fg']:
                        style_cmds.append(('TEXTCOLOR', coord, coord, st['fg']))
                    if st['bold']:
                        style_cmds.append(('FONTNAME', coord, coord, 'Helvetica-Bold'))
                    align_map = {'right': 'RIGHT', 'center': 'CENTER', 'left': 'LEFT'}
                    rl_align = align_map.get(st['align'], 'LEFT')
                    style_cmds.append(('ALIGN', coord, coord, rl_align))

            table = Table(
                table_data,
                colWidths=col_widths,
                rowHeights=row_heights,
                repeatRows=1
            )
            table.setStyle(TableStyle(style_cmds))
            all_elements.append(table)

        if not all_elements:
            return jsonify({'error': 'No data found'}), 400

        doc = SimpleDocTemplate(
            pdf_buffer,
            pagesize=page_size,
            leftMargin=margin,
            rightMargin=margin,
            topMargin=margin,
            bottomMargin=margin + 12
        )
        doc.build(all_elements)
        pdf_buffer.seek(0)

        out_name = (file.filename or 'output').rsplit('.', 1)[0] + '.pdf'
        return send_file(
            pdf_buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=out_name
        )

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
