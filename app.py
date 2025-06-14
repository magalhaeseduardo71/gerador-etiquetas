from flask import Flask, request, send_file, render_template_string
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.graphics.barcode import code128

app = Flask(__name__)

HTML = """
<!doctype html>
<title>Gerar Etiquetas</title>
<h1>Upload do arquivo Excel para gerar etiquetas</h1>
<form method=post enctype=multipart/form-data>
  <input type=file name=file>
  <input type=submit value=Gerar>
</form>
"""

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if not file:
            return 'Nenhum arquivo enviado!'
        df = pd.read_excel(file)
        coluna_endereco = df.columns[0]
        pdf_file = "etiquetas_ajustadas.pdf"
        largura_mm, altura_mm = 60, 25
        largura_pt, altura_pt = largura_mm * mm, altura_mm * mm
        c = canvas.Canvas(pdf_file, pagesize=(largura_pt, altura_pt))
        for i, row in df.iterrows():
            endereco_completo = str(row[coluna_endereco]).strip()
            endereco_visual = endereco_completo[5:]
            font_size = 32
            c.setFont("Helvetica-Bold", font_size)
            text_width = c.stringWidth(endereco_visual, "Helvetica-Bold", font_size)
            y_text = altura_pt - (altura_pt * 0.45)
            c.drawString((largura_pt - text_width) / 2, y_text, endereco_visual)
            espacamento = 3 * mm
            bar_height = 9 * mm
            barcode = code128.Code128(endereco_completo, barHeight=bar_height, barWidth=0.8)
            barcode_width = barcode.width
            x_barcode = (largura_pt - barcode_width) / 2
            y_barcode = y_text - espacamento - bar_height
            barcode.drawOn(c, x_barcode, y_barcode)
            c.showPage()
        c.save()
        return send_file(pdf_file, as_attachment=True)
    return render_template_string(HTML)

import os

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)

