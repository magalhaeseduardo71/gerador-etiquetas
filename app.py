from flask import Flask, request, send_file, render_template_string
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.graphics.barcode import code128

app = Flask(__name__)

HTML = """
<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <title>Gerador de Etiquetas</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body {
      min-height: 100vh;
      margin: 0;
      padding: 0;
      font-family: 'Segoe UI', Arial, sans-serif;
      background: #f8fafc;
      display: flex;
      align-items: center;
      justify-content: center;
      flex-direction: column;
    }
    .card {
      background: #fff;
      border-radius: 18px;
      box-shadow: 0 2px 16px 0 rgba(50,60,70,0.07);
      padding: 32px 26px 28px 26px;
      max-width: 370px;
      width: 90%;
      display: flex;
      flex-direction: column;
      align-items: center;
      margin-top: 32px;
    }
    h1 {
      font-size: 1.7em;
      margin-bottom: 6px;
      color: #1d2431;
      font-weight: 700;
      text-align: center;
    }
    .desc {
      color: #555c6c;
      font-size: 1.1em;
      margin-bottom: 18px;
      text-align: center;
    }
    input[type="file"] {
      border: 1px solid #d1d7e2;
      border-radius: 8px;
      padding: 8px;
      background: #f6f8fa;
      color: #333;
      margin-bottom: 18px;
      width: 100%;
      box-sizing: border-box;
      font-size: 1em;
    }
    button {
      padding: 12px 0;
      background: #116cff;
      color: #fff;
      font-weight: bold;
      border: none;
      border-radius: 8px;
      font-size: 1.1em;
      width: 100%;
      cursor: pointer;
      transition: background 0.2s;
      margin-top: 5px;
      margin-bottom: 2px;
      box-shadow: 0 2px 4px 0 rgba(17, 108, 255, 0.04);
    }
    button:hover {
      background: #0045a3;
    }
    .creditos {
      margin-top: 40px;
      color: #94a3b8;
      font-size: 0.93em;
    }
    @media (max-width: 500px) {
      .card { padding: 16px 5px; }
      h1 { font-size: 1.12em; }
      .desc { font-size: 1em; }
    }
  </style>
</head>
<body>
  <div class="card">
    <h1>Gerador de Etiquetas 2.0</h1>
    <div class="desc">
      Envie sua planilha Excel (.xlsx)<br><br>
      <span style="font-size:0.93em; color:#116cff; font-weight:500">Atenção:</span> Os endereço deve estar na primeira coluna da planilha.
    </div>
    <form method="post" enctype="multipart/form-data">
      <input type="file" name="file" required accept=".xlsx">
      <button type="submit">Gerar PDF de Etiquetas</button>
    </form>
  </div>
  <div class="creditos">
    &copy; 2025 - Gerador de Etiquetas | Desenvolvido por Eduardo Magalhães
  </div>
</body>
</html>
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
            font_size = 28
            c.setFont("Helvetica-Bold", font_size)
            text_width = c.stringWidth(endereco_visual, "Helvetica-Bold", font_size)
            y_text = altura_pt - (altura_pt * 0.45)
            c.drawString((largura_pt - text_width) / 2, y_text, endereco_visual)
            espacamento = 3 * mm
            bar_height = 9 * mm
            barcode = code128.Code128(endereco_completo, barHeight=bar_height, barWidth=0.9)
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

