from flask import Flask, request, send_file, render_template_string
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.graphics.barcode import code128
import os

app = Flask(__name__)

HTML = """
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <title>Gerador de Etiquetas</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body {
      margin: 0;
      padding: 0;
      min-height: 100vh;
      font-family: 'Segoe UI', Arial, sans-serif;
      background: #181b23;
      color: #e5eaf5;
      display: flex;
      flex-direction: column;
    }
    header {
      background: #10121a;
      padding: 24px 0 18px 0;
      text-align: center;
      letter-spacing: 1px;
      box-shadow: 0 2px 12px #0008;
    }
    header h1 {
      margin: 0;
      color: #36a2f5;
      font-size: 2.1em;
      font-weight: 800;
      text-shadow: 0 1px 8px #00408066;
    }
    main {
      flex: 1;
      display: flex;
      justify-content: center;
      align-items: center;
    }
    .card {
      background: #232735;
      border-radius: 14px;
      box-shadow: 0 4px 32px #0009;
      padding: 36px 28px 28px 28px;
      max-width: 370px;
      width: 100%;
      margin: 36px 0 24px 0;
      display: flex;
      flex-direction: column;
      align-items: center;
    }
    .card h2 {
      margin-top: 0;
      margin-bottom: 16px;
      color: #36a2f5;
      font-size: 1.22em;
      font-weight: 600;
      letter-spacing: 0.7px;
      text-align: center;
    }
    .desc {
      color: #bfc9d7;
      margin-bottom: 22px;
      text-align: center;
      font-size: 1em;
    }
    form {
      width: 100%;
      display: flex;
      flex-direction: column;
      align-items: center;
    }
    input[type="file"] {
      background: #181b23;
      color: #e5eaf5;
      border: 2px solid #36a2f5;
      border-radius: 7px;
      padding: 12px 10px;
      margin-bottom: 20px;
      width: 100%;
      font-size: 1em;
      transition: border 0.2s;
    }
    input[type="file"]:focus {
      border-color: #5cc0ff;
      outline: none;
    }
    button {
      padding: 13px 0;
      background: linear-gradient(90deg, #2077d7 60%, #36a2f5 100%);
      color: #fff;
      font-weight: bold;
      border: none;
      border-radius: 8px;
      font-size: 1.12em;
      width: 100%;
      cursor: pointer;
      margin-bottom: 4px;
      box-shadow: 0 2px 8px #005ec880;
      letter-spacing: 0.5px;
      transition: background 0.17s;
    }
    button:hover {
      background: linear-gradient(90deg, #36a2f5 60%, #2077d7 100%);
    }
    footer {
      background: #10121a;
      color: #bfc9d7;
      text-align: center;
      padding: 18px 10px 14px 10px;
      font-size: 1em;
      letter-spacing: 0.4px;
      box-shadow: 0 -2px 14px #0005;
    }
    @media (max-width: 600px) {
      .card {
        padding: 14px 4vw 14px 4vw;
        margin: 18px 0 12px 0;
      }
      header h1 { font-size: 1.16em; }
      .card h2 { font-size: 1em; }
    }
  </style>
</head>
<body>
  <header>
    <h1>Gerador de Etiquetas 2.0</h1>
  </header>
  <main>
    <div class="card">
      <h2>Envie sua planilha Excel (.xlsx)</h2>
      <div class="desc">
        O sistema irá gerar um PDF pronto para impressão das etiquetas.<br>
        <span style="color:#5cc0ff;font-size:0.95em;">Dica: Coloque os endereços na primeira coluna da planilha.</span>
      </div>
      <form method="post" enctype="multipart/form-data">
        <input type="file" name="file" required accept=".xlsx">
        <button type="submit">Gerar PDF</button>
      </form>
    </div>
  </main>
  <footer>
    &copy; 2025 Gerador de Etiquetas &mdash; Desenvolvido por Eduardo Magalhães
  </footer>
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
            barcode = code128.Code128(endereco_completo, barHeight=bar_height, barWidth=0.8)
            barcode_width = barcode.width
            x_barcode = (largura_pt - barcode_width) / 2
            y_barcode = y_text - espacamento - bar_height
            barcode.drawOn(c, x_barcode, y_barcode)
            c.showPage()
        c.save()
        return send_file(pdf_file, as_attachment=True)
    return render_template_string(HTML)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
