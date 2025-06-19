from flask import Flask, request, send_file, render_template_string
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.graphics.barcode import code128
import os

app = Flask(__name__)

HTML = """
<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<title>Gerar Etiquetas – O Varejão Auto Peças</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  body {
    margin:0;
    font-family: Arial, sans-serif;
    background:#f5f5f5;
    color:#333;
  }
  header {
    background: #f8aa00;
    padding: 12px 20px;
    display: flex;
    align-items: center;
    justify-content: space-between;
  }
  header .logo {
    font-weight: bold;
    font-size: 1.4em;
    color: #fff;
  }
  header nav a {
    margin-left: 16px;
    color: #fff;
    text-decoration: none;
    font-size: 1em;
  }
  .hero {
    background: #fff;
    padding: 60px 20px;
    text-align: center;
    border-bottom: 2px solid #eee;
  }
  .hero h1 {
    margin: 0;
    font-size: 2em;
    color: #6b4700;
  }
  .container {
    max-width: 480px;
    margin: 40px auto;
    background: #fff;
    padding: 24px;
    border-radius: 8px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
  }
  .container h2 {
    margin-top:0;
    color:#6b4700;
    font-size:1.3em;
  }
  input[type="file"] {
    border:1px solid #ccc;
    border-radius:4px;
    padding:8px;
    width:100%;
    margin:12px 0;
  }
  button {
    background: #f8aa00;
    color: #fff;
    border: none;
    padding: 12px 0;
    width:100%;
    font-size:1.1em;
    cursor: pointer;
    border-radius:4px;
  }
  button:hover {
    background: #d99800;
  }
  footer {
    background:#333;
    color:#fff;
    text-align:center;
    padding:16px 20px;
    font-size:0.9em;
  }
</style>
</head>
<body>

<header>
  <div class="logo">O Varejão Auto Peças</div>
  <nav>
    <a href="#">Home</a>
    <a href="#">Produtos</a>
    <a href="#">Contato</a>
  </nav>
</header>

<section class="hero">
  <h1>Gerador de Etiquetas com Código de Barras</h1>
</section>

<div class="container">
  <h2>Envie sua planilha Excel (.xlsx)</h2>
  <form method="post" enctype="multipart/form-data">
    <input type="file" name="file" required accept=".xlsx">
    <button type="submit">Gerar PDF</button>
  </form>
</div>

<footer>
  &copy; 2025 O Varejão Auto Peças – Todos os direitos reservados
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
            font_size = 32
            c.setFont("Helvetica-Bold", font_size)
            text_width = c.stringWidth(endereco_visual, "Helvetica-Bold", font_size)
            y_text = altura_pt - (altura_pt * 0.45)
            c.drawString((largura_pt - text_width) / 2, y_text, endereco_visual)
            espacamento = 3 * mm
            bar_height = 9 * mm
            barcode = code128.Code128(endereco_completo, barHeight=bar_height, barWidth=1)
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
