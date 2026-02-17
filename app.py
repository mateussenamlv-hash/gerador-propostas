from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import os
import subprocess
from datetime import datetime

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.docx")
OUTPUT_DOCX = os.path.join(BASE_DIR, "contrato_gerado.docx")
OUTPUT_PDF = os.path.join(BASE_DIR, "contrato_gerado.pdf")


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/gerar-pdf", methods=["POST"])
def gerar_pdf():

    doc = DocxTemplate(TEMPLATE_PATH)

    cliente = request.form.get("cliente")
    cpf = request.form.get("cpf")
    modelo = request.form.get("modelo")
    franquia = request.form.get("franquia")
    contrato = request.form.get("contrato")
    excedente = request.form.get("excedente")
    valor = request.form.get("valor")
    validade = request.form.get("validade")

    data_atual = datetime.now().strftime("%d/%m/%Y")

    imagem = request.files.get("imagem")

    if imagem and imagem.filename != "":
        imagem_template = InlineImage(doc, imagem, width=Mm(50))
    else:
        imagem_template = ""

    context = {
        "CLIENTE": cliente,
        "CPF": cpf,
        "MODELO": modelo,
        "FRANQUIA": franquia,
        "CONTRATO": contrato,
        "EXCEDENTE": excedente,
        "VALOR": valor,
        "VALIDADE": validade,
        "DATA": data_atual,
        "IMAGEM": imagem_template,
    }

    doc.render(context)
    doc.save(OUTPUT_DOCX)

    # 🔥 CONVERSÃO VIA LIBREOFFICE
    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        OUTPUT_DOCX,
        "--outdir", BASE_DIR
    ])

    return send_file(OUTPUT_PDF, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
