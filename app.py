from flask import Flask, render_template, request, send_file
from docx import Document
from datetime import datetime
import os
import subprocess

app = Flask(__name__)

# Caminhos do projeto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "template.docx")

@app.route("/")
def home():
    return render_template("index.html")


@app.route("/gerar-pdf", methods=["POST"])
def gerar_pdf():
    dados = {
        "{{CLIENTE}}": request.form.get("cliente", ""),
        "{{CPF}}": request.form.get("cpf", ""),
        "{{MODELO}}": request.form.get("modelo", ""),
        "{{FRANQUIA}}": request.form.get("franquia", ""),
        "{{CONTRATO}}": request.form.get("contrato", ""),
        "{{EXCEDENTE}}": request.form.get("excedente", ""),
        "{{VALOR}}": request.form.get("valor", ""),
        "{{VALIDADE}}": request.form.get("validade", ""),
        "{{DATA}}": datetime.now().strftime("%d/%m/%Y"),
    }

    doc = Document(TEMPLATE_PATH)

    # Substitui todos os textos
    for p in doc.paragraphs:
        for chave, valor in dados.items():
            if chave in p.text:
                p.text = p.text.replace(chave, valor)

    # Caminhos de saída
    docx_path = os.path.join(BASE_DIR, "proposta.docx")
    pdf_path = os.path.join(BASE_DIR, "proposta.pdf")

    # Salva DOCX temporário
    doc.save(docx_path)

    # Converte para PDF
    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        docx_path,
        "--outdir", BASE_DIR
    ])

    return send_file(pdf_path, as_attachment=True)


# Necessário para Railway
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
