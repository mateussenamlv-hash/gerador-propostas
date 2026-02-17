from flask import Flask, render_template, request, send_file
from docx import Document
from datetime import datetime
import os

app = Flask(__name__)

# Caminho correto do projeto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "template.docx")

# Valores fixos (não aparecem no formulário)
CONTRATO_FIXO = "36 meses"
EXCEDENTE_FIXO = "R$ 0,10 por página"


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/gerar-pdf", methods=["POST"])
def gerar_pdf():
    # Pega dados do formulário (sem contrato e excedente)
    dados = {
        "cliente": request.form.get("cliente", ""),
        "cpf": request.form.get("cpf", ""),
        "modelo": request.form.get("modelo", ""),
        "franquia": request.form.get("franquia", ""),
        "contrato": CONTRATO_FIXO,
        "excedente": EXCEDENTE_FIXO,
        "valor": request.form.get("valor", ""),
        "validade": request.form.get("validade", ""),
        "data": datetime.now().strftime("%d/%m/%Y"),
    }

    # Abre o template
    doc = Document(TEMPLATE_PATH)

    # Substitui placeholders ({{ CLIENTE }}, etc.)
    for p in doc.paragraphs:
        for chave, valor in dados.items():
            p.text = p.text.replace(f"{{{{ {chave.upper()} }}}}", str(valor))

    # Salva arquivo final
    output_path = os.path.join(BASE_DIR, "proposta.docx")
    doc.save(output_path)

    # Retorna o arquivo gerado (mantendo seu comportamento atual)
    return send_file(output_path, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
