from flask import Flask, render_template, request, send_file
from docx import Document
from datetime import datetime
import os

app = Flask(__name__)

# Caminho correto do projeto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "template.docx")

# Página inicial (teste)
@app.route("/")
def home():
    return render_template("index.html")


@app.route("/gerar-pdf", methods=["POST"])
def gerar_pdf():
    dados = {
        "cliente": request.form["cliente"],
        "cpf": request.form["cpf"],
        "modelo": request.form["modelo"],
        "franquia": request.form["franquia"],
        "contrato": request.form["contrato"],
        "excedente": request.form["excedente"],
        "valor": request.form["valor"],
        "validade": request.form["validade"],
        "data": datetime.now().strftime("%d/%m/%Y")
    }

    doc = Document(TEMPLATE_PATH)

    for p in doc.paragraphs:
        for chave, valor in dados.items():
            p.text = p.text.replace(f"{{{{{chave.upper()}}}}}", valor)

    output_path = os.path.join(BASE_DIR, "proposta.docx")
    doc.save(output_path)

    return send_file(output_path, as_attachment=True)


# ESSA PARTE É O QUE FAZ FUNCIONAR NO RAILWAY
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
