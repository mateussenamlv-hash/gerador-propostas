from flask import Flask, render_template, request, send_file
from docx import Document
from datetime import datetime
import os

app = Flask(__name__)

# Caminho correto do projeto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "templet.doc")

# Página inicial (teste)
@app.route("/")
def home():
    return "Servidor online"

# Exemplo de rota para gerar documento
@app.route("/gerar", methods=["POST"])
def gerar_documento():
    nome = request.form.get("nome", "Cliente")

    doc = Document(TEMPLATE_PATH)

    for p in doc.paragraphs:
        if "{{NOME}}" in p.text:
            p.text = p.text.replace("{{NOME}}", nome)

    output_path = os.path.join(BASE_DIR, "proposta.docx")
    doc.save(output_path)

    return send_file(output_path, as_attachment=True)

# ESSA PARTE É O QUE FAZ FUNCIONAR NO RAILWAY
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
