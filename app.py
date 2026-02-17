from flask import Flask, render_template, request, send_file
from docx import Document
from datetime import datetime
import os

app = Flask(__name__)

TEMPLATE = "template.docx"
OUTPUT = "proposta.pdf"

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        doc = Document(TEMPLATE)

        dados = {
            "{{CLIENTE}}": request.form["cliente"],
            "{{CPF}}": request.form["cpf"],
            "{{MODELO}}": request.form["modelo"],
            "{{FRANQUIA}}": request.form["franquia"],
            "{{CONTRATO}}": request.form["contrato"],
            "{{EXCEDENTE}}": request.form["excedente"],
            "{{VALOR}}": request.form["valor"],
            "{{VALIDADE}}": request.form["validade"],
            "{{DATA}}": datetime.now().strftime("%d/%m/%Y"),
        }

        for p in doc.paragraphs:
            for chave, valor in dados.items():
                if chave in p.text:
                    p.text = p.text.replace(chave, valor)

        doc.save("temp.docx")

        os.system("soffice --headless --convert-to pdf temp.docx --outdir .")

        return send_file("temp.pdf", as_attachment=True)

    return render_template("index.html")

app.run(host="0.0.0.0", port=8080)
