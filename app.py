from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import os
import subprocess
from datetime import datetime
import uuid

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.docx")


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/gerar-pdf", methods=["POST"])
def gerar_pdf():
    try:
        doc = DocxTemplate(TEMPLATE_PATH)

        # ===== PEGAR DADOS DO FORM =====
        cliente = request.form.get("cliente")
        cpf = request.form.get("cpf")
        modelo = request.form.get("modelo")
        franquia = request.form.get("franquia")
        contrato = request.form.get("contrato")
        excedente = request.form.get("excedente")
        valor = request.form.get("valor")
        validade = request.form.get("validade")

        data_atual = datetime.now().strftime("%d/%m/%Y")

        # ===== IMAGEM OPCIONAL =====
        imagem = request.files.get("imagem")

        if imagem and imagem.filename != "":
            imagem_template = InlineImage(doc, imagem, width=Mm(50))
        else:
            imagem_template = ""

        # ===== CONTEXTO PARA O TEMPLATE =====
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

        # ===== SALVAR DOCX TEMPORÁRIO =====
        unique_id = str(uuid.uuid4())
        docx_path = os.path.join(BASE_DIR, f"{unique_id}.docx")
        doc.save(docx_path)

        # ===== CONVERTER PARA PDF (LIBREOFFICE) =====
        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                docx_path,
                "--outdir", BASE_DIR,
            ],
            check=True,
        )

        pdf_path = docx_path.replace(".docx", ".pdf")

        # ===== NOME PERSONALIZADO DO PDF =====
        nome_cliente = cliente.replace(" ", "_")
        nome_final = f"Proposta_{nome_cliente}.pdf"

        return send_file(
            pdf_path,
            as_attachment=True,
            download_name=nome_final
        )

    except Exception as e:
        return f"Erro interno: {str(e)}"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
