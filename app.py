from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import os
import subprocess
from datetime import datetime
from decimal import Decimal
import psycopg2
from psycopg2.extras import RealDictCursor

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_PROPOSTA_PATH = os.path.join(BASE_DIR, "template.docx")
TEMPLATE_CONTRATO_PATH = os.path.join(BASE_DIR, "contrato_template.docx")

# =========================
# BANCO (Postgres Railway)
# =========================

DATABASE_URL = os.environ.get("DATABASE_URL")


def db_conn():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)


def init_db():
    # NÃO cria tabela porque você já criou manualmente
    pass


def cleanup_old_proposals(days=15):
    conn = db_conn()
    cur = conn.cursor()
    cur.execute(f"""
        DELETE FROM propostas
        WHERE created_at < NOW() - INTERVAL '{days} days'
    """)
    conn.commit()
    conn.close()


def save_proposta(cliente, cpf, modelo, franquia, valor):
    conn = db_conn()
    cur = conn.cursor()

    # ⚠️ NÃO inserimos ID porque ele é SERIAL no banco
    cur.execute("""
        INSERT INTO propostas
        (created_at, status, cliente, cpf, modelo, franquia, valor)
        VALUES (NOW(), %s, %s, %s, %s, %s, %s)
    """, (
        "pendente",
        cliente,
        cpf,
        modelo,
        franquia,
        valor
    ))

    conn.commit()
    conn.close()


def get_recent_proposals(limit=50):
    conn = db_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT * FROM propostas
        ORDER BY created_at DESC
        LIMIT %s
    """, (limit,))
    rows = cur.fetchall()
    conn.close()
    return rows


def get_proposta_by_id(proposal_id):
    conn = db_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM propostas WHERE id = %s", (proposal_id,))
    row = cur.fetchone()
    conn.close()
    return row


# =========================
# AUXILIARES
# =========================

def parse_money(valor_str: str) -> Decimal:
    s = (valor_str or "").strip()
    s = s.replace("R$", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "")
        s = s.replace(",", ".")
    return Decimal(s)


def format_money_ptbr(valor: Decimal) -> str:
    valor = valor.quantize(Decimal("0.01"))
    return f"{valor:.2f}".replace(".", ",")


def hoje_por_extenso():
    meses = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
             "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
    agora = datetime.now()
    return f"{agora.day} de {meses[agora.month-1]} de {agora.year}"


# =========================
# ROTAS
# =========================

@app.route("/")
def home():
    return render_template("index.html")


@app.route("/proposta")
def proposta_form():
    return render_template("proposta.html")


@app.route("/gerar-pdf", methods=["POST"])
def gerar_pdf():
    try:
        cleanup_old_proposals()

        doc = DocxTemplate(TEMPLATE_PROPOSTA_PATH)

        cliente = request.form.get("cliente")
        cpf = request.form.get("cpf")
        modelo = request.form.get("modelo")
        franquia = request.form.get("franquia")
        valor_input = request.form.get("valor")

        valor_dec = parse_money(valor_input)
        valor_formatado = format_money_ptbr(valor_dec)

        # salva no banco
        save_proposta(cliente, cpf, modelo, franquia, valor_formatado)

        imagem = request.files.get("imagem")
        imagem_template = InlineImage(doc, imagem, height=Mm(45)) if imagem and imagem.filename else ""

        context = {
            "CLIENTE": cliente,
            "CPF": cpf,
            "MODELO": modelo,
            "FRANQUIA": franquia,
            "VALOR": valor_formatado,
            "DATA": hoje_por_extenso(),
            "IMAGEM": imagem_template,
        }

        doc.render(context)

        unique_name = f"proposta_{datetime.now().timestamp()}"
        docx_path = os.path.join(BASE_DIR, f"{unique_name}.docx")
        doc.save(docx_path)

        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", BASE_DIR],
            check=True,
        )

        pdf_path = docx_path.replace(".docx", ".pdf")
        nome_final = f"Proposta_{(cliente or 'Cliente').replace(' ','_')}.pdf"

        return send_file(pdf_path, as_attachment=True, download_name=nome_final)

    except Exception as e:
        return f"Erro: {str(e)}"


@app.route("/propostas-recentes")
def propostas_recentes():
    cleanup_old_proposals()
    rows = get_recent_proposals()
    return render_template("propostas_recentes.html", propostas=rows)


@app.route("/contrato")
def contrato_form():
    proposal_id = request.args.get("from")
    prefill = get_proposta_by_id(proposal_id) if proposal_id else None
    return render_template("contrato.html", prefill=prefill)


@app.route("/gerar-contrato", methods=["POST"])
def gerar_contrato():
    try:
        doc = DocxTemplate(TEMPLATE_CONTRATO_PATH)

        context = {key.upper(): value for key, value in request.form.items()}
        context["DATA_ASSINATURA"] = hoje_por_extenso()

        doc.render(context)

        unique_name = f"contrato_{datetime.now().timestamp()}"
        docx_path = os.path.join(BASE_DIR, f"{unique_name}.docx")
        doc.save(docx_path)

        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", BASE_DIR],
            check=True,
        )

        pdf_path = docx_path.replace(".docx", ".pdf")
        nome_final = f"Contrato_{context.get('DENOMINACAO','Cliente').replace(' ','_')}.pdf"

        return send_file(pdf_path, as_attachment=True, download_name=nome_final)

    except Exception as e:
        return f"Erro: {str(e)}"


# =========================
# STARTUP
# =========================

init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
