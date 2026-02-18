from flask import Flask, render_template, request, send_file, redirect, url_for
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import os
import subprocess
from datetime import datetime, timedelta, timezone
import uuid
from decimal import Decimal, InvalidOperation

from sqlalchemy import create_engine, Column, Integer, String, DateTime, Text
from sqlalchemy.orm import sessionmaker, declarative_base

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_PROPOSTA_PATH = os.path.join(BASE_DIR, "template.docx")
TEMPLATE_CONTRATO_PATH = os.path.join(BASE_DIR, "contrato_template.docx")

# ==========================================================
# BANCO DE DADOS (PostgreSQL na Railway)
# - Railway geralmente fornece DATABASE_URL
# - Para dev local sem Postgres, cai em SQLite
# ==========================================================
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()

if DATABASE_URL:
    # Railway costuma vir com postgres:// (precisa ser postgresql:// pro SQLAlchemy)
    if DATABASE_URL.startswith("postgres://"):
        DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)
else:
    DATABASE_URL = "sqlite:///" + os.path.join(BASE_DIR, "app.db")

engine = create_engine(DATABASE_URL, pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
Base = declarative_base()


class Proposta(Base):
    __tablename__ = "propostas"

    id = Column(Integer, primary_key=True)
    created_at = Column(DateTime, nullable=False, index=True)
    status = Column(String(30), nullable=False, default="pendente")

    cliente = Column(String(200), nullable=False)
    cpf = Column(String(50), nullable=False)
    modelo = Column(String(200), nullable=False)
    franquia = Column(String(100), nullable=False)
    valor = Column(String(120), nullable=False)     # já formatado "250,00 (duzentos...)"
    validade = Column(String(120), nullable=False)

    # opcional, se quiser usar no futuro
    observacao = Column(Text, nullable=True)


class Contrato(Base):
    __tablename__ = "contratos"

    id = Column(Integer, primary_key=True)
    created_at = Column(DateTime, nullable=False, index=True)
    proposta_id = Column(Integer, nullable=True, index=True)

    denominacao = Column(String(200), nullable=False)
    cpf_cnpj = Column(String(80), nullable=False)
    endereco = Column(String(250), nullable=False)
    telefone = Column(String(80), nullable=False)
    email = Column(String(200), nullable=False)

    equipamento = Column(String(250), nullable=False)
    acessorios = Column(Text, nullable=False)

    data_inicio = Column(String(60), nullable=False)     # por extenso
    data_termino = Column(String(60), nullable=False)    # por extenso

    franquia_formatada = Column(String(60), nullable=False)
    franquia_extenso = Column(String(120), nullable=False)

    valor_mensal_formatado = Column(String(60), nullable=False)
    valor_mensal_extenso = Column(String(180), nullable=False)

    data_assinatura = Column(String(60), nullable=False)


Base.metadata.create_all(engine)

# ==========================================================
# FUNÇÕES AUXILIARES
# ==========================================================
def agora_utc() -> datetime:
    return datetime.now(timezone.utc)


def limpar_propostas_antigas(db):
    """Apaga propostas com mais de 15 dias."""
    limite = agora_utc() - timedelta(days=15)
    db.query(Proposta).filter(Proposta.created_at < limite).delete()
    db.commit()


def data_por_extenso(data_str: str) -> str:
    """Converte 'DD/MM/AAAA' -> '11 de Fevereiro de 2026'."""
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    dt = datetime.strptime(data_str.strip(), "%d/%m/%Y")
    return f"{dt.day} de {meses[dt.month - 1]} de {dt.year}"


def numero_formatado_ptbr(n: int) -> str:
    """1000 -> '1.000'"""
    return f"{n:,}".replace(",", ".")


def numero_por_extenso_pt(n: int) -> str:
    """Inteiro por extenso (pt-BR) até bilhões."""
    if n == 0:
        return "zero"

    unidades = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"]
    dez_a_dezenove = ["dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"]
    dezenas = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"]
    centenas = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"]

    def ext_ate_999(x: int) -> str:
        if x == 0:
            return ""
        if x == 100:
            return "cem"
        c = x // 100
        r = x % 100
        parts = []
        if c:
            parts.append(centenas[c])
        if r:
            if r < 10:
                parts.append(unidades[r])
            elif 10 <= r < 20:
                parts.append(dez_a_dezenove[r - 10])
            else:
                d = r // 10
                u = r % 10
                if u:
                    parts.append(f"{dezenas[d]} e {unidades[u]}")
                else:
                    parts.append(dezenas[d])
        if len(parts) == 2 and (" e " not in parts[1]):
            return f"{parts[0]} e {parts[1]}"
        return " ".join(parts)

    def ext_grupo(valor: int, singular: str, plural: str) -> str:
        if valor == 0:
            return ""
        if valor == 1:
            return f"um {singular}"
        return f"{numero_por_extenso_pt(valor)} {plural}"

    bilhoes = n // 1_000_000_000
    n = n % 1_000_000_000
    milhoes = n // 1_000_000
    n = n % 1_000_000
    milhares = n // 1000
    resto = n % 1000

    partes = []
    if bilhoes:
        partes.append(ext_grupo(bilhoes, "bilhão", "bilhões"))
    if milhoes:
        partes.append(ext_grupo(milhoes, "milhão", "milhões"))
    if milhares:
        if milhares == 1:
            partes.append("mil")
        else:
            partes.append(f"{ext_ate_999(milhares)} mil")
    if resto:
        partes.append(ext_ate_999(resto))

    resultado = ""
    for i, p in enumerate(partes):
        if resultado == "":
            resultado = p
        else:
            if i == len(partes) - 1:
                if resto and resto < 100:
                    resultado = f"{resultado} e {p}"
                else:
                    resultado = f"{resultado} {p}"
            else:
                resultado = f"{resultado} {p}"
    return resultado.strip()


def parse_money(valor_str: str) -> Decimal:
    """Aceita '250', '250,00', '250.00', 'R$ 250,00'."""
    s = (valor_str or "").strip()
    s = s.replace("R$", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "")
        s = s.replace(",", ".")
    return Decimal(s)


def format_money_ptbr(valor: Decimal) -> str:
    """200.00 -> '200,00'"""
    valor = valor.quantize(Decimal("0.01"))
    s = f"{valor:.2f}"
    return s.replace(".", ",")


def dinheiro_por_extenso(valor: Decimal) -> str:
    """Ex: 200.00 -> 'duzentos reais' ; 200.50 -> 'duzentos reais e cinquenta centavos'"""
    valor = valor.quantize(Decimal("0.01"))
    reais = int(valor)
    centavos = int((valor - Decimal(reais)) * 100)

    texto_reais = numero_por_extenso_pt(reais)
    moeda = "real" if reais == 1 else "reais"

    if centavos == 0:
        return f"{texto_reais} {moeda}"

    texto_cent = numero_por_extenso_pt(centavos)
    cent = "centavo" if centavos == 1 else "centavos"
    return f"{texto_reais} {moeda} e {texto_cent} {cent}"


def hoje_por_extenso() -> str:
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    # usa data local do servidor (Railway) — normalmente ok
    agora = datetime.now()
    return f"{agora.day} de {meses[agora.month - 1]} de {agora.year}"


# ==========================================================
# ROTAS
# ==========================================================
@app.route("/")
def home():
    db = SessionLocal()
    try:
        limpar_propostas_antigas(db)
    finally:
        db.close()
    return render_template("index.html")


@app.route("/propostas-recentes")
def propostas_recentes():
    db = SessionLocal()
    try:
        limpar_propostas_antigas(db)
        limite = agora_utc() - timedelta(days=15)
        propostas = (
            db.query(Proposta)
            .filter(Proposta.created_at >= limite)
            .order_by(Proposta.created_at.desc())
            .all()
        )
        return render_template("propostas_recentes.html", propostas=propostas)
    finally:
        db.close()


@app.route("/propostas-recentes/excluir/<int:proposta_id>", methods=["POST"])
def excluir_proposta(proposta_id: int):
    db = SessionLocal()
    try:
        p = db.query(Proposta).filter(Proposta.id == proposta_id).first()
        if p:
            db.delete(p)
            db.commit()
        return redirect(url_for("propostas_recentes"))
    finally:
        db.close()


# --------- PROPOSTA ----------
@app.route("/proposta")
def proposta_form():
    return render_template("proposta.html")


@app.route("/gerar-pdf", methods=["POST"])
def gerar_pdf():
    db = SessionLocal()
    try:
        limpar_propostas_antigas(db)

        doc = DocxTemplate(TEMPLATE_PROPOSTA_PATH)

        cliente = request.form.get("cliente")
        cpf = request.form.get("cpf")
        modelo = request.form.get("modelo")
        franquia = request.form.get("franquia")
        valor_input = request.form.get("valor")
        validade = request.form.get("validade")

        data_atual = hoje_por_extenso()

        # VALOR FORMATADO + EXTENSO
        valor_dec = parse_money(valor_input)
        valor_formatado = format_money_ptbr(valor_dec)

        valor_reais_int = int(valor_dec.quantize(Decimal("0.01")))
        centavos = int((valor_dec.quantize(Decimal("0.01")) - Decimal(valor_reais_int)) * 100)

        if centavos == 0:
            valor_extenso = numero_por_extenso_pt(valor_reais_int)
        else:
            valor_extenso = dinheiro_por_extenso(valor_dec)

        valor_final = f"{valor_formatado} ({valor_extenso})"

        # IMAGEM (mantém upload exatamente como estava, só reduz tamanho)
        imagem = request.files.get("imagem")
        if imagem and imagem.filename != "":
            imagem_template = InlineImage(doc, imagem, height=Mm(40))  # menor pra evitar 2 páginas
        else:
            imagem_template = ""

        context = {
            "CLIENTE": cliente,
            "CPF": cpf,
            "MODELO": modelo,
            "FRANQUIA": franquia,
            "VALOR": valor_final,
            "VALIDADE": validade,
            "DATA": data_atual,
            "IMAGEM": imagem_template,
        }

        doc.render(context)

        unique_id = str(uuid.uuid4())
        docx_path = os.path.join(BASE_DIR, f"{unique_id}.docx")
        doc.save(docx_path)

        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", BASE_DIR],
            check=True,
        )

        pdf_path = docx_path.replace(".docx", ".pdf")

        # SALVA NO BANCO (15 dias)
        proposta = Proposta(
            created_at=agora_utc(),
            status="pendente",
            cliente=cliente or "",
            cpf=cpf or "",
            modelo=modelo or "",
            franquia=franquia or "",
            valor=valor_final,
            validade=validade or "",
        )
        db.add(proposta)
        db.commit()

        nome_cliente = (cliente or "Cliente").replace(" ", "_")
        nome_final = f"Proposta_{nome_cliente}.pdf"

        return send_file(pdf_path, as_attachment=True, download_name=nome_final)

    except (InvalidOperation, ValueError):
        return "Erro: confira o campo VALOR (ex: 250 ou 250,00)."
    except Exception as e:
        return f"Erro interno: {str(e)}"
    finally:
        db.close()


# --------- CONTRATO ----------
@app.route("/contrato")
def contrato_form():
    proposta_id = request.args.get("proposta_id")
    prefill = {}

    if proposta_id:
        db = SessionLocal()
        try:
            limpar_propostas_antigas(db)
            p = db.query(Proposta).filter(Proposta.id == int(proposta_id)).first()
            if p:
                # reaproveita o que existe na proposta
                prefill = {
                    "denominacao": p.cliente,
                    "cpf_cnpj": p.cpf,
                    "equipamento": p.modelo,
                    "franquia_total": p.franquia,     # se você quiser manter só número, ajuste no formulário
                    "valor_mensal": p.valor.split(" ")[0].strip() if p.valor else "",  # pega "250,00"
                    "proposta_id": str(p.id),
                }
        finally:
            db.close()

    return render_template("contrato.html", prefill=prefill)


@app.route("/gerar-contrato", methods=["POST"])
def gerar_contrato():
    db = SessionLocal()
    try:
        limpar_propostas_antigas(db)

        doc = DocxTemplate(TEMPLATE_CONTRATO_PATH)

        proposta_id = request.form.get("proposta_id")  # pode vir vazio

        denominacao = request.form.get("denominacao")
        cpf_cnpj = request.form.get("cpf_cnpj")
        endereco = request.form.get("endereco")
        telefone = request.form.get("telefone")
        email = request.form.get("email")

        equipamento = request.form.get("equipamento")
        acessorios = request.form.get("acessorios")

        data_inicio_input = request.form.get("data_inicio")
        data_termino_input = request.form.get("data_termino")

        franquia_input = request.form.get("franquia_total")   # ex: 1000
        valor_mensal_input = request.form.get("valor_mensal") # ex: 200

        data_inicio = data_por_extenso(data_inicio_input)
        data_termino = data_por_extenso(data_termino_input)

        # franquia: aceita "1000" ou "1.000" (limpa pontos)
        franquia_clean = str(franquia_input).replace(".", "").strip()
        franquia_int = int(franquia_clean)
        franquia_formatada = numero_formatado_ptbr(franquia_int)
        franquia_extenso = numero_por_extenso_pt(franquia_int)

        valor_mensal_dec = parse_money(valor_mensal_input)
        valor_mensal_formatado = format_money_ptbr(valor_mensal_dec)
        valor_mensal_extenso = dinheiro_por_extenso(valor_mensal_dec)

        data_assinatura = hoje_por_extenso()

        context = {
            "DENOMINACAO": denominacao,
            "CPF_CNPJ": cpf_cnpj,
            "ENDERECO": endereco,
            "TELEFONE": telefone,
            "EMAIL": email,
            "EQUIPAMENTO": equipamento,
            "ACESSORIOS": acessorios,
            "DATA_INICIO": data_inicio,
            "DATA_TERMINO": data_termino,
            "FRANQUIA_FORMATADA": franquia_formatada,
            "FRANQUIA_EXTENSO": franquia_extenso,
            "VALOR_MENSAL_FORMATADO": valor_mensal_formatado,
            "VALOR_MENSAL_EXTENSO": valor_mensal_extenso,
            "DATA_ASSINATURA": data_assinatura,
        }

        doc.render(context)

        unique_id = str(uuid.uuid4())
        docx_path = os.path.join(BASE_DIR, f"{unique_id}.docx")
        doc.save(docx_path)

        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", BASE_DIR],
            check=True,
        )

        pdf_path = docx_path.replace(".docx", ".pdf")

        # salva contrato no banco (opcional, mas útil)
        contrato = Contrato(
            created_at=agora_utc(),
            proposta_id=int(proposta_id) if proposta_id and proposta_id.isdigit() else None,
            denominacao=denominacao or "",
            cpf_cnpj=cpf_cnpj or "",
            endereco=endereco or "",
            telefone=telefone or "",
            email=email or "",
            equipamento=equipamento or "",
            acessorios=acessorios or "",
            data_inicio=data_inicio,
            data_termino=data_termino,
            franquia_formatada=franquia_formatada,
            franquia_extenso=franquia_extenso,
            valor_mensal_formatado=valor_mensal_formatado,
            valor_mensal_extenso=valor_mensal_extenso,
            data_assinatura=data_assinatura,
        )
        db.add(contrato)

        # marca proposta como "contrato_gerado"
        if proposta_id and proposta_id.isdigit():
            p = db.query(Proposta).filter(Proposta.id == int(proposta_id)).first()
            if p:
                p.status = "contrato_gerado"

        db.commit()

        nome_cliente = (denominacao or "Cliente").replace(" ", "_")
        nome_final = f"Contrato_{nome_cliente}.pdf"

        return send_file(pdf_path, as_attachment=True, download_name=nome_final)

    except (ValueError, InvalidOperation):
        return "Erro: confira os campos numéricos (franquia e valor mensal) e as datas (DD/MM/AAAA)."
    except Exception as e:
        return f"Erro interno: {str(e)}"
    finally:
        db.close()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
