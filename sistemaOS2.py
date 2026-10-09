from flask import (
    Flask,
    render_template_string,
    request,
    redirect,
    url_for,
    send_file,
    flash,
)
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
from io import BytesIO
from pathlib import Path
from threading import Lock
import pandas as pd

app = Flask(__name__)

app.config["SECRET_KEY"] = "troque-por-uma-chave-secreta"
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///os.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

FORMATO_DATA = "%d/%m/%Y %H:%M:%S"
PREFIXO_ANO_OS = "26"
NUMERO_OS_INICIAL = 1730

# Evita conflitos entre gerações simultâneas dentro
# deste processo do aplicativo.
os_lock = Lock()


# =====================================================
# BANCO DE DADOS
# =====================================================

class OrdemServico(db.Model):
    __tablename__ = "ordens_servico"

    id = db.Column(db.Integer, primary_key=True)

    numero_os = db.Column(db.String(20), unique=True, nullable=False)
    solicitante = db.Column(db.String(100), nullable=False)
    setor = db.Column(db.String(100), default="")
    titulo = db.Column(db.String(200), default="")
    descricao = db.Column(db.Text, default="")
    prioridade = db.Column(db.String(30), default="Média")
    responsavel = db.Column(db.String(100), default="")

    abertura = db.Column(db.String(50), nullable=False)
    finalizacao = db.Column(db.String(50), default="")

    status = db.Column(db.String(50), default="Aberta")
    tempo_total = db.Column(db.String(50), default="")


with app.app_context():
    db.create_all()


# =====================================================
# FUNÇÕES AUXILIARES
# =====================================================

def agora_formatado():
    return datetime.now().strftime(FORMATO_DATA)


def gerar_numero_os():
    """
    Gera o próximo número usando os registros existentes.
    Funciona também quando há OS antigas no banco.
    """
    ordens = OrdemServico.query.with_entities(
        OrdemServico.numero_os
    ).all()

    maior_numero = NUMERO_OS_INICIAL - 1

    for (numero_os,) in ordens:
        try:
            parte_numero = numero_os.split("-")[0]
            maior_numero = max(maior_numero, int(parte_numero))
        except (ValueError, AttributeError):
            continue

    return f"{maior_numero + 1}-{PREFIXO_ANO_OS}"


def calcular_tempo_total(abertura, finalizacao):
    """
    Calcula o tempo decorrido sem perder os dias completos.
    """
    inicio = datetime.strptime(abertura, FORMATO_DATA)
    fim = datetime.strptime(finalizacao, FORMATO_DATA)

    segundos = max(0, int((fim - inicio).total_seconds()))

    dias, resto = divmod(segundos, 86400)
    horas, resto = divmod(resto, 3600)
    minutos, _ = divmod(resto, 60)

    partes = []

    if dias:
        partes.append(f"{dias}d")

    if horas:
        partes.append(f"{horas}h")

    partes.append(f"{minutos}min")

    return " ".join(partes)


# =====================================================
# HTML DO SISTEMA
# =====================================================

HTML = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">

    <title>CENTRAL O.S JAMAICA</title>

    <style>
        * {
            box-sizing: border-box;
        }

        body {
            font-family: Arial, Helvetica, sans-serif;
            background: #f1f5f9;
            color: #1e293b;
            margin: 0;
            padding: 24px;
        }

        .container {
            max-width: 1600px;
            margin: auto;
        }

        .cabecalho {
            background: linear-gradient(135deg, #000000 0%, #7c2d12 55%, #f97316 100%);
            color: white;
            padding: 26px;
            border-radius: 14px;
            margin-bottom: 22px;
        }

        .cabecalho h1 {
            margin: 0 0 8px;
        }

        .cabecalho p {
            margin: 0;
            color: #dbeafe;
        }

        .resumo {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 16px;
            margin-bottom: 22px;
        }

        .resumo-card {
            background: white;
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 2px 8px #0f172a0d;
        }

        .resumo-card span {
            display: block;
            color: #64748b;
            font-size: 14px;
            margin-bottom: 8px;
        }

        .resumo-card strong {
            font-size: 27px;
        }

        .card {
            background: white;
            padding: 24px;
            border-radius: 12px;
            margin-bottom: 22px;
            box-shadow: 0 2px 8px #0f172a0d;
        }

        .card h2 {
            margin-top: 0;
            font-size: 21px;
        }

        .form-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 14px 18px;
        }

        .campo {
            min-width: 0;
        }

        .campo.completo {
            grid-column: 1 / -1;
        }

        label {
            display: block;
            font-weight: bold;
            font-size: 14px;
            margin-bottom: 7px;
        }

        input, textarea, select {
            display: block;
            width: 100%;
            padding: 11px 12px;
            border: 1px solid #cbd5e1;
            border-radius: 7px;
            font: inherit;
            background: white;
            color: #0f172a;
        }

        input:focus, textarea:focus, select:focus {
            outline: 2px solid #93c5fd;
            border-color: #2563eb;
        }

        textarea {
            min-height: 95px;
            resize: vertical;
        }

        button, .botao {
            display: inline-block;
            padding: 10px 15px;
            border: none;
            border-radius: 7px;
            font-size: 14px;
            font-weight: bold;
            text-decoration: none;
            text-align: center;
            cursor: pointer;
            transition: opacity .15s;
        }

        button:hover, .botao:hover {
            opacity: .86;
        }

        .salvar {
            background: #15803d;
            color: white;
        }

        .excel {
            background: #166534;
            color: white;
        }

        .pesquisar {
            background: #1d4ed8;
            color: white;
        }

        .limpar {
            background: #e2e8f0;
            color: #334155;
        }

        .finalizar {
            background: #2563eb;
            color: white;
        }

        .excluir {
            background: #dc2626;
            color: white;
        }

        .acoes-topo {
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 14px;
        }

        .busca {
            display: flex;
            gap: 9px;
            flex: 1;
            min-width: 260px;
        }

        .busca input {
            min-width: 0;
            flex: 1;
        }

        .tabela-wrapper {
            width: 100%;
            overflow-x: auto;
        }

        table {
            width: 100%;
            min-width: 1050px;
            border-collapse: collapse;
            background: white;
        }

        th, td {
            padding: 12px 10px;
            border-bottom: 1px solid #e2e8f0;
            text-align: left;
            font-size: 13px;
            vertical-align: middle;
        }

        th {
            background: #f8fafc;
            color: #475569;
            white-space: nowrap;
        }

        tbody tr:hover {
            background: #f8fafc;
        }

        .status {
            display: inline-block;
            padding: 5px 9px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: bold;
            white-space: nowrap;
        }

        .aberta {
            background: #fef3c7;
            color: #92400e;
        }

        .finalizada {
            background: #dcfce7;
            color: #166534;
        }

        .prioridade {
            font-weight: bold;
        }

        .prioridade-critica {
            color: #b91c1c;
        }

        .prioridade-alta {
            color: #c2410c;
        }

        .prioridade-media {
            color: #a16207;
        }

        .prioridade-baixa {
            color: #15803d;
        }

        .acoes {
            display: flex;
            gap: 6px;
            flex-wrap: wrap;
        }

        .acoes form {
            margin: 0;
        }

        .acoes button {
            padding: 7px 9px;
            font-size: 12px;
        }

        .mensagem {
            padding: 13px 16px;
            margin-bottom: 18px;
            border-radius: 8px;
            background: #dbeafe;
            color: #1e40af;
        }

        .vazio {
            text-align: center;
            padding: 30px;
            color: #64748b;
        }

        .rodape {
            text-align: center;
            color: #64748b;
            font-size: 12px;
            padding: 10px;
        }

        @media (max-width: 800px) {
            body {
                padding: 12px;
            }

            .resumo {
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }

            .form-grid {
                grid-template-columns: 1fr;
            }

            .campo.completo {
                grid-column: auto;
            }

            .card {
                padding: 16px;
            }

            .acoes-topo {
                align-items: stretch;
            }

            .busca {
                min-width: 100%;
            }
        }

        @media (max-width: 420px) {
            .resumo-card strong {
                font-size: 22px;
            }

            .busca {
                flex-wrap: wrap;
            }

            .busca input {
                flex-basis: 100%;
            }
        }
    </style>
</head>

<body>
<div class="container">

    <header class="cabecalho">
        <h1>CENTRAL DE ORDENS DE SERVIÇOS DE MANUTENÇÃO JAMAICA</h1>
        <p>Gerenciamento, acompanhamento e controle de O.S.</p>
    </header>

    {% with mensagens = get_flashed_messages() %}
        {% for mensagem in mensagens %}
            <div class="mensagem">{{ mensagem }}</div>
        {% endfor %}
    {% endwith %}

    <section class="resumo">
        <div class="resumo-card">
            <span>Total de O.S.</span>
            <strong>{{ total }}</strong>
        </div>

        <div class="resumo-card">
            <span>O.S. abertas</span>
            <strong>{{ abertas }}</strong>
        </div>

        <div class="resumo-card">
            <span>O.S. finalizadas</span>
            <strong>{{ finalizadas }}</strong>
        </div>

        <div class="resumo-card">
            <span>Próxima O.S.</span>
            <strong style="font-size: 23px">{{ proxima }}</strong>
        </div>
    </section>

    <section class="card">
        <h2>Abrir nova Ordem de Serviço</h2>

        <form action="{{ url_for('criar') }}" method="POST">
            <div class="form-grid">

                <div class="campo">
                    <label>Número da O.S.</label>
                    <input value="{{ proxima }}" readonly>
                </div>

                <div class="campo">
                    <label>Solicitante *</label>
                    <input
                        name="solicitante"
                        maxlength="100"
                        required
                        placeholder="Nome do solicitante">
                </div>

                <div class="campo">
                    <label>Setor</label>
                    <input
                        name="setor"
                        maxlength="100"
                        placeholder="Setor solicitante">
                </div>

                <div class="campo">
                    <label>Título</label>
                    <input
                        name="titulo"
                        maxlength="200"
                        placeholder="Resumo do problema">
                </div>

                <div class="campo">
                    <label>Prioridade</label>
                    <select name="prioridade">
                        <option>Baixa</option>
                        <option>Média</option>
                        <option>Alta</option>
                        <option>Crítica</option>
                    </select>
                </div>

                <div class="campo">
                    <label>Responsável</label>
                    <input
                        name="responsavel"
                        maxlength="100"
                        placeholder="Responsável pelo atendimento">
                </div>

                <div class="campo completo">
                    <label>Descrição</label>
                    <textarea
                        name="descricao"
                        placeholder="Descreva o serviço solicitado"></textarea>
                </div>

                <div class="campo completo">
                    <button class="salvar" type="submit">
                        + Abrir O.S.
                    </button>
                </div>

            </div>
        </form>
    </section>

    <section class="card">
        <div class="acoes-topo">
            <form class="busca" action="{{ url_for('dashboard') }}" method="GET">
                <input
                    name="busca"
                    value="{{ busca }}"
                    placeholder="Pesquisar por número, solicitante, setor ou título">

                <button class="pesquisar" type="submit">Pesquisar</button>

                <a class="botao limpar" href="{{ url_for('dashboard') }}">
                    Limpar
                </a>
            </form>

            <a class="botao excel" href="{{ url_for('excel') }}">
                Exportar Excel
            </a>
        </div>
    </section>

    <section class="card">
        <h2>Ordens de Serviço registradas</h2>

        <div class="tabela-wrapper">
            <table>
                <thead>
                    <tr>
                        <th>OS</th>
                        <th>Solicitante</th>
                        <th>Setor</th>
                        <th>Título</th>
                        <th>Prioridade</th>
                        <th>Responsável</th>
                        <th>Status</th>
                        <th>Abertura</th>
                        <th>Finalização</th>
                        <th>Tempo total</th>
                        <th>Ações</th>
                    </tr>
                </thead>

                <tbody>
                {% for os in ordens %}
                    <tr>
                        <td><strong>{{ os.numero_os }}</strong></td>
                        <td>{{ os.solicitante }}</td>
                        <td>{{ os.setor or '-' }}</td>
                        <td>{{ os.titulo or '-' }}</td>

                        <td>
                            <span class="prioridade prioridade-{{ os.prioridade|lower|replace('í', 'i')|replace('é', 'e') }}">
                                {{ os.prioridade }}
                            </span>
                        </td>

                        <td>{{ os.responsavel or '-' }}</td>

                        <td>
                            {% if os.status == 'Finalizada' %}
                                <span class="status finalizada">Finalizada</span>
                            {% else %}
                                <span class="status aberta">Aberta</span>
                            {% endif %}
                        </td>

                        <td>{{ os.abertura }}</td>
                        <td>{{ os.finalizacao or '-' }}</td>
                        <td>{{ os.tempo_total or '-' }}</td>

                        <td>
                            <div class="acoes">
                                {% if os.status != 'Finalizada' %}
                                    <form
                                        action="{{ url_for('finalizar', id=os.id) }}"
                                        method="POST"
                                        onsubmit="return confirm('Deseja finalizar esta O.S.?')">

                                        <button class="finalizar" type="submit">
                                            Finalizar
                                        </button>
                                    </form>
                                {% endif %}

                                <form
                                    action="{{ url_for('excluir', id=os.id) }}"
                                    method="POST"
                                    onsubmit="return confirm('Tem certeza que deseja excluir a O.S. {{ os.numero_os }}? Esta ação não pode ser desfeita.')">

                                    <button class="excluir" type="submit">
                                        Excluir
                                    </button>
                                </form>
                            </div>
                        </td>
                    </tr>
                {% else %}
                    <tr>
                        <td colspan="11" class="vazio">
                            Nenhuma Ordem de Serviço encontrada.
                        </td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
    </section>

    <footer class="rodape">
        Central de O.S. — Sistema de controle de serviços
    </footer>

</div>
</body>
</html>
"""


# =====================================================
# DASHBOARD
# =====================================================

@app.route("/")
def dashboard():
    busca = request.args.get("busca", "").strip()

    consulta = OrdemServico.query

    if busca:
        termo = f"%{busca}%"

        consulta = consulta.filter(
            db.or_(
                OrdemServico.numero_os.ilike(termo),
                OrdemServico.solicitante.ilike(termo),
                OrdemServico.setor.ilike(termo),
                OrdemServico.titulo.ilike(termo),
            )
        )

    ordens = consulta.order_by(
        OrdemServico.id.desc()
    ).all()

    abertas = OrdemServico.query.filter_by(
        status="Aberta"
    ).count()

    finalizadas = OrdemServico.query.filter_by(
        status="Finalizada"
    ).count()

    total = OrdemServico.query.count()

    return render_template_string(
        HTML,
        ordens=ordens,
        abertas=abertas,
        finalizadas=finalizadas,
        total=total,
        proxima=gerar_numero_os(),
        busca=busca,
    )


# =====================================================
# CRIAR O.S.
# =====================================================

@app.route("/criar", methods=["POST"])
def criar():
    solicitante = request.form.get(
        "solicitante", ""
    ).strip()

    if not solicitante:
        flash("Informe o nome do solicitante.")
        return redirect(url_for("dashboard"))

    prioridades_validas = ["Baixa", "Média", "Alta", "Crítica"]

    prioridade = request.form.get("prioridade", "Média")

    if prioridade not in prioridades_validas:
        prioridade = "Média"

    try:
        with os_lock:
            nova = OrdemServico(
                numero_os=gerar_numero_os(),
                solicitante=solicitante,
                setor=request.form.get("setor", "").strip(),
                titulo=request.form.get("titulo", "").strip(),
                descricao=request.form.get("descricao", "").strip(),
                prioridade=prioridade,
                responsavel=request.form.get("responsavel", "").strip(),
                abertura=agora_formatado(),
                finalizacao="",
                status="Aberta",
                tempo_total="",
            )

            db.session.add(nova)
            db.session.commit()

        flash(f"O.S. {nova.numero_os} criada com sucesso!")

    except Exception:
        db.session.rollback()
        app.logger.exception("Erro ao criar ordem de serviço.")
        flash("Não foi possível criar a O.S. Verifique o erro no terminal.")

    return redirect(url_for("dashboard"))


# =====================================================
# FINALIZAR O.S.
# =====================================================

@app.route("/finalizar/<int:id>", methods=["POST"])
def finalizar(id):
    os = db.session.get(OrdemServico, id)

    if os is None:
        flash("Ordem de Serviço não encontrada.")
        return redirect(url_for("dashboard"))

    if os.status == "Finalizada":
        flash(f"A O.S. {os.numero_os} já está finalizada.")
        return redirect(url_for("dashboard"))

    try:
        finalizacao = agora_formatado()

        os.finalizacao = finalizacao
        os.status = "Finalizada"
        os.tempo_total = calcular_tempo_total(
            os.abertura,
            os.finalizacao,
        )

        db.session.commit()
        flash(f"O.S. {os.numero_os} finalizada com sucesso!")

    except Exception:
        db.session.rollback()
        app.logger.exception("Erro ao finalizar ordem de serviço.")
        flash("Não foi possível finalizar a O.S.")

    return redirect(url_for("dashboard"))


# =====================================================
# EXCLUIR O.S.
# =====================================================

@app.route("/excluir/<int:id>", methods=["POST"])
def excluir(id):
    os = db.session.get(OrdemServico, id)

    if os is None:
        flash("Ordem de Serviço não encontrada.")
        return redirect(url_for("dashboard"))

    numero = os.numero_os

    try:
        db.session.delete(os)
        db.session.commit()
        flash(f"O.S. {numero} excluída com sucesso!")

    except Exception:
        db.session.rollback()
        app.logger.exception("Erro ao excluir ordem de serviço.")
        flash("Não foi possível excluir a O.S.")

    return redirect(url_for("dashboard"))


# =====================================================
# EXPORTAR PARA EXCEL
# =====================================================

@app.route("/excel")
def excel():
    ordens = OrdemServico.query.order_by(
        OrdemServico.id.asc()
    ).all()

    colunas = [
        "OS",
        "Solicitante",
        "Setor",
        "Título",
        "Descrição",
        "Prioridade",
        "Responsável",
        "Status",
        "Abertura",
        "Finalização",
        "Tempo total",
    ]

    dados = []

    for os in ordens:
        dados.append({
            "OS": os.numero_os,
            "Solicitante": os.solicitante,
            "Setor": os.setor or "",
            "Título": os.titulo or "",
            "Descrição": os.descricao or "",
            "Prioridade": os.prioridade,
            "Responsável": os.responsavel or "",
            "Status": os.status,
            "Abertura": os.abertura,
            "Finalização": os.finalizacao or "",
            "Tempo total": os.tempo_total or "",
        })

    df = pd.DataFrame(dados, columns=colunas)

    memoria = BytesIO()

    try:
        with pd.ExcelWriter(
            memoria,
            engine="openpyxl",
        ) as writer:
            df.to_excel(
                writer,
                index=False,
                sheet_name="Ordens de Serviço",
            )

            planilha = writer.sheets["Ordens de Serviço"]

            # Congela o cabeçalho.
            planilha.freeze_panes = "A2"

            # Ativa os filtros.
            planilha.auto_filter.ref = planilha.dimensions

            # Ajusta a largura das colunas.
            for coluna in planilha.columns:
                letra = coluna[0].column_letter

                maior = max(
                    len(str(celula.value or ""))
                    for celula in coluna
                )

                planilha.column_dimensions[letra].width = min(
                    max(maior + 2, 12),
                    40,
                )

        memoria.seek(0)

        nome_arquivo = (
            f"ordens_servico_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        )

        return send_file(
            memoria,
            as_attachment=True,
            download_name=nome_arquivo,
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    except Exception:
        app.logger.exception("Erro ao exportar ordens para Excel.")
        flash("Não foi possível gerar o Excel. Verifique o terminal.")
        return redirect(url_for("dashboard"))


# =====================================================
# EXECUTAR A APLICAÇÃO
# =====================================================

if __name__ == "__main__":
    app.run(
        debug=True,
        host="0.0.0.0",
        port=5000,
    )
