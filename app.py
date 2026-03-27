import os
import re
import traceback
from datetime import datetime

from flask import Flask, redirect, render_template_string, request, url_for

from gsheet_utils import append_to_sheet


DEFAULT_COLOR = "#0f766e"
DATE_FIELDS = {"data_inicio", "encerramento"}

TEMPLATE_FORM = r'''
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
    <title>Cadastro de Cursos</title>
    <style>
        :root {
            --bg: #f4efe7;
            --card: rgba(255, 252, 248, 0.94);
            --card-strong: #fffaf4;
            --ink: #1f2933;
            --muted: #6b7280;
            --line: #eadfce;
            --accent: #0f766e;
            --accent-strong: #115e59;
            --danger: #b42318;
            --danger-bg: #fef3f2;
            --success: #166534;
            --success-bg: #ecfdf3;
            --shadow: 0 24px 60px rgba(55, 65, 81, 0.12);
        }

        * {
            box-sizing: border-box;
        }

        html,
        body {
            margin: 0;
            min-height: 100%;
            font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
            background:
                radial-gradient(circle at top left, rgba(15, 118, 110, 0.14), transparent 30%),
                radial-gradient(circle at bottom right, rgba(244, 96, 54, 0.16), transparent 28%),
                linear-gradient(135deg, #f8f5ef 0%, #f4efe7 48%, #efe7dc 100%);
            color: var(--ink);
        }

        body {
            padding: 24px 14px 40px;
        }

        .shell {
            width: min(980px, 100%);
            margin: 0 auto;
        }

        .form-card {
            background: var(--card);
            border: 1px solid rgba(255, 255, 255, 0.65);
            border-radius: 28px;
            box-shadow: var(--shadow);
            backdrop-filter: blur(10px);
        }

        .form-card {
            padding: 28px;
        }

        .section-title {
            margin: 0 0 6px;
            font-size: 1.55rem;
            text-align: center;
        }

        .section-copy {
            margin: 0 0 24px;
            color: var(--muted);
            line-height: 1.5;
            text-align: center;
        }

        .banner {
            margin-bottom: 18px;
            padding: 14px 16px;
            border-radius: 18px;
            font-size: 0.96rem;
            line-height: 1.5;
        }

        .banner.success {
            background: var(--success-bg);
            color: var(--success);
            border: 1px solid #abefc6;
        }

        .banner.error {
            background: var(--danger-bg);
            color: var(--danger);
            border: 1px solid #fecdca;
        }

        .form-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 18px 16px;
        }

        .field,
        .field.full {
            min-width: 0;
        }

        .field.full {
            grid-column: 1 / -1;
        }

        label {
            display: block;
            margin-bottom: 8px;
            font-size: 0.95rem;
            font-weight: 700;
        }

        input,
        textarea {
            width: 100%;
            border: 1px solid var(--line);
            border-radius: 16px;
            background: #fff;
            padding: 14px 15px;
            color: var(--ink);
            font-size: 1rem;
            transition: border-color 0.2s ease, box-shadow 0.2s ease, transform 0.2s ease;
        }

        input:focus,
        textarea:focus {
            outline: none;
            border-color: rgba(15, 118, 110, 0.7);
            box-shadow: 0 0 0 4px rgba(15, 118, 110, 0.12);
            transform: translateY(-1px);
        }

        textarea {
            min-height: 110px;
            resize: vertical;
        }

        .field.error input,
        .field.error textarea {
            border-color: #f97066;
            box-shadow: 0 0 0 4px rgba(249, 112, 102, 0.12);
        }

        .error-text {
            display: block;
            margin-top: 6px;
            color: var(--danger);
            font-size: 0.88rem;
        }

        .palette-label {
            margin-bottom: 12px;
        }

        .color-picker {
            display: flex;
            align-items: center;
            gap: 16px;
            padding: 16px;
            border: 1px solid var(--line);
            border-radius: 18px;
            background: #fff;
        }

        .color-picker input[type="color"] {
            width: 72px;
            min-width: 72px;
            height: 72px;
            padding: 0;
            border: 0;
            border-radius: 18px;
            background: transparent;
            cursor: pointer;
        }

        .color-picker input[type="color"]::-webkit-color-swatch-wrapper {
            padding: 0;
        }

        .color-picker input[type="color"]::-webkit-color-swatch {
            border: 3px solid rgba(255, 255, 255, 0.95);
            border-radius: 18px;
            box-shadow: 0 0 0 1px rgba(31, 41, 51, 0.12);
        }

        .color-meta {
            display: grid;
            gap: 6px;
        }

        .color-meta strong {
            font-size: 1rem;
        }

        .color-code {
            font-family: Consolas, "Courier New", monospace;
            font-size: 0.95rem;
            color: var(--accent-strong);
        }

        .ticket-preview {
            position: relative;
            margin-top: 14px;
            padding: 18px;
            border-radius: 22px;
            color: #fff;
            background: linear-gradient(135deg, var(--ticket-color, #0f766e) 0%, color-mix(in srgb, var(--ticket-color, #0f766e) 68%, #ffffff 32%) 100%);
            box-shadow: 0 18px 36px rgba(15, 118, 110, 0.22);
            overflow: hidden;
        }

        .ticket-preview::before,
        .ticket-preview::after {
            content: "";
            position: absolute;
            width: 28px;
            height: 28px;
            border-radius: 999px;
            background: var(--bg);
            top: 50%;
            transform: translateY(-50%);
            opacity: 0.95;
        }

        .ticket-preview::before {
            left: -14px;
        }

        .ticket-preview::after {
            right: -14px;
        }

        .ticket-badge {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 7px 12px;
            border-radius: 999px;
            background: rgba(255, 255, 255, 0.16);
            border: 1px solid rgba(255, 255, 255, 0.22);
            font-size: 0.78rem;
            font-weight: 800;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }

        .ticket-title {
            margin: 14px 0 6px;
            font-size: 1.15rem;
            font-weight: 800;
        }

        .ticket-copy {
            margin: 0;
            max-width: 320px;
            font-size: 0.95rem;
            line-height: 1.45;
            color: rgba(255, 255, 255, 0.92);
        }

        .ticket-footer {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            margin-top: 18px;
            padding-top: 14px;
            border-top: 1px dashed rgba(255, 255, 255, 0.3);
        }

        .ticket-level {
            font-size: 0.82rem;
            font-weight: 700;
            letter-spacing: 0.04em;
            text-transform: uppercase;
        }

        .ticket-stars {
            font-size: 1rem;
            letter-spacing: 0.18em;
            filter: drop-shadow(0 2px 6px rgba(0, 0, 0, 0.2));
        }

        .actions {
            margin-top: 26px;
            display: flex;
            justify-content: flex-end;
        }

        .submit-button {
            border: 0;
            border-radius: 999px;
            background: linear-gradient(135deg, #0f766e 0%, #115e59 100%);
            color: #fff;
            padding: 15px 24px;
            font-size: 1rem;
            font-weight: 700;
            cursor: pointer;
            box-shadow: 0 16px 32px rgba(15, 118, 110, 0.24);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }

        .submit-button:hover {
            transform: translateY(-2px);
            box-shadow: 0 18px 36px rgba(15, 118, 110, 0.28);
        }

        .helper {
            margin-top: 10px;
            color: var(--muted);
            font-size: 0.88rem;
        }

        @media (max-width: 760px) {
            body {
                padding: 16px 10px 28px;
            }

            .form-card {
                padding: 20px;
                border-radius: 22px;
            }

            .form-grid {
                grid-template-columns: minmax(0, 1fr);
            }

            .actions {
                justify-content: stretch;
            }

            .submit-button {
                width: 100%;
            }

            .color-picker {
                flex-direction: column;
                align-items: center;
                text-align: center;
                padding: 14px;
            }

            .color-picker input[type="color"] {
                width: min(100%, 84px);
                min-width: 84px;
                height: 84px;
            }

            .color-meta {
                justify-items: center;
            }

            .ticket-preview {
                padding: 16px;
                border-radius: 20px;
            }

            .ticket-copy {
                max-width: none;
            }

            .ticket-footer {
                flex-direction: column;
                align-items: flex-start;
            }
        }
    </style>
</head>
<body>
    <div class="shell">
        <section class="form-card">
            <h2 class="section-title">FORMULÁRIO</h2>
            <p class="section-copy">Preencha os dados exatamente como devem ser registrados. Para endereço e CEP, use a referência do Maps, como você pediu.</p>

            {% if success_message %}
            <div class="banner success">{{ success_message }}</div>
            {% endif %}

            {% if save_error %}
            <div class="banner error">{{ save_error }}</div>
            {% endif %}

            <form method="POST" action="{{ url_for('home') }}" novalidate>
                <div class="form-grid">
                    <div class="field {% if errors.get('nome_local') %}error{% endif %}">
                        <label for="nome_local">Nome do local *</label>
                        <input type="text" id="nome_local" name="nome_local" maxlength="120" placeholder="Ex.: Polo Campo Grande" value="{{ form_data.get('nome_local', '') }}">
                        {% if errors.get('nome_local') %}<span class="error-text">{{ errors.get('nome_local') }}</span>{% endif %}
                    </div>

                    <div class="field {% if errors.get('regiao') %}error{% endif %}">
                        <label for="regiao">Região *</label>
                        <input type="text" id="regiao" name="regiao" maxlength="120" placeholder="Ex.: Zona Oeste" value="{{ form_data.get('regiao', '') }}">
                        {% if errors.get('regiao') %}<span class="error-text">{{ errors.get('regiao') }}</span>{% endif %}
                    </div>

                    <div class="field full {% if errors.get('endereco_completo') %}error{% endif %}">
                        <label for="endereco_completo">Endereço completo *</label>
                        <textarea id="endereco_completo" name="endereco_completo" placeholder="Digite o endereço completo conforme o Maps">{{ form_data.get('endereco_completo', '') }}</textarea>
                        {% if errors.get('endereco_completo') %}<span class="error-text">{{ errors.get('endereco_completo') }}</span>{% endif %}
                    </div>

                    <div class="field {% if errors.get('cep') %}error{% endif %}">
                        <label for="cep">CEP *</label>
                        <input type="text" id="cep" name="cep" maxlength="9" inputmode="numeric" placeholder="00000-000" value="{{ form_data.get('cep', '') }}">
                        {% if errors.get('cep') %}<span class="error-text">{{ errors.get('cep') }}</span>{% endif %}
                    </div>

                    <div class="field {% if errors.get('cursos') %}error{% endif %}">
                        <label for="cursos">Cursos *</label>
                        <input type="text" id="cursos" name="cursos" maxlength="160" placeholder="Ex.: Informática Básica" value="{{ form_data.get('cursos', '') }}">
                        {% if errors.get('cursos') %}<span class="error-text">{{ errors.get('cursos') }}</span>{% endif %}
                    </div>

                    <div class="field {% if errors.get('local_turma') %}error{% endif %}">
                        <label for="local_turma">Local da turma *</label>
                        <input type="text" id="local_turma" name="local_turma" maxlength="120" placeholder="Ex.: Sala 02" value="{{ form_data.get('local_turma', '') }}">
                        {% if errors.get('local_turma') %}<span class="error-text">{{ errors.get('local_turma') }}</span>{% endif %}
                    </div>

                    <div class="field {% if errors.get('horario') %}error{% endif %}">
                        <label for="horario">Horário *</label>
                        <input type="text" id="horario" name="horario" maxlength="80" placeholder="Ex.: 18h às 20h" value="{{ form_data.get('horario', '') }}">
                        {% if errors.get('horario') %}<span class="error-text">{{ errors.get('horario') }}</span>{% endif %}
                    </div>

                    <div class="field {% if errors.get('vagas') %}error{% endif %}">
                        <label for="vagas">Vagas *</label>
                        <input type="number" id="vagas" name="vagas" min="1" max="9999" step="1" placeholder="Ex.: 30" value="{{ form_data.get('vagas', '') }}">
                        {% if errors.get('vagas') %}<span class="error-text">{{ errors.get('vagas') }}</span>{% endif %}
                    </div>

                    <div class="field {% if errors.get('turma') %}error{% endif %}">
                        <label for="turma">Turma *</label>
                        <input type="text" id="turma" name="turma" maxlength="120" placeholder="Ex.: Turma A" value="{{ form_data.get('turma', '') }}">
                        {% if errors.get('turma') %}<span class="error-text">{{ errors.get('turma') }}</span>{% endif %}
                    </div>

                    <div class="field {% if errors.get('dias_aula') %}error{% endif %}">
                        <label for="dias_aula">Dias de aula *</label>
                        <input type="text" id="dias_aula" name="dias_aula" maxlength="120" placeholder="Ex.: Segunda e quarta" value="{{ form_data.get('dias_aula', '') }}">
                        {% if errors.get('dias_aula') %}<span class="error-text">{{ errors.get('dias_aula') }}</span>{% endif %}
                    </div>

                    <div class="field {% if errors.get('data_inicio') %}error{% endif %}">
                        <label for="data_inicio">Data de início *</label>
                        <input type="date" id="data_inicio" name="data_inicio" value="{{ form_data.get('data_inicio', '') }}">
                        {% if errors.get('data_inicio') %}<span class="error-text">{{ errors.get('data_inicio') }}</span>{% endif %}
                    </div>

                    <div class="field {% if errors.get('encerramento') %}error{% endif %}">
                        <label for="encerramento">Encerramento *</label>
                        <input type="date" id="encerramento" name="encerramento" value="{{ form_data.get('encerramento', '') }}">
                        {% if errors.get('encerramento') %}<span class="error-text">{{ errors.get('encerramento') }}</span>{% endif %}
                    </div>

                    <div class="field full {% if errors.get('cor_ficha') %}error{% endif %}">
                        <label class="palette-label" for="cor_ficha">Cor da ficha *</label>
                        <div class="color-picker">
                            <input type="color" id="cor_ficha" name="cor_ficha" value="{{ form_data.get('cor_ficha', '') or default_color }}">
                            <div class="color-meta">
                                <strong>Escolha qualquer cor</strong>
                                <span class="color-code" id="cor_ficha_valor">{{ form_data.get('cor_ficha', '') or default_color }}</span>
                            </div>
                        </div>
                        <div class="ticket-preview" id="ticket_preview" style="--ticket-color: {{ form_data.get('cor_ficha', '') or default_color }};">
                            <div class="ticket-badge">Ficha desbloqueada</div>
                            <div class="ticket-title">Prévia da ficha</div>
                            <p class="ticket-copy">A cor escolhida vira a identidade visual da ficha.</p>
                            <div class="ticket-footer">
                                <span class="ticket-level">Nível visual da ficha</span>
                                <span class="ticket-stars">★ ★ ★</span>
                            </div>
                        </div>
                        {% if errors.get('cor_ficha') %}<span class="error-text">{{ errors.get('cor_ficha') }}</span>{% endif %}
                        <div class="helper">Clique no seletor para abrir a paleta do navegador e escolher a cor que quiser.</div>
                    </div>
                </div>

                <div class="actions">
                    <button type="submit" class="submit-button">Salvar informações</button>
                </div>
            </form>
        </section>
    </div>

    <script>
        document.addEventListener('DOMContentLoaded', function() {
            const cepInput = document.getElementById('cep');
            const colorInput = document.getElementById('cor_ficha');
            const colorValue = document.getElementById('cor_ficha_valor');
            const ticketPreview = document.getElementById('ticket_preview');

            function formatCep(value) {
                const digits = (value || '').replace(/\D/g, '').slice(0, 8);
                if (digits.length <= 5) {
                    return digits;
                }
                return digits.slice(0, 5) + '-' + digits.slice(5);
            }

            cepInput.addEventListener('input', function(event) {
                event.target.value = formatCep(event.target.value);
            });

            colorInput.addEventListener('input', function(event) {
                colorValue.textContent = event.target.value;
                ticketPreview.style.setProperty('--ticket-color', event.target.value);
            });
        });
    </script>
</body>
</html>
'''


app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "chave-secreta-para-sessao")


def parse_date(value):
    cleaned = (value or "").strip()
    if not cleaned:
        return "", False

    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            parsed = datetime.strptime(cleaned, fmt)
            return parsed.strftime("%d/%m/%Y"), True
        except ValueError:
            continue

    return cleaned, False


def get_default_form_data(source=None):
    form_data = {
        "nome_local": "",
        "regiao": "",
        "endereco_completo": "",
        "cep": "",
        "cursos": "",
        "local_turma": "",
        "horario": "",
        "vagas": "",
        "turma": "",
        "dias_aula": "",
        "data_inicio": "",
        "encerramento": "",
        "cor_ficha": "",
    }

    if not source:
        return form_data

    for key in form_data:
        form_data[key] = (source.get(key, "") or "").strip()

    if not form_data["cor_ficha"]:
        form_data["cor_ficha"] = DEFAULT_COLOR

    return form_data


def validate_form_data(form_data):
    errors = {}

    required_messages = {
        "nome_local": "Informe o nome do local.",
        "regiao": "Informe a região.",
        "endereco_completo": "Informe o endereço completo.",
        "cep": "Informe o CEP.",
        "cursos": "Informe o curso.",
        "local_turma": "Informe o local da turma.",
        "horario": "Informe o horário.",
        "vagas": "Informe a quantidade de vagas.",
        "turma": "Informe a turma.",
        "dias_aula": "Informe os dias de aula.",
        "data_inicio": "Informe a data de início.",
        "encerramento": "Informe o encerramento.",
        "cor_ficha": "Escolha a cor da ficha.",
    }

    for field_name, message in required_messages.items():
        if not form_data[field_name]:
            errors[field_name] = message

    cep = form_data["cep"]
    if cep and not re.fullmatch(r"\d{5}-\d{3}", cep):
        errors["cep"] = "Use o formato 00000-000."

    vagas = form_data["vagas"]
    if vagas:
        if not vagas.isdigit():
            errors["vagas"] = "Digite apenas números inteiros para vagas."
        elif int(vagas) <= 0:
            errors["vagas"] = "A quantidade de vagas deve ser maior que zero."

    for field_name in DATE_FIELDS:
        value = form_data[field_name]
        if value:
            _, valid = parse_date(value)
            if not valid:
                label = "data de início" if field_name == "data_inicio" else "encerramento"
                errors[field_name] = f"Informe um {label} válido."

    if form_data["cor_ficha"] and not re.fullmatch(r"#[0-9a-fA-F]{6}", form_data["cor_ficha"]):
        errors["cor_ficha"] = "Escolha uma cor válida usando a paleta."

    return errors


def build_sheet_row(form_data):
    data_inicio, _ = parse_date(form_data["data_inicio"])
    encerramento, _ = parse_date(form_data["encerramento"])
    return [
        form_data["nome_local"],
        form_data["regiao"],
        form_data["endereco_completo"],
        form_data["cep"],
        form_data["cursos"],
        form_data["local_turma"],
        form_data["horario"],
        form_data["vagas"],
        form_data["turma"],
        form_data["dias_aula"],
        data_inicio,
        encerramento,
        form_data["cor_ficha"],
    ]


def render_form(form_data=None, errors=None, success_message="", save_error=""):
    return render_template_string(
        TEMPLATE_FORM,
        default_color=DEFAULT_COLOR,
        errors=errors or {},
        form_data=form_data or get_default_form_data(),
        save_error=save_error,
        success_message=success_message,
    )


@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "GET":
        return render_form()

    form_data = get_default_form_data(request.form)
    errors = validate_form_data(form_data)
    if errors:
        return render_form(form_data=form_data, errors=errors)

    try:
        append_to_sheet(build_sheet_row(form_data))
    except Exception as exc:
        print("Erro ao salvar na planilha:", exc)
        traceback.print_exc()
        return render_form(
            form_data=form_data,
            save_error="Não foi possível salvar na planilha agora. Os dados continuam no formulário para você tentar novamente.",
        )

    return render_form(
        form_data=get_default_form_data(),
        success_message="Informações salvas com sucesso na planilha.",
    )


@app.route("/inscricao", methods=["GET", "POST"])
@app.route("/curso", methods=["GET", "POST"])
@app.route("/revisao", methods=["GET", "POST"])
@app.route("/wizard", methods=["GET", "POST"])
@app.route("/confirmacao", methods=["GET", "POST"])
def legacy_routes():
    if request.method == "POST" and request.path == "/inscricao":
        return home()
    return redirect(url_for("home"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
