import os
import re
import traceback
from datetime import datetime

from flask import Flask, redirect, render_template_string, request, url_for
from werkzeug.utils import secure_filename

from gsheet_utils import append_to_sheet


DEFAULT_COLOR = "#0f766e"

COURSE_EMOJIS = {
    "designer": "💅", "unha": "💅", "costura": "🧵", "corte": "✂️",
    "recepcion": "🛎️", "logística": "📦", "logistica": "📦",
    "administrativo": "📂", "cozinha": "🍳", "barbeiro": "💈",
    "idoso": "👵", "eletric": "⚡", "inteligência": "🤖",
    "marketing": "📱", "lazer": "🎈", "porteiro": "🔑", "portaria": "🔑",
    "informática": "💻", "informatica": "💻",
}

def get_course_emoji(course_name):
    lower = course_name.lower()
    for key, emoji in COURSE_EMOJIS.items():
        if key in lower:
            return emoji
    return "📚"


TEMPLATE_FORM = r'''
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Formulário de Inscrições</title>
    <style>
        :root {
            --bg: #f4efe7;
            --card: rgba(255, 252, 248, 0.94);
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
        * { box-sizing: border-box; }
        html, body { margin: 0; min-height: 100%; font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif; background: radial-gradient(circle at top left, rgba(15,118,110,0.14),transparent 30%), radial-gradient(circle at bottom right, rgba(244,96,54,0.16),transparent 28%), linear-gradient(135deg,#f8f5ef 0%,#f4efe7 48%,#efe7dc 100%); color: var(--ink); }
        body { padding: 24px 14px 40px; }
        .shell { width: min(980px, 100%); margin: 0 auto; display: flex; flex-direction: column; gap: 24px; }

        /* Cards */
        .form-card { background: var(--card); border: 1px solid rgba(255,255,255,0.65); border-radius: 28px; box-shadow: var(--shadow); backdrop-filter: blur(10px); padding: 28px; }
        .card-title { margin: 0 0 4px; font-size: 1.1rem; font-weight: 700; color: var(--accent-strong); display: flex; align-items: center; gap: 8px; }
        .card-title .step-badge { background: var(--accent); color: #fff; border-radius: 999px; width: 26px; height: 26px; display: inline-flex; align-items: center; justify-content: center; font-size: 0.82rem; font-weight: 800; flex-shrink: 0; }
        .card-desc { margin: 0 0 20px; color: var(--muted); font-size: 0.9rem; }

        /* Page title */
        .page-title { text-align: center; margin: 0; font-size: 1.55rem; }
        .page-copy { text-align: center; margin: 0; color: var(--muted); line-height: 1.5; }

        /* Grid */
        .form-grid { display: grid; grid-template-columns: repeat(2, minmax(0,1fr)); gap: 18px 16px; }
        .field { min-width: 0; }
        .field.full { grid-column: 1 / -1; }
        label { display: block; margin-bottom: 8px; font-size: 0.95rem; font-weight: 700; }
        input, textarea, select { width: 100%; border: 1px solid var(--line); border-radius: 16px; background: #fff; padding: 14px 15px; color: var(--ink); font-size: 1rem; transition: border-color .2s, box-shadow .2s, transform .2s; font-family: inherit; }
        input:focus, textarea:focus, select:focus { outline: none; border-color: rgba(15,118,110,0.7); box-shadow: 0 0 0 4px rgba(15,118,110,0.12); transform: translateY(-1px); }
        textarea { min-height: 90px; resize: vertical; }
        .field.error input, .field.error textarea, .field.error select { border-color: #f97066; box-shadow: 0 0 0 4px rgba(249,112,102,0.12); }
        .error-text { display: block; margin-top: 6px; color: var(--danger); font-size: 0.88rem; }
        .helper { margin-top: 8px; color: var(--muted); font-size: 0.88rem; }

        /* Banners */
        .banner { margin-bottom: 18px; padding: 14px 16px; border-radius: 18px; font-size: 0.96rem; line-height: 1.5; }
        .banner.success { background: var(--success-bg); color: var(--success); border: 1px solid #abefc6; }
        .banner.error { background: var(--danger-bg); color: var(--danger); border: 1px solid #fecdca; }

        /* Course list */
        .turmas-list { display: flex; flex-direction: column; gap: 20px; }
        .turma-item { background: #fffaf4; border: 1px solid var(--line); border-radius: 22px; padding: 20px; position: relative; }
        .turma-item-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; }
        .turma-num { font-size: 0.8rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.08em; color: var(--accent); }
        .remove-btn { background: none; border: none; cursor: pointer; color: var(--muted); font-size: 1.3rem; padding: 2px 8px; border-radius: 8px; line-height: 1; transition: color .2s, background .2s; }
        .remove-btn:hover { color: var(--danger); background: var(--danger-bg); }

        .add-btn { display: flex; align-items: center; gap: 8px; background: none; border: 2px dashed var(--line); border-radius: 18px; color: var(--accent); font-size: 0.95rem; font-weight: 700; padding: 14px 20px; cursor: pointer; width: 100%; justify-content: center; transition: border-color .2s, background .2s; }
        .add-btn:hover { border-color: var(--accent); background: rgba(15,118,110,0.04); }

        /* Color picker */
        .color-picker { display: flex; align-items: center; gap: 16px; padding: 16px; border: 1px solid var(--line); border-radius: 18px; background: #fff; }
        .color-picker input[type="color"] { width: 72px; min-width: 72px; height: 72px; padding: 0; border: 0; border-radius: 18px; background: transparent; cursor: pointer; }
        .color-picker input[type="color"]::-webkit-color-swatch-wrapper { padding: 0; }
        .color-picker input[type="color"]::-webkit-color-swatch { border: 3px solid rgba(255,255,255,0.95); border-radius: 18px; box-shadow: 0 0 0 1px rgba(31,41,51,0.12); }
        .color-meta { display: grid; gap: 6px; }
        .color-code { font-family: Consolas, monospace; font-size: 0.95rem; color: var(--accent-strong); }

        /* Actions */
        .actions { margin-top: 26px; display: flex; justify-content: flex-end; }
        .submit-button { border: 0; border-radius: 999px; background: linear-gradient(135deg, #0f766e 0%, #115e59 100%); color: #fff; padding: 15px 36px; font-size: 1rem; font-weight: 700; cursor: pointer; box-shadow: 0 16px 32px rgba(15,118,110,0.24); transition: transform .2s, box-shadow .2s; }
        .submit-button:hover { transform: translateY(-2px); box-shadow: 0 18px 36px rgba(15,118,110,0.28); }

        @media (max-width: 760px) {
            body { padding: 16px 10px 28px; }
            .form-card { padding: 20px; border-radius: 22px; }
            .form-grid { grid-template-columns: minmax(0,1fr); }
            .actions { justify-content: stretch; }
            .submit-button { width: 100%; }
            .color-picker { flex-direction: column; align-items: center; text-align: center; }
            .color-picker input[type="color"] { width: min(100%,84px); min-width: 84px; height: 84px; }
        }
    </style>
</head>
<body>
<div class="shell">

    {% if success_message %}
    <div class="banner success">{{ success_message }}</div>
    {% endif %}
    {% if save_error %}
    <div class="banner error">{{ save_error }}</div>
    {% endif %}

    <div style="text-align:center;">
        <h2 class="page-title">Formulário de Link de Inscrições</h2>
        <p class="page-copy">Preencha os dados do projeto e adicione os cursos/turmas disponíveis.</p>
    </div>

    <form method="POST" action="{{ url_for('home') }}" enctype="multipart/form-data" novalidate>

        <!-- BLOCO 1: DADOS DO PROJETO -->
        <div class="form-card" style="margin-bottom:0;">
            <div class="card-title"><span class="step-badge">1</span> Dados do Projeto</div>
            <p class="card-desc">Informações gerais que aparecem no topo do link de inscrições.</p>
            <div class="form-grid">
                <div class="field {% if errors.get('nome_projeto') %}error{% endif %}">
                    <label for="nome_projeto">Nome do projeto *</label>
                    <input type="text" id="nome_projeto" name="nome_projeto" maxlength="120" placeholder="Ex.: EDUCATECH" value="{{ form_data.get('nome_projeto','') }}">
                    {% if errors.get('nome_projeto') %}<span class="error-text">{{ errors.get('nome_projeto') }}</span>{% endif %}
                </div>
                <div class="field {% if errors.get('titulo') %}error{% endif %}">
                    <label for="titulo">Título *</label>
                    <input type="text" id="titulo" name="titulo" maxlength="200" placeholder="Ex.: TRANSFORME SUA CARREIRA AGORA!" value="{{ form_data.get('titulo','') }}">
                    {% if errors.get('titulo') %}<span class="error-text">{{ errors.get('titulo') }}</span>{% endif %}
                </div>
                <div class="field full {% if errors.get('subtitulo') %}error{% endif %}">
                    <label for="subtitulo">Subtítulo / Slogan *</label>
                    <textarea id="subtitulo" name="subtitulo" placeholder="Ex.: Conectando vidas, transformando pessoas.">{{ form_data.get('subtitulo','') }}</textarea>
                    {% if errors.get('subtitulo') %}<span class="error-text">{{ errors.get('subtitulo') }}</span>{% endif %}
                </div>
                <div class="field full {% if errors.get('beneficios') %}error{% endif %}">
                    <label for="beneficios">Benefícios / Diferenciais *</label>
                    <textarea id="beneficios" name="beneficios" placeholder="- 100% Gratuito&#10;- Certificado de Conclusão&#10;- Material didático incluso" style="min-height:120px;">{{ form_data.get('beneficios','') }}</textarea>
                    <div class="helper">Liste um benefício por linha, começando com hífen (-).</div>
                    {% if errors.get('beneficios') %}<span class="error-text">{{ errors.get('beneficios') }}</span>{% endif %}
                </div>
                <div class="field full {% if errors.get('logo') %}error{% endif %}">
                    <label for="logo">Logo do projeto</label>
                    <div style="display:flex;align-items:center;gap:18px;flex-wrap:wrap;">
                        <input type="file" id="logo" name="logo" accept="image/*" onchange="previewLogo(event)">
                        <div id="logo-preview-container" style="min-width:120px;min-height:80px;display:flex;align-items:center;justify-content:center;border:1px dashed #ccc;border-radius:12px;background:#fafafa;">
                            {% if form_data.get('logo') %}
                                <img id="logo-preview" src="/{{ form_data.get('logo') }}" alt="Prévia do logo" style="max-width:120px;max-height:80px;border-radius:8px;">
                            {% else %}
                                <span id="logo-preview-placeholder" style="color:#aaa;font-size:0.9rem;">Prévia do logo</span>
                            {% endif %}
                        </div>
                    </div>
                    <div class="helper">Opcional. Formatos aceitos: png, jpg, gif, webp.</div>
                    {% if errors.get('logo') %}<span class="error-text">{{ errors.get('logo') }}</span>{% endif %}
                </div>
                <div class="field full {% if errors.get('cor_ficha') %}error{% endif %}">
                    <label>Cor da ficha *</label>
                    <div class="color-picker">
                        <input type="color" id="cor_ficha" name="cor_ficha" value="{{ form_data.get('cor_ficha','') or default_color }}">
                        <div class="color-meta">
                            <strong>Escolha qualquer cor</strong>
                            <span class="color-code" id="cor_ficha_valor">{{ form_data.get('cor_ficha','') or default_color }}</span>
                        </div>
                    </div>
                    {% if errors.get('cor_ficha') %}<span class="error-text">{{ errors.get('cor_ficha') }}</span>{% endif %}
                </div>
            </div>
        </div>

        <!-- BLOCO 2: LOCAIS -->
        <div class="form-card" style="margin-top:24px; margin-bottom:0;">
            <div class="card-title"><span class="step-badge">2</span> Locais de Aula</div>
            <p class="card-desc">Cadastre todos os locais onde ocorrerão as aulas (nome, região e endereço completo).</p>
            <div id="locais-list" class="turmas-list">
                {% set locais = form_data.get('locais', [{}]) %}
                {% for local in locais %}
                <div class="turma-item" data-local>
                    <div class="turma-item-header">
                        <span class="turma-num">Local {{ loop.index }}</span>
                        {% if loop.length > 1 %}<button type="button" class="remove-btn" onclick="removeItem(this, 'local')">✕</button>{% endif %}
                    </div>
                    <div class="form-grid">
                        <div class="field"><label>Nome do local *</label><input type="text" name="local_nome[]" maxlength="120" placeholder="Ex.: Polo Campo Grande" value="{{ local.get('nome','') }}"></div>
                        <div class="field"><label>Região *</label><input type="text" name="local_regiao[]" maxlength="120" placeholder="Ex.: Zona Oeste" value="{{ local.get('regiao','') }}"></div>
                        <div class="field full"><label>Endereço completo *</label><textarea name="local_endereco[]" placeholder="Digite o endereço completo conforme o Maps">{{ local.get('endereco','') }}</textarea></div>
                    </div>
                </div>
                {% endfor %}
            </div>
            <button type="button" class="add-btn" style="margin-top:16px;" onclick="addLocal()">+ Adicionar local</button>
            {% if errors.get('locais') %}<span class="error-text" style="margin-top:8px;display:block;">{{ errors.get('locais') }}</span>{% endif %}
        </div>

        <!-- BLOCO 3: CURSOS E TURMAS -->
        <div class="form-card" style="margin-top:24px; margin-bottom:0;">
            <div class="card-title"><span class="step-badge">3</span> Cursos e Turmas</div>
            <p class="card-desc">Adicione cada turma com seu curso, local, horário, vagas e datas.</p>
            <div id="turmas-list" class="turmas-list">
                {% set turmas = form_data.get('turmas', [{}]) %}
                {% for turma in turmas %}
                <div class="turma-item" data-turma>
                    <div class="turma-item-header">
                        <span class="turma-num">Turma {{ loop.index }}</span>
                        {% if loop.length > 1 %}<button type="button" class="remove-btn" onclick="removeItem(this, 'turma')">✕</button>{% endif %}
                    </div>
                    <div class="form-grid">
                        <div class="field"><label>Curso *</label><input type="text" name="turma_curso[]" maxlength="160" placeholder="Ex.: Designer de Unha" value="{{ turma.get('curso','') }}"></div>
                        <div class="field"><label>Local da turma *</label><input type="text" name="turma_local[]" maxlength="160" placeholder="Ex.: Polo Campo Grande — Sala 02" value="{{ turma.get('local','') }}"></div>
                        <div class="field"><label>Horário *</label><input type="text" name="turma_horario[]" maxlength="80" placeholder="Ex.: 9h30 às 11h30" value="{{ turma.get('horario','') }}"></div>
                        <div class="field"><label>Vagas *</label><input type="number" name="turma_vagas[]" min="1" max="9999" placeholder="Ex.: 30" value="{{ turma.get('vagas','') }}"></div>
                        <div class="field"><label>Dias de aula *</label><input type="text" name="turma_dias[]" maxlength="120" placeholder="Ex.: Terça e Quinta" value="{{ turma.get('dias','') }}"></div>
                        <div class="field"><label>Data de início *</label><input type="date" name="turma_inicio[]" value="{{ turma.get('inicio','') }}"></div>
                        <div class="field"><label>Encerramento *</label><input type="date" name="turma_encerramento[]" value="{{ turma.get('encerramento','') }}"></div>
                        <div class="field full"><label>Endereço da turma *</label><textarea name="turma_endereco[]" placeholder="📍Endereço completo conforme o Maps + CEP">{{ turma.get('endereco','') }}</textarea></div>
                    </div>
                </div>
                {% endfor %}
            </div>
            <button type="button" class="add-btn" style="margin-top:16px;" onclick="addTurma()">+ Adicionar turma</button>
            {% if errors.get('turmas') %}<span class="error-text" style="margin-top:8px;display:block;">{{ errors.get('turmas') }}</span>{% endif %}
        </div>

        <div class="form-card" style="margin-top:24px;">
            <div class="actions" style="margin-top:0;">
                <button type="submit" class="submit-button">Salvar informações</button>
            </div>
        </div>

    </form>
</div>

<script>
document.getElementById('cor_ficha').addEventListener('input', function(e) {
    document.getElementById('cor_ficha_valor').textContent = e.target.value;
});

const LOCAL_TEMPLATE = `
<div class="turma-item" data-local>
  <div class="turma-item-header">
    <span class="turma-num">Local __N__</span>
    <button type="button" class="remove-btn" onclick="removeItem(this,'local')">✕</button>
  </div>
  <div class="form-grid">
    <div class="field"><label>Nome do local *</label><input type="text" name="local_nome[]" maxlength="120" placeholder="Ex.: Polo Campo Grande"></div>
    <div class="field"><label>Região *</label><input type="text" name="local_regiao[]" maxlength="120" placeholder="Ex.: Zona Oeste"></div>
    <div class="field full"><label>Endereço completo *</label><textarea name="local_endereco[]" placeholder="Digite o endereço completo conforme o Maps"></textarea></div>
  </div>
</div>`;

const TURMA_TEMPLATE = `
<div class="turma-item" data-turma>
  <div class="turma-item-header">
    <span class="turma-num">Turma __N__</span>
    <button type="button" class="remove-btn" onclick="removeItem(this,'turma')">✕</button>
  </div>
  <div class="form-grid">
    <div class="field"><label>Curso *</label><input type="text" name="turma_curso[]" maxlength="160" placeholder="Ex.: Designer de Unha"></div>
    <div class="field"><label>Local da turma *</label><input type="text" name="turma_local[]" maxlength="160" placeholder="Ex.: Polo Campo Grande — Sala 02"></div>
    <div class="field"><label>Horário *</label><input type="text" name="turma_horario[]" maxlength="80" placeholder="Ex.: 9h30 às 11h30"></div>
    <div class="field"><label>Vagas *</label><input type="number" name="turma_vagas[]" min="1" max="9999" placeholder="Ex.: 30"></div>
    <div class="field"><label>Dias de aula *</label><input type="text" name="turma_dias[]" maxlength="120" placeholder="Ex.: Terça e Quinta"></div>
    <div class="field"><label>Data de início *</label><input type="date" name="turma_inicio[]"></div>
    <div class="field"><label>Encerramento *</label><input type="date" name="turma_encerramento[]"></div>
    <div class="field full"><label>Endereço da turma *</label><textarea name="turma_endereco[]" placeholder="📍Endereço completo conforme o Maps + CEP"></textarea></div>
  </div>
</div>`;

function addLocal() {
    const list = document.getElementById('locais-list');
    const n = list.querySelectorAll('[data-local]').length + 1;
    const div = document.createElement('div');
    div.innerHTML = LOCAL_TEMPLATE.replace('__N__', n);
    list.appendChild(div.firstElementChild);
    renumber();
}

function addTurma() {
    const list = document.getElementById('turmas-list');
    const n = list.querySelectorAll('[data-turma]').length + 1;
    const div = document.createElement('div');
    div.innerHTML = TURMA_TEMPLATE.replace('__N__', n);
    list.appendChild(div.firstElementChild);
    renumber();
}

function removeItem(btn, type) {
    btn.closest('[data-' + type + ']').remove();
    renumber();
}

function renumber() {
    document.querySelectorAll('[data-local]').forEach((el, i) => {
        el.querySelector('.turma-num').textContent = 'Local ' + (i+1);
    });
    document.querySelectorAll('[data-turma]').forEach((el, i) => {
        el.querySelector('.turma-num').textContent = 'Turma ' + (i+1);
    });
}

function previewLogo(event) {
    const input = event.target;
    const container = document.getElementById('logo-preview-container');
    let preview = document.getElementById('logo-preview');
    let placeholder = document.getElementById('logo-preview-placeholder');
    if (input.files && input.files[0]) {
        const reader = new FileReader();
        reader.onload = function(e) {
            if (!preview) {
                preview = document.createElement('img');
                preview.id = 'logo-preview';
                preview.style.cssText = 'max-width:120px;max-height:80px;border-radius:8px;';
                container.innerHTML = '';
                container.appendChild(preview);
            }
            preview.src = e.target.result;
        };
        reader.readAsDataURL(input.files[0]);
        if (placeholder) placeholder.style.display = 'none';
    } else {
        if (preview) preview.remove();
        if (placeholder) placeholder.style.display = '';
    }
}
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
            return datetime.strptime(cleaned, fmt).strftime("%d/%m/%Y"), True
        except ValueError:
            continue
    return cleaned, False


def build_formatted_text(form_data):
    """Monta o texto formatado completo para colar em uma única célula da planilha."""
    lines = []

    lines.append("PROJETO:")
    lines.append(form_data["nome_projeto"].upper())
    lines.append("")
    lines.append("TÍTULO:")
    lines.append(form_data["titulo"].upper())
    lines.append("")
    lines.append("SUBTÍTULO:")
    lines.append(form_data["subtitulo"])
    lines.append("")

    # Lista de cursos únicos
    turmas = form_data.get("turmas", [])
    cursos_unicos = []
    seen = set()
    for t in turmas:
        c = t.get("curso", "").strip()
        if c and c.upper() not in seen:
            seen.add(c.upper())
            cursos_unicos.append(c)

    if cursos_unicos:
        lines.append("CURSOS DISPONÍVEIS:")
        for c in cursos_unicos:
            emoji = get_course_emoji(c)
            lines.append(f"{emoji} {c.upper()}")
        lines.append("")

    # Benefícios
    beneficios = form_data.get("beneficios", "").strip()
    if beneficios:
        lines.append("BENEFÍCIOS:")
        for linha in beneficios.splitlines():
            linha = linha.strip()
            if linha:
                if not linha.startswith("-"):
                    linha = "- " + linha
                lines.append(linha)
        lines.append("")

    # Opções por curso
    opcao_num = 0
    sub_map = {}  # curso -> sub-número
    for turma in turmas:
        curso = turma.get("curso", "").strip()
        if not curso:
            continue
        curso_up = curso.upper()

        if curso_up not in sub_map:
            opcao_num += 1
            sub_map[curso_up] = {"opcao": opcao_num, "sub": 0}
            lines.append(f"OPÇÃO {opcao_num}")
            lines.append(f"CURSO = {curso_up}")

        sub_map[curso_up]["sub"] += 1
        sub = sub_map[curso_up]["sub"]
        opcao = sub_map[curso_up]["opcao"]

        lines.append(f"###OPÇÃO {opcao}.{sub}")

        local = turma.get("local", "").strip()
        horario = turma.get("horario", "").strip()
        dias = turma.get("dias", "").strip()
        inicio, _ = parse_date(turma.get("inicio", ""))
        encerramento, _ = parse_date(turma.get("encerramento", ""))
        endereco = turma.get("endereco", "").strip()
        vagas = turma.get("vagas", "").strip()

        if local:
            lines.append(f"LOCAL = {local.upper()}")
        if dias and horario:
            lines.append(f"DIA / HORÁRIO = {dias} | {horario}")
        elif horario:
            lines.append(f"HORÁRIO = {horario}")
        if vagas:
            lines.append(f"VAGAS = {vagas}")
        if inicio:
            lines.append(f"DATA DE INÍCIO = {inicio}")
        if encerramento:
            lines.append(f"ENCERRAMENTO = {encerramento}")
        if endereco:
            addr = endereco if endereco.startswith("📍") else f"📍{endereco}"
            lines.append(f"ENDEREÇO = {addr}")

    lines.append("")
    lines.append(f"COR DA FICHA = {form_data.get('cor_ficha','')}")

    return "\n".join(lines)


def get_default_form_data():
    return {
        "nome_projeto": "",
        "titulo": "",
        "subtitulo": "",
        "beneficios": "",
        "cor_ficha": DEFAULT_COLOR,
        "locais": [{}],
        "turmas": [{}],
    }


def parse_lists_from_request(req):
    """Extrai locais e turmas do request como listas de dicts."""
    nomes = req.form.getlist("local_nome[]")
    regioes = req.form.getlist("local_regiao[]")
    enderecos = req.form.getlist("local_endereco[]")

    locais = []
    for i in range(len(nomes)):
        locais.append({
            "nome": nomes[i].strip() if i < len(nomes) else "",
            "regiao": regioes[i].strip() if i < len(regioes) else "",
            "endereco": enderecos[i].strip() if i < len(enderecos) else "",
        })

    t_cursos = req.form.getlist("turma_curso[]")
    t_locais = req.form.getlist("turma_local[]")
    t_horarios = req.form.getlist("turma_horario[]")
    t_vagas = req.form.getlist("turma_vagas[]")
    t_dias = req.form.getlist("turma_dias[]")
    t_inicios = req.form.getlist("turma_inicio[]")
    t_encerramentos = req.form.getlist("turma_encerramento[]")
    t_enderecos = req.form.getlist("turma_endereco[]")

    turmas = []
    for i in range(len(t_cursos)):
        turmas.append({
            "curso": t_cursos[i].strip() if i < len(t_cursos) else "",
            "local": t_locais[i].strip() if i < len(t_locais) else "",
            "horario": t_horarios[i].strip() if i < len(t_horarios) else "",
            "vagas": t_vagas[i].strip() if i < len(t_vagas) else "",
            "dias": t_dias[i].strip() if i < len(t_dias) else "",
            "inicio": t_inicios[i].strip() if i < len(t_inicios) else "",
            "encerramento": t_encerramentos[i].strip() if i < len(t_encerramentos) else "",
            "endereco": t_enderecos[i].strip() if i < len(t_enderecos) else "",
        })

    return locais, turmas


def validate(form_data):
    errors = {}

    for field, msg in [
        ("nome_projeto", "Informe o nome do projeto."),
        ("titulo", "Informe o título."),
        ("subtitulo", "Informe o subtítulo/slogan."),
        ("beneficios", "Informe os benefícios."),
        ("cor_ficha", "Escolha a cor da ficha."),
    ]:
        if not form_data.get(field, "").strip():
            errors[field] = msg

    if form_data.get("cor_ficha") and not re.fullmatch(r"#[0-9a-fA-F]{6}", form_data["cor_ficha"]):
        errors["cor_ficha"] = "Escolha uma cor válida usando a paleta."

    locais = form_data.get("locais", [])
    for i, local in enumerate(locais):
        if not local.get("nome") or not local.get("regiao") or not local.get("endereco"):
            errors["locais"] = f"Preencha todos os campos obrigatórios do Local {i+1}."
            break

    turmas = form_data.get("turmas", [])
    if not turmas:
        errors["turmas"] = "Adicione ao menos uma turma."
    for i, t in enumerate(turmas):
        required = ["curso", "local", "horario", "vagas", "dias", "inicio", "encerramento", "endereco"]
        missing = [f for f in required if not t.get(f)]
        if missing:
            errors["turmas"] = f"Preencha todos os campos da Turma {i+1}."
            break

    return errors


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

    locais, turmas = parse_lists_from_request(request)

    # Processa upload do logo do projeto
    logo_path = ""
    logo_error = ""
    logo_file = request.files.get("logo")
    if logo_file and logo_file.filename:
        filename = secure_filename(logo_file.filename)
        ext = os.path.splitext(filename)[1].lower()
        if ext not in [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"]:
            logo_error = "Apenas imagens são permitidas (png, jpg, jpeg, gif, bmp, webp)."
        else:
            os.makedirs("static", exist_ok=True)
            save_path = os.path.join("static", filename)
            try:
                logo_file.save(save_path)
                logo_path = save_path
            except Exception:
                logo_error = "Erro ao salvar a imagem. Tente novamente."

    form_data = {
        "nome_projeto": request.form.get("nome_projeto", "").strip(),
        "titulo": request.form.get("titulo", "").strip(),
        "subtitulo": request.form.get("subtitulo", "").strip(),
        "beneficios": request.form.get("beneficios", "").strip(),
        "cor_ficha": request.form.get("cor_ficha", DEFAULT_COLOR).strip(),
        "logo": logo_path,
        "locais": locais,
        "turmas": turmas,
    }

    errors = validate(form_data)
    if logo_error:
        errors["logo"] = logo_error
    if errors:
        return render_form(form_data=form_data, errors=errors)

    texto_formatado = build_formatted_text(form_data)

    # Monta a linha da planilha:
    # Coluna A = nome do projeto, B = cor, C = logo, D = texto completo formatado
    row = [
        form_data["nome_projeto"],
        form_data["cor_ficha"],
        form_data.get("logo", ""),
        texto_formatado,
    ]

    try:
        append_to_sheet(row)
    except Exception as exc:
        print("Erro ao salvar na planilha:", exc)
        traceback.print_exc()
        return render_form(
            form_data=form_data,
            save_error="Não foi possível salvar na planilha agora. Os dados continuam no formulário para tentar novamente.",
        )

    return render_form(
        form_data=get_default_form_data(),
        success_message="Informações salvas com sucesso na planilha!",
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
