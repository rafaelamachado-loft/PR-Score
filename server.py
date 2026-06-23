#!/usr/bin/env python3
"""
PR Score — Servidor local
Uso:
  pip install flask anthropic requests beautifulsoup4 openpyxl
  export ANTHROPIC_API_KEY="sk-ant-api03-..."
  python server.py
  Abra: http://localhost:5000
"""

import os, json, re, time
from flask import Flask, request, jsonify, send_from_directory
import anthropic
import requests
from bs4 import BeautifulSoup

app = Flask(__name__, static_folder=".")

CLAUDE_MODEL = "claude-sonnet-4-6"
_client = None

def get_client():
    global _client
    if _client is None:
        key = os.environ.get("ANTHROPIC_API_KEY", "")
        if not key:
            raise ValueError("ANTHROPIC_API_KEY não configurada. Execute: export ANTHROPIC_API_KEY='sk-ant-...'")
        _client = anthropic.Anthropic(api_key=key)
    return _client


# ── serve o HTML ────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory(".", "pr_score_local.html")


# ── teste de conexão ────────────────────────────────────────
@app.route("/api/ping")
def ping():
    try:
        c = get_client()
        c.messages.create(model=CLAUDE_MODEL, max_tokens=5,
                          messages=[{"role": "user", "content": "OK"}])
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ── buscar texto do artigo ──────────────────────────────────
@app.route("/api/fetch-text", methods=["POST"])
def fetch_text():
    url = request.json.get("url", "")
    if not url.startswith("http"):
        return jsonify({"text": ""})
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=12)
        soup = BeautifulSoup(r.text, "html.parser")
        for tag in soup(["script", "style", "nav", "header", "footer", "aside"]):
            tag.decompose()
        text = " ".join(p.get_text().strip() for p in soup.find_all("p"))
        return jsonify({"text": text[:4500]})
    except Exception as e:
        return jsonify({"text": "", "error": str(e)})


PROMPT_AVALIACAO = (
    "Você é um especialista em avaliação de PR Score para empresas do mercado imobiliário brasileiro.\n"
    "\n"
    "Sua tarefa é avaliar a matéria a seguir e preencher todos os campos solicitados.\n"
    "\n"
    "EMPRESA MONITORADA NESTA MATÉRIA: {empresa}\n"
    "TÍTULO: {titulo}\n"
    "VEÍCULO: {veiculo}\n"
    "TIPO DE MÍDIA: {tipo_midia}\n"
    "CONTEÚDO DA MATÉRIA:\n"
    "{conteudo}\n"
    "\n"
    "---\n"
    "\n"
    "## FILTRO INICIAL\n"
    "Se a matéria for release empresarial explícito, conteúdo patrocinado ou segmento fora do core imobiliário: "
    "protagonismo = 'Excluído', explicar em obs. EXCEÇÃO: releases de terceiros citando a marca organicamente: MANTER.\n"
    "\n"
    "## CRITÉRIOS PRINCIPAIS (não-cumulativos)\n"
    "Verifique se pelo menos UM é atendido. Se nenhum: protagonismo = 'Menção', soma = 0.\n"
    "Se algum: protagonismo = 'Destaque', soma começa em 1.\n"
    "\n"
    "1. Porta-voz Ativo: 3+ frases com aspas diretas OU artigo assinado pelo porta-voz.\n"
    "2. Espaço Qualificado: 5+ frases OU 8+ linhas referenciando a marca.\n"
    "3. Densidade: trechos da marca ocupam 30%+ do texto.\n"
    "4. Fontes de Dados: marca citada como fonte/pesquisa.\n"
    "5. Colunas Curtas: nota ≤300 caracteres onde a marca é protagonista.\n"
    "6. Lives: 1.000+ visualizações. [SINALIZAR COMO HUMANO]\n"
    "7. Anúncios Governamentais: citação em anúncio de governo federal/estadual/capital.\n"
    "\n"
    "## CRITÉRIOS EXTRAS (cumulativos, +1 cada)\n"
    "8. Protagonismo Editorial: +1 por local: título, chapéu, subtítulo, linha fina (máx +4).\n"
    "9. Volume de Menções: 6+ menções nominais: +1.\n"
    "10. CTA/Link: 'Leia mais' direcionando para matéria da marca: +1 por ocorrência.\n"
    "11. Recursos Visuais: +1 por TIPO diferente de imagem. [VER HUMANO se não confirmável]\n"
    "12. Gráficos/Tabelas: +1 por TIPO diferente. [VER HUMANO se não confirmável]\n"
    "13. Páginas Impressas: +1 por página física adicional (só jornais/revistas físicas).\n"
    "14. Capas e Redes Sociais: [SEMPRE ver humano]\n"
    "15. Audiovisual: [ver humano se rádio/TV/podcast; senão 0]\n"
    "16. Protocolo 10+ trechos: [SEMPRE ver humano]\n"
    "\n"
    "Responda EXCLUSIVAMENTE em JSON (sem texto antes ou depois):\n"
    "{{\n"
    '  "protagonismo": "Destaque | Menção | Destaque Negativo | Menção Negativa | Excluído",\n'
    '  "criterio_ativado": "1 a 7 ou nenhum",\n'
    '  "criterio_8": 0, "criterio_9": 0, "criterio_10": 0,\n'
    '  "criterio_11": 0, "criterio_12": 0, "criterio_13": 0,\n'
    '  "criterio_14": "ver humano", "criterio_15": 0, "criterio_16": "ver humano",\n'
    '  "soma": 0,\n'
    '  "pr_tec_produto": false, "pr_puro_imobis": false,\n'
    '  "data": true,\n'
    '  "data_carona": false,\n'
    '  "tema": "", "subtema": "", "produto": "", "obs": "", "cidade": "",\n'
    '  "confianca": "alta"\n'
    "}}\n"
    "\n"
    "Regras: criterio_14 e criterio_16 são SEMPRE 'ver humano'. "
    "soma = 1 (se Destaque) + soma numérica critérios 8-16 (ignora 'ver humano'). "
    "Se paywall: obs = 'PAYWALL — leitura incompleta', confianca = 'baixa'. "
    "pr_tec_produto e pr_puro_imobis são false para empresas que não são Loft."
)


# ── avaliar matéria ─────────────────────────────────────────
@app.route("/api/evaluate", methods=["POST"])
def evaluate():
    body = request.json or {}
    empresa  = body.get("empresa", "")
    titulo   = body.get("titulo", "")
    veiculo  = body.get("veiculo", "")
    midia    = body.get("midia", "")
    conteudo = body.get("conteudo", "") or f"[Texto indisponível — avalie pelo título: {titulo}]"

    prompt = PROMPT_AVALIACAO.format(
        empresa=empresa, titulo=titulo, veiculo=veiculo,
        tipo_midia=midia, conteudo=conteudo
    )

    for tentativa in range(3):
        try:
            c   = get_client()
            msg = c.messages.create(
                model=CLAUDE_MODEL, max_tokens=900,
                messages=[{"role": "user", "content": prompt}]
            )
            raw = msg.content[0].text.strip()
            raw = re.sub(r"```json\s*", "", raw)
            raw = re.sub(r"```\s*",    "", raw)
            match = re.search(r"\{[\s\S]*\}", raw)
            if not match:
                raise ValueError("JSON não encontrado na resposta")
            res = json.loads(match.group(0))
            res["dado_proativo"] = res.pop("data", False)
            return jsonify({"ok": True, "result": res})
        except Exception as e:
            if tentativa < 2:
                time.sleep(2 ** (tentativa + 1))
            else:
                return jsonify({"ok": False, "error": str(e)}), 500

    return jsonify({"ok": False, "error": "Falha após 3 tentativas"}), 500


if __name__ == "__main__":
    print("=" * 55)
    print("🟠  PR Score — Servidor local")
    print("=" * 55)
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        print("⚠️  ANTHROPIC_API_KEY não encontrada!")
        print("   Execute: export ANTHROPIC_API_KEY='sk-ant-...'")
    else:
        masked = api_key[:12] + "..." + api_key[-4:]
        print(f"✅  API Key: {masked}")
    print(f"✅  Modelo:  {CLAUDE_MODEL}")
    print("🌐  Abrindo em: http://localhost:5000")
    print("=" * 55)
    app.run(host="0.0.0.0", port=5000, debug=False)
