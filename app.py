import io
import os
import time
import html
from datetime import datetime, timedelta, timezone
import requests
import openpyxl
from flask import Flask, request, jsonify, send_file

# static_folder=None impede exposição de app.py, .env ou outros arquivos da raiz
app = Flask(__name__, static_folder=None)

PUBLIC_FILE_ID = os.environ.get("PUBLIC_FILE_ID", "").strip()
APPS_SCRIPT_URL = os.environ.get("APPS_SCRIPT_URL", "").strip()
URL_DOWNLOAD_PUBLICO = f"https://drive.google.com/uc?export=download&confirm=t&id={PUBLIC_FILE_ID}"

# Cache de 30 segundos em memória para evitar downloads repetidos do Drive
_CACHE_PLANILHA = {
    "timestamp": 0.0,
    "dados": None
}
CACHE_TTL_SEGUNDOS = 30


def _parse_data_excel(valor):
    if isinstance(valor, datetime):
        return valor
    texto = str(valor or "").strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(texto, fmt)
        except ValueError:
            continue
    return None


def _formatar_data(valor):
    dt = _parse_data_excel(valor)
    if dt:
        return dt.strftime("%d/%m/%Y %H:%M")
    return str(valor or "").strip()


def _agora_local():
    # Planilha em horário de Belém (UTC-3); o servidor (ex.: Render) roda em UTC
    return datetime.now(timezone(timedelta(hours=-3))).replace(tzinfo=None)


def _calcular_duracao_minutos(inicio_val):
    dt_inicio = _parse_data_excel(inicio_val)
    if not dt_inicio:
        return 0.0
    diff = (_agora_local() - dt_inicio).total_seconds() / 60.0
    return round(max(diff, 0.0), 2)


def _inferir_categoria(frota, cat_planilha=""):
    cat_up = str(cat_planilha or "").strip().upper()
    if cat_up:
        return cat_up
    cod = str(frota or "").strip().upper()
    if cod.startswith("CR"):
        return "EQUIPAMENTO INFRA"
    if cod.startswith("EH0"):
        return "EQUIPAMENTO AUXILIAR"
    if cod.startswith(("C221", "EH720")):
        return "EQUIPAMENTO DRYSTACK"
    if cod.startswith("C23"):
        return "EQUIPAMENTO ALIMENTACAO"
    return "EQUIPAMENTO MINA"


def _ler_planilha_drive(forcar=False):
    agora = time.time()
    if not forcar and _CACHE_PLANILHA["dados"] is not None and (agora - _CACHE_PLANILHA["timestamp"] < CACHE_TTL_SEGUNDOS):
        return _CACHE_PLANILHA["dados"]

    resposta = requests.get(URL_DOWNLOAD_PUBLICO, timeout=25)
    resposta.raise_for_status()
    wb = openpyxl.load_workbook(io.BytesIO(resposta.content), data_only=True, read_only=True)

    em_manutencao = []
    encerrados_recentes = []

    for nome_aba in wb.sheetnames:
        ws = wb[nome_aba]
        linha_cabecalho = None
        cabecalhos = []

        for idx_linha, row in enumerate(ws.iter_rows(min_row=1, max_row=12, values_only=True), start=1):
            valores = [str(c).strip().upper() if c is not None else "" for c in row]
            tem_frota = any(k in v for v in valores for k in ("FROTA", "EQUIP", "TAG", "NOME"))
            tem_inicio = any(k in v for v in valores for k in ("INICIO", "INÍCIO"))
            tem_fim = any("FIM" in v for v in valores)
            if tem_frota and tem_inicio and tem_fim:
                linha_cabecalho = idx_linha
                cabecalhos = valores
                break

        if not linha_cabecalho:
            continue

        idx_equip = next((i for i, h in enumerate(cabecalhos) if any(k in h for k in ("FROTA", "EQUIP", "TAG", "NOME"))), 0)
        idx_inicio = next((i for i, h in enumerate(cabecalhos) if any(k in h for k in ("INICIO", "INÍCIO"))), None)
        idx_fim = next((i for i, h in enumerate(cabecalhos) if "FIM" in h), None)
        # Prioriza COMENTÁRIO sobre DESCRIÇÃO para trazer o defeito real
        idx_comentario = next((i for i, h in enumerate(cabecalhos) if "COMENT" in h), None)
        if idx_comentario is None:
            idx_comentario = next((i for i, h in enumerate(cabecalhos) if any(k in h for k in ("FALHA", "MOTIVO", "DESCR"))), None)
        idx_esp = next((i for i, h in enumerate(cabecalhos) if "ESPEC" in h), None)
        if idx_esp is None:
            idx_esp = next((i for i, h in enumerate(cabecalhos) if h == "DESCRICAO" or h == "DESCRIÇÃO"), None)
        idx_mec = next((i for i, h in enumerate(cabecalhos) if any(k in h for k in ("EXECUT", "MECANIC", "MECÂNIC", "NOME_OPERADOR"))), None)
        idx_mat = next((i for i, h in enumerate(cabecalhos) if "MATRIC" in h or "MATRÍC" in h or h == "OPERADOR"), None)
        idx_hor = next((i for i, h in enumerate(cabecalhos) if "HORIM" in h or "HORÍM" in h), None)
        idx_tipo = next((i for i, h in enumerate(cabecalhos) if "TIPO" in h or h == "ESTADO"), None)
        idx_cat = next((i for i, h in enumerate(cabecalhos) if "CATEG" in h or "GRUPO" in h or "FROTA_DESC" in h), None)
        idx_local = next((i for i, h in enumerate(cabecalhos) if h == "LOCAL"), None)
        idx_dur = next((i for i, h in enumerate(cabecalhos) if "DURAC" in h or "DURAÇ" in h or "MINUT" in h), None)

        for row in ws.iter_rows(min_row=linha_cabecalho + 1, values_only=True):
            if not row:
                continue

            equip_val = row[idx_equip] if idx_equip is not None and idx_equip < len(row) else None
            inicio_val = row[idx_inicio] if idx_inicio is not None and idx_inicio < len(row) else None
            fim_val = row[idx_fim] if idx_fim is not None and idx_fim < len(row) else None

            equip_str = str(equip_val).strip() if equip_val is not None else ""
            inicio_str = _formatar_data(inicio_val)
            fim_str = _formatar_data(fim_val) if fim_val not in (None, "") else ""

            if not equip_str or not inicio_str:
                continue

            comentario_val = row[idx_comentario] if idx_comentario is not None and idx_comentario < len(row) else None
            esp_val = row[idx_esp] if idx_esp is not None and idx_esp < len(row) else None
            mec_val = row[idx_mec] if idx_mec is not None and idx_mec < len(row) else None
            mat_val = row[idx_mat] if idx_mat is not None and idx_mat < len(row) else None
            hor_val = row[idx_hor] if idx_hor is not None and idx_hor < len(row) else None
            tipo_val = row[idx_tipo] if idx_tipo is not None and idx_tipo < len(row) else None
            cat_val = row[idx_cat] if idx_cat is not None and idx_cat < len(row) else None
            local_val = row[idx_local] if idx_local is not None and idx_local < len(row) else None
            dur_val = row[idx_dur] if idx_dur is not None and idx_dur < len(row) else None

            try:
                duracao_min = round(float(dur_val), 2) if dur_val not in (None, "") else _calcular_duracao_minutos(inicio_val)
            except (ValueError, TypeError):
                duracao_min = _calcular_duracao_minutos(inicio_val)

            item = {
                "equipamento": equip_str,
                "categoria": _inferir_categoria(equip_str, cat_val),
                "inicio": inicio_str,
                "fim": fim_str if fim_str not in ("None", "nan", "-") else "",
                "duracao_min": duracao_min,
                "especialidade": str(esp_val).strip() if esp_val not in (None, "") else "MANUTENÇÃO",
                "mecanico": str(mec_val).strip() if mec_val not in (None, "") else "-",
                "matricula": str(mat_val).strip() if mat_val not in (None, "") else "-",
                "horimetro": str(hor_val).strip() if hor_val not in (None, "") else "-",
                "tipo": str(tipo_val).strip().capitalize() if tipo_val not in (None, "") else "Corretiva",
                "local": str(local_val).strip() if local_val not in (None, "") else "",
                "motivo": str(comentario_val).strip() if comentario_val not in (None, "") else "SEM COMENTÁRIO"
            }

            if item["fim"] == "":
                em_manutencao.append(item)
            else:
                encerrados_recentes.append(item)

    wb.close()

    resultado = {
        "total": len(em_manutencao),
        "equipamentos": em_manutencao,
        "encerrados_recentes": sorted(
            encerrados_recentes,
            key=lambda i: _parse_data_excel(i["fim"]) or datetime.min,
            reverse=True
        )[:50]
    }
    _CACHE_PLANILHA["timestamp"] = agora
    _CACHE_PLANILHA["dados"] = resultado
    return resultado


@app.route("/")
def index():
    return send_file("index.html")


@app.route("/api/em-manutencao", methods=["GET"])
def listar_em_manutencao():
    if not PUBLIC_FILE_ID:
        return jsonify({
            "erro": "Variável PUBLIC_FILE_ID não configurada no servidor.",
            "total": 0,
            "equipamentos": [],
            "encerrados_recentes": []
        }), 500

    forcar = request.args.get("refresh") == "1"
    try:
        dados = _ler_planilha_drive(forcar=forcar)
        return jsonify(dados)
    except Exception as e:
        return jsonify({
            "erro": f"Falha ao ler planilha do Drive: {str(e)}",
            "total": 0,
            "equipamentos": [],
            "encerrados_recentes": []
        }), 500


def _enviar_para_apps_script(tipo, payload):
    if not APPS_SCRIPT_URL:
        return jsonify({
            "ok": False,
            "erro": "APPS_SCRIPT_URL não configurada no Render. Log não gravado no Drive."
        }), 503

    dados = dict(payload or {})
    dados["tipo"] = tipo
    try:
        res = requests.post(APPS_SCRIPT_URL, json=dados, timeout=20)
        res.raise_for_status()
        try:
            return jsonify({"ok": True, **res.json()})
        except ValueError:
            return jsonify({"ok": True, "resposta": res.text})
    except Exception as e:
        return jsonify({
            "ok": False,
            "erro": f"Falha na comunicação com Google Apps Script: {str(e)}"
        }), 502


@app.route("/api/log/entrada", methods=["POST"])
def log_entrada():
    return _enviar_para_apps_script("entrada", request.json)


@app.route("/api/log/retirada", methods=["POST"])
def log_retirada():
    return _enviar_para_apps_script("retirada", request.json)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
