import io
import os
import re
import time
from datetime import datetime, timedelta, timezone
import requests
import openpyxl
from flask import Flask, request, jsonify, send_from_directory

app = Flask(__name__, static_folder=".")


def _limpar_id(valor):
    valor = (valor or "").strip().strip('"').strip("'")
    m = re.search(r"/d/([A-Za-z0-9_-]{10,})", valor) or re.search(r"[?&]id=([A-Za-z0-9_-]{10,})", valor)
    return m.group(1) if m else valor


PUBLIC_FILE_ID = _limpar_id(os.environ.get("PUBLIC_FILE_ID", ""))
APPS_SCRIPT_URL = os.environ.get("APPS_SCRIPT_URL", "").strip()
LOCAL_XLSM_PATH = os.environ.get("LOCAL_XLSM_PATH", "/sdcard/Download/manuntecao.xlsm").strip()

_CACHE_PLANILHA = {
    "timestamp": 0.0,
    "dados": None,
    "historico_por_frota": {}
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


def _baixar_conteudo_drive(file_id):
    url = f"https://drive.google.com/uc?export=download&confirm=t&id={file_id}"
    session = requests.Session()
    resposta = session.get(url, timeout=30)
    if resposta.status_code in (403, 404):
        raise RuntimeError(f"O Drive respondeu HTTP {resposta.status_code}. Confira o compartilhamento.")
    resposta.raise_for_status()

    for k, v in resposta.cookies.items():
        if k.startswith("download_warning"):
            resposta = session.get(url, params={"confirm": v}, timeout=30)
            resposta.raise_for_status()
            break

    if not resposta.content.startswith(b"PK"):
        if b"confirm=" in resposta.content:
            m = re.search(r'confirm=([A-Za-z0-9_-]+)', resposta.text)
            if m:
                resposta = session.get(url, params={"confirm": m.group(1)}, timeout=30)
                resposta.raise_for_status()

    if not resposta.content.startswith(b"PK"):
        trecho = re.sub(r"\s+", " ", resposta.text[:150]) if resposta.content else "(vazio)"
        raise RuntimeError(f"O Drive não devolveu uma planilha (veio página de aviso). Trecho: {trecho}")

    return resposta.content


def _processar_workbook(wb):
    em_manutencao = []
    historico_por_frota = {}
    abas_lidas = 0

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
        abas_lidas += 1

        idx_equip = next((i for i, h in enumerate(cabecalhos) if any(k in h for k in ("FROTA", "EQUIP", "TAG"))), 0)
        idx_inicio = next((i for i, h in enumerate(cabecalhos) if any(k in h for k in ("INICIO", "INÍCIO"))), None)
        idx_fim = next((i for i, h in enumerate(cabecalhos) if "FIM" in h), None)
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
                "duracaoMin": duracao_min,
                "especialidade": str(esp_val).strip() if esp_val not in (None, "") else "MANUTENÇÃO",
                "mecanico": str(mec_val).strip() if mec_val not in (None, "") else "-",
                "matricula": str(mat_val).strip() if mat_val not in (None, "") else "-",
                "horimetro": str(hor_val).strip() if hor_val not in (None, "") else "-",
                "tipo": str(tipo_val).strip().capitalize() if tipo_val not in (None, "") else "Corretiva",
                "local": str(local_val).strip() if local_val not in (None, "") else "OFICINA FAGUNDES",
                "motivo": str(comentario_val).strip() if comentario_val not in (None, "") else "SEM COMENTÁRIO"
            }

            if item["fim"] == "":
                em_manutencao.append(item)
            else:
                # Agrupa no histórico por equipamento
                if equip_str not in historico_por_frota:
                    historico_por_frota[equip_str] = []
                historico_por_frota[equip_str].append({
                    "data": item["fim"],
                    "inicio": item["inicio"],
                    "fim": item["fim"],
                    "duracaoMin": item["duracaoMin"],
                    "tipo": item["tipo"],
                    "especialidade": item["especialidade"],
                    "falha": item["motivo"],
                    "servico": item["especialidade"],
                    "pecas": "-",
                    "horimetro": item["horimetro"],
                    "executante": f"{item['mecanico']} ({item['matricula']})" if item['matricula'] != '-' else item['mecanico']
                })

    wb.close()
    return em_manutencao, historico_por_frota, abas_lidas


def _obter_dados_planilha(forcar=False):
    agora = time.time()
    if not forcar and _CACHE_PLANILHA["dados"] is not None and (agora - _CACHE_PLANILHA["timestamp"] < CACHE_TTL_SEGUNDOS):
        return _CACHE_PLANILHA["dados"], _CACHE_PLANILHA["historico_por_frota"]

    wb = None
    fonte = "nenhuma"

    if PUBLIC_FILE_ID:
        try:
            conteudo = _baixar_conteudo_drive(PUBLIC_FILE_ID)
            wb = openpyxl.load_workbook(io.BytesIO(conteudo), data_only=True, read_only=True)
            fonte = "google_drive"
        except Exception as e:
            if not os.path.exists(LOCAL_XLSM_PATH):
                raise RuntimeError(f"Falha ao ler planilha do Google Drive: {str(e)}")
            print(f"Aviso Google Drive: {e}. Usando fallback local.")

    if wb is None and os.path.exists(LOCAL_XLSM_PATH):
        try:
            wb = openpyxl.load_workbook(LOCAL_XLSM_PATH, data_only=True, read_only=True)
            fonte = f"local ({LOCAL_XLSM_PATH})"
        except Exception as e:
            raise RuntimeError(f"Erro ao ler arquivo local: {str(e)}")

    if wb is None:
        raise RuntimeError("Nenhuma planilha encontrada. Configure PUBLIC_FILE_ID ou verifique o arquivo local.")

    em_manutencao, historico_por_frota, abas_lidas = _processar_workbook(wb)
    if abas_lidas == 0:
        raise RuntimeError("Nenhuma aba tem as colunas esperadas (equipamento, inicio e fim).")

    resultado = {
        "total": len(em_manutencao),
        "equipamentos": em_manutencao,
        "fonte": fonte
    }

    _CACHE_PLANILHA["timestamp"] = agora
    _CACHE_PLANILHA["dados"] = resultado
    _CACHE_PLANILHA["historico_por_frota"] = historico_por_frota

    return resultado, historico_por_frota


@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/api/em-manutencao", methods=["GET"])
def listar_em_manutencao():
    forcar = request.args.get("refresh") == "1"
    try:
        dados, _ = _obter_dados_planilha(forcar=forcar)
        return jsonify(dados)
    except Exception as e:
        return jsonify({"erro": str(e), "total": 0, "equipamentos": []}), 500


@app.route("/api/historico/<frota>", methods=["GET"])
def obter_historico(frota):
    frota_limpa = frota.strip().upper()
    try:
        _, historico = _obter_dados_planilha(forcar=False)
        registros = historico.get(frota_limpa, [])
        return jsonify({
            "frota": frota_limpa,
            "total": len(registros),
            "registros": registros[:50]  # Retorna as últimas 50 manutenções reais dessa máquina
        })
    except Exception as e:
        return jsonify({"erro": str(e), "frota": frota_limpa, "registros": []}), 500


@app.route("/api/log/entrada", methods=["POST"])
def log_entrada():
    dados = dict(request.json or {})
    dados["tipo"] = "entrada"
    if not APPS_SCRIPT_URL:
        return jsonify({"ok": True, "status": "simulado", "mensagem": "Log aceito localmente."}), 200
    try:
        res = requests.post(APPS_SCRIPT_URL, json=dados, timeout=25)
        return jsonify({"ok": True, **res.json()})
    except Exception as e:
        return jsonify({"ok": False, "erro": str(e)}), 502


@app.route("/api/log/retirada", methods=["POST"])
def log_retirada():
    dados = dict(request.json or {})
    dados["tipo"] = "retirada"
    if not APPS_SCRIPT_URL:
        return jsonify({"ok": True, "status": "simulado", "mensagem": "Log aceito localmente."}), 200
    try:
        res = requests.post(APPS_SCRIPT_URL, json=dados, timeout=25)
        return jsonify({"ok": True, **res.json()})
    except Exception as e:
        return jsonify({"ok": False, "erro": str(e)}), 502


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
