import io
import os
import requests
import openpyxl
from flask import Flask, request, jsonify, send_from_directory

app = Flask(__name__, static_folder=".")

PUBLIC_FILE_ID = os.environ.get("PUBLIC_FILE_ID", "").strip()
APPS_SCRIPT_URL = os.environ.get("APPS_SCRIPT_URL", "").strip()
URL_DOWNLOAD_PUBLICO = f"https://drive.google.com/uc?export=download&id={PUBLIC_FILE_ID}"


@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/api/em-manutencao", methods=["GET"])
def listar_em_manutencao():
    if not PUBLIC_FILE_ID:
        return jsonify({"total": 0, "equipamentos": []})

    try:
        resposta = requests.get(URL_DOWNLOAD_PUBLICO, timeout=15)
        resposta.raise_for_status()
        wb = openpyxl.load_workbook(io.BytesIO(resposta.content), data_only=True, read_only=True)
        ws = wb.active

        cabecalhos = [
            str(c.value).strip() if c.value else ""
            for c in next(ws.iter_rows(min_row=1, max_row=1))
        ]
        idx_equip = next((i for i, h in enumerate(cabecalhos) if "EQUIP" in h.upper() or "TAG" in h.upper()), 0)
        idx_inicio = next((i for i, h in enumerate(cabecalhos) if "INICIO" in h.upper() or "INÍCIO" in h.upper()), None)
        idx_fim = next((i for i, h in enumerate(cabecalhos) if "FIM" in h.upper()), None)
        idx_motivo = next((i for i, h in enumerate(cabecalhos) if "FALHA" in h.upper() or "MOTIVO" in h.upper() or "DESCR" in h.upper()), None)
        idx_mec = next((i for i, h in enumerate(cabecalhos) if "MECANIC" in h.upper() or "MECÂNIC" in h.upper()), None)

        em_manutencao = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            equipamento = row[idx_equip] if idx_equip is not None and idx_equip < len(row) else None
            inicio = row[idx_inicio] if idx_inicio is not None and idx_inicio < len(row) else None
            fim = row[idx_fim] if idx_fim is not None and idx_fim < len(row) else None

            if equipamento and inicio and (fim is None or str(fim).strip() == ""):
                em_manutencao.append({
                    "equipamento": str(equipamento).strip(),
                    "inicio": str(inicio),
                    "motivo": str(row[idx_motivo]).strip() if idx_motivo is not None and row[idx_motivo] else "Em atendimento",
                    "mecanico": str(row[idx_mec]).strip() if idx_mec is not None and row[idx_mec] else "-"
                })

        wb.close()
        return jsonify({"total": len(em_manutencao), "equipamentos": em_manutencao})
    except Exception as e:
        return jsonify({"erro": str(e), "total": 0, "equipamentos": []}), 500


@app.route("/api/log/entrada", methods=["POST"])
def log_entrada():
    dados = request.json or {}
    dados["tipo"] = "entrada"
    res = requests.post(APPS_SCRIPT_URL, json=dados, timeout=20)
    return jsonify(res.json())


@app.route("/api/log/retirada", methods=["POST"])
def log_retirada():
    dados = request.json or {}
    dados["tipo"] = "retirada"
    res = requests.post(APPS_SCRIPT_URL, json=dados, timeout=20)
    return jsonify(res.json())


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
