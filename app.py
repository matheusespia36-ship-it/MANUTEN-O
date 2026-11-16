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
        return jsonify({"total": 0, "equipamentos": [], "aviso": "PUBLIC_FILE_ID não configurado"})

    try:
        resposta = requests.get(URL_DOWNLOAD_PUBLICO, timeout=20)
        resposta.raise_for_status()
        wb = openpyxl.load_workbook(io.BytesIO(resposta.content), data_only=True, read_only=True)

        em_manutencao = []

        for nome_aba in wb.sheetnames:
            ws = wb[nome_aba]
            linha_cabecalho = None
            cabecalhos = []

            # Procura o cabeçalho nas primeiras 10 linhas da aba
            for idx_linha, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
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
            idx_motivo = next((i for i, h in enumerate(cabecalhos) if any(k in h for k in ("COMENT", "FALHA", "MOTIVO", "DESCR"))), None)
            idx_mec = next((i for i, h in enumerate(cabecalhos) if any(k in h for k in ("EXECUT", "MECANIC", "MECÂNIC"))), None)
            idx_esp = next((i for i, h in enumerate(cabecalhos) if "ESPEC" in h), None)

            for row in ws.iter_rows(min_row=linha_cabecalho + 1, values_only=True):
                if not row:
                    continue

                equipamento = row[idx_equip] if idx_equip is not None and idx_equip < len(row) else None
                inicio = row[idx_inicio] if idx_inicio is not None and idx_inicio < len(row) else None
                fim = row[idx_fim] if idx_fim is not None and idx_fim < len(row) else None

                equip_str = str(equipamento).strip() if equipamento is not None else ""
                inicio_str = str(inicio).strip() if inicio is not None else ""
                fim_str = str(fim).strip() if fim is not None else ""

                # Considera em manutenção toda linha com Frota + Início preenchidos e Fim vazio
                if equip_str and inicio_str and fim_str in ("", "None", "nan", "-"):
                    motivo_val = row[idx_motivo] if idx_motivo is not None and idx_motivo < len(row) else None
                    mec_val = row[idx_mec] if idx_mec is not None and idx_mec < len(row) else None
                    esp_val = row[idx_esp] if idx_esp is not None and idx_esp < len(row) else None

                    em_manutencao.append({
                        "equipamento": equip_str,
                        "inicio": inicio_str,
                        "motivo": str(motivo_val).strip() if motivo_val is not None and str(motivo_val).strip() else "EM MANUTENÇÃO",
                        "mecanico": str(mec_val).strip() if mec_val is not None and str(mec_val).strip() else "-",
                        "especialidade": str(esp_val).strip() if esp_val is not None and str(esp_val).strip() else ""
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
