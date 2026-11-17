import io
import json
import os
import re
import secrets
import threading
import time
from functools import wraps
from datetime import datetime, timedelta, timezone
import requests
import openpyxl
from flask import Flask, request, jsonify, send_from_directory, session
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__, static_folder=None)


def _limpar_id(valor):
    valor = (valor or "").strip().strip('"').strip("'")
    m = re.search(r"/d/([A-Za-z0-9_-]{10,})", valor) or re.search(r"[?&]id=([A-Za-z0-9_-]{10,})", valor)
    return m.group(1) if m else valor


PUBLIC_FILE_ID = _limpar_id(os.environ.get("PUBLIC_FILE_ID", ""))
APPS_SCRIPT_URL = os.environ.get("APPS_SCRIPT_URL", "").strip()
LOCAL_XLSM_PATH = os.environ.get("LOCAL_XLSM_PATH", "/sdcard/Download/manuntecao.xlsm").strip()

USERS_FILE = os.environ.get("USERS_FILE", os.path.join(os.path.dirname(os.path.abspath(__file__)), "usuarios.json"))
USERS_SYNC_TOKEN = os.environ.get("USERS_SYNC_TOKEN", "").strip()


def _secret_key():
    chave = os.environ.get("SECRET_KEY", "").strip()
    if chave:
        return chave
    caminho = USERS_FILE + ".secret"
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            valor = f.read().strip()
            if valor:
                return valor
    except OSError:
        pass
    valor = secrets.token_hex(32)
    try:
        with open(caminho, "w", encoding="utf-8") as f:
            f.write(valor)
    except OSError:
        pass
    return valor


app.secret_key = _secret_key()
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=bool(os.environ.get("RENDER")),
    PERMANENT_SESSION_LIFETIME=timedelta(hours=12),
)

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
    req_session = requests.Session()
    resposta = req_session.get(url, timeout=30)
    if resposta.status_code in (403, 404):
        raise RuntimeError(f"O Drive respondeu HTTP {resposta.status_code}. Confira o compartilhamento.")
    resposta.raise_for_status()

    for k, v in resposta.cookies.items():
        if k.startswith("download_warning"):
            resposta = req_session.get(url, params={"confirm": v}, timeout=30)
            resposta.raise_for_status()
            break

    if not resposta.content.startswith(b"PK"):
        if b"confirm=" in resposta.content:
            m = re.search(r'confirm=([A-Za-z0-9_-]+)', resposta.text)
            if m:
                resposta = req_session.get(url, params={"confirm": m.group(1)}, timeout=30)
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


# =============================================================================
# USUÁRIOS E ACESSOS (COM PERSISTÊNCIA REMOTA NO APPS SCRIPT / GOOGLE DRIVE)
# =============================================================================
PERFIS = {
    "ENCARREGADO": ["quadro", "os"],
    "MECANICO": ["quadro", "minhas-os"],
    "CONTROLADOR": ["quadro", "pendencias"],
    "ADMINISTRADOR": ["quadro", "os", "minhas-os", "pendencias", "acessos"],
}
LOGIN_RE = re.compile(r"^[a-z0-9._-]{3,30}$")
SENHA_MIN = 3

_users_lock = threading.RLock()
_users_cache = {"carregado": False, "usuarios": {}}
_tentativas = {}


def _agora_iso():
    return _agora_local().strftime("%Y-%m-%d %H:%M:%S")


def _admin_padrao():
    return {
        "login": "admin",
        "nome": "Administrador Geral",
        "matricula": "0001",
        "perfil": "ADMINISTRADOR",
        "senha_hash": generate_password_hash("123"),
        "ativo": True,
        "trocar_senha": True,
        "criado_em": _agora_iso(),
        "ultimo_acesso": "",
    }


def _ler_arquivo_usuarios():
    try:
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            dados = json.load(f)
        return dados if isinstance(dados, dict) else None
    except (OSError, ValueError):
        return None


def _apps_script_usuarios(acao, payload=None):
    if not APPS_SCRIPT_URL:
        print("[USUÁRIOS] APPS_SCRIPT_URL não configurada no ambiente!")
        return None
    corpo = {"tipo": acao, "token": USERS_SYNC_TOKEN}
    if payload is not None:
        corpo["usuarios"] = payload
    try:
        res = requests.post(APPS_SCRIPT_URL, json=corpo, timeout=25)
        res.raise_for_status()
        data = res.json()
        if data.get("status") == "sucesso":
            return data
        print(f"[USUÁRIOS] Apps Script recusou ({acao}): {data}")
    except Exception as e:
        print(f"[USUÁRIOS] Erro ao conectar no Apps Script ({acao}): {e}")
    return None


def _carregar_usuarios():
    """Lê usuários do Apps Script (persistente no Render) ou do arquivo local."""
    with _users_lock:
        if _users_cache["carregado"]:
            return _users_cache["usuarios"]

        usuarios = None
        remoto = _apps_script_usuarios("usuarios_carregar")
        if remoto and isinstance(remoto.get("usuarios"), dict) and len(remoto["usuarios"]) > 0:
            usuarios = remoto["usuarios"]
            _persistir_usuarios(usuarios, remoto=False)

        if not usuarios:
            usuarios = _ler_arquivo_usuarios()

        if not usuarios:
            usuarios = {"admin": _admin_padrao()}
            _persistir_usuarios(usuarios, remoto=True)

        _users_cache["usuarios"] = usuarios
        _users_cache["carregado"] = True
        return usuarios


def _persistir_usuarios(usuarios, remoto=True):
    try:
        tmp = USERS_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(usuarios, f, ensure_ascii=False, indent=2)
        os.replace(tmp, USERS_FILE)
    except OSError as e:
        print(f"Aviso: não foi possível gravar {USERS_FILE}: {e}")
    if remoto:
        res = _apps_script_usuarios("usuarios_salvar", usuarios)
        if not res:
            print("[CRÍTICO] Usuários salvos apenas localmente! Ao reiniciar o servidor serão perdidos.")


def _usuario_publico(u):
    return {
        "login": u["login"],
        "nome": u["nome"],
        "matricula": u.get("matricula", ""),
        "perfil": u["perfil"],
        "abas": PERFIS.get(u["perfil"], ["quadro"]),
        "ativo": bool(u.get("ativo", True)),
        "trocarSenha": bool(u.get("trocar_senha", False)),
        "criadoEm": u.get("criado_em", ""),
        "ultimoAcesso": u.get("ultimo_acesso", ""),
    }


def _usuario_logado():
    login = session.get("login")
    if not login:
        return None
    u = _carregar_usuarios().get(login)
    if not u or not u.get("ativo", True):
        session.clear()
        return None
    return u


def exigir_login(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not _usuario_logado():
            return jsonify({"erro": "Sessão expirada. Entre novamente."}), 401
        return f(*args, **kwargs)
    return wrapper


def exigir_admin(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        u = _usuario_logado()
        if not u:
            return jsonify({"erro": "Sessão expirada. Entre novamente."}), 401
        if u["perfil"] != "ADMINISTRADOR":
            return jsonify({"erro": "Apenas o administrador pode fazer isso."}), 403
        return f(*args, **kwargs)
    return wrapper


def _qtd_admins_ativos(usuarios, ignorar=None):
    return sum(
        1 for l, u in usuarios.items()
        if l != ignorar and u["perfil"] == "ADMINISTRADOR" and u.get("ativo", True)
    )


def _validar_senha(senha):
    if not isinstance(senha, str) or len(senha) < SENHA_MIN:
        return f"A senha precisa ter pelo menos {SENHA_MIN} caracteres."
    return None


@app.route("/api/login", methods=["POST"])
def api_login():
    dados = request.get_json(silent=True) or {}
    login = str(dados.get("login", "")).strip().lower()
    senha = str(dados.get("senha", ""))
    chave = f"{request.remote_addr}|{login}"
    falhas, ate = _tentativas.get(chave, (0, 0))
    if ate > time.time():
        restante = int(ate - time.time()) // 60 + 1
        return jsonify({"erro": f"Muitas tentativas. Aguarde {restante} min e tente de novo."}), 429

    usuarios = _carregar_usuarios()
    u = usuarios.get(login)
    hash_ref = u["senha_hash"] if u else generate_password_hash("invalido")
    ok = check_password_hash(hash_ref, senha) and u is not None
    if not ok:
        falhas += 1
        _tentativas[chave] = (falhas, time.time() + 300 if falhas >= 5 else 0)
        return jsonify({"erro": "Usuário ou senha inválidos."}), 401
    if not u.get("ativo", True):
        return jsonify({"erro": "Este acesso está bloqueado. Fale com o administrador."}), 403

    _tentativas.pop(chave, None)
    with _users_lock:
        u["ultimo_acesso"] = _agora_iso()
        _persistir_usuarios(usuarios, remoto=False)
    session.clear()
    session.permanent = True
    session["login"] = login
    return jsonify({"usuario": _usuario_publico(u)})


@app.route("/api/logout", methods=["POST"])
def api_logout():
    session.clear()
    return jsonify({"ok": True})


@app.route("/api/me", methods=["GET"])
def api_me():
    u = _usuario_logado()
    if not u:
        return jsonify({"usuario": None})
    return jsonify({"usuario": _usuario_publico(u)})


@app.route("/api/status-persistencia", methods=["GET"])
@exigir_admin
def api_status_persistencia():
    teste = _apps_script_usuarios("usuarios_carregar")
    return jsonify({
        "apps_script_configurado": bool(APPS_SCRIPT_URL),
        "token_configurado": bool(USERS_SYNC_TOKEN),
        "conectado_drive": teste is not None,
        "qtd_usuarios_remoto": len(teste.get("usuarios", {})) if teste else 0
    })


@app.route("/api/me/senha", methods=["POST"])
@exigir_login
def api_trocar_propria_senha():
    u = _usuario_logado()
    dados = request.get_json(silent=True) or {}
    if not check_password_hash(u["senha_hash"], str(dados.get("senhaAtual", ""))):
        return jsonify({"erro": "Senha atual incorreta."}), 400
    nova = str(dados.get("novaSenha", ""))
    erro = _validar_senha(nova)
    if erro:
        return jsonify({"erro": erro}), 400
    if nova == str(dados.get("senhaAtual", "")):
        return jsonify({"erro": "A nova senha precisa ser diferente da atual."}), 400
    with _users_lock:
        u["senha_hash"] = generate_password_hash(nova)
        u["trocar_senha"] = False
        _persistir_usuarios(_carregar_usuarios())
    return jsonify({"usuario": _usuario_publico(u)})


@app.route("/api/usuarios", methods=["GET"])
@exigir_admin
def api_listar_usuarios():
    usuarios = _carregar_usuarios()
    lista = sorted(usuarios.values(), key=lambda x: (x["perfil"] != "ADMINISTRADOR", x["nome"].lower()))
    return jsonify({"usuarios": [_usuario_publico(u) for u in lista]})


@app.route("/api/usuarios/mecanicos", methods=["GET"])
@exigir_login
def api_listar_mecanicos():
    usuarios = _carregar_usuarios()
    lista = [
        {"login": u["login"], "nome": u["nome"], "matricula": u.get("matricula", "")}
        for u in usuarios.values() if u["perfil"] == "MECANICO" and u.get("ativo", True)
    ]
    lista.sort(key=lambda x: x["nome"].lower())
    return jsonify({"mecanicos": lista})


@app.route("/api/usuarios", methods=["POST"])
@exigir_admin
def api_criar_usuario():
    dados = request.get_json(silent=True) or {}
    login = str(dados.get("login", "")).strip().lower()
    nome = str(dados.get("nome", "")).strip()
    perfil = str(dados.get("perfil", "")).strip().upper()
    matricula = str(dados.get("matricula", "")).strip()[:20]
    senha = str(dados.get("senha", ""))

    if not LOGIN_RE.match(login):
        return jsonify({"erro": "Login: 3 a 30 caracteres, só letras minúsculas, números, ponto, hífen ou _."}), 400
    if len(nome) < 3:
        return jsonify({"erro": "Informe o nome completo."}), 400
    if perfil not in PERFIS:
        return jsonify({"erro": "Perfil inválido."}), 400
    erro = _validar_senha(senha)
    if erro:
        return jsonify({"erro": erro}), 400

    with _users_lock:
        usuarios = _carregar_usuarios()
        if login in usuarios:
            return jsonify({"erro": "Já existe um usuário com esse login."}), 409
        usuarios[login] = {
            "login": login,
            "nome": nome[:80],
            "matricula": matricula,
            "perfil": perfil,
            "senha_hash": generate_password_hash(senha),
            "ativo": True,
            "trocar_senha": bool(dados.get("trocarSenha", True)),
            "criado_em": _agora_iso(),
            "ultimo_acesso": "",
        }
        _persistir_usuarios(usuarios)
        return jsonify({"usuario": _usuario_publico(usuarios[login])}), 201


@app.route("/api/usuarios/<login>", methods=["PUT"])
@exigir_admin
def api_editar_usuario(login):
    login = login.strip().lower()
    dados = request.get_json(silent=True) or {}
    atual = _usuario_logado()
    with _users_lock:
        usuarios = _carregar_usuarios()
        u = usuarios.get(login)
        if not u:
            return jsonify({"erro": "Usuário não encontrado."}), 404

        if "nome" in dados:
            nome = str(dados["nome"]).strip()
            if len(nome) < 3:
                return jsonify({"erro": "Informe o nome completo."}), 400
            u["nome"] = nome[:80]
        if "matricula" in dados:
            u["matricula"] = str(dados["matricula"]).strip()[:20]

        novo_perfil = str(dados.get("perfil", u["perfil"])).strip().upper()
        novo_ativo = bool(dados.get("ativo", u.get("ativo", True)))
        if novo_perfil not in PERFIS:
            return jsonify({"erro": "Perfil inválido."}), 400

        era_admin_ativo = u["perfil"] == "ADMINISTRADOR" and u.get("ativo", True)
        segue_admin_ativo = novo_perfil == "ADMINISTRADOR" and novo_ativo
        if era_admin_ativo and not segue_admin_ativo and _qtd_admins_ativos(usuarios, ignorar=login) == 0:
            return jsonify({"erro": "Precisa existir pelo menos um administrador ativo."}), 400
        if login == atual["login"] and not novo_ativo:
            return jsonify({"erro": "Você não pode bloquear o seu próprio acesso."}), 400

        u["perfil"] = novo_perfil
        u["ativo"] = novo_ativo
        _persistir_usuarios(usuarios)
        return jsonify({"usuario": _usuario_publico(u)})


@app.route("/api/usuarios/<login>/senha", methods=["POST"])
@exigir_admin
def api_redefinir_senha(login):
    login = login.strip().lower()
    dados = request.get_json(silent=True) or {}
    nova = str(dados.get("novaSenha", ""))
    erro = _validar_senha(nova)
    if erro:
        return jsonify({"erro": erro}), 400
    with _users_lock:
        usuarios = _carregar_usuarios()
        u = usuarios.get(login)
        if not u:
            return jsonify({"erro": "Usuário não encontrado."}), 404
        u["senha_hash"] = generate_password_hash(nova)
        u["trocar_senha"] = bool(dados.get("trocarSenha", True))
        _persistir_usuarios(usuarios)
        return jsonify({"usuario": _usuario_publico(u)})


@app.route("/api/usuarios/<login>", methods=["DELETE"])
@exigir_admin
def api_excluir_usuario(login):
    login = login.strip().lower()
    atual = _usuario_logado()
    with _users_lock:
        usuarios = _carregar_usuarios()
        u = usuarios.get(login)
        if not u:
            return jsonify({"erro": "Usuário não encontrado."}), 404
        if login == atual["login"]:
            return jsonify({"erro": "Você não pode excluir o seu próprio acesso."}), 400
        if u["perfil"] == "ADMINISTRADOR" and u.get("ativo", True) and _qtd_admins_ativos(usuarios, ignorar=login) == 0:
            return jsonify({"erro": "Precisa existir pelo menos um administrador ativo."}), 400
        del usuarios[login]
        _persistir_usuarios(usuarios)
        return jsonify({"ok": True})


@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/api/em-manutencao", methods=["GET"])
@exigir_login
def listar_em_manutencao():
    forcar = request.args.get("refresh") == "1"
    try:
        dados, _ = _obter_dados_planilha(forcar=forcar)
        return jsonify(dados)
    except Exception as e:
        return jsonify({"erro": str(e), "total": 0, "equipamentos": []}), 500


@app.route("/api/historico/<frota>", methods=["GET"])
@exigir_login
def obter_historico(frota):
    frota_limpa = frota.strip().upper()
    try:
        _, historico = _obter_dados_planilha(forcar=False)
        registros = historico.get(frota_limpa, [])
        return jsonify({
            "frota": frota_limpa,
            "total": len(registros),
            "registros": registros[:50]
        })
    except Exception as e:
        return jsonify({"erro": str(e), "frota": frota_limpa, "registros": []}), 500


@app.route("/api/log/entrada", methods=["POST"])
@exigir_login
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
@exigir_login
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
