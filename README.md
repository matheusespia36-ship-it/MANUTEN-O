# Painel de Manutenção QHA — Fagundes Construção e Mineração

Painel web de controle de manutenção de equipamentos (QHA), Ordens de Serviço e fila de pendências CCO.

## Tecnologias
- **Backend:** Python + Flask + Gunicorn
- **Frontend:** HTML + Tailwind CSS (via CDN)
- **Dados:** Planilha `.xlsm` hospedada no Google Drive (pública)

## Variáveis de Ambiente (obrigatório no Render)

| Variável | Descrição |
|---|---|
| `PUBLIC_FILE_ID` | ID do arquivo `.xlsm` no Google Drive (compartilhado como "qualquer pessoa com o link") |
| `APPS_SCRIPT_URL` | (Opcional) URL do Google Apps Script para gravar logs de OS |

## Como obter o PUBLIC_FILE_ID

1. Abra o arquivo `manuntecao.xlsm` no Google Drive
2. Clique em **Compartilhar → Qualquer pessoa com o link → Visualizador**
3. Copie o link — o ID é a parte entre `/d/` e `/view`:  
   `https://drive.google.com/file/d/**SEU_ID_AQUI**/view`

## Deploy no Render

1. Suba este repositório no GitHub
2. Acesse [render.com](https://render.com) → New → Web Service
3. Conecte o repositório GitHub
4. Em **Environment Variables**, adicione `PUBLIC_FILE_ID`
5. Build command: `pip install -r requirements.txt`
6. Start command: `gunicorn app:app`

## Usuários e acessos

Existe apenas um acesso inicial:

| Login | Senha inicial | Perfil |
|---|---|---|
| `admin` | `123` | Administrador (acesso total) |

Entre com ele, abra a aba **Gestão de Acessos** e crie os demais usuários
(Encarregado, Mecânico, Controlador CCO ou outro Administrador). Dá para editar,
redefinir senha, bloquear/ativar e excluir. O painel avisa para trocar a senha
inicial (botão **Minha Senha**).

As senhas ficam criptografadas no servidor (nunca no HTML). Os mecânicos cadastrados
aparecem sozinhos na lista "Mecânico Destinado" ao abrir uma OS.

### Para os usuários não sumirem quando o Render reinicia

O disco do Render grátis é apagado a cada reinício. Para guardar os usuários no Google Drive:

1. No `codigo_google_apps_script.js`, troque o valor de `USERS_SYNC_TOKEN` por um texto secreto longo e reimplante o script.
2. No Render, crie a variável `USERS_SYNC_TOKEN` com **o mesmo texto** (e mantenha `APPS_SCRIPT_URL`).

Variáveis extras (opcionais): `SECRET_KEY` (o `render.yaml` já gera uma), `USERS_FILE` (caminho do arquivo local de usuários).
