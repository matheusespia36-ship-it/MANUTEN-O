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

## Usuários do sistema

| Login | Senha | Perfil |
|---|---|---|
| `encarregado` | `123` | Abre OS e vê QHA |
| `clecio` | `123` | Mecânico — vê suas OS |
| `controlador` | `123` | CCO — vê QHA + fila de pendências |
| `admin` | `123` | Acesso total |
