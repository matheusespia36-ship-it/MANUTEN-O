/**
 * GOOGLE APPS SCRIPT PARA O PAINEL DE MANUTENÇÃO (QHA)
 * 
 * Como instalar no Google Drive:
 * 1. Abra o Google Drive (https://drive.google.com)
 * 2. Crie uma pasta chamada "Logs_Manutencao" (ou o nome que preferir)
 * 3. Vá em https://script.google.com/home e clique em "Novo projeto"
 * 4. Cole este código abaixo
 * 5. Clique em "Implantar" -> "Nova implantação"
 * 6. Tipo: "Aplicativo da Web"
 * 7. Executar como: "Eu"
 * 8. Quem pode acessar: "Qualquer pessoa" (mesmo anônima)
 * 9. Copie a URL gerada (terminada em /exec) e configure no APPS_SCRIPT_URL
 *
 * USUÁRIOS (Gestão de Acessos): para os acessos criados pelo administrador não
 * se perderem quando o Render reinicia, defina abaixo o mesmo valor que você
 * colocar na variável USERS_SYNC_TOKEN do Render (qualquer texto longo e secreto).
 */
var USERS_SYNC_TOKEN = "TROQUE-POR-UM-TEXTO-SECRETO-LONGO";

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function _abaUsuarios() {
  var pastas = DriveApp.getFoldersByName("Logs_Manutencao");
  var pasta = pastas.hasNext() ? pastas.next() : DriveApp.createFolder("Logs_Manutencao");
  var arquivos = pasta.getFilesByName("Usuarios_Painel_QHA");
  var ss;
  if (arquivos.hasNext()) {
    ss = SpreadsheetApp.open(arquivos.next());
  } else {
    ss = SpreadsheetApp.create("Usuarios_Painel_QHA");
    DriveApp.getFileById(ss.getId()).moveTo(pasta);
  }
  return ss.getSheets()[0];
}

function doPost(e) {
  try {
    var rawData = e.postData.contents;
    var dados = JSON.parse(rawData);
    var tipo = dados.tipo || "registro";

    // Persistência dos usuários (guarda o JSON na célula A1; só guarda hashes de senha)
    if (tipo === "usuarios_carregar" || tipo === "usuarios_salvar") {
      if (!USERS_SYNC_TOKEN || USERS_SYNC_TOKEN.indexOf("TROQUE") === 0 || dados.token !== USERS_SYNC_TOKEN) {
        return _json({ status: "erro", mensagem: "Token inválido" });
      }
      var abaUsers = _abaUsuarios();
      if (tipo === "usuarios_salvar") {
        abaUsers.getRange("A1").setValue(JSON.stringify(dados.usuarios || {}));
        return _json({ status: "sucesso" });
      }
      var bruto = abaUsers.getRange("A1").getValue();
      return _json({ status: "sucesso", usuarios: bruto ? JSON.parse(bruto) : {} });
    }
    
    // Obter ou criar a pasta Logs_Manutencao
    var pastas = DriveApp.getFoldersByName("Logs_Manutencao");
    var pastaLogs = pastas.hasNext() ? pastas.next() : DriveApp.createFolder("Logs_Manutencao");
    
    // Obter ou criar a planilha de histórico/logs
    var arquivos = pastaLogs.getFilesByName("Historico_Logs_Manutencao");
    var spreadsheet;
    if (arquivos.hasNext()) {
      spreadsheet = SpreadsheetApp.open(arquivos.next());
    } else {
      spreadsheet = SpreadsheetApp.create("Historico_Logs_Manutencao");
      var file = DriveApp.getFileById(spreadsheet.getId());
      file.moveTo(pastaLogs);
      
      var abaPadrao = spreadsheet.getActiveSheet();
      abaPadrao.setName("Logs");
      abaPadrao.appendRow(["Data/Hora", "Tipo", "ID_OS", "Equipamento", "Categoria/Frota", "Mecanico/Executante", "Servico/Detalhes", "Foto_Drive_URL"]);
      abaPadrao.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#00379e").setFontColor("#ffffff");
    }
    
    var aba = spreadsheet.getSheetByName("Logs") || spreadsheet.getSheets()[0];
    
    // Se enviou imagem em base64, salvar como arquivo de imagem no Drive
    var fotoUrl = "";
    if (dados.imagem_base64 && dados.imagem_base64.indexOf("data:image") !== -1) {
      try {
        var partes = dados.imagem_base64.split(",");
        var meta = partes[0];
        var base64Content = partes[1];
        var mimeType = meta.split(":")[1].split(";")[0];
        var ext = mimeType.split("/")[1] || "png";
        
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Content), mimeType, "evidencia_" + (dados.equipamento || "OS") + "_" + (new Date().getTime()) + "." + ext);
        var arquivoFoto = pastaLogs.createFile(blob);
        arquivoFoto.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fotoUrl = arquivoFoto.getUrl();
      } catch (errFoto) {
        fotoUrl = "Erro ao salvar imagem: " + errFoto.toString();
      }
    }
    
    var timestamp = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss");
    
    aba.appendRow([
      timestamp,
      tipo.toUpperCase(),
      dados.id_os || "-",
      dados.equipamento || "-",
      dados.frota || "-",
      dados.mecanico || dados.usuario || "-",
      dados.servico || dados.falha || "-",
      fotoUrl
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "sucesso",
      mensagem: "Registro gravado com sucesso no Google Drive",
      foto_url: fotoUrl,
      timestamp: timestamp
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "erro",
      mensagem: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "online",
    servico: "Webhook Google Apps Script - Manutenção Fagundes"
  })).setMimeType(ContentService.MimeType.JSON);
}
