/**
 * S.I.M - Sistema Interno de Manutenção
 * Sistema criado por Clodoaldo Antunes Garcia
 * Criado em 03/11/2025 com direitos autorais reservado ao desenvolvedor 
 * Direitos de uso reservados ao uso do Oscar Inn Eco Resort
 * Desenvolvido em AppsScript
 * Ultima atualização 07/01/2026.
 */


// ⚠️ SUBSTITUA PELO ID DA SUA PLANILHA GOOGLE SE NECESSÁRIO ⚠️
const SPREADSHEET_ID = '1buORzDvwtwOTHk2xN1JpPLhsHiH-k-nDMxMB8XCZuDw'; 

// Nomes das abas (folhas) - Case-sensitive!
const SHEET_USUARIOS = 'Usuarios';
const SHEET_OS = 'OrdensServico';
const SHEET_SETORES = 'Setores';
const SHEET_DIARIO = 'DiarioTecnico';
const SHEET_CONFIGURACOES = 'Configuracoes';
const SHEET_LOGS = 'Logs'; 
const SHEET_MENSAGENS = 'Mensagens'; 
const SHEET_CHAT = 'ChatGlobal';
const FOLDER_NAME = 'SIM_Anexos_OS'; 

function doGet(e) {
  return HtmlService.createTemplateFromFile('PaginaPrincipal')
      .evaluate()
      .setTitle('Sistema de Ordens de Serviço')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ----------------------------------------
 * FUNÇÕES DE CHAT GLOBAL
 * ----------------------------------------
 */
function getGlobalChatHistory(userEmail) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_CHAT);
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_CHAT);
      // Cabeçalho com coluna adicional para attachments
      sheet.appendRow(['Timestamp', 'Email', 'Nome', 'Mensagem', 'Contexto', 'Attachments']);
      sheet.setFrozenRows(1);
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; 

    const startRow = Math.max(2, lastRow - 49);
    const numRows = lastRow - startRow + 1;
    const numCols = Math.max(6, sheet.getLastColumn());
    
    const data = sheet.getRange(startRow, 1, numRows, numCols).getValues();
    
    return data.map(row => {
      let timeStr = "";
      if (row[0] instanceof Date) {
        timeStr = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "dd/MM HH:mm");
      } else {
        timeStr = String(row[0]);
      }
      return {
        timestamp: timeStr,
        fromEmail: row[1],
        fromName: row[2],
        message: row[3],
        context: row[4] || '',
        image: row[5] || '',
        isMe: (String(row[1]).toLowerCase() === String(userEmail).toLowerCase())
      };
    });
  } catch (e) {
    Logger.log('Erro Chat: ' + e.message);
    return [];
  }
}

function sendChatMessage(email, context, message) {
  if (!message || message.trim() === "") return { status: 'falha' };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_CHAT);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_CHAT);
      sheet.appendRow(['Timestamp', 'Email', 'Nome', 'Mensagem', 'Contexto', 'Attachments']);
    }
    const userData = getUserDataByEmail(email);
    const nome = userData ? userData.nome : email;
    sheet.appendRow([new Date(), email, nome, message, context, '']);
    return { status: 'sucesso' };
  } catch (e) {
    return { status: 'falha', message: e.message };
  }
}

/**
 * Faz upload de arquivos enviados via chat para uma pasta específica no Drive
 * Espera files como [{ base64: "data:...;base64,....", name: "arquivo.ext" }, ...]
 * Retorna string com URLs separados por \n
 */
function uploadChatFilesToDrive(files, prefix) {
  try {
    if (!files || !files.length) return '';
    const folderName = FOLDER_NAME + '_Chat';
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const urls = [];
    files.forEach((f, idx) => {
      if (!f || !f.base64) return;
      const parts = String(f.base64).split(',');
      const b64 = parts.length > 1 ? parts[1] : parts[0];
      const contentTypeMatch = String(f.base64).match(/^data:(.*?);base64,/);
      const contentType = contentTypeMatch ? contentTypeMatch[1] : 'application/octet-stream';
      const blob = Utilities.newBlob(Utilities.base64Decode(b64), contentType, `${prefix || 'chat'}_${new Date().getTime()}_${idx}_${(f.name || 'file')}`);
      const file = folder.createFile(blob);
      urls.push(file.getUrl());
    });
    return urls.join('\n');
  } catch (e) {
    Logger.log('uploadChatFilesToDrive error: ' + e.message);
    return '';
  }
}

/**
 * Função aprimorada para enviar mensagens de chat que podem incluir:
 * - message: texto
 * - files: array de { base64, name } (arquivos em dataURL/base64)
 *
 * Salva no ChatGlobal: Timestamp, Email, Nome, Mensagem, Contexto, Attachments(URLs separados por \n)
 */
function sendChatMessageAdvanced(email, context, message, files) {
  if ((!message || message.trim() === "") && (!files || files.length === 0)) return { status: 'falha', message: 'Mensagem e anexos vazios.' };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_CHAT);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_CHAT);
      sheet.appendRow(['Timestamp', 'Email', 'Nome', 'Mensagem', 'Contexto', 'Attachments']);
    }
    const userData = getUserDataByEmail(email);
    const nome = userData ? userData.nome : email;

    // Se houver arquivos, faz upload e obtém URLs
    let attachmentsUrls = '';
    if (files && files.length) {
      attachmentsUrls = uploadChatFilesToDrive(files, `CHAT_${email.replace(/[^a-z0-9]/gi,'')}`);
    }

    sheet.appendRow([new Date(), email, nome, message || '', context || '', attachmentsUrls || '']);
    return { status: 'sucesso' };
  } catch (e) {
    return { status: 'falha', message: e.message };
  }
}

/**
 * ----------------------------------------
 * LOGS E AUDITORIA
 * ----------------------------------------
 */
function logAction(email, acao, detalhes = '') {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_LOGS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_LOGS);
      sheet.appendRow(['Timestamp', 'Email Utilizador', 'Ação', 'Detalhes']);
      sheet.setFrozenRows(1);
    }
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, 4).setValues([[new Date(), email, acao, detalhes]]);
    if (sheet.getLastRow() > 1001) {
      sheet.deleteRows(1002, sheet.getLastRow() - 1001);
    }
  } catch (e) {
    Logger.log('Log error: ' + e.message);
  }
}

function getActionLogs(callerEmail) {
  if (verificarNivel(callerEmail).toLowerCase() !== 'admin') return [];
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_LOGS);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    const numRowsToGet = Math.min(sheet.getLastRow() - 1, 100);
    const data = sheet.getRange(2, 1, numRowsToGet, 4).getValues();
    return data.map(row => {
      if (row[0] instanceof Date) row[0] = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      return row;
    });
  } catch (e) { return []; }
}

/**
 * ----------------------------------------
 * AUTENTICAÇÃO E USUÁRIOS
 * ----------------------------------------
 */
function getUserDataByEmail(email) {
  if (!email) return null;
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USUARIOS);
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
        return { email: data[i][0], senha: data[i][1], nome: data[i][2], nivel: data[i][3] };
      }
    }
  } catch (e) { return null; }
  return null;
}

function verificarNivel(email) {
  const u = getUserDataByEmail(email);
  return u ? u.nivel : null;
}

function verificarLogin(nome, senha) {
  if (!nome || !senha) return { status: 'falha', mensagem: 'Dados incompletos.' };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USUARIOS);
    if (!sheet) return { status: 'falha', mensagem: 'Base de dados não encontrada.' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === nome.toLowerCase().trim() && String(data[i][1]) === senha) {
        return { 
          status: 'sucesso', 
          email: data[i][0], 
          nome: data[i][2], 
          nivel: data[i][3], 
          popUpConfig: getConfiguracaoPopUp() 
        };
      }
    }
    return { status: 'falha', mensagem: 'Credenciais inválidas.' };
  } catch (e) { return { status: 'falha', mensagem: 'Erro servidor.' }; }
}

function getAllUsers(callerEmail) {
  if (verificarNivel(callerEmail).toLowerCase() !== 'admin') return [];
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USUARIOS);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues().map(r => [r[0], r[2], r[3]]);
  } catch (e) { return []; }
}

function saveUser(callerEmail, u) {
  if (verificarNivel(callerEmail).toLowerCase() !== 'admin') return { status: 'falha', message: 'Negado.' };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USUARIOS);
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === u.email.toLowerCase()) {
        sheet.getRange(i+1, 3).setValue(u.nome);
        sheet.getRange(i+1, 4).setValue(u.nivel);
        if(u.senha && u.senha.length>=4) sheet.getRange(i+1, 2).setValue(u.senha);
        found = true; break;
      }
    }
    if (!found) sheet.appendRow([u.email, u.senha, u.nome, u.nivel]);
    logAction(callerEmail, 'Save User', u.email);
    return { status: 'sucesso', message: 'Salvo.' };
  } catch (e) { return { status: 'falha', message: e.message }; }
}

function deleteUser(callerEmail, email) {
  if (verificarNivel(callerEmail).toLowerCase() !== 'admin') return { status: 'falha', message: 'Negado.' };
  if (callerEmail.toLowerCase() === email.toLowerCase()) return { status: 'falha', message: 'Auto-exclusão proibida.' };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USUARIOS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
        sheet.deleteRow(i+1);
        return { status: 'sucesso', message: 'Excluído.' };
      }
    }
    return { status: 'falha', message: 'Não encontrado.' };
  } catch (e) { return { status: 'falha', message: e.message }; }
}

function changeUserPassword(email, oldP, newP) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USUARIOS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
        if (String(data[i][1]) === oldP) {
          sheet.getRange(i+1, 2).setValue(newP);
          return { status: 'sucesso', message: 'Senha alterada.' };
        } else return { status: 'falha', message: 'Senha atual incorreta.' };
      }
    }
  } catch (e) { return { status: 'falha', message: 'Erro.' }; }
}

/**
 * ----------------------------------------
 * MENSAGENS E CONFIG
 * ----------------------------------------
 */
function sendMessageToUser(sender, recipient, msg) {
  if (verificarNivel(sender).toLowerCase() !== 'admin') return { status: 'falha', message: 'Negado.' };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_MENSAGENS);
    if(!sheet) { sheet=ss.insertSheet(SHEET_MENSAGENS); sheet.appendRow(['Ts','DeEmail','DeNome','Para','Msg','Lido']); }
    const senderData = getUserDataByEmail(sender);
    sheet.appendRow([new Date(), sender, senderData.nome, recipient, msg, 'NÃO']);
    return { status: 'sucesso', message: 'Enviada.' };
  } catch (e) { return { status: 'falha', message: e.message }; }
}

function getMessagesForUser(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_MENSAGENS);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
    return data.filter(r => String(r[3]).toLowerCase() === email.toLowerCase())
               .map((r, i) => [
                   r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "dd/MM HH:mm") : r[0],
                   r[2], r[4], i+2
               ]).reverse();
  } catch (e) { return []; }
}

function checkForUnreadMessages(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_MENSAGENS);
    if (!sheet || sheet.getLastRow() <= 1) return false;
    const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
    return data.some(r => String(r[3]).toLowerCase() === email.toLowerCase() && String(r[5]).toUpperCase() === 'NÃO');
  } catch (e) { return false; }
}

function markMessagesAsRead(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_MENSAGENS);
    if (!sheet || sheet.getLastRow() <= 1) return;
    const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
    data.forEach((r, i) => {
      if(String(r[3]).toLowerCase() === email.toLowerCase() && String(r[5]).toUpperCase() === 'NÃO') {
        sheet.getRange(i+2, 6).setValue('SIM');
      }
    });
  } catch(e) {}
}

function getConfiguracaoPopUp() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_CONFIGURACOES);
    if(!sheet) return { habilitado: 'FALSE', mensagem: '' };
    const data = sheet.getDataRange().getValues();
    let c = { habilitado: 'FALSE', mensagem: '', imageUrl: '' };
    data.forEach(r => {
      if(r[0]==='POPUP_HABILITADO') c.habilitado=String(r[1]);
      if(r[0]==='POPUP_MENSAGEM') c.mensagem=String(r[1]);
      if(r[0]==='POPUP_IMAGE_URL') c.imageUrl=String(r[1]);
    });
    return c;
  } catch(e) { return { habilitado: 'FALSE' }; }
}

function saveConfiguracaoPopUp(caller, c) {
  if (verificarNivel(caller).toLowerCase() !== 'admin') return { status: 'falha' };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_CONFIGURACOES);
    if(!sheet) { sheet=ss.insertSheet(SHEET_CONFIGURACOES); sheet.appendRow(['Key','Val']); }
    const data = sheet.getDataRange().getValues();
    const set = (k, v) => {
        let f=false; 
        for(let i=0; i<data.length; i++) { if(data[i][0]===k) { sheet.getRange(i+1, 2).setValue(v); f=true; break; } }
        if(!f) sheet.appendRow([k,v]);
    };
    set('POPUP_HABILITADO', c.habilitado); set('POPUP_MENSAGEM', c.mensagem); set('POPUP_IMAGE_URL', c.imageUrl);
    return { status: 'sucesso', message: 'Salvo.' };
  } catch(e) { return { status: 'falha', message: e.message }; }
}

/**
 * ----------------------------------------
 * ORDENS DE SERVIÇO (O CORAÇÃO DO SISTEMA)
 * ----------------------------------------
 */
function uploadMultipleFilesToDrive(files, osId) {
  try {
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
    return files.map((f, i) => {
       const b64 = f.base64.split(',')[1] || f.base64;
       const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', `OS_${osId}_${i}_${f.name}`);
       return folder.createFile(blob).getUrl();
    }).join('\n');
  } catch (e) { return ''; }
}

function gerarOS(data) {
  const nivel = verificarNivel(data.solicitanteEmail);
  if (!nivel) return { status: 'falha', message: 'Negado.' };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_OS);
    const lastRow = sheet.getLastRow();
    const newId = (lastRow > 1) ? Number(sheet.getRange(lastRow, 1).getValue()) + 1 : 1;
    
    let url = '';
    if(data.filesData && data.filesData.length) url = uploadMultipleFilesToDrive(data.filesData, newId);
    
    const tec = (['recepcao','solicitante'].includes(nivel.toLowerCase())) ? "" : data.tecnico;
    const obs = `Criado em ${new Date().toLocaleString('pt-BR')} por ${data.solicitanteNome}`;
    
    sheet.appendRow([newId, new Date(), data.solicitanteEmail, data.solicitanteNome, data.setor, data.descricao, 'Aberta', tec, obs, data.prioridade, new Date(), url]);
    logAction(data.solicitanteEmail, 'Nova OS', `#${newId}`);
    return { status: 'sucesso', mensagem: `OS #${newId} gerada!` };
  } catch (e) { return { status: 'falha', mensagem: e.message }; }
}

function atualizarOS(id, st, obs, tec, prio, editor, files) {
  if(!['admin','gerente','tecnico'].includes(verificarNivel(editor).toLowerCase())) return { status: 'falha', message: 'Negado.' };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_OS);
    const data = sheet.getDataRange().getValues();
    let row = -1;
    for(let i=1; i<data.length; i++) { if(data[i][0] == id) { row = i+1; break; } }
    if(row === -1) return { status: 'falha', message: 'Não encontrado.' };
    
    let url = '';
    if(files && files.length) url = uploadMultipleFilesToDrive(files, id);
    
    sheet.getRange(row, 7).setValue(st);
    sheet.getRange(row, 8).setValue(tec);
    sheet.getRange(row, 11).setValue(new Date());
    if(prio && ['admin','gerente'].includes(verificarNivel(editor).toLowerCase())) sheet.getRange(row, 10).setValue(prio);
    
    if(url) {
        const old = sheet.getRange(row, 12).getValue();
        sheet.getRange(row, 12).setValue(old ? old + '\n' + url : url);
    }
    if(obs) {
        const oldObs = sheet.getRange(row, 9).getValue();
        sheet.getRange(row, 9).setValue(oldObs + `\n[${new Date().toLocaleString()}] ${obs}`);
    }
    logAction(editor, 'Update OS', `#${id}`);
    return { status: 'sucesso', message: 'Atualizado.' };
  } catch(e) { return { status: 'falha', message: e.message }; }
}

/**
 * ----------------------------------------
 * CORREÇÃO: LEITURA SEGURA DE DADOS (COM CONVERSÃO DE DATA)
 * ----------------------------------------
 */
function filtrarOSPainel(filtros) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_OS);
    
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    
    // Se a planilha estiver vazia (só cabeçalho ou nem isso), retorna vazio
    if (lastRow <= 1) return [];

    // Busca dados brutos. Limitamos a 12 colunas para garantir estrutura.
    let data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    
    // Filtro por Status
    // CORREÇÃO: Verifica se é diferente de 'Todos' E 'Todas' para cobrir inconsistências
    if (filtros.status === 'Ativas') {
        data = data.filter(r => r[6] !== 'Concluída' && r[6] !== 'Cancelada');
    } else if (filtros.status && filtros.status !== 'Todos' && filtros.status !== 'Todas') {
        data = data.filter(r => r[6] === filtros.status);
    }
    
    // Filtro por Setor
    if (filtros.setor && filtros.setor !== 'Todos') data = data.filter(r => r[4] === filtros.setor);
    
    // Filtro por Data (Hoje)
    if (filtros.data === 'Hoje') {
       const today = new Date(); today.setHours(0,0,0,0);
       data = data.filter(r => { 
         if (r[1] instanceof Date) {
           const d = new Date(r[1]); d.setHours(0,0,0,0); 
           return d.getTime() === today.getTime(); 
         }
         return false;
       });
    }

    // --- NOVA ORDENAÇÃO: Prioridade Alta PRIMEIRO, depois ID decrescente ---
    data.sort((a, b) => {
        // Coluna 9 (índice 9) é a Prioridade (J)
        const prioA = String(a[9] || '').toLowerCase();
        const prioB = String(b[9] || '').toLowerCase();
        
        // Se A é Alta e B não, A vem primeiro (-1)
        if (prioA === 'alta' && prioB !== 'alta') return -1;
        // Se B é Alta e A não, B vem primeiro (1)
        if (prioA !== 'alta' && prioB === 'alta') return 1;
        
        // Se ambos são iguais (ambos alta ou ambos não alta), ordena por ID decrescente
        return (Number(b[0]) || 0) - (Number(a[0]) || 0);
    });
    
    // CRUCIAL: Converter todas as datas para String antes de retornar ao Client
    return data.map(r => {
        // Coluna 2 (Index 1): Data Criação
        if(r[1] instanceof Date) r[1] = Utilities.formatDate(r[1], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        // Coluna 11 (Index 10): Data Atualização
        if(r[10] instanceof Date) r[10] = Utilities.formatDate(r[10], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        return r;
    });
  } catch (e) { 
    Logger.log("Erro no Painel: " + e.message);
    return []; 
  }
}

function filtrarOS(filtros) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_OS);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    let data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    
    if (filtros.status !== 'Todos') data = data.filter(r => r[6] === filtros.status);
    
    // ALTERAÇÃO AQUI: Troca de filtro Técnico por Setor (Coluna 4)
    if (filtros.setor && filtros.setor !== 'Todos') data = data.filter(r => r[4] === filtros.setor);
    
    if (filtros.dataInicio) {
        const d1 = new Date(filtros.dataInicio); d1.setHours(0,0,0,0);
        data = data.filter(r => (r[1] instanceof Date) && r[1].getTime() >= d1.getTime());
    }
    if (filtros.dataFim) {
        const d2 = new Date(filtros.dataFim); d2.setHours(23,59,59,999);
        data = data.filter(r => (r[1] instanceof Date) && r[1].getTime() <= d2.getTime());
    }

    // Ordenação Decrescente por ID (Coluna 0)
    data.sort((a, b) => {
        return (Number(b[0]) || 0) - (Number(a[0]) || 0);
    });
    
    // CRUCIAL: Converter todas as datas para String
    return data.map(r => {
        if(r[1] instanceof Date) r[1] = Utilities.formatDate(r[1], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        if(r[10] instanceof Date) r[10] = Utilities.formatDate(r[10], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        return r;
    });
  } catch(e) { 
    Logger.log("Erro no Relatório: " + e.message);
    return []; 
  }
}

function getLatestOSID() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_OS);
    const lastRow = sheet.getLastRow();
    if(lastRow <= 1) return 0;
    return sheet.getRange(lastRow, 1).getValue();
  } catch(e) { return 0; }
}

function getUpdatesForTecnico(nome, lastCheck) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_OS);
    const lastRow = sheet.getLastRow();
    
    if(lastRow <= 1) return { updateFound: false, newTimestamp: new Date().toISOString() };
    
    // Coluna 11 (K) é a de Data de Atualização
    const lastModCol = sheet.getRange(2, 11, lastRow-1, 1).getValues();
    let max = new Date(0);
    
    lastModCol.forEach(r => { 
      if(r[0] instanceof Date && r[0] > max) max = r[0]; 
    });
    
    if (max.getTime() > new Date(lastCheck).getTime()) {
        // Alterado para false para silenciar o aviso visual no Frontend
        return { updateFound: false, updateType: 'local', newTimestamp: max.toISOString() };
    }
    return { updateFound: false, newTimestamp: new Date().toISOString() };
  } catch(e) { return { updateFound: false }; }
}

function getTecnicos() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USUARIOS);
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow-1, 4).getValues();
    return data.filter(r => String(r[3]).toLowerCase()==='tecnico').map(r => r[2]);
  } catch(e) { return []; }
}

function getSetores() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_SETORES);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    return sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat().filter(String);
  } catch(e) { return []; }
}

function filtrarRotinas(f) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_DIARIO);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    // Pega dados brutos (4 colunas: Data, Técnico, Descrição, Anexos)
    let rawData = sheet.getRange(2, 1, sheet.getLastRow()-1, 4).getValues();
    
    // Anexa o índice da linha REAL (index + 2) antes de filtrar ou ordenar
    // Isso garante que Edição/Exclusão afetem a linha correta na planilha
    let data = rawData.map((r, i) => {
        r.push(i + 2); // Adiciona index na posição 4
        return r;
    });
    
    // Filtro por Técnico
    if(f.tecnico !== 'Todos') {
        data = data.filter(r => r[1] === f.tecnico);
    }

    // Ordenação Decrescente por Data (Coluna 0) - Melhorada para lidar com Strings e Datas
    data.sort((a, b) => {
        // Função auxiliar para obter timestamp seguro
        const getTime = (val) => {
            if (val instanceof Date) return val.getTime();
            const d = new Date(val);
            return isNaN(d.getTime()) ? 0 : d.getTime();
        };

        const dateA = getTime(a[0]);
        const dateB = getTime(b[0]);
        
        return dateB - dateA;
    });
    
    // Formatação da Data para String (Incluindo Horário)
    return data.map(r => {
        // Se já for data, formata
        if(r[0] instanceof Date) {
            r[0] = Utilities.formatDate(r[0], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        } 
        // Se for string que pode ser convertida, converte e formata para padronizar
        else {
            const d = new Date(r[0]);
            if (!isNaN(d.getTime())) {
                r[0] = Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
            }
        }
        return r;
    });
  } catch(e) { return []; }
}

function salvarDiarioTecnico(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let url = data.filesData ? uploadMultipleFilesToDrive(data.filesData, 'Diario') : '';
    ss.getSheetByName(SHEET_DIARIO).appendRow([data.data, data.tecnicoNome, data.descricao, url]);
    return { status: 'sucesso', message: 'Salvo.' };
  } catch(e) { return { status: 'falha' }; }
}

function atualizarRotina(c, row, desc) {
  if (verificarNivel(c).toLowerCase() !== 'admin') return { status: 'falha' };
  try { SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DIARIO).getRange(row, 3).setValue(desc); return {status:'sucesso'}; } catch(e){return {status:'falha'};}
}

function excluirRotina(c, row) {
  if (verificarNivel(c).toLowerCase() !== 'admin') return { status: 'falha' };
  try { SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DIARIO).deleteRow(row); return {status:'sucesso'}; } catch(e){return {status:'falha'};}
}

function getDashboardStatistics(e, d1, d2) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const osSheet = ss.getSheetByName(SHEET_OS);
    const diarioSheet = ss.getSheetByName(SHEET_DIARIO);
    const userSheet = ss.getSheetByName(SHEET_USUARIOS);
    
    if(!osSheet || osSheet.getLastRow() <= 1) return { total:0, chartData:[] };

    const osData = osSheet.getRange(2, 1, osSheet.getLastRow()-1, 10).getValues(); // Inclui técnico na col 8 (index 7)
    let total=0, abertas=0, andamento=0, concluidas=0;
    
    // Contadores para o gráfico de técnicos
    let techStats = {}; // { "Nome": { os: 0, diario: 0 } }
    
    // Processa OS
    osData.forEach(r => {
        if (r[1] instanceof Date) {
          const dt = r[1];
          if((!d1 || dt >= new Date(d1)) && (!d2 || dt <= new Date(d2))) {
              total++;
              const status = String(r[6]);
              if(status==='Aberta') abertas++;
              else if(status==='Em Andamento') andamento++;
              else if(status==='Concluída') concluidas++;
              
              const techName = String(r[7] || '').trim();
              if (techName && techName !== '-' && techName !== '') {
                  if (!techStats[techName]) techStats[techName] = { os: 0, diario: 0 };
                  techStats[techName].os++;
              }
          }
        }
    });
    
    // Processa Diário Técnico
    if (diarioSheet && diarioSheet.getLastRow() > 1) {
        const diarioData = diarioSheet.getRange(2, 1, diarioSheet.getLastRow()-1, 2).getValues();
        diarioData.forEach(r => {
            const dt = (r[0] instanceof Date) ? r[0] : new Date(r[0]);
            if (!isNaN(dt.getTime())) {
                if((!d1 || dt >= new Date(d1)) && (!d2 || dt <= new Date(d2))) {
                    const techName = String(r[1] || '').trim();
                    if (techName) {
                        if (!techStats[techName]) techStats[techName] = { os: 0, diario: 0 };
                        techStats[techName].diario++;
                    }
                }
            }
        });
    }

    // Prepara dados para o gráfico de técnicos
    let techPerformanceData = [['Técnico', 'Atendimentos OS', 'Registros Diário']];
    Object.keys(techStats).forEach(name => {
        techPerformanceData.push([name, techStats[name].os, techStats[name].diario]);
    });

    return { 
        total: total, 
        abertas: abertas, 
        andamento: andamento, 
        concluidas: concluidas, 
        chartData: [['Status','Quantidade'], ['Aberta',abertas], ['Andamento',andamento], ['Concluída',concluidas]],
        techPerformanceData: techPerformanceData
    };
  } catch(e) { 
    Logger.log("Erro Stats: " + e.message);
    return { total:0, chartData:[], techPerformanceData: [] }; 
  }
}
