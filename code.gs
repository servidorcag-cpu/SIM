/**
 * S.I.M - Sistema Interno de Manutenção
 * Sistema criado por Clodoaldo Antunes Garcia
 * Criado em 03/11/2025 com direitos autorais reservado ao desenvolvedor 
 * Direitos de uso reservados ao uso do Oscar Inn Eco Resort
 * Desenvolvido em AppsScript / aistudio
 * Ultima atualização: Março 2026. Versão 2.20 (COMPLETA E OTIMIZADA)
 */

// --- CONFIGURAÇÕES GLOBAIS ---
const SPREADSHEET_ID = '1buORzDvwtwOTHk2xN1JpPLhsHiH-k-nDMxMB8XCZuDw'; 
const SYSTEM_EXPIRATION_DATE = new Date('2027-01-01T00:00:00.000-03:00'); 
const SHEET_USUARIOS = 'Usuarios';
const SHEET_OS = 'OrdensServico';
const SHEET_SETORES = 'Setores';
const SHEET_DIARIO = 'DiarioTecnico';
const SHEET_CONFIGURACOES = 'Configuracoes';
const SHEET_LOGS = 'Logs'; 
const SHEET_MENSAGENS = 'Mensagens'; 
const SHEET_CHAT = 'ChatGlobal';
const SHEET_DAILY_STATS = 'RelatorioDiarioTecnicos';
const FOLDER_NAME = 'SIM_Anexos_OS'; 

function doGet(e) {
  const today = new Date();
  if (today >= SYSTEM_EXPIRATION_DATE) {
    return HtmlService.createHtmlOutput('<html><body><div style="text-align:center; margin-top:50px;"><h1>Sistema Expirado</h1><p>Contate o desenvolvedor.</p></div></body></html>');
  }
  return HtmlService.createTemplateFromFile('PaginaPrincipal').evaluate().setTitle('S.I.M - Manutenção').setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// =================================================================
// FUNÇÃO AUXILIAR DE DATAS (PREVINE ERROS DE FILTRO)
// =================================================================
function parseDateSafe(val) {
  if (val instanceof Date) return val;
  if (!val) return null;
  const d = new Date(val);
  if (!isNaN(d.getTime())) return d;
  if (typeof val === 'string' && val.includes('/')) {
    const parts = val.split(' ')[0].split('/'); 
    if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0], 0, 0, 0); 
  }
  return null;
}

// =================================================================
// DADOS INICIAIS E CACHE
// =================================================================
function getInitialData(userEmail) {
  try {
    const user = getUserDataByEmail(userEmail);
    if (!user) return { error: "Usuário não encontrado." };
    const defaultFilters = { status: "Ativas", setor: "Todos", data: "Todas", tecnico: "Todos", keyword: "" };
    
    // Calcular dados de expiração
    const today = new Date();
    const timeDiff = SYSTEM_EXPIRATION_DATE.getTime() - today.getTime();
    const daysDiff = Math.ceil(timeDiff / (1000 * 3600 * 24));
    const expirationDateFormatted = Utilities.formatDate(SYSTEM_EXPIRATION_DATE, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const expirationWarningDays = (daysDiff <= 30 && daysDiff > 0) ? daysDiff : null;
    
    return {
      setores: getSetores(), 
      tecnicos: getTecnicos(), 
      chatHistory: getGlobalChatHistory(userEmail),
      hasUnreadPrivateMessages: checkForUnreadMessages(userEmail),
      painelOS: (['tecnico', 'admin', 'gerente', 'recepcao'].includes(user.nivel.toLowerCase())) ? filtrarOSPainel(defaultFilters, 1, 50) : { items: [], totalItems: 0 },
      user: user, 
      allUserNames: getAllUserNames(),
      expirationDateFormatted: expirationDateFormatted,
      expirationWarningDays: expirationWarningDays
    };
  } catch (e) { return { error: e.message }; }
}

function getSetores() {
  const cache = CacheService.getScriptCache();
  if (cache.get('cache_setores')) return JSON.parse(cache.get('cache_setores'));
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_SETORES);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues().flat().filter(String);
    cache.put('cache_setores', JSON.stringify(data), 21600); return data;
  } catch(e) { return []; }
}

function getTecnicos() {
  const cache = CacheService.getScriptCache();
  if (cache.get('cache_tecnicos')) return JSON.parse(cache.get('cache_tecnicos'));
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USUARIOS);
    const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 4).getValues();
    const tecnicos = data.filter(r => String(r[3]).toLowerCase()==='tecnico').map(r => r[2]);
    cache.put('cache_tecnicos', JSON.stringify(tecnicos), 21600); return tecnicos;
  } catch(e) { return []; }
}

function getAllUserNames() {
  const cache = CacheService.getScriptCache();
  if (cache.get('cache_usernames')) return JSON.parse(cache.get('cache_usernames'));
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USUARIOS);
    const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues().flat().filter(String);
    cache.put('cache_usernames', JSON.stringify(data), 21600); return data;
  } catch (e) { return []; }
}

function limparCacheUsuarios() {
  CacheService.getScriptCache().remove('cache_tecnicos');
  CacheService.getScriptCache().remove('cache_usernames');
}

// =================================================================
// LOGS E AUDITORIA
// =================================================================
function logAction(email, acao, detalhes = '') {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_LOGS);
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, 4).setValues([[new Date(), email, acao, detalhes]]);
  } catch (e) { }
}

function getActionLogs(callerEmail, filtros) {
  if (verificarNivel(callerEmail).toLowerCase() !== 'admin') return [];
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_LOGS);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();

    if (filtros && filtros.usuarioEmail && filtros.usuarioEmail !== 'Todos') {
      data = data.filter(r => String(r[1]).toLowerCase() === String(filtros.usuarioEmail).toLowerCase());
    }
    
    if (filtros && filtros.dataInicio) {
      const parts = filtros.dataInicio.split('-');
      const d1 = new Date(parts[0], parts[1]-1, parts[2], 0,0,0);
      data = data.filter(r => parseDateSafe(r[0]) && parseDateSafe(r[0]).getTime() >= d1.getTime());
    }
    if (filtros && filtros.dataFim) {
      const parts = filtros.dataFim.split('-');
      const d2 = new Date(parts[0], parts[1]-1, parts[2], 23,59,59);
      data = data.filter(r => parseDateSafe(r[0]) && parseDateSafe(r[0]).getTime() <= d2.getTime());
    }
    
    data.sort((a, b) => (parseDateSafe(b[0])?.getTime() || 0) - (parseDateSafe(a[0])?.getTime() || 0));
    return data.slice(0, 200).map(row => {
      row[0] = parseDateSafe(row[0]) ? Utilities.formatDate(parseDateSafe(row[0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss") : row[0];
      return row;
    });
  } catch (e) { return []; }
}

// =================================================================
// AUTENTICAÇÃO E USUÁRIOS
// =================================================================
function getUserDataByEmail(email) {
  if (!email) return null;
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USUARIOS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === email.toLowerCase()) return { email: data[i][0], senha: data[i][1], nome: data[i][2], nivel: data[i][3] };
    }
  } catch (e) { return null; }
  return null;
}

function verificarNivel(email) {
  const u = getUserDataByEmail(email); return u ? u.nivel : null;
}

function verificarLogin(nome, senha) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USUARIOS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === nome.toLowerCase().trim() && String(data[i][1]) === senha) {
        return { status: 'sucesso', email: data[i][0], nome: data[i][2], nivel: data[i][3], popUpConfig: getConfiguracaoPopUp() };
      }
    }
    return { status: 'falha', mensagem: 'Credenciais inválidas.' };
  } catch (e) { return { status: 'falha', mensagem: 'Erro servidor.' }; }
}

function getAllUsers(callerEmail) {
  if (verificarNivel(callerEmail).toLowerCase() !== 'admin') return [];
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USUARIOS);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  } catch (e) { return []; }
}

function saveUser(callerEmail, u) {
  if (verificarNivel(callerEmail).toLowerCase() !== 'admin') return { status: 'falha', message: 'Acesso negado.' };
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USUARIOS);
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === u.email.toLowerCase()) {
        sheet.getRange(i+1, 3).setValue(u.nome);
        sheet.getRange(i+1, 4).setValue(u.nivel);
        if(u.senha && u.senha.length >= 4) sheet.getRange(i+1, 2).setValue(u.senha);
        found = true; break;
      }
    }
    if (!found) sheet.appendRow([u.email, u.senha, u.nome, u.nivel]);
    logAction(callerEmail, 'Salvar Utilizador', u.email);
    limparCacheUsuarios(); 
    return { status: 'sucesso', message: 'Utilizador salvo.' };
  } catch (e) { return { status: 'falha', message: e.message }; }
}

function deleteUser(callerEmail, email) {
  if (verificarNivel(callerEmail).toLowerCase() !== 'admin') return { status: 'falha', message: 'Acesso negado.' };
  if (callerEmail.toLowerCase() === email.toLowerCase()) return { status: 'falha', message: 'Não pode excluir a si próprio.' };
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USUARIOS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
        sheet.deleteRow(i+1);
        limparCacheUsuarios();
        return { status: 'sucesso', message: 'Excluído com sucesso.' };
      }
    }
    return { status: 'falha', message: 'Utilizador não encontrado.' };
  } catch (e) { return { status: 'falha', message: e.message }; }
}

function changeUserPassword(email, oldPass, newPass) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_USUARIOS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
        if (String(data[i][1]) !== oldPass) return { status: 'falha', message: 'Senha atual incorreta.' };
        sheet.getRange(i+1, 2).setValue(newPass);
        logAction(email, 'Alterar Senha');
        return { status: 'sucesso', message: 'Senha alterada com sucesso.' };
      }
    }
    return { status: 'falha', message: 'Utilizador não encontrado.' };
  } catch(e) { return { status: 'falha', message: e.message }; }
}

// =================================================================
// GESTÃO DE O.S. (CRIAR, LER, ATUALIZAR)
// =================================================================
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
  if (!nivel) return { status: 'falha', message: 'Acesso negado.' };
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_OS);
    const lastRow = sheet.getLastRow();
    const newId = (lastRow > 1) ? Number(sheet.getRange(lastRow, 1).getValue()) + 1 : 1;
    
    let url = '';
    if(data.filesData && data.filesData.length) url = uploadMultipleFilesToDrive(data.filesData, newId);
    
    const tec = (['recepcao','solicitante'].includes(nivel.toLowerCase())) ? "" : data.tecnico;
    const obs = `Criada em ${new Date().toLocaleString('pt-PT')} por ${data.solicitanteNome}`;
    
    const dataMovimentacao = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    sheet.appendRow([newId, new Date(), data.solicitanteEmail, data.solicitanteNome, data.setor, data.descricao, 'Aberta', tec, obs, data.prioridade, new Date(), url, '', '', dataMovimentacao]);
    logAction(data.solicitanteEmail, 'Gerar OS', `#${newId}`);
    return { status: 'sucesso', mensagem: `O.S. #${newId} gerada com sucesso!` };
  } catch (e) { return { status: 'falha', mensagem: e.message }; }
}

function atualizarOS(id, st, obs, tec, prio, editor, files) {
  if(!['admin','gerente','tecnico'].includes(verificarNivel(editor).toLowerCase())) return { status: 'falha', message: 'Acesso negado.' };
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_OS);
    const lastRow = sheet.getLastRow();
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let row = -1;
    for(let i=0; i<ids.length; i++) { if(ids[i][0] == id) { row = i+2; break; } }
    if(row === -1) return { status: 'falha', message: 'O.S. não encontrada.' };
    
    let url = '';
    if(files && files.length) url = uploadMultipleFilesToDrive(files, id);
    
    const dataMovimentacao = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    sheet.getRange(row, 7).setValue(st);
    sheet.getRange(row, 8).setValue(tec);
    sheet.getRange(row, 11).setValue(new Date());
    sheet.getRange(row, 15).setValue(dataMovimentacao);
    if(prio && ['admin','gerente'].includes(verificarNivel(editor).toLowerCase())) sheet.getRange(row, 10).setValue(prio);
    
    if(url) {
        const old = sheet.getRange(row, 12).getValue();
        sheet.getRange(row, 12).setValue(old ? old + '\n' + url : url);
    }
    if(obs) {
        const oldObs = sheet.getRange(row, 9).getValue();
        sheet.getRange(row, 9).setValue(oldObs + `\n[${new Date().toLocaleString('pt-PT')}] ${obs}`);
    }
    logAction(editor, 'Atualizar OS', `#${id}`);
    return { status: 'sucesso', message: 'O.S. atualizada.' };
  } catch(e) { return { status: 'falha', message: e.message }; }
}

function getOSById(id) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_OS);
    if (!sheet || sheet.getLastRow() <= 1) return null;
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    for (const row of data) {
      if (row[0] == id) {
        row[1] = parseDateSafe(row[1]) ? Utilities.formatDate(parseDateSafe(row[1]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : row[1];
        row[10] = parseDateSafe(row[10]) ? Utilities.formatDate(parseDateSafe(row[10]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : row[10];
        return row;
      }
    }
  } catch(e) { return null; }
  return null;
}

function marcarOSComoVista(id, nomeUsuario) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_OS);
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for(let i=0; i<ids.length; i++) {
      if(ids[i][0] == id) {
        const cell = sheet.getRange(i+2, 13);
        let current = String(cell.getValue() || '');
        if(!current.includes(nomeUsuario)) {
           cell.setValue(current ? current + ', ' + nomeUsuario : nomeUsuario);
        }
        break;
      }
    }
  } catch(e) {}
}

function getUpdatesForTecnico(nome, lastCheck) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_OS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { updateFound: false };

    const startRow = Math.max(2, lastRow - 99);
    const numRows = lastRow - startRow + 1;
    const data = sheet.getRange(startRow, 1, numRows, 11).getValues();
    const lastCheckDate = new Date(lastCheck);
    let mostRecentUpdate = null;

    for (const row of data) {
      if (row[7] === nome && parseDateSafe(row[10]) && parseDateSafe(row[10]).getTime() > lastCheckDate.getTime()) {
        if (!mostRecentUpdate || parseDateSafe(row[10]).getTime() > mostRecentUpdate.updateTimestamp.getTime()) {
          mostRecentUpdate = { osId: row[0], updateTimestamp: parseDateSafe(row[10]) };
        }
      }
    }
    if (mostRecentUpdate) return { updateFound: true, osId: mostRecentUpdate.osId, newTimestamp: mostRecentUpdate.updateTimestamp.toISOString() };
    return { updateFound: false, newTimestamp: new Date().toISOString() };
  } catch (e) { return { updateFound: false }; }
}

function checkForAssignedOS(tecnicoNome, lastCheckedIdStr) {
  try {
    const lastCheckedId = Number(lastCheckedIdStr) || 0;
    if (!tecnicoNome) return null;

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_OS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null;

    const startRow = Math.max(2, lastRow - 29);
    const numRows = lastRow - startRow + 1;
    const data = sheet.getRange(startRow, 1, numRows, 8).getValues(); 

    let newlyAssignedOS = null;
    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      if (String(row[7]).trim() === tecnicoNome && Number(row[0]) > lastCheckedId) {
        if (!newlyAssignedOS || Number(row[0]) > newlyAssignedOS.id) {
            newlyAssignedOS = { id: Number(row[0]), solicitante: row[3], setor: row[4], descricao: String(row[5]).substring(0, 50) + '...' };
        }
      }
    }
    return newlyAssignedOS; 
  } catch (e) { return null; }
}


// =================================================================
// ORDENS DE SERVIÇO (PAINEL E RELATÓRIOS)
// =================================================================
function filtrarOSPainel(filtros, page = 1, pageSize = 50) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_OS);
    if (!sheet || sheet.getLastRow() <= 1) return { items: [], totalItems: 0 };
    let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    
    data = data.filter(r => r[0] && r[0] !== ''); 
    
    if (filtros.status === 'Ativas') data = data.filter(r => r[6] !== 'Concluída' && r[6] !== 'Cancelada');
    else if (filtros.status && filtros.status !== 'Todos' && filtros.status !== 'Todas') data = data.filter(r => r[6] === filtros.status);
    
    if (filtros.setor && filtros.setor !== 'Todos') data = data.filter(r => r[4] === filtros.setor);
    if (filtros.tecnico && filtros.tecnico !== 'Todos') data = data.filter(r => (filtros.tecnico === 'Nenhum') ? (!r[7]) : r[7] === filtros.tecnico);

    // --- BUSCA RÁPIDA (KEYWORD) ---
    if (filtros.keyword && filtros.keyword.trim() !== "") {
      const kw = filtros.keyword.toLowerCase().trim();
      data = data.filter(r => 
        String(r[0]).toLowerCase().includes(kw) || // ID
        String(r[3]).toLowerCase().includes(kw) || // Solicitante
        String(r[5]).toLowerCase().includes(kw) || // Descrição
        String(r[8]).toLowerCase().includes(kw)    // Observações
      );
    }

    const totalItems = data.length;
    
    // --- ORDENAÇÃO: 1º Prioridade Alta, 2º ID (Mais Recente) ---
    data.sort((a, b) => {
      const prioA = String(a[9]).trim() === 'Alta' ? 1 : 0;
      const prioB = String(b[9]).trim() === 'Alta' ? 1 : 0;
      
      if (prioA !== prioB) {
        return prioB - prioA; // O.S. com 'Alta' (1) vão para o topo
      }
      // Se tiverem a mesma prioridade, ordena pela O.S. mais recente (ID decrescente)
      return (Number(b[0]) || 0) - (Number(a[0]) || 0);
    });

    const paginatedData = data.slice((page - 1) * pageSize, ((page - 1) * pageSize) + pageSize);

    return { items: paginatedData.map(r => {
        r[1] = parseDateSafe(r[1]) ? Utilities.formatDate(parseDateSafe(r[1]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : r[1];
        r[10] = parseDateSafe(r[10]) ? Utilities.formatDate(parseDateSafe(r[10]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : r[10];
        return r;
    }), totalItems: totalItems };
  } catch (e) { return { items: [], totalItems: 0 }; }
}

function filtrarOS(filtros) {
  return getRelatorioOS(filtros);
}

function getRelatorioOS(filtros) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_OS);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    data = data.filter(r => r[0] && r[0] !== ''); 
    
    if (filtros.status === 'Ativas') {
        data = data.filter(r => String(r[6]).trim() !== 'Concluída' && String(r[6]).trim() !== 'Cancelada');
    } else if (filtros.status && filtros.status !== 'Todos' && filtros.status !== 'Todas') {
        data = data.filter(r => String(r[6]).trim() === filtros.status);
    }
    
    if (filtros.setor && filtros.setor !== 'Todos') data = data.filter(r => String(r[4]).trim() === filtros.setor);
    if (filtros.tecnico && filtros.tecnico !== 'Todos') {
      data = data.filter(r => (filtros.tecnico === 'Nenhum') ? (!r[7] || String(r[7]).trim() === '') : String(r[7]).trim() === filtros.tecnico);
    }

    if (filtros.dataInicio) {
      const d1Parts = filtros.dataInicio.split('-'); 
      const d1 = new Date(d1Parts[0], d1Parts[1] - 1, d1Parts[2], 0, 0, 0);
      data = data.filter(r => parseDateSafe(r[1]) && parseDateSafe(r[1]).getTime() >= d1.getTime());
    }
    if (filtros.dataFim) {
      const d2Parts = filtros.dataFim.split('-');
      const d2 = new Date(d2Parts[0], d2Parts[1] - 1, d2Parts[2], 23, 59, 59, 999);
      data = data.filter(r => parseDateSafe(r[1]) && parseDateSafe(r[1]).getTime() <= d2.getTime());
    }

    data.sort((a, b) => (Number(b[0]) || 0) - (Number(a[0]) || 0)); 
    
    return data.map(r => {
        r[1] = parseDateSafe(r[1]) ? Utilities.formatDate(parseDateSafe(r[1]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : r[1];
        r[10] = parseDateSafe(r[10]) ? Utilities.formatDate(parseDateSafe(r[10]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : r[10];
        return r;
    });
  } catch (e) { return []; }
}

function filtrarRotinas(filtros) {
  return getRelatorioDiario(filtros);
}

function getOverdueOS() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_OS);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
    const now = new Date();
    const sevenDaysAgo = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));
    
    return data
      .filter(r => {
        const status = String(r[6]).trim();
        const dataAbertura = parseDateSafe(r[1]);
        return (status === 'Aberta' || status === 'Pendente' || status === 'Em Andamento') && 
               dataAbertura && dataAbertura < sevenDaysAgo;
      })
      .map(r => r[0]); // Retorna apenas os IDs
  } catch (e) { return []; }
}

function getDashboardStatistics(e, d1_str, d2_str) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const osSheet = ss.getSheetByName(SHEET_OS);
    if(!osSheet || osSheet.getLastRow() <= 1) return { total:0, abertas:0, andamento:0, concluidas:0, pendente:0, avgResolutionTime: 'N/A', chartData:[['Status', 'Quantidade', { role: 'style' }],['Aberta', 0, 'color: #38bdf8'],['Andamento', 0, 'color: #f59e0b'],['Pendente', 0, 'color: #ef4444'],['Concluída', 0, 'color: #2cb67d']], sectorChartData:[['Setor', 'Quantidade']], techPerformanceData: [['Técnico', 'O.S. Atendidas', 'Registos no Diário']], yesterdayTechPerformanceChartData: [['Técnico', 'Concluídas']] };

    let filterStartDate = null;
    if (d1_str) { const p = d1_str.split('-'); filterStartDate = new Date(p[0], p[1]-1, p[2], 0,0,0); }
    let filterEndDate = null;
    if (d2_str) { const p = d2_str.split('-'); filterEndDate = new Date(p[0], p[1]-1, p[2], 23,59,59); }

    // FETCH ATÉ A COLUNA O (ÍNDICE 14) PARA GARANTIR A LEITURA DA DATA DE MOVIMENTAÇÃO
    const osData = osSheet.getRange(2, 1, osSheet.getLastRow() - 1, 15).getValues(); 
    let total=0, abertas=0, andamento=0, concluidas=0, pendente=0;
    let totalResolutionTime = 0, resolvedCount = 0;
    let techStats = {}, sectorStats = {}; 

    osData.forEach(r => {
        const creationDate = parseDateSafe(r[1]);
        if (creationDate) {
          if((!filterStartDate || creationDate >= filterStartDate) && (!filterEndDate || creationDate <= filterEndDate)) {
              total++;
              const status = String(r[6]).trim();
              const techName = String(r[7] || '').trim();
              const setorName = String(r[4] || 'Não Informado').trim(); 

              if(status==='Aberta') abertas++;
              else if(status==='Em Andamento') andamento++;
              else if(status==='Concluída') concluidas++;
              else if(status==='Pendente') pendente++;

              if (status !== 'Cancelada' && status !== '') sectorStats[setorName] = (sectorStats[setorName] || 0) + 1;

              const endDate = parseDateSafe(r[10]);
              if (status === 'Concluída' && endDate) {
                  const resolutionTime = endDate.getTime() - creationDate.getTime();
                  if (resolutionTime >= 0) { totalResolutionTime += resolutionTime; resolvedCount++; }
              }
              if (techName && techName !== '-' && techName !== '') {
                  if (!techStats[techName]) techStats[techName] = { os: 0, diario: 0 };
                  techStats[techName].os++;
              }
          }
        }
    });

    // Calcular Tempo Médio de Resolução
    let avgResText = "N/A";
    if (resolvedCount > 0) {
      const avgMs = totalResolutionTime / resolvedCount;
      const days = Math.floor(avgMs / (24 * 60 * 60 * 1000));
      const hours = Math.floor((avgMs % (24 * 60 * 60 * 1000)) / (60 * 60 * 1000));
      avgResText = days > 0 ? `${days}d ${hours}h` : `${hours}h`;
    }
    
    // --- LER DIÁRIO TÉCNICO (Para o gráfico de performance) ---
    try {
        const diarioSheet = ss.getSheetByName(SHEET_DIARIO);
        if (diarioSheet && diarioSheet.getLastRow() > 1) {
            const diarioData = diarioSheet.getRange(2, 1, diarioSheet.getLastRow() - 1, 2).getValues();
            diarioData.forEach(r => {
                const date = parseDateSafe(r[0]);
                if (date && (!filterStartDate || date >= filterStartDate) && (!filterEndDate || date <= filterEndDate)) {
                    const techName = String(r[1]).trim();
                    if (techName) {
                        if (!techStats[techName]) techStats[techName] = { os: 0, diario: 0 };
                        techStats[techName].diario++;
                    }
                }
            });
        }
    } catch(e) { console.error("Erro ao ler diário no dashboard: " + e); }

    const techPerformanceData = [['Técnico', 'O.S. Atendidas', 'Registos no Diário']];
    Object.keys(techStats).forEach(tech => {
        techPerformanceData.push([tech, techStats[tech].os || 0, techStats[tech].diario || 0]);
    });

    // =========================================================
    // LER ESTATÍSTICAS DE ONTEM (Baseado na Coluna G e Coluna O)
    // =========================================================
    let yesterdayChartData = [['Técnico', 'Concluídas']];
    try {
        let ontemStats = {};
        const tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';
        const hoje = new Date();
        const onte = new Date(hoje.getTime() - (24 * 60 * 60 * 1000));
        const onteString = Utilities.formatDate(onte, tz, "dd/MM/yyyy");

        osData.forEach(r => {
            const status = String(r[6] || '').trim(); // Coluna G (Índice 6)
            const techName = String(r[7] || '').trim(); // Coluna H
            const dataMovimentacao = r[14]; // Coluna O (Índice 14)
            
            let movString = "";
            if (dataMovimentacao instanceof Date) {
                movString = Utilities.formatDate(dataMovimentacao, tz, "dd/MM/yyyy");
            } else if (dataMovimentacao) {
                movString = String(dataMovimentacao).trim();
            }

            // Conta apenas as que estão Concluídas E onde a última movimentação foi ontem
            if ((status === 'Concluída' || status === 'Concluida') && movString === onteString && techName && techName !== '-' && techName !== 'Não Atribuído') {
                ontemStats[techName] = (ontemStats[techName] || 0) + 1;
            }
        });

        Object.keys(ontemStats).forEach(tech => {
            yesterdayChartData.push([tech, ontemStats[tech]]);
        });

        if (yesterdayChartData.length === 1) {
            yesterdayChartData.push(['Sem atividade', 0]);
        }
    } catch(e) { 
        console.error("Erro ao ler estatísticas de ontem: " + e); 
        yesterdayChartData.push(['Erro', 0]);
    }

    return { 
        total: total, abertas: abertas, andamento: andamento, concluidas: concluidas, pendente: pendente,
        avgResolutionTime: avgResText,
        chartData: [
            ['Status', 'Quantidade', { role: 'style' }],
            ['Aberta', abertas, 'color: #38bdf8'],      
            ['Andamento', andamento, 'color: #f59e0b'], 
            ['Pendente', pendente, 'color: #ef4444'],   
            ['Concluída', concluidas, 'color: #2cb67d'] 
        ],
        sectorChartData: [['Setor', 'Quantidade'], ...Object.keys(sectorStats).map(s => [s, sectorStats[s]])],
        techPerformanceData: techPerformanceData,
        yesterdayTechPerformanceChartData: yesterdayChartData
    };
  } catch(e) { 
    return { total:0, abertas:0, andamento:0, concluidas:0, pendente:0, avgResolutionTime: 'N/A', chartData:[['Status', 'Quantidade', { role: 'style' }],['Aberta', 0, 'color: #38bdf8'],['Andamento', 0, 'color: #f59e0b'],['Pendente', 0, 'color: #ef4444'],['Concluída', 0, 'color: #2cb67d']], sectorChartData:[['Setor', 'Quantidade']], techPerformanceData: [['Técnico', 'O.S. Atendidas', 'Registos no Diário']], yesterdayTechPerformanceChartData: [['Técnico', 'Concluídas']] }; 
  }
}

function _recordYesterdayStats(stats) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_DAILY_STATS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_DAILY_STATS);
      sheet.appendRow(['Data', 'Técnico', 'Concluídas', 'Pendentes', 'Em Andamento']);
      sheet.setFrozenRows(1);
    }

    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const yesterdayString = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "yyyy-MM-dd");

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const range = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for(let i = 0; i < range.length; i++){
          if(range[i][0] instanceof Date && Utilities.formatDate(range[i][0], Session.getScriptTimeZone(), "yyyy-MM-dd") === yesterdayString){
              return; // Já gravado
          }
      }
    }
    
    const dateToRecord = new Date(yesterdayString);
    Object.keys(stats).forEach(tech => {
       sheet.appendRow([dateToRecord, tech, stats[tech].concluidas || 0, stats[tech].pendentes || 0, stats[tech].emAndamento || 0]);
    });
  } catch (error) { console.error("Erro ao gravar estatísticas: " + error); }
}


// =================================================================
// DIÁRIO TÉCNICO E ROTINAS
// =================================================================
function salvarDiarioTecnico(dateValue, tecnicoNome, desc, files) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DIARIO);
    if (!sheet) return { status: 'falha', message: 'Planilha não encontrada.' };
    
    let url = '';
    if(files && files.length) {
      const folder = DriveApp.getFoldersByName(FOLDER_NAME).hasNext() ? DriveApp.getFoldersByName(FOLDER_NAME).next() : DriveApp.createFolder(FOLDER_NAME);
      url = files.map((f, i) => {
         const b64 = f.base64.split(',')[1] || f.base64;
         return folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', `DIARIO_${new Date().getTime()}_${i}_${f.name}`)).getUrl();
      }).join('\n');
    }
    
    const parts = dateValue.split('-');
    const recordDate = new Date(parts[0], parts[1]-1, parts[2]);
    sheet.appendRow([recordDate, tecnicoNome, desc, url]);
    logAction(tecnicoNome, 'Salvar Diário');
    return { status: 'sucesso', message: 'Registro salvo.' };
  } catch (e) { return { status: 'falha', message: e.message }; }
}

function getRelatorioDiario(filtros) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DIARIO);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    data = data.map((r, i) => [...r, i + 2]); // Adiciona o índice da linha (i+2) como 5ª coluna

    if (filtros.tecnico && filtros.tecnico !== 'Todos') data = data.filter(r => String(r[1]).trim() === filtros.tecnico);

    if (filtros.dataInicio) {
      const parts = filtros.dataInicio.split('-');
      const d1 = new Date(parts[0], parts[1] - 1, parts[2], 0, 0, 0);
      data = data.filter(r => parseDateSafe(r[0]) && parseDateSafe(r[0]).getTime() >= d1.getTime());
    }
    if (filtros.dataFim) {
      const parts = filtros.dataFim.split('-');
      const d2 = new Date(parts[0], parts[1] - 1, parts[2], 23, 59, 59, 999);
      data = data.filter(r => parseDateSafe(r[0]) && parseDateSafe(r[0]).getTime() <= d2.getTime());
    }

    data.sort((a, b) => (parseDateSafe(b[0])?.getTime() || 0) - (parseDateSafe(a[0])?.getTime() || 0));
    return data.map(r => {
        r[0] = parseDateSafe(r[0]) ? Utilities.formatDate(parseDateSafe(r[0]), Session.getScriptTimeZone(), "dd/MM/yyyy") : r[0];
        return r;
    });
  } catch (e) { return []; }
}

function getRotinaByRowIndex(rowIndex) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DIARIO);
    if (!sheet || sheet.getLastRow() < rowIndex) return null;
    return sheet.getRange(rowIndex, 1, 1, 4).getValues()[0];
  } catch(e) { return null; }
}

function atualizarRotina(email, rowIndex, desc) {
  if (verificarNivel(email).toLowerCase() !== 'admin') return { status: 'falha', message: 'Acesso negado.' };
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DIARIO);
    sheet.getRange(rowIndex, 3).setValue(desc);
    logAction(email, 'Atualizar Rotina', `Row ${rowIndex}`);
    return { status: 'sucesso' };
  } catch(e) { return { status: 'falha', message: e.message }; }
}

function excluirRotina(email, rowIndex) {
  if (verificarNivel(email).toLowerCase() !== 'admin') return { status: 'falha', message: 'Acesso negado.' };
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DIARIO);
    sheet.deleteRow(rowIndex);
    logAction(email, 'Excluir Rotina', `Row ${rowIndex}`);
    return { status: 'sucesso' };
  } catch(e) { return { status: 'falha', message: e.message }; }
}


// =================================================================
// MENSAGENS E CHAT GLOBAL
// =================================================================
function getConfiguracaoPopUp() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CONFIGURACOES);
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

function saveConfiguracaoPopUp(email, config) {
  if (verificarNivel(email).toLowerCase() !== 'admin') return { status: 'falha', message: 'Acesso negado.' };
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CONFIGURACOES);
    const data = sheet.getDataRange().getValues();
    let keysFound = { 'POPUP_HABILITADO': false, 'POPUP_MENSAGEM': false, 'POPUP_IMAGE_URL': false };
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'POPUP_HABILITADO') { sheet.getRange(i+1, 2).setValue(config.habilitado); keysFound['POPUP_HABILITADO'] = true; }
      if (data[i][0] === 'POPUP_MENSAGEM') { sheet.getRange(i+1, 2).setValue(config.mensagem); keysFound['POPUP_MENSAGEM'] = true; }
      if (data[i][0] === 'POPUP_IMAGE_URL') { sheet.getRange(i+1, 2).setValue(config.imageUrl); keysFound['POPUP_IMAGE_URL'] = true; }
    }
    
    if(!keysFound['POPUP_HABILITADO']) sheet.appendRow(['POPUP_HABILITADO', config.habilitado]);
    if(!keysFound['POPUP_MENSAGEM']) sheet.appendRow(['POPUP_MENSAGEM', config.mensagem]);
    if(!keysFound['POPUP_IMAGE_URL']) sheet.appendRow(['POPUP_IMAGE_URL', config.imageUrl]);
    
    logAction(email, 'Configuração Pop-Up Atualizada');
    return { status: 'sucesso' };
  } catch(e) { return { status: 'falha', message: e.message }; }
}

function sendMessageToUser(sender, recipient, content) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_MENSAGENS);
    if (!sheet) return { status: 'falha', message: 'Tabela de mensagens não encontrada.' };
    sheet.appendRow([new Date(), sender, content, recipient, '', 'NÃO']);
    return { status: 'sucesso', message: 'Mensagem enviada com sucesso!' };
  } catch (e) { return { status: 'falha', message: e.message }; }
}

function getMessagesForUser(email) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_MENSAGENS);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    let messages = data.filter(r => String(r[3]).toLowerCase() === email.toLowerCase());
    messages.sort((a,b) => (parseDateSafe(b[0])?.getTime() || 0) - (parseDateSafe(a[0])?.getTime() || 0));
    return messages.map((r, index) => {
      r[0] = parseDateSafe(r[0]) ? Utilities.formatDate(parseDateSafe(r[0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : r[0];
      return [r[0], r[1], r[2], index + 2]; // Passa a data formatada, Remetente, Msg e a Linha real (aproximada para UI)
    });
  } catch(e) { return []; }
}

function checkForUnreadMessages(email) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_MENSAGENS);
    const lastRow = sheet.getLastRow();
    if (!sheet || lastRow <= 1) return false;
    const startRow = Math.max(2, lastRow - 49);
    const numRows = lastRow - startRow + 1;
    const data = sheet.getRange(startRow, 1, numRows, 6).getValues();
    return data.some(r => String(r[3]).toLowerCase() === email.toLowerCase() && String(r[5]).toUpperCase() === 'NÃO');
  } catch (e) { return false; }
}

function markMessagesAsRead(email) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_MENSAGENS);
    const lastRow = sheet.getLastRow();
    if (!sheet || lastRow <= 1) return;
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for(let i=0; i<data.length; i++) {
      if(String(data[i][3]).toLowerCase() === email.toLowerCase() && String(data[i][5]).toUpperCase() === 'NÃO') {
        sheet.getRange(i+2, 6).setValue('SIM');
      }
    }
  } catch(e) {}
}

function getGlobalChatHistory(userEmail) {
  try {
    let sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CHAT);
    if (!sheet) {
      sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(SHEET_CHAT);
      sheet.appendRow(['Timestamp', 'Email', 'Nome', 'Mensagem', 'Contexto', 'Attachments', 'ReplyTo', 'ParentID']);
      sheet.setFrozenRows(1); return [];
    }
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; 

    const startRow = Math.max(2, lastRow - 99); 
    const numRows = lastRow - startRow + 1;
    const numCols = Math.max(8, sheet.getLastColumn());
    const data = sheet.getRange(startRow, 1, numRows, numCols).getValues();
    
    const messagesMap = new Map();
    const rootMessages = [];

    data.forEach((row, index) => {
      const messageId = startRow + index;
      let timeStr = parseDateSafe(row[0]) ? Utilities.formatDate(parseDateSafe(row[0]), Session.getScriptTimeZone(), "dd/MM HH:mm") : String(row[0]);
      let replyTo = null;
      if (row[6]) { try { replyTo = JSON.parse(row[6]); } catch (e) {} }

      const message = {
        id: messageId, timestamp: timeStr, fromEmail: row[1], fromName: row[2],
        message: row[3], context: row[4] || '', image: row[5] || '',
        replyTo: replyTo, parentID: row[7] ? Number(row[7]) : null,
        isMe: (String(row[1]).toLowerCase() === String(userEmail).toLowerCase()),
        replies: []
      };
      messagesMap.set(messageId, message);
    });

    messagesMap.forEach(message => {
      if (message.parentID && messagesMap.has(message.parentID)) {
        messagesMap.get(message.parentID).replies.push(message);
      } else {
        rootMessages.push(message);
      }
    });
    return rootMessages;
  } catch (e) { return []; }
}

function uploadChatFilesToDrive(files, prefix) {
  try {
    if (!files || !files.length) return '';
    const folderName = FOLDER_NAME + '_Chat';
    const folder = DriveApp.getFoldersByName(folderName).hasNext() ? DriveApp.getFoldersByName(folderName).next() : DriveApp.createFolder(folderName);
    const urls = [];
    files.forEach((f, idx) => {
      if (!f || !f.base64) return;
      const b64 = f.base64.split(',').length > 1 ? f.base64.split(',')[1] : f.base64.split(',')[0];
      const contentTypeMatch = String(f.base64).match(/^data:(.*?);base64,/);
      const contentType = contentTypeMatch ? contentTypeMatch[1] : 'application/octet-stream';
      const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), contentType, `${prefix || 'chat'}_${new Date().getTime()}_${idx}_${(f.name || 'file')}`));
      urls.push(file.getUrl());
    });
    return urls.join('\n');
  } catch (e) { return ''; }
}

function sendChatMessageAdvanced(email, context, message, files, replyTo) {
  if ((!message || message.trim() === "") && (!files || files.length === 0)) return { status: 'falha', message: 'Mensagem e anexos vazios.' };
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CHAT);
    const userData = getUserDataByEmail(email);
    const nome = userData ? userData.nome : email;
    let attachmentsUrls = files && files.length ? uploadChatFilesToDrive(files, `CHAT_${email.replace(/[^a-z0-9]/gi,'')}`) : '';
    
    sheet.appendRow([new Date(), email, nome, message || '', context || '', attachmentsUrls || '', replyTo ? JSON.stringify(replyTo) : '', replyTo ? replyTo.id : '']);
    return { status: 'sucesso' };
  } catch (e) { return { status: 'falha', message: e.message }; }
}
