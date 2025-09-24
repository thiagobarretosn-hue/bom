/**
 * @OnlyCurrentDoc
 * SCRIPT DE AUTOMAÇÃO DE RELATÓRIOS DINÂMICOS - VERSÃO COM PAINEL UNIFICADO
 *
 * REVISÃO 23.0 (REFINAMENTO FINAL DE UI/UX):
 * - MELHORIA DE LAYOUT: O "Painel de Seleção" e o "Painel de Pré-visualização" agora são
 * desenhados dinamicamente, um ao lado do outro, eliminando espaços vazios e criando um
 * "Painel Unificado" coeso e orgânico.
 * - HARMONIA VISUAL: Aplicado Row Banding em toda a área dos painéis para unificar
 * visualmente a interface e melhorar a legibilidade.
 * - CONFIABILIDADE: Otimizada a lógica de redesenho dos painéis para garantir que o fluxo
 * permaneça 100% automático, rápido e estável.
 * - CÓDIGO COMPLETO: Esta versão contém todas as funcionalidades revisadas e prontas para uso.
 */

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('🔧 Relatórios Dinâmicos')
      .addItem('⚙️ Painel de Controle', 'openConfigSidebar')
      .addSeparator()
      .addItem('📊 Processar Dados', 'runProcessingWithFeedback')
      .addItem('📄 Exportar Todos PDFs', 'exportAllPDFsWithFeedback')
      .addSeparator()
      .addItem('🧪 Diagnóstico', 'testSystem')
      .addItem('🔧 Recriar Config', 'forceCreateConfig')
      .addToUi();
    ensureConfigExists();
  } catch (error) {
    console.error('Erro na inicialização:', error);
  }
}

function openConfigSidebar() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ConfigSidebar.html')
                           .setTitle('Painel de Controle')
                           .setWidth(320);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error)    {
    console.error('Erro ao abrir a sidebar:', error);
    SpreadsheetApp.getUi().alert("Erro ao carregar o Painel de Controle. Verifique se o arquivo 'ConfigSidebar.html' existe.");
  }
}

// =================================================================
// FUNÇÕES UTILITÁRIAS E CACHE
// =================================================================

const CACHE_EXPIRATION = 600; // Cache expira em 10 minutos

function getUniqueColumnValues(sheet, columnIndex) {
    const cache = CacheService.getScriptCache();
    const cacheKey = `unique_${sheet.getSheetId()}_${columnIndex}`;
    const cached = cache.get(cacheKey);
    if (cached) {
        return JSON.parse(cached);
    }
    
    const values = sheet.getRange(2, columnIndex, sheet.getLastRow()).getValues().flat();
    const uniqueValues = [...new Set(values.filter(v => v))].sort((a,b) => String(a).localeCompare(String(b), undefined, {numeric: true}));
    
    cache.put(cacheKey, JSON.stringify(uniqueValues), CACHE_EXPIRATION);
    return uniqueValues;
}


function getColumnIndex(colConfig) {
  if (!colConfig || typeof colConfig !== 'string') return -1;
  const letter = colConfig.charAt(0).toUpperCase();
  return letter.charCodeAt(0) - 64;
}

function getColumnHeader(colConfig) {
  if (!colConfig || typeof colConfig !== 'string') return '';
  const parts = colConfig.split(' - ');
  return parts.length > 1 ? parts[1].trim() : '';
}

function formatVersion(inputValue) {
    if (inputValue === null || inputValue === undefined) return '';
    const str = String(inputValue).trim();
    if (str === '') return '';
    const num = parseInt(str, 10);
    return (isNaN(num) || num < 1) ? str : (num < 10 ? `0${num}` : `${num}`);
}

function extractFolderIdFromUrl(url) {
  if (typeof url !== 'string' || url.trim() === '') return '';
  const match = url.match(/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : url;
}

// =================================================================
// GERENCIAMENTO DA ABA 'CONFIG'
// =================================================================

function ensureConfigExists() {
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config')) {
    forceCreateConfig();
  }
}

function forceCreateConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Config';
  let configSheet = ss.getSheetByName(sheetName);
  if (configSheet) ss.deleteSheet(configSheet);
  configSheet = ss.insertSheet(sheetName, 0);
  configSheet.clear();

  const styles = {
    bgHeader: '#2c3e50', bgSection: '#3498db', bgInput: '#ecf0f1',
    fontLight: '#ffffff', fontDark: '#2c3e50', fontSubtle: '#7f8c8d',
    border: '#bdc3c7', fontFamily: 'Inter'
  };

  const configData = [
    ['CONFIGURAÇÃO', 'VALOR', 'DESCRIÇÃO'],
    ['', '', ''],
    [' 📊 AGRUPAMENTO', '', 'Defina as colunas para agrupar e os painéis serão atualizados automaticamente.'],
    ['Aba Origem', '', 'Aba que contém todos os dados brutos.'],
    ['Agrupar por Nível 1', '', 'Coluna principal de agrupamento (obrigatório).'],
    ['Agrupar por Nível 2', '', 'Subnível de agrupamento (opcional).'],
    ['Agrupar por Nível 3', '', 'Terceiro nível de agrupamento (opcional).'],
    ['', '', ''],
    [' 📋 DADOS BOMS', '', 'Mapeamento das colunas para a criação das listas de materiais (BOMs).'],
    ['Coluna 1', '', 'Coluna para o ID principal ou identificador do item.'],
    ['Coluna 2', 'I - DESC', 'Coluna para a descrição do item.'],
    ['Coluna 3', 'L - UPC', 'Coluna para o código de barras ou UPC.'],
    ['Coluna 4', 'K - UOM', 'Coluna para a unidade de medida (Unit of Measure).'],
    ['Coluna 5', 'N - PROJECT', 'Coluna com a quantidade a ser somada.'],
    ['', '', ''],
    [' 🏷️ CABEÇALHO', '', 'Informações que aparecerão no cabeçalho de cada relatório gerado.'],
    ['Project', 'HG1 BE', 'Nome principal do projeto.'],
    ['BOM', 'RISERS JS', 'Nome da lista de materiais (Bill of Materials).'],
    ['KOJO Prefixo', 'HG1.PLB.RGH.JS.BE.UNT.RISER', 'Prefixo fixo para o código KOJO.'],
    ['Engenheiro', 'WANDERSON', 'Nome do engenheiro responsável.'],
    ['Versão', '', 'Versão do relatório (ex: 01, 02, 03...).'],
    ['', '', ''],
    [' 💾 SALVAMENTO', '', 'Configurações para exportação e salvamento dos arquivos.'],
    ['Pasta Drive ID', '', 'Cole o LINK COMPLETO da pasta ou apenas o ID.'],
    ['Pasta Nome', '', 'Nome da pasta a ser criada caso o link/ID não seja fornecido.'],
    ['PDF Prefixo', 'HG1.PLB.RGH.JS.BE.UNT.RISER', 'Prefixo para os nomes dos arquivos PDF exportados.']
  ];
  
  configSheet.getRange(1, 1, configData.length, 3).setValues(configData);
  configSheet.getRange("A1:C" + configData.length).setFontFamily(styles.fontFamily).setVerticalAlignment('middle').setFontColor(styles.fontDark);
  configSheet.setColumnWidth(1, 220).setColumnWidth(2, 300).setColumnWidth(3, 350);
  configSheet.getRange('A1:C1').merge().setValue('⚙️ PAINEL DE CONFIGURAÇÃO | RELATÓRIOS DINÂMICOS').setBackground(styles.bgHeader).setFontColor(styles.fontLight).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  configSheet.setRowHeight(1, 40);
  const sections = { 3: { endRow: 7 }, 9: { endRow: 14 }, 16: { endRow: 21 }, 23: { endRow: 26 } };
  for (const startRow in sections) {
    const s = sections[startRow], start = parseInt(startRow);
    configSheet.getRange(start, 1, 1, 3).merge().setBackground(styles.bgSection).setFontColor(styles.fontLight).setFontSize(11).setFontWeight('bold').setHorizontalAlignment('left');
    configSheet.setRowHeight(start, 30);
    for (let r = start + 1; r <= s.endRow; r++) {
      configSheet.getRange(r, 2).setBackground(styles.bgInput).setHorizontalAlignment('center');
      configSheet.getRange(r, 1).setFontWeight('500');
      configSheet.getRange(r, 3).setFontStyle('italic').setFontColor(styles.fontSubtle);
      configSheet.setRowHeight(r, 28);
    }
    configSheet.getRange(start, 1, s.endRow - start + 1, 3).setBorder(true, true, true, true, null, null, styles.border, SpreadsheetApp.BorderStyle.SOLID);
  }
  configSheet.getRange(21, 2).setNumberFormat('@STRING@');
  configSheet.setFrozenRows(1);
  configSheet.getRange('E1').setValue('PAINEL UNIFICADO DE AGRUPAMENTO E PRÉ-VISUALIZAÇÃO').setFontWeight('bold').setFontFamily(styles.fontFamily);
  configSheet.getRange('E2').setValue('O painel é atualizado automaticamente. Edite o Sufixo KOJO na área de pré-visualização.').setFontStyle('italic').setFontFamily(styles.fontFamily);
  SpreadsheetApp.flush();
  Utilities.sleep(100);
  updateConfigDropdownsAuto();
}

function getConfigValue(key) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config').getRange("A1:B" + SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config').getLastRow()).getValues().find(row => row[0].toString().trim() === key)?.[1] || null;
}

function updateConfigDropdownsAuto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) return;

  const sourceSheets = ss.getSheets().map(s => s.getName()).filter(n => n !== 'Config');
  if (sourceSheets.length > 0) {
    configSheet.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(sourceSheets, true).setAllowInvalid(false).build());
  }

  const sourceSheet = ss.getSheetByName(getConfigValue('Aba Origem'));
  if (!sourceSheet || sourceSheet.getLastColumn() === 0) return;
  
  const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const columnOptions = headers.map((h, i) => `${String.fromCharCode(65 + i)} - ${h || `Coluna ${String.fromCharCode(65 + i)}`}`);
  const columnRule = SpreadsheetApp.newDataValidation().requireValueInList(columnOptions, true).setAllowInvalid(false).build();
  [5, 6, 7, 10, 11, 12, 13, 14].forEach(row => configSheet.getRange(row, 2).setDataValidation(columnRule));
}

function onEdit(e) {
    if (!e || !e.source) return;
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    if (sheet.getName() !== 'Config') return;
    const col = range.getColumn();
    const row = range.getRow();

    // Se a edição for na configuração de agrupamento (coluna B, linhas 4-7)
    if (col === 2 && row >= 4 && row <= 7) { 
        SpreadsheetApp.getActiveSpreadsheet().toast('Detectada alteração, atualizando...', 'Aguarde', 3);
        Utilities.sleep(200);
        updateConfigDropdownsAuto();
        const nextCol = updateGroupingPanel();
        updatePreviewPanel(nextCol);
    }
    
    // Se a edição for um checkbox no Painel de Seleção
    if (col >= 6 && (col - 6) % 3 === 0 && row >= 5) {
        const nextCol = updateGroupingPanel(); // Redesenha para manter a consistência
        updatePreviewPanel(nextCol);
    }
}


function updateGroupingPanel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  ss.toast('Atualizando Painel de Agrupamento...', 'Aguarde');
  
  const selections = getSelectionsFromPanel(configSheet);
  configSheet.getRange('E4:Z1000').clear();

  const sourceSheet = ss.getSheetByName(getConfigValue('Aba Origem'));
  if (!sourceSheet) return 5; // Retorna a coluna inicial se não houver aba

  const groupConfigs = [getConfigValue('Agrupar por Nível 1'), getConfigValue('Agrupar por Nível 2'), getConfigValue('Agrupar por Nível 3')];
  let currentColumn = 5;

  for (let i = 0; i < groupConfigs.length; i++) {
    const config = groupConfigs[i];
    if (config && config.trim() !== '') {
      const colIndex = getColumnIndex(config);
      if (colIndex === -1) continue;
      
      const colHeader = getColumnHeader(config);
      const levelId = `NÍVEL ${i+1}`;
      
      configSheet.getRange(4, currentColumn, 1, 2).merge().setValue(`${levelId}: ${colHeader.toUpperCase()}`).setBackground('#3498db').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
      
      const uniqueValues = getUniqueColumnValues(sourceSheet, colIndex);

      if (uniqueValues.length > 0) {
        configSheet.getRange(5, currentColumn, uniqueValues.length, 1).setValues(uniqueValues.map(v => [v]));
        const savedSelections = selections[levelId] || new Set();
        const checkboxes = uniqueValues.map(v => [savedSelections.has(v)]);
        configSheet.getRange(5, currentColumn + 1, uniqueValues.length, 1).insertCheckboxes().setValues(checkboxes);
      }
      configSheet.setColumnWidth(currentColumn, 150).setColumnWidth(currentColumn + 1, 50);
      currentColumn += 2; // Move para a próxima coluna disponível
    }
  }
  return currentColumn; // Retorna onde o próximo painel deve começar
}

function getSelectionsFromPanel(sheet) {
    const selections = {};
    for (let i = 0; i < 3; i++) {
        const startCol = 5 + (i * 2); // Ajustado para a nova largura
        const header = sheet.getRange(4, startCol).getValue();
        if (header) {
            const levelId = header.split(':')[0];
            const lastDataRow = sheet.getRange(sheet.getMaxRows(), startCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
            if (lastDataRow >= 5) {
                const numItems = lastDataRow - 4;
                const values = sheet.getRange(5, startCol, numItems, 1).getValues().flat();
                const checkboxes = sheet.getRange(5, startCol + 1, numItems, 1).getValues().flat();
                selections[levelId] = new Set(values.filter((_, index) => checkboxes[index] === true));
            }
        }
    }
    return selections;
}

// =================================================================
// LÓGICA DE PRÉ-VISUALIZAÇÃO E PROCESSAMENTO
// =================================================================

function updatePreviewPanel(previewStartCol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  ss.toast('Atualizando pré-visualização...', 'Aguarde');

  const combinations = getSelectedCombinationsFromPanel();
  
  const existingSuffixes = new Map();
  const lastPreviewRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  if (lastPreviewRow >= 5) {
      configSheet.getRange(5, previewStartCol, lastPreviewRow - 4, 3).getValues().forEach(row => {
          if (row[0]) existingSuffixes.set(row[0], row[2]);
      });
  }

  if (combinations.length === 0) {
    ss.toast('Nenhuma combinação selecionada.', 'Aviso', 5);
    return;
  }
  
  configSheet.getRange(4, previewStartCol, 1, 3).setValues([['COMBINAÇÃO GERADA', 'CRIAR?', 'SUFIXO KOJO FINAL']]).setBackground('#2c3e50').setFontColor('#ffffff').setFontWeight('bold');

  const tableData = combinations.map(combo => [combo, true, existingSuffixes.get(combo) || combo]);
  
  configSheet.getRange(5, previewStartCol, tableData.length, 3).setValues(tableData);
  configSheet.getRange(5, previewStartCol + 1, tableData.length, 1).insertCheckboxes();
  configSheet.getRange(5, previewStartCol + 2, tableData.length, 1).setBackground('#ecf0f1');
  
  configSheet.setColumnWidth(previewStartCol, 250).setColumnWidth(previewStartCol + 1, 80).setColumnWidth(previewStartCol + 2, 250);

  // Aplica o Row Banding para unificar visualmente
  const lastPanelCol = configSheet.getRange(4, 4).getNextDataCell(SpreadsheetApp.Direction.RIGHT).getColumn();
  const bandingRange = configSheet.getRange(4, 5, configSheet.getLastRow(), lastPanelCol - 4);
  bandingRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  ss.toast(`Pré-visualização atualizada!`, 'Sucesso', 5);
}

function getSelectedCombinationsFromPanel() {
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    let selectedGroups = [];
    
    for (let i = 0; i < 3; i++) {
        const startCol = 5 + (i * 2);
        const header = configSheet.getRange(4, startCol).getValue();
        if (header) {
            const lastDataRow = configSheet.getRange(configSheet.getMaxRows(), startCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
            if(lastDataRow >= 5) {
                const numItems = lastDataRow - 4;
                const values = configSheet.getRange(5, startCol, numItems, 1).getValues().flat();
                const checkboxes = configSheet.getRange(5, startCol + 1, numItems, 1).getValues().flat();
                const selected = values.filter((_, index) => checkboxes[index] === true);
                if (selected.length > 0) selectedGroups.push(selected);
            }
        }
    }

    if (selectedGroups.length === 0) return [];
    if (selectedGroups.length === 1) return selectedGroups[0].map(String);
    return selectedGroups.reduce((a, b) => a.flatMap(x => b.map(y => `${x}.${y}`)));
}


function runProcessingWithFeedback() {
    SpreadsheetApp.getActiveSpreadsheet().toast('Iniciando processamento...', 'Aguarde', -1);
    const result = runProcessing();
    if (result.success) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Processamento concluído! ${result.created} relatórios foram gerados.`, 'Sucesso!', 5);
    } else {
        SpreadsheetApp.getUi().alert('Falha no Processamento', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

function runProcessing() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const previewStartCol = 5 + (getConfigValue('Agrupar por Nível 3') ? 6 : getConfigValue('Agrupar por Nível 2') ? 4 : getConfigValue('Agrupar por Nível 1') ? 2 : 0);
  
  const lastPreviewRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  
  if (lastPreviewRow < 5) {
    return { success: false, message: 'Nenhuma combinação encontrada na pré-visualização.' };
  }
  
  const previewData = configSheet.getRange(5, previewStartCol, lastPreviewRow - 4, 3).getValues();
  const combinationsToProcess = previewData.filter(row => row[1] === true).map(row => ({ combination: row[0], kojoSuffix: row[2] }));

  if (combinationsToProcess.length === 0) {
    return { success: false, message: 'Nenhum relatório selecionado para criação.' };
  }
  
  const sourceSheet = ss.getSheetByName(getConfigValue('Aba Origem'));
  if (!sourceSheet) {
      return { success: false, message: 'Aba de origem não encontrada.' };
  }

  const groupConfigs = [getConfigValue('Agrupar por Nível 1'), getConfigValue('Agrupar por Nível 2'), getConfigValue('Agrupar por Nível 3')].filter(Boolean);
  const groupIndices = groupConfigs.map(getColumnIndex);
  const bomCols = { c1: getColumnIndex(getConfigValue('Coluna 1')), c2: getColumnIndex(getConfigValue('Coluna 2')), c3: getColumnIndex(getConfigValue('Coluna 3')), c4: getColumnIndex(getConfigValue('Coluna 4')), c5: getColumnIndex(getConfigValue('Coluna 5')) };

  const dataMap = new Map();
  combinationsToProcess.forEach(item => dataMap.set(item.combination, []));

  const allData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();

  for (const row of allData) {
      const rowCombination = groupIndices.map(index => row[index - 1]).join('.');
      if (dataMap.has(rowCombination)) {
          dataMap.get(rowCombination).push([ row[bomCols.c1 - 1], row[bomCols.c2 - 1], row[bomCols.c3 - 1], row[bomCols.c4 - 1], parseFloat(row[bomCols.c5 - 1]) || 0 ]);
      }
  }
  
  let createdCount = 0;
  for (const item of combinationsToProcess) {
      const { combination, kojoSuffix } = item;
      const rawData = dataMap.get(combination);
      if (!rawData || rawData.length === 0) continue;
      
      const processedData = groupAndSumData(rawData);
      const sanitizedName = combination.toString().replace(/[\/\\\?\*\[\]:]/g, '_').substring(0, 100);
      let targetSheet = ss.getSheetByName(sanitizedName) || ss.insertSheet(sanitizedName);
      targetSheet.clear();

      createAndFormatReport(targetSheet, combination, kojoSuffix, processedData);
      createdCount++;
  }
  
  return { success: true, created: createdCount };
}

function groupAndSumData(data) {
  const grouped = {};
  data.forEach(row => {
    const key = `${row[0]}|${row[1]}|${row[2]}|${row[3]}`;
    if (grouped[key]) {
      grouped[key][4] += row[4];
    } else {
      grouped[key] = [...row];
    }
  });
  return Object.values(grouped).sort((a, b) => String(a[0]).localeCompare(String(b[0]), undefined, { numeric: true }));
}

// =================================================================
// FORMATAÇÃO E EXPORTAÇÃO
// =================================================================

function createAndFormatReport(sheet, combination, kojoSuffix, data) {
    const config = { project: getConfigValue('Project'), bom: getConfigValue('BOM'), kojoPrefix: getConfigValue('KOJO Prefixo'), engineer: getConfigValue('Engenheiro'), version: formatVersion(getConfigValue('Versão')) };
    const headers = { h1: getColumnHeader(getConfigValue('Coluna 1')), h2: getColumnHeader(getConfigValue('Coluna 2')), h3: getColumnHeader(getConfigValue('Coluna 3')), h4: getColumnHeader(getConfigValue('Coluna 4')), h5: 'QTY' };
    const headerValues = createHeaderData(config, kojoSuffix);
    sheet.getRange(1, 1, headerValues.length, 2).setValues(headerValues);
    formatHeader(sheet, headerValues.length);
    const dataStartRow = headerValues.length + 2;
    const finalData = [[headers.h1, headers.h2, headers.h3, headers.h4, headers.h5]].concat(data);
    sheet.getRange(dataStartRow, 1, finalData.length, 5).setValues(finalData);
    formatDataTable(sheet, dataStartRow, finalData.length);
    setupColumnWidths(sheet);
    protectHeader(sheet, dataStartRow - 1);
}

function createHeaderData(config, kojoSuffix) {
  const lastUpdate = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'MM/dd/yyyy');
  const bomKojoComplete = `${config.kojoPrefix}.${kojoSuffix}`;
  return [
    ['PROJECT:', config.project], ['BOM:', config.bom], ['BOM KOJO:', bomKojoComplete],
    ['ENG.:', config.engineer], ['VERSION:', config.version], ['LAST UPDATE:', lastUpdate]
  ];
}

function formatHeader(sheet, headerLength) {
  sheet.getRange(1, 2, headerLength, 1).setNumberFormat('@STRING@');
  for (let r = 1; r <= headerLength; r++) sheet.getRange(r, 2, 1, 4).merge();
  sheet.getRange(1, 1, headerLength, 5).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
}

function formatDataTable(sheet, startRow, numRows) {
  if (numRows <= 1) return;
  sheet.getRange(startRow, 1, numRows, 5).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  sheet.getRange(startRow, 1, 1, 5).setFontWeight('bold');
}

function setupColumnWidths(sheet) {
  sheet.setColumnWidth(1, 105).setColumnWidth(2, 570).setColumnWidth(3, 105).setColumnWidth(4, 105).setColumnWidth(5, 105);
}

function protectHeader(sheet, headerEndRow) {
  try {
    const protection = sheet.getRange(1, 1, headerEndRow, 5).protect();
    protection.setDescription('Cabeçalho protegido').removeEditors(protection.getEditors());
  } catch (e) { console.warn('Erro ao proteger cabeçalho:', e.message); }
}

function exportAllPDFsWithFeedback() {
    SpreadsheetApp.getActiveSpreadsheet().toast('Iniciando exportação de todos os PDFs...', 'Aguarde', -1);
    const result = exportAllPDFs();
     if (result.success) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Exportação concluída! ${result.exported} PDFs salvos na pasta "${result.folder}".`, 'Sucesso!', 5);
    } else {
        SpreadsheetApp.getUi().alert('Falha na Exportação', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

function exportAllPDFs() {
    return exportSelectedPDFs(getReportSheetNames());
}

function exportSelectedPDFs(selectedSheets) {
  if (!selectedSheets || selectedSheets.length === 0) {
    return { success: false, message: 'Nenhuma aba foi selecionada.' };
  }
  try {
    const folder = getOrCreateOutputFolder();
    if (!folder) return { success: false, message: 'Pasta de destino não configurada ou inválida.' };
    const prefix = getConfigValue('PDF Prefixo') || 'Relatorio';
    selectedSheets.forEach(sheetName => {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (sheet) exportSheetToPdf(sheet, `${prefix}_${sheetName}`, folder);
    });
    return { success: true, exported: selectedSheets.length, folder: folder.getName() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getOrCreateOutputFolder() {
  const folderInput = getConfigValue('Pasta Drive ID');
  const folderName = getConfigValue('Pasta Nome');
  try {
    if (folderInput) return DriveApp.getFolderById(extractFolderIdFromUrl(folderInput));
    if (folderName) {
      const folders = DriveApp.getFoldersByName(folderName);
      return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    }
    const defaultName = `${SpreadsheetApp.getActiveSpreadsheet().getName()} - PDFs`;
    const folders = DriveApp.getFoldersByName(defaultName);
    return folders.hasNext() ? folders.next() : DriveApp.createFolder(defaultName);
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Não foi possível acessar a pasta de destino: ${e.message}`);
    return null;
  }
}

function exportSheetToPdf(sheet, pdfName, folder) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?gid=${sheet.getSheetId()}&format=pdf&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;
  const params = { headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }, muteHttpExceptions: true };
  for (let i = 0, delay = 1000; i < 5; i++) {
    const response = UrlFetchApp.fetch(url, params);
    if (response.getResponseCode() == 200) {
      folder.createFile(response.getBlob().setName(`${pdfName}.pdf`));
      return;
    }
    if (response.getResponseCode() == 429) {
      Utilities.sleep(delay + Math.floor(Math.random() * 1000));
      delay *= 2;
    } else {
      throw new Error(`Falha na exportação com código ${response.getResponseCode()}`);
    }
  }
  throw new Error(`Não foi possível exportar a aba "${sheet.getName()}" devido a limites do servidor.`);
}

function getReportSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName()).filter(name => name !== 'Config');
}

function testSystem() {
    const combinations = getSelectedCombinationsFromPanel();
    const message = `Diagnóstico do Sistema:\n- Modo de Agrupamento: Interativo (3 Níveis)\n- Combinações Selecionadas no Painel: ${combinations.length}\n- Aba de Origem: ${getConfigValue('Aba Origem') || "Não configurada"}\n\nO sistema está operando com a lógica de agrupamento mais recente e otimizações de cache.`;
    SpreadsheetApp.getUi().alert('Diagnóstico', message, SpreadsheetApp.getUi().ButtonSet.OK);
    return { "Modo de Agrupamento": "Interativo (3 Níveis)", "Combinações Selecionadas": combinations.length, "Aba de Origem": getConfigValue('Aba Origem') || "Não configurada" };
}

