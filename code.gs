/**
 * @OnlyCurrentDoc
 * SCRIPT DE AUTOMAÇÃO DE RELATÓRIOS DINÂMICOS - VERSÃO COM PAINEL UNIFICADO
 *
 * REVISÃO 31.0 (CACHE INTELIGENTE):
 * - LIMPEZA AUTOMÁTICA DE CACHE: O script agora detecta edições na 'Aba Origem' e limpa
 * automaticamente o cache de dados. Isso garante que os painéis e relatórios
 * sempre utilizarão as informações mais recentes, evitando inconsistências.
 * - OTIMIZAÇÃO DO GATILHO onEdit: A função foi reestruturada para lidar com edições tanto
 * na aba de 'Config' (para atualizar a UI) quanto na 'Aba Origem' (para limpar o cache).
 */

// =================================================================
// CONFIGURAÇÕES GLOBAIS E CONSTANTES
// =================================================================
const CONSTANTS = {
  SHEETS: {
    CONFIG: 'Config',
  },
  CONFIG_KEYS: {
    SOURCE_SHEET: 'Aba Origem',
    GROUP_L1: 'Agrupar por Nível 1',
    GROUP_L2: 'Agrupar por Nível 2',
    GROUP_L3: 'Agrupar por Nível 3',
    COL_1: 'Coluna 1',
    COL_2: 'Coluna 2',
    COL_3: 'Coluna 3',
    COL_4: 'Coluna 4',
    COL_5: 'Coluna 5',
    PROJECT: 'Project',
    BOM: 'BOM',
    KOJO_PREFIX: 'KOJO Prefixo',
    ENGINEER: 'Engenheiro',
    VERSION: 'Versão',
    DRIVE_FOLDER_ID: 'Pasta Drive ID',
    DRIVE_FOLDER_NAME: 'Pasta Nome',
    PDF_PREFIX: 'PDF Prefixo',
  },
  UI: {
    MENU_NAME: '🔧 Relatórios Dinâmicos',
    SIDEBAR_TITLE: 'Painel de Controle',
    SIDEBAR_FILE: 'ConfigSidebar.html',
  },
  COLORS: {
    HEADER_BG: '#2c3e50',
    SECTION_BG: '#3498db',
    INPUT_BG: '#ecf0f1',
    FONT_LIGHT: '#ffffff',
    FONT_DARK: '#2c3e50',
    FONT_SUBTLE: '#7f8c8d',
    BORDER: '#bdc3c7',
    FONT_FAMILY: 'Inter',
    PANEL_EMPTY_BG: '#95a5a6',
    PANEL_ERROR_BG: '#c0392b'
  },
  CACHE_EXPIRATION_SECONDS: 180, // Cache expira em 3 minutos
  COMBINATION_DELIMITER: '|||', // Delimitador para evitar conflito com dados que contenham "."
};


function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu(CONSTANTS.UI.MENU_NAME)
      .addItem('⚙️ Painel de Controle', 'openConfigSidebar')
      .addSeparator()
      .addItem('📊 Processar Dados', 'runProcessingWithFeedback')
      .addItem('📄 Exportar Todos PDFs', 'exportAllPDFsWithFeedback')
      .addSeparator()
      .addItem('🗑️ Limpar Relatórios Antigos', 'clearOldReports')
      .addSeparator()
      .addItem('🔄 Forçar Atualização (Limpar Cache)', 'forceRefreshPanels')
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
    const html = HtmlService.createHtmlOutputFromFile(CONSTANTS.UI.SIDEBAR_FILE)
                           .setTitle(CONSTANTS.UI.SIDEBAR_TITLE)
                           .setWidth(320);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    console.error('Erro ao abrir a sidebar:', error);
    SpreadsheetApp.getUi().alert(`Erro ao carregar o Painel de Controle. Verifique se o arquivo '${CONSTANTS.UI.SIDEBAR_FILE}' existe.`);
  }
}

// =================================================================
// FUNÇÕES UTILITÁRIAS, CACHE E LEITURA DE CONFIG
// =================================================================

function getAllConfigValues() {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'all_config_values';
    const cached = cache.get(cacheKey);
    if (cached) {
        return JSON.parse(cached);
    }

    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONSTANTS.SHEETS.CONFIG);
    if (!configSheet) return {};
    
    const values = configSheet.getRange("A1:B" + configSheet.getLastRow()).getValues();
    const config = {};
    values.forEach(row => {
        const key = row[0].toString().trim();
        if (key) {
            config[key] = row[1];
        }
    });
    
    cache.put(cacheKey, JSON.stringify(config), CONSTANTS.CACHE_EXPIRATION_SECONDS / 2);
    return config;
}

function getUniqueColumnValues(sheet, columnIndex) {
    const cache = CacheService.getScriptCache();
    const cacheKey = `unique_${sheet.getSheetId()}_${columnIndex}`;
    const cached = cache.get(cacheKey);
    if (cached) {
        return JSON.parse(cached);
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const values = sheet.getRange(2, columnIndex, lastRow - 1).getValues().flat();
    const uniqueValues = [...new Set(values.filter(v => v))].sort((a,b) => String(a).localeCompare(String(b), undefined, {numeric: true}));
    
    cache.put(cacheKey, JSON.stringify(uniqueValues), CONSTANTS.CACHE_EXPIRATION_SECONDS);
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

function extractFolderIdFromurl(url) {
  if (typeof url !== 'string' || url.trim() === '') return '';
  const match = url.match(/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : url;
}

// =================================================================
// GERENCIAMENTO DA ABA 'CONFIG'
// =================================================================

function ensureConfigExists() {
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONSTANTS.SHEETS.CONFIG)) {
    forceCreateConfig();
  }
}

function forceCreateConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = CONSTANTS.SHEETS.CONFIG;
  let configSheet = ss.getSheetByName(sheetName);
  if (configSheet) ss.deleteSheet(configSheet);
  configSheet = ss.insertSheet(sheetName, 0);
  configSheet.clear();

  const styles = CONSTANTS.COLORS;

  const configData = [
    ['CONFIGURAÇÃO', 'VALOR', 'DESCRIÇÃO'],
    ['', '', ''],
    [' 📊 AGRUPAMENTO', '', 'Defina as colunas para agrupar e os painéis serão atualizados automaticamente.'],
    [CONSTANTS.CONFIG_KEYS.SOURCE_SHEET, '', 'Aba que contém todos os dados brutos.'],
    [CONSTANTS.CONFIG_KEYS.GROUP_L1, '', 'Coluna principal de agrupamento (obrigatório).'],
    [CONSTANTS.CONFIG_KEYS.GROUP_L2, '', 'Subnível de agrupamento (opcional).'],
    [CONSTANTS.CONFIG_KEYS.GROUP_L3, '', 'Terceiro nível de agrupamento (opcional).'],
    ['', '', ''],
    [' 📋 DADOS BOMS', '', 'Mapeamento das colunas para a criação das listas de materiais (BOMs).'],
    [CONSTANTS.CONFIG_KEYS.COL_1, '', 'Coluna para o ID principal ou identificador do item.'],
    [CONSTANTS.CONFIG_KEYS.COL_2, 'I - DESC', 'Coluna para a descrição do item.'],
    [CONSTANTS.CONFIG_KEYS.COL_3, 'L - UPC', 'Coluna para o código de barras ou UPC.'],
    [CONSTANTS.CONFIG_KEYS.COL_4, 'K - UOM', 'Coluna para a unidade de medida (Unit of Measure).'],
    [CONSTANTS.CONFIG_KEYS.COL_5, 'N - PROJECT', 'Coluna com a quantidade a ser somada.'],
    ['', '', ''],
    [' 🏷️ CABEÇALHO', '', 'Informações que aparecerão no cabeçalho de cada relatório gerado.'],
    [CONSTANTS.CONFIG_KEYS.PROJECT, 'HG1 BE', 'Nome principal do projeto.'],
    [CONSTANTS.CONFIG_KEYS.BOM, 'RISERS JS', 'Nome da lista de materiais (Bill of Materials).'],
    [CONSTANTS.CONFIG_KEYS.KOJO_PREFIX, 'HG1.PLB.RGH.JS.BE.UNT.RISER', 'Prefixo fixo para o código KOJO.'],
    [CONSTANTS.CONFIG_KEYS.ENGINEER, 'WANDERSON', 'Nome do engenheiro responsável.'],
    [CONSTANTS.CONFIG_KEYS.VERSION, '', 'Versão do relatório (ex: 01, 02, 03...).'],
    ['', '', ''],
    [' 💾 SALVAMENTO', '', 'Configurações para exportação e salvamento dos arquivos.'],
    [CONSTANTS.CONFIG_KEYS.DRIVE_FOLDER_ID, '', 'Cole o LINK COMPLETO da pasta ou apenas o ID.'],
    [CONSTANTS.CONFIG_KEYS.DRIVE_FOLDER_NAME, '', 'Nome da pasta a ser criada caso o link/ID não seja fornecido.'],
    [CONSTANTS.CONFIG_KEYS.PDF_PREFIX, 'HG1.PLB.RGH.JS.BE.UNT.RISER', 'Prefixo para os nomes dos arquivos PDF exportados.']
  ];
  
  configSheet.getRange(1, 1, configData.length, 3).setValues(configData);
  configSheet.getRange("A1:C" + configData.length).setFontFamily(styles.FONT_FAMILY).setVerticalAlignment('middle').setFontColor(styles.FONT_DARK);
  configSheet.setColumnWidth(1, 220).setColumnWidth(2, 300).setColumnWidth(3, 350);
  configSheet.getRange('A1:C1').merge().setValue('⚙️ PAINEL DE CONFIGURAÇÃO | RELATÓRIOS DINÂMICOS').setBackground(styles.HEADER_BG).setFontColor(styles.FONT_LIGHT).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  
  const sections = { 3: { endRow: 7 }, 9: { endRow: 14 }, 16: { endRow: 21 }, 23: { endRow: 26 } };
  for (const startRow in sections) {
    const s = sections[startRow], start = parseInt(startRow);
    configSheet.getRange(start, 1, 1, 3).merge().setBackground(styles.SECTION_BG).setFontColor(styles.FONT_LIGHT).setFontSize(11).setFontWeight('bold').setHorizontalAlignment('left');
    configSheet.setRowHeight(start, 30);
    for (let r = start + 1; r <= s.endRow; r++) {
      configSheet.getRange(r, 2).setBackground(styles.INPUT_BG).setHorizontalAlignment('center');
      configSheet.getRange(r, 1).setFontWeight('500');
      configSheet.getRange(r, 3).setFontStyle('italic').setFontColor(styles.FONT_SUBTLE);
      configSheet.setRowHeight(r, 28);
    }
    configSheet.getRange(start, 1, s.endRow - start + 1, 3).setBorder(true, true, true, true, null, null, styles.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }
  configSheet.getRange(21, 2).setNumberFormat('@STRING@');
  configSheet.setFrozenRows(1);
  
  configSheet.getRange('E1:M1').merge().setValue('PAINEL UNIFICADO DE AGRUPAMENTO E PRÉ-VISUALIZAÇÃO')
      .setBackground(styles.HEADER_BG).setFontColor(styles.FONT_LIGHT).setFontSize(14)
      .setFontWeight('bold').setHorizontalAlignment('center').setFontFamily(styles.FONT_FAMILY);
  configSheet.setRowHeight(1, 40);

  SpreadsheetApp.flush();
  Utilities.sleep(100);
  updateConfigDropdownsAuto();
}

function updateConfigDropdownsAuto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONSTANTS.SHEETS.CONFIG);
  if (!configSheet) return;
  const config = getAllConfigValues();

  const sourceSheets = ss.getSheets().map(s => s.getName()).filter(n => n !== CONSTANTS.SHEETS.CONFIG);
  if (sourceSheets.length > 0) {
    configSheet.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(sourceSheets, true).setAllowInvalid(false).build());
  }

  const sourceSheet = ss.getSheetByName(config[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET]);
  if (!sourceSheet || sourceSheet.getLastColumn() === 0) return;
  
  const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const columnOptions = headers.map((h, i) => `${String.fromCharCode(65 + i)} - ${h || `Coluna ${String.fromCharCode(65 + i)}`}`);
  const columnRule = SpreadsheetApp.newDataValidation().requireValueInList(columnOptions, true).setAllowInvalid(false).build();
  [5, 6, 7, 10, 11, 12, 13, 14].forEach(row => configSheet.getRange(row, 2).setDataValidation(columnRule));
}

/**
 * ATUALIZADO: Lógica expandida para limpar o cache automaticamente quando a aba de origem é editada.
 */
function onEdit(e) {
    if (!e || !e.source) return;
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    const range = e.range;

    // Pega as configurações para saber qual é a aba de origem.
    // A função `getAllConfigValues` usa cache, então esta chamada é rápida.
    const config = getAllConfigValues();
    const sourceSheetName = config[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET];

    // Cenário 1: A aba de origem de dados foi editada.
    // Limpa o cache para garantir que a próxima operação use dados atualizados.
    if (sheetName === sourceSheetName) {
        clearSourceDataCache(config, sheet);
        SpreadsheetApp.getActiveSpreadsheet().toast('Fonte de dados modificada. Cache atualizado.', 'Informação', 3);
        return; // Sai após limpar o cache.
    }

    // Cenário 2: A aba 'Config' foi editada.
    // Lógica existente para atualizar os painéis de IU.
    if (sheetName === CONSTANTS.SHEETS.CONFIG) {
        CacheService.getScriptCache().remove('all_config_values');
        const col = range.getColumn();
        const row = range.getRow();

        // Se a configuração de agrupamento ou aba de origem for alterada
        if (col === 2 && row >= 4 && row <= 7) {
            Utilities.sleep(200);
            const updatedConfig = getAllConfigValues(); // Re-busca as configs que acabaram de ser alteradas.
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const sourceSheet = ss.getSheetByName(updatedConfig[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET]);

            switch (row) {
                case 4: // Aba Origem mudou: redesenha tudo
                    SpreadsheetApp.getActiveSpreadsheet().toast('Aba de origem alterada, recriando painel...', 'Aguarde', 3);
                    updateConfigDropdownsAuto();
                    updateGroupingPanel();
                    break;
                case 5: // Nível 1 mudou: redesenha apenas o painel 1
                    SpreadsheetApp.getActiveSpreadsheet().toast('Atualizando Nível 1...', 'Aguarde', 2);
                    updateSingleLevelPanel(1, updatedConfig, sourceSheet);
                    break;
                case 6: // Nível 2 mudou: redesenha apenas o painel 2
                    SpreadsheetApp.getActiveSpreadsheet().toast('Atualizando Nível 2...', 'Aguarde', 2);
                    updateSingleLevelPanel(2, updatedConfig, sourceSheet);
                    break;
                case 7: // Nível 3 mudou: redesenha apenas o painel 3
                    SpreadsheetApp.getActiveSpreadsheet().toast('Atualizando Nível 3...', 'Aguarde', 2);
                    updateSingleLevelPanel(3, updatedConfig, sourceSheet);
                    break;
            }
            updatePreviewPanel(11); // A pré-visualização sempre atualiza
        }

        // Se uma checkbox no painel for clicada, atualiza apenas a pré-visualização
        if ((col === 6 || col === 8 || col === 10) && row >= 4 && range.getWidth() === 1 && range.getDataValidation() && range.getDataValidation().getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
            SpreadsheetApp.getActiveSpreadsheet().toast('Atualizando combinações...', 'Aguarde', 2);
            updatePreviewPanel(11);
        }
    }
}

/**
 * Limpa o cache relacionado aos dados da aba de origem.
 * Chamado automaticamente quando a aba de origem é editada.
 * @param {Object} config O objeto de configuração atual.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet A instância da aba de origem.
 */
function clearSourceDataCache(config, sourceSheet) {
    const cache = CacheService.getScriptCache();
    cache.remove('all_config_values');

    if (sourceSheet) {
        const sourceSheetId = sourceSheet.getSheetId();
        const groupConfigs = [
            config[CONSTANTS.CONFIG_KEYS.GROUP_L1],
            config[CONSTANTS.CONFIG_KEYS.GROUP_L2],
            config[CONSTANTS.CONFIG_KEYS.GROUP_L3]
        ].filter(Boolean);

        const cacheKeysToRemove = groupConfigs.map(conf => {
            const colIndex = getColumnIndex(conf);
            if (colIndex !== -1) {
                return `unique_${sourceSheetId}_${colIndex}`;
            }
            return null;
        }).filter(Boolean);

        if (cacheKeysToRemove.length > 0) {
            cache.removeAll(cacheKeysToRemove);
        }
    }
}


// Função updateSingleLevelPanel - MELHORADA
function updateSingleLevelPanel(level, config, sourceSheet) {
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONSTANTS.SHEETS.CONFIG);
    const startCol = 5 + (level - 1) * 2;
    const levelId = `NÍVEL ${level}`;
    const groupConfigKey = `Agrupar por Nível ${level}`;
    const groupConfig = config[groupConfigKey];

    configSheet.getRange(3, startCol, configSheet.getMaxRows() - 2, 2).clear();
    configSheet.setColumnWidth(startCol, 150).setColumnWidth(startCol + 1, 50);
    const headerRange = configSheet.getRange(3, startCol, 1, 2).merge();

    if (sourceSheet && groupConfig && groupConfig.trim() !== '') {
        const colIndex = getColumnIndex(groupConfig);
        if (colIndex === -1) {
            headerRange.setValue(`${levelId}: ERRO`).setBackground(CONSTANTS.COLORS.PANEL_ERROR_BG).setFontColor(CONSTANTS.COLORS.FONT_LIGHT).setFontWeight('bold').setHorizontalAlignment('center');
            configSheet.getRange(4, startCol).setValue('Config. inválida.').setFontColor(CONSTANTS.COLORS.FONT_SUBTLE).setFontStyle('italic');
        } else {
            const colHeader = getColumnHeader(groupConfig);
            headerRange.setValue(`${levelId}: ${colHeader.toUpperCase()}`).setBackground(CONSTANTS.COLORS.SECTION_BG).setFontColor(CONSTANTS.COLORS.FONT_LIGHT).setFontWeight('bold').setHorizontalAlignment('center');
            
            const uniqueValues = getUniqueColumnValues(sourceSheet, colIndex);
            if (uniqueValues.length > 0) {
                // MELHORIA: Mantém o tipo original dos valores
                const valuesFormatted = uniqueValues.map(v => [v === null || v === undefined || v === '' ? '' : v]);
                configSheet.getRange(4, startCol, valuesFormatted.length, 1).setValues(valuesFormatted);
                configSheet.getRange(4, startCol + 1, valuesFormatted.length, 1).insertCheckboxes();
            }
        }
    } else {
        headerRange.setValue(levelId).setBackground(CONSTANTS.COLORS.PANEL_EMPTY_BG).setFontColor(CONSTANTS.COLORS.FONT_LIGHT).setFontWeight('bold').setHorizontalAlignment('center');
        const message = sourceSheet ? '(Não configurado)' : '(Selecione Aba Origem)';
        configSheet.getRange(4, startCol).setValue(message).setFontColor(CONSTANTS.COLORS.FONT_SUBTLE).setFontStyle('italic');
    }

    const lastRowInPanel = configSheet.getRange(configSheet.getMaxRows(), startCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    const effectiveLastRow = Math.max(3, lastRowInPanel);
    configSheet.getRange(3, startCol, effectiveLastRow - 2, 2).setBorder(true, true, true, true, true, true, CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
}


function updateGroupingPanel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getAllConfigValues();
  const sourceSheet = ss.getSheetByName(config[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET]);
  
  updateSingleLevelPanel(1, config, sourceSheet);
  updateSingleLevelPanel(2, config, sourceSheet);
  updateSingleLevelPanel(3, config, sourceSheet);
}

// Função getPanelSelections - MELHORADA
function getPanelSelections(sheet) {
    const selections = {};
    for (let i = 0; i < 3; i++) {
        const startCol = 5 + (i * 2);
        const headerRange = sheet.getRange(3, startCol);
        if (headerRange.isBlank()) continue;

        const header = headerRange.getValue();
        const levelId = header.split(':')[0].trim();
        
        const lastDataRow = sheet.getRange(sheet.getMaxRows(), startCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
        if (lastDataRow >= 4) {
            const numItems = lastDataRow - 3;
            const values = sheet.getRange(4, startCol, numItems, 1).getValues().flat();
            const checkboxes = sheet.getRange(4, startCol + 1, numItems, 1).getValues().flat();
            
            // CORREÇÃO: Mantém valores originais (número/texto) sem forçar conversão
            const selectedValues = values
                .map((v, index) => ({value: v, checked: checkboxes[index]}))
                .filter(item => item.checked === true)
                .map(item => {
                    const val = item.value;
                    return val === null || val === undefined || val === '' ? '' : val;
                });
            
            selections[levelId] = new Set(selectedValues);
        } else {
            selections[levelId] = new Set();
        }
    }
    return selections;
}

// =================================================================
// LÓGICA DE PRÉ-VISUALIZAÇÃO, PROCESSAMENTO E LIMPEZA
// =================================================================

function updatePreviewPanel(previewStartCol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONSTANTS.SHEETS.CONFIG);

  const combinations = getSelectedCombinationsFromPanel();
  
  const existingSuffixes = new Map();
  const lastPreviewRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  if (lastPreviewRow >= 4) {
      configSheet.getRange(4, previewStartCol, lastPreviewRow - 3, 3).getValues().forEach(row => {
          if (row[0]) existingSuffixes.set(row[0], row[2]);
      });
  }
  configSheet.getRange(3, previewStartCol, configSheet.getMaxRows() - 2, 3).clear();

  configSheet.getRange(3, previewStartCol, 1, 3).setValues([['COMBINAÇÃO GERADA', 'CRIAR?', 'SUFIXO KOJO FINAL']]).setBackground(CONSTANTS.COLORS.HEADER_BG).setFontColor(CONSTANTS.COLORS.FONT_LIGHT).setFontWeight('bold');

  if (combinations.length > 0) {
    const tableData = combinations.map(combo => [combo, true, existingSuffixes.get(combo) || combo]);
    
    configSheet.getRange(4, previewStartCol, tableData.length, 3).setValues(tableData);
    configSheet.getRange(4, previewStartCol + 1, tableData.length, 1).insertCheckboxes();
    configSheet.getRange(4, previewStartCol + 2, tableData.length, 1).setBackground(CONSTANTS.COLORS.INPUT_BG);
  }
  
  configSheet.setColumnWidth(previewStartCol, 250).setColumnWidth(previewStartCol + 1, 80).setColumnWidth(previewStartCol + 2, 250);

  const lastCombinationRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const effectiveLastRow = Math.max(3, lastCombinationRow);
  configSheet.getRange(3, previewStartCol, effectiveLastRow - 2, 3).setBorder(true, true, true, true, true, true, CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  
  ss.toast(`Pré-visualização atualizada!`, 'Sucesso', 3);
}

// Função getSelectedCombinationsFromPanel - CORRIGIDA
function getSelectedCombinationsFromPanel() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(CONSTANTS.SHEETS.CONFIG);
    const config = getAllConfigValues();
    
    const sourceSheet = ss.getSheetByName(config[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET]);
    if (!sourceSheet) return [];

    const groupConfigs = [
        config[CONSTANTS.CONFIG_KEYS.GROUP_L1], 
        config[CONSTANTS.CONFIG_KEYS.GROUP_L2], 
        config[CONSTANTS.CONFIG_KEYS.GROUP_L3]
    ].filter(Boolean);
    
    if (groupConfigs.length === 0) return [];

    const groupIndices = groupConfigs.map(getColumnIndex);
    const panelSelections = getPanelSelections(configSheet);
    
    const activeSelectionLevels = Object.keys(panelSelections).filter(level => panelSelections[level].size > 0);
    if (activeSelectionLevels.length === 0) return [];

    const allData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
    const existingCombinations = new Set();

    allData.forEach(row => {
        const combinationParts = groupIndices.map(index => {
            const value = row[index - 1];
            // CORREÇÃO: Converte consistentemente, mantendo diferença entre número e texto
            return value === null || value === undefined || value === '' ? '' : String(value).trim();
        });
        
        if (combinationParts.every(part => part !== '')) {
            existingCombinations.add(combinationParts.join(CONSTANTS.COMBINATION_DELIMITER));
        }
    });

    const finalCombinations = [...existingCombinations].filter(combo => {
        const parts = combo.split(CONSTANTS.COMBINATION_DELIMITER);
        return parts.every((part, i) => {
            const levelId = `NÍVEL ${i + 1}`;
            if (!panelSelections[levelId]) return false;
            
            // CORREÇÃO: Compara considerando tipos mistos (número/texto)
            const normalizedPart = part.trim();
            return Array.from(panelSelections[levelId]).some(selected => 
                String(selected).trim() === normalizedPart
            );
        });
    });

    return finalCombinations.sort((a, b) => {
        // MELHORIA: Ordenação que lida com números e texto
        const aParts = a.split(CONSTANTS.COMBINATION_DELIMITER);
        const bParts = b.split(CONSTANTS.COMBINATION_DELIMITER);
        for (let i = 0; i < Math.min(aParts.length, bParts.length); i++) {
            const aNum = parseFloat(aParts[i]);
            const bNum = parseFloat(bParts[i]);
            if (!isNaN(aNum) && !isNaN(bNum)) {
                if (aNum !== bNum) return aNum - bNum;
            } else {
                const comparison = aParts[i].localeCompare(bParts[i], undefined, {numeric: true});
                if (comparison !== 0) return comparison;
            }
        }
        return aParts.length - bParts.length;
    });
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

// Função runProcessing - CORRIGIDA para lidar com tipos mistos
function runProcessing() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getAllConfigValues();
  const configSheet = ss.getSheetByName(CONSTANTS.SHEETS.CONFIG);
  
  const previewStartCol = 11;
  
  const lastPreviewRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  
  if (lastPreviewRow < 4) return { success: false, message: 'Nenhuma combinação encontrada na pré-visualização.' };
  
  const previewData = configSheet.getRange(4, previewStartCol, lastPreviewRow - 3, 3).getValues();
  const combinationsToProcess = previewData.filter(row => row[1] === true).map(row => ({ combination: row[0], kojoSuffix: row[2] }));

  if (combinationsToProcess.length === 0) return { success: false, message: 'Nenhum relatório selecionado para criação.' };
  
  const sourceSheet = ss.getSheetByName(config[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET]);
  if (!sourceSheet) return { success: false, message: 'Aba de origem não encontrada.' };

  const groupConfigs = [config[CONSTANTS.CONFIG_KEYS.GROUP_L1], config[CONSTANTS.CONFIG_KEYS.GROUP_L2], config[CONSTANTS.CONFIG_KEYS.GROUP_L3]].filter(Boolean);
  const groupIndices = groupConfigs.map(getColumnIndex);
  const bomCols = {
      c1: getColumnIndex(config[CONSTANTS.CONFIG_KEYS.COL_1]), c2: getColumnIndex(config[CONSTANTS.CONFIG_KEYS.COL_2]),
      c3: getColumnIndex(config[CONSTANTS.CONFIG_KEYS.COL_3]), c4: getColumnIndex(config[CONSTANTS.CONFIG_KEYS.COL_4]),
      c5: getColumnIndex(config[CONSTANTS.CONFIG_KEYS.COL_5])
  };

  const dataMap = new Map();
  combinationsToProcess.forEach(item => dataMap.set(item.combination, []));

  const allData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();

  for (const row of allData) {
      // CORREÇÃO: Normaliza valores mantendo consistência com a geração de combinações
      const rowCombination = groupIndices.map(index => {
          const value = row[index - 1];
          return value === null || value === undefined || value === '' ? '' : String(value).trim();
      }).join(CONSTANTS.COMBINATION_DELIMITER);
      
      if (dataMap.has(rowCombination)) {
          dataMap.get(rowCombination).push([  
              row[bomCols.c1 - 1], row[bomCols.c2 - 1], row[bomCols.c3 - 1], 
              row[bomCols.c4 - 1], parseFloat(row[bomCols.c5 - 1]) || 0 
          ]);
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
  // AJUSTE: Ordena os dados pela segunda coluna (DESC) em ordem alfabética.
  return Object.values(grouped).sort((a, b) => String(a[1]).localeCompare(String(b[1]), undefined, { numeric: false }));
}

function clearOldReports() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirmação', 'Tem a certeza de que deseja apagar TODAS as abas de relatório? Esta ação não pode ser desfeita.', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getAllConfigValues();
    const protectedSheets = [CONSTANTS.SHEETS.CONFIG, config[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET]].filter(Boolean);
    
    let deletedCount = 0;
    ss.getSheets().forEach(sheet => {
      const sheetName = sheet.getName();
      if (!protectedSheets.includes(sheetName)) {
        ss.deleteSheet(sheet);
        deletedCount++;
      }
    });
    
    if (deletedCount > 0) {
      ui.alert('Limpeza Concluída', `${deletedCount} abas de relatório foram removidas.`, ui.ButtonSet.OK);
    } else {
      ui.alert('Limpeza', 'Nenhuma aba de relatório para remover.', ui.ButtonSet.OK);
    }
  }
}

// =================================================================
// FORMATAÇÃO E EXPORTAÇÃO
// =================================================================

function createAndFormatReport(sheet, combination, kojoSuffix, data) {
    const config = getAllConfigValues();
    const reportConfig = { 
        project: config[CONSTANTS.CONFIG_KEYS.PROJECT], 
        bom: config[CONSTANTS.CONFIG_KEYS.BOM], 
        kojoPrefix: config[CONSTANTS.CONFIG_KEYS.KOJO_PREFIX], 
        engineer: config[CONSTANTS.CONFIG_KEYS.ENGINEER], 
        version: formatVersion(config[CONSTANTS.CONFIG_KEYS.VERSION]) 
    };
    const headers = { 
        h1: getColumnHeader(config[CONSTANTS.CONFIG_KEYS.COL_1]), 
        h2: getColumnHeader(config[CONSTANTS.CONFIG_KEYS.COL_2]), 
        h3: getColumnHeader(config[CONSTANTS.CONFIG_KEYS.COL_3]), 
        h4: getColumnHeader(config[CONSTANTS.CONFIG_KEYS.COL_4]), 
        h5: 'QTY' 
    };
    const headerValues = createHeaderData(reportConfig, kojoSuffix);
    sheet.getRange(1, 1, headerValues.length, 2).setValues(headerValues);
    formatHeader(sheet, headerValues.length);
    const dataStartRow = headerValues.length + 2;
    const finalData = [[headers.h1, headers.h2, headers.h3, headers.h4, headers.h5]].concat(data);
    sheet.getRange(dataStartRow, 1, finalData.length, 5).setValues(finalData);
    formatDataTable(sheet, dataStartRow, finalData.length);
    setupColumnWidths(sheet);
    protectHeader(sheet, dataStartRow - 1);
}

function createHeaderData(reportConfig, kojoSuffix) {
  const lastUpdate = Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'MM/dd/yyyy');
  const bomKojoComplete = `${reportConfig.kojoPrefix}.${kojoSuffix}`;
  return [
    ['PROJECT:', reportConfig.project], ['BOM:', reportConfig.bom], ['BOM KOJO:', bomKojoComplete],
    ['ENG.:', reportConfig.engineer], ['VERSION:', reportConfig.version], ['LAST UPDATE:', lastUpdate]
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
    const config = getAllConfigValues();
    const prefix = config[CONSTANTS.CONFIG_KEYS.PDF_PREFIX] || 'Relatorio';
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
  const config = getAllConfigValues();
  const folderInput = config[CONSTANTS.CONFIG_KEYS.DRIVE_FOLDER_ID];
  const folderName = config[CONSTANTS.CONFIG_KEYS.DRIVE_FOLDER_NAME];
  try {
    if (folderInput) return DriveApp.getFolderById(extractFolderIdFromurl(folderInput));
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
  const config = getAllConfigValues();
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName()).filter(name => name !== CONSTANTS.SHEETS.CONFIG && name !== config[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET]);
}

/**
 * NOVA FUNÇÃO: Limpa o cache manualmente e força a atualização dos painéis.
 */
function forceRefreshPanels() {
    SpreadsheetApp.getActiveSpreadsheet().toast('Limpando cache e atualizando painéis...', 'Aguarde', -1);
    
    // Remove o cache que armazena os valores de configuração.
    CacheService.getScriptCache().remove('all_config_values');
    
    // Para limpar os caches de valores únicos dinâmicos, precisamos obter o ID da planilha e os índices das colunas.
    const config = getAllConfigValues(); // Isso buscará uma nova cópia das configurações.
    const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET]);
    
    if (sourceSheet) {
        const sourceSheetId = sourceSheet.getSheetId();
        const groupConfigs = [
            config[CONSTANTS.CONFIG_KEYS.GROUP_L1],
            config[CONSTANTS.CONFIG_KEYS.GROUP_L2],
            config[CONSTANTS.CONFIG_KEYS.GROUP_L3]
        ].filter(Boolean);
        
        const cacheKeysToRemove = groupConfigs.map(conf => {
            const colIndex = getColumnIndex(conf);
            if (colIndex !== -1) {
                return `unique_${sourceSheetId}_${colIndex}`;
            }
            return null;
        }).filter(Boolean);
        
        if (cacheKeysToRemove.length > 0) {
            CacheService.getScriptCache().removeAll(cacheKeysToRemove);
        }
    }
    
    // Agora, força a atualização dos painéis com dados novos.
    updateGroupingPanel();
    updatePreviewPanel(11);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Painéis atualizados com os dados mais recentes!', 'Sucesso', 5);
}

function testSystem() {
    const config = getAllConfigValues();
    const combinations = getSelectedCombinationsFromPanel();
    const message = `Diagnóstico do Sistema:\n- Modo de Agrupamento: Interativo (3 Níveis)\n- Combinações Selecionadas no Painel: ${combinations.length}\n- Aba de Origem: ${config[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET] || "Não configurada"}\n\nO sistema está operando com a lógica de agrupamento mais recente e otimizações de cache.`;
    SpreadsheetApp.getUi().alert('Diagnóstico', message, SpreadsheetApp.getUi().ButtonSet.OK);
    return { "Modo de Agrupamento": "Interativo (3 Níveis)", "Combinações Selecionadas": combinations.length, "Aba de Origem": config[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET] || "Não configurada" };
}

