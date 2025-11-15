/**
 * @OnlyCurrentDoc
 * SCRIPT DE AUTOMAÇÃO DE RELATÓRIOS DINÂMICOS - VERSÃO COM PAINEL UNIFICADO
 *
 * REVISÃO 32.0 (CLASSIFICAÇÃO DINÂMICA):
 * - NOVA FUNCIONALIDADE: Adicionada uma nova secção 'OPÇÕES DE RELATÓRIO' na aba 'Config'.
 * - CONTROLO DE CLASSIFICAÇÃO: O utilizador pode agora escolher dinamicamente a coluna
 * e a ordem (Ascendente/Descendente) para a classificação dos dados nos relatórios gerados.
 * - INTEGRAÇÃO SEGURA: A lógica de classificação foi centralizada e atualizada para não
 * afetar outras funcionalidades do script.
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
    // NOVO: Chaves para a configuração de classificação
    SORT_BY: 'CLASSIFICAR POR',
    SORT_ORDER: 'ORDEM',
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


// ============= ADICIONAR AO MENU onOpen() =============
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu(CONSTANTS.UI.MENU_NAME)
      .addItem('⚙️ Painel de Controle', 'openConfigSidebar')
      .addSeparator()
      .addItem('📊 Processar Dados', 'runProcessingWithFeedback')
      .addItem('🔧 Fixadores → Fonte', 'abrirSeletorFixadores')  // NOVO
      .addItem('🔧 Fixadores → Relatórios', 'calcularFixadoresRelatoriosComFeedback')  // NOVO
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
    // Usa getDisplayValues() para garantir que a formatação (ex: "01") seja lida como texto.
    const values = configSheet.getRange("A1:B" + configSheet.getLastRow()).getDisplayValues();
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
    return (isNaN(num) || num < 1) ?
    str : (num < 10 ? `0${num}` : `${num}`);
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

// ALTERADO: Adicionada nova secção para classificação
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
    [' ⚙️ OPÇÕES DE RELATÓRIO', '', 'Defina opções adicionais para a geração dos relatórios.'], // NOVO
    [CONSTANTS.CONFIG_KEYS.SORT_BY, 'Coluna 2', 'Coluna a ser usada para classificar os itens no relatório.'], // NOVO
    [CONSTANTS.CONFIG_KEYS.SORT_ORDER, 'Ascendente (A-Z, 0-9)', 'A ordem de classificação (ascendente ou descendente).'], // NOVO
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
  
  // ALTERADO: Atualizados os números das linhas para incluir a nova secção
  const sections = { 3: { endRow: 7 }, 9: { endRow: 14 }, 16: { endRow: 21 }, 23: { endRow: 26 }, 28: { endRow: 31 } };

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
  configSheet.getRange(21, 2).setNumberFormat('@STRING@'); // Linha da Versão
  configSheet.setFrozenRows(1);
  
  configSheet.getRange('E1:M1').merge().setValue('PAINEL UNIFICADO DE AGRUPAMENTO E PRÉ-VISUALIZAÇÃO')
      .setBackground(styles.HEADER_BG).setFontColor(styles.FONT_LIGHT).setFontSize(14)
      .setFontWeight('bold').setHorizontalAlignment('center').setFontFamily(styles.FONT_FAMILY);
  configSheet.setRowHeight(1, 40);

  SpreadsheetApp.flush();
  Utilities.sleep(100);
  updateConfigDropdownsAuto();
}

// ALTERADO: Adicionados dropdowns para as novas opções de classificação
function updateConfigDropdownsAuto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONSTANTS.SHEETS.CONFIG);
  if (!configSheet) return;
  const config = getAllConfigValues();

  // Dropdown para Aba Origem
  const sourceSheets = ss.getSheets().map(s => s.getName()).filter(n => n !== CONSTANTS.SHEETS.CONFIG);
  if (sourceSheets.length > 0) {
    configSheet.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(sourceSheets, true).setAllowInvalid(false).build());
  }

  const sourceSheet = ss.getSheetByName(config[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET]);
  if (!sourceSheet || sourceSheet.getLastColumn() === 0) return;
  
  // Dropdowns para Colunas
  const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const columnOptions = headers.map((h, i) => `${String.fromCharCode(65 + i)} - ${h || `Coluna ${String.fromCharCode(65 + i)}`}`);
  const columnRule = SpreadsheetApp.newDataValidation().requireValueInList(columnOptions, true).setAllowInvalid(false).build();
  [5, 6, 7, 10, 11, 12, 13, 14].forEach(row => configSheet.getRange(row, 2).setDataValidation(columnRule));

  // NOVO: Dropdowns para as opções de classificação
  const bomColumnOptions = [
      CONSTANTS.CONFIG_KEYS.COL_1, 
      CONSTANTS.CONFIG_KEYS.COL_2, 
      CONSTANTS.CONFIG_KEYS.COL_3, 
      CONSTANTS.CONFIG_KEYS.COL_4,
      CONSTANTS.CONFIG_KEYS.COL_5
  ];
  const sortColumnRule = SpreadsheetApp.newDataValidation().requireValueInList(bomColumnOptions, true).setAllowInvalid(false).build();
  configSheet.getRange(24, 2).setDataValidation(sortColumnRule); // Célula B24 para 'CLASSIFICAR POR'

  const sortOrderOptions = ['Ascendente (A-Z, 0-9)', 'Descendente (Z-A, 9-0)'];
  const sortOrderRule = SpreadsheetApp.newDataValidation().requireValueInList(sortOrderOptions, true).setAllowInvalid(false).build();
  configSheet.getRange(25, 2).setDataValidation(sortOrderRule); // Célula B25 para 'ORDEM'
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
    const config = getAllConfigValues();
    const sourceSheetName = config[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET];
    
    // Cenário 1: A aba de origem de dados foi editada.
    if (sheetName === sourceSheetName) {
        clearSourceDataCache(config, sheet);
        
        return; // Sai após limpar o cache.
    }

    // Cenário 2: A aba 'Config' foi editada.
    if (sheetName === CONSTANTS.SHEETS.CONFIG) {
        CacheService.getScriptCache().remove('all_config_values');
        const col = range.getColumn();
        const row = range.getRow();

        // Formata a versão automaticamente ao ser editada (célula B21).
        if (col === 2 && row === 21) {
            const currentValue = range.getValue();
            const formattedValue = formatVersion(currentValue);
            if (String(currentValue) !== String(formattedValue)) {
                range.setValue(formattedValue);
            }
        }

        // Se a configuração de agrupamento ou aba de origem for alterada
        if (col === 2 && row >= 4 && row <= 7) {
            Utilities.sleep(200);
            const updatedConfig = getAllConfigValues();
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const sourceSheet = ss.getSheetByName(updatedConfig[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET]);

            switch (row) {
                case 4: // Aba Origem mudou: redesenha tudo
                    SpreadsheetApp.getActiveSpreadsheet().toast('Aba de origem alterada, recriando painel...', 'Aguarde', 3);
                    updateConfigDropdownsAuto();
                    updateGroupingPanel();
                    break;
                case 5: // Nível 1 mudou
                    SpreadsheetApp.getActiveSpreadsheet().toast('Atualizando Nível 1...', 'Aguarde', 2);
                    updateSingleLevelPanel(1, updatedConfig, sourceSheet);
                    break;
                case 6: // Nível 2 mudou
                    SpreadsheetApp.getActiveSpreadsheet().toast('Atualizando Nível 2...', 'Aguarde', 2);
                    updateSingleLevelPanel(2, updatedConfig, sourceSheet);
                    break;
                case 7: // Nível 3 mudou
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
  const internalDataCol = previewStartCol + 3; // Coluna N para dados internos

  const combinations = getSelectedCombinationsFromPanel();
  const existingSuffixes = new Map();
  const lastPreviewRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  
  if (lastPreviewRow >= 4) {
      configSheet.getRange(4, previewStartCol, lastPreviewRow - 3, 3).getValues().forEach(row => {
          if (row[0]) existingSuffixes.set(row[0], row[2]);
      });
  }
  configSheet.getRange(3, previewStartCol, configSheet.getMaxRows() - 2, 4).clear(); // Limpa 4 colunas

  configSheet.getRange(3, previewStartCol, 1, 3).setValues([['COMBINAÇÃO GERADA', 'CRIAR?', 'SUFIXO KOJO FINAL']]).setBackground(CONSTANTS.COLORS.HEADER_BG).setFontColor(CONSTANTS.COLORS.FONT_LIGHT).setFontWeight('bold');
  
  if (combinations.length > 0) {
    const tableData = combinations.map(combo => {
      const displayCombo = combo.replace(/\|\|\|/g, '.');
      const kojoSuffix = existingSuffixes.get(combo) || displayCombo;
      return [displayCombo, true, kojoSuffix, combo]; // Adiciona o combo original na 4ª coluna
    });
    configSheet.getRange(4, previewStartCol, tableData.length, 4).setValues(tableData); // Escreve 4 colunas
    configSheet.getRange(4, previewStartCol + 1, tableData.length, 1).insertCheckboxes();
    configSheet.getRange(4, previewStartCol + 2, tableData.length, 1).setBackground(CONSTANTS.COLORS.INPUT_BG);
  }
  
  configSheet.setColumnWidth(previewStartCol, 250).setColumnWidth(previewStartCol + 1, 80).setColumnWidth(previewStartCol + 2, 250);
  configSheet.hideColumns(internalDataCol); // Oculta a coluna com os dados internos

  const lastCombinationRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const effectiveLastRow = Math.max(3, lastCombinationRow);
  configSheet.getRange(3, previewStartCol, effectiveLastRow - 2, 3).setBorder(true, true, true, true, true, true, CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  
  ss.toast(`Pré-visualização atualizada!`, 'Sucesso', 3);
}

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
            
            const normalizedPart = part.trim();
            return Array.from(panelSelections[levelId]).some(selected => 
                String(selected).trim() === normalizedPart
            );
        });
    });

    return finalCombinations.sort((a, b) => {
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

// ALTERADO: Função principal de processamento agora lê as configurações de classificação
function runProcessing() {
  CacheService.getScriptCache().remove('all_config_values');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getAllConfigValues();
  const configSheet = ss.getSheetByName(CONSTANTS.SHEETS.CONFIG);
  
  const previewStartCol = 11;
  const lastPreviewRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  
  if (lastPreviewRow < 4) return { success: false, message: 'Nenhuma combinação encontrada na pré-visualização.' };
  
  const previewData = configSheet.getRange(4, previewStartCol, lastPreviewRow - 3, 4).getValues();
  const combinationsToProcess = previewData.filter(row => row[1] === true).map(row => ({ 
      displayCombination: String(row[0]),
      combination: String(row[3]),
      kojoSuffix: row[2] 
  }));
  
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

  // NOVO: Lê as configurações de classificação da aba 'Config'
  const sortColumnName = config[CONSTANTS.CONFIG_KEYS.SORT_BY];
  const sortOrder = config[CONSTANTS.CONFIG_KEYS.SORT_ORDER] === 'Descendente (Z-A, 9-0)' ? 'desc' : 'asc';

  const bomColumnNameToIndex = {
      [CONSTANTS.CONFIG_KEYS.COL_1]: 0, [CONSTANTS.CONFIG_KEYS.COL_2]: 1, [CONSTANTS.CONFIG_KEYS.COL_3]: 2,
      [CONSTANTS.CONFIG_KEYS.COL_4]: 3, [CONSTANTS.CONFIG_KEYS.COL_5]: 4,
  };
  const sortIndex = bomColumnNameToIndex[sortColumnName] ?? 1; // Se não encontrar, classifica pela Coluna 2 por padrão

  const sortConfig = { sortIndex, sortOrder };

  const dataMap = new Map();
  combinationsToProcess.forEach(item => dataMap.set(item.combination, []));

  const allData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
  for (const row of allData) {
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
      const { displayCombination, combination, kojoSuffix } = item;
      const rawData = dataMap.get(combination);
      if (!rawData || rawData.length === 0) continue;
      
      // ALTERADO: Passa a configuração de classificação para a função de processamento de dados
      const processedData = groupAndSumData(rawData, sortConfig);
      
      const sanitizedName = displayCombination.toString().replace(/[\/\\\?\*\[\]:]/g, '_').substring(0, 100);
      let targetSheet = ss.getSheetByName(sanitizedName) || ss.insertSheet(sanitizedName);
      targetSheet.clear();

      createAndFormatReport(targetSheet, combination, kojoSuffix, processedData);
      createdCount++;
  }
  
  return { success: true, created: createdCount };
}

// ALTERADO: Esta função agora aceita uma configuração de classificação e a aplica
function groupAndSumData(data, sortConfig) {
  const grouped = {};
  data.forEach(row => {
    const key = `${row[0]}|${row[1]}|${row[2]}|${row[3]}`; // Chave de agrupamento baseada nas 4 primeiras colunas
    if (grouped[key]) {
      grouped[key][4] += row[4]; // Soma a 5ª coluna (quantidade)
    } else {
      grouped[key] = [...row];
    }
  });

  const groupedData = Object.values(grouped);

  // NOVO: Lógica de classificação dinâmica
  const { sortIndex, sortOrder } = sortConfig;
  const direction = sortOrder === 'asc' ? 1 : -1;

  groupedData.sort((a, b) => {
    const valA = a[sortIndex];
    const valB = b[sortIndex];

    // Trata valores nulos ou vazios para que fiquem no final
    if (valA === null || valA === undefined || valA === '') return 1;
    if (valB === null || valB === undefined || valB === '') return -1;

    // Tenta comparar como números primeiro, especialmente útil para a coluna de quantidade (QTY)
    const aIsNum = !isNaN(parseFloat(valA)) && isFinite(valA);
    const bIsNum = !isNaN(parseFloat(valB)) && isFinite(valB);

    if (aIsNum && bIsNum) {
      return (parseFloat(valA) - parseFloat(valB)) * direction;
    }

    // Se não forem números, usa a comparação de texto inteligente (localeCompare)
    return String(valA).localeCompare(String(valB), undefined, { numeric: true }) * direction;
  });
  
  return groupedData;
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
// FORMATAÇÃO E EXPORTAÇÃO (Nenhuma alteração necessária aqui)
// =================================================================

function createAndFormatReport(sheet, combination, kojoSuffix, data) {
    const config = getAllConfigValues();
    const reportConfig = { 
        project: config[CONSTANTS.CONFIG_KEYS.PROJECT], 
        bom: config[CONSTANTS.CONFIG_KEYS.BOM], 
        kojoPrefix: config[CONSTANTS.CONFIG_KEYS.KOJO_PREFIX], 
        engineer: config[CONSTANTS.CONFIG_KEYS.ENGINEER], 
        version: config[CONSTANTS.CONFIG_KEYS.VERSION]
    };
    
    const headers = { 
        h1: getColumnHeader(config[CONSTANTS.CONFIG_KEYS.COL_1]), 
        h2: getColumnHeader(config[CONSTANTS.CONFIG_KEYS.COL_2]), 
        h3: getColumnHeader(config[CONSTANTS.CONFIG_KEYS.COL_3]), 
        h4: getColumnHeader(config[CONSTANTS.CONFIG_KEYS.COL_4]), 
        h5: 'QTY' 
    };

    const headerValues = createHeaderData(reportConfig, kojoSuffix);
    
    formatHeader(sheet, headerValues.length);
    sheet.getRange(1, 1, headerValues.length, 2).setValues(headerValues);

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

function forceRefreshPanels() {
    SpreadsheetApp.getActiveSpreadsheet().toast('Limpando cache e atualizando painéis...', 'Aguarde', -1);
    CacheService.getScriptCache().remove('all_config_values');
    
    const config = getAllConfigValues();
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
    
    updateGroupingPanel();
    updatePreviewPanel(11);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Painéis atualizados com os dados mais recentes!', 'Sucesso', 5);
}

// ============= CONFIGURAÇÃO DE FIXADORES ATUALIZADA =============
const FIXADOR_CONFIG = {
  RISER: {
    interval: 10,
    clamps: {
      '1/2': 'RISER CLAMP 1 IN. METAL',
      '3/4': 'RISER CLAMP 1 IN. METAL',
      '1': 'RISER CLAMP 1 IN. METAL',
      '1-1/4': 'RISER CLAMP 1-1/4 IN. METAL',
      '1-1/2': 'RISER CLAMP 1-1/2 IN. METAL',
      '2': 'RISER CLAMP 2 IN. METAL',
      '2-1/2': 'RISER CLAMP 2 IN. METAL',
      '3': 'RISER CLAMP 3 IN. METAL',
      '4': 'RISER CLAMP 4 IN. METAL',
      '6': 'RISER CLAMP 6 IN. METAL',
      '8': 'RISER CLAMP 8 IN. METAL',
      '10': 'RISER CLAMP 12 IN. METAL',
      '12': 'RISER CLAMP 12 IN. METAL'
    },
    materials: [
      { desc: 'NUT 3/8 IN. METAL', qtyFormula: 4 },
      { desc: 'FENDER WASHER 3/8 X 1-1/2', qtyFormula: 4 },
      { desc: 'ANCHOR DROP-IN 3/8 IN. X 3/4 IN. LONG HDI-P (W/ AUTO SET TOOL) [HILTI 409499]', qtyFormula: 2 },
      { desc: 'PLTD STEEL ALL THREAD ROD 3/8 IN. X 6 FT.', qtyFormula: 2 }
    ]
  },
  LOOP: {
    interval: 3,
    hangs: {
      '1/2': 'LOOP HANG 1/2 IN. HANGER',
      '3/4': 'LOOP HANG 3/4 IN. HANGER',
      '1': 'LOOP HANG 1 IN. METAL',
      '1-1/4': 'LOOP HANG 1-1/4 IN. METAL',
      '1-1/2': 'LOOP HANG 1-1/2 IN. HANGER',
      '2': 'LOOP HANG 2 IN. METAL',
      '2-1/2': 'LOOP HANG 2-1/2 IN. METAL',
      '3': 'LOOP HANG 3 IN. METAL',
      '4': 'LOOP HANG 4 IN. METAL',
      '6': 'LOOP HANG 6 IN. METAL',
      '8': 'LOOP HANG 6 IN. METAL',
      '10': 'LOOP HANG 12 IN. METAL',
      '12': 'LOOP HANG 12 IN. METAL'
    },
    materials: [
      { desc: 'NUT 3/8 IN. METAL', qtyFormula: 2 },
      { desc: 'FENDER WASHER 3/8 X 1-1/2', qtyFormula: 2 },
      { desc: 'ANCHOR DROP-IN 3/8 IN. X 3/4 IN. LONG HDI-P (W/ AUTO SET TOOL) [HILTI 409499]', qtyFormula: 1 },
      { desc: 'PLTD STEEL ALL THREAD ROD 3/8 IN. X 6 FT.', qtyFormula: 1 }
    ]
  }
};

// ============= FUNÇÕES AUXILIARES =============
function arredondarInteligente(value) {
  if (value <= 0) return 0;
  if (value < 1) return 1;
  const base = Math.floor(value);
  const decimal = value - base;
  return decimal <= 0.3 ? base : base + 1;
}

function extrairDiametro(desc) {
  const patterns = [
    /PIPE\s+(\d+\/\d+)\s+IN/i,
    /PIPE\s+(\d+)\s+IN/i,
    /PIPE\s+(\d+-\d+\/\d+)\s+IN/i,
    /(\d+\/\d+)\s*IN/i,
    /(\d+-\d+\/\d+)\s*IN/i,
    /(\d+)\s*IN/i
  ];
  
  for (const pattern of patterns) {
    const match = desc.match(pattern);
    if (match) return normalizarDiametro(match[1]);
  }
  return null;
}

function normalizarDiametro(diam) {
  const normalized = diam.replace(/\s+/g, '');
  const mapping = {
    '1/2': '1/2', '3/4': '3/4', '1': '1', '1-1/4': '1-1/4',
    '1-1/2': '1-1/2', '2': '2', '2-1/2': '2-1/2', '3': '3',
    '4': '4', '6': '6', '8': '8', '10': '10', '12': '12'
  };
  return mapping[normalized] || normalized;
}

function validarTipoFixacao(section) {
  const sectionUpper = String(section).toUpperCase();
  return sectionUpper.includes('RISER') || sectionUpper.includes('COLGANTE');
}

// ============= ABERTURA DO SELETOR =============
function abrirSeletorFixadores() {
  const html = HtmlService.createHtmlOutputFromFile('FixadorSelector')
    .setWidth(600)
    .setHeight(700)
    .setTitle('Seletor de Fixadores');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Seleção de Tubos');
}

// ============= LEITURA DE PIPES ELEGÍVEIS =============
function getPipesElegiveis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getAllConfigValues();
  const sourceSheet = ss.getSheetByName(config[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET]);
  
  if (!sourceSheet) return [];
  
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) return [];
  
  const data = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();
  const pipes = [];
  
  data.forEach((row, idx) => {
    const section = String(row[1] || '');
    const desc = String(row[8] || '');
    const qty = parseFloat(row[9]) || 0;
    const uom = String(row[10] || '');
    
    if (validarTipoFixacao(section) && desc.toUpperCase().includes('PIPE') && qty > 0) {
      const diameter = extrairDiametro(desc);
      const isRiser = section.toUpperCase().includes('RISER');
      const config = isRiser ? FIXADOR_CONFIG.RISER : FIXADOR_CONFIG.LOOP;
      const itemMap = isRiser ? config.clamps : config.hangs;
      
      if (diameter && itemMap[diameter]) {
        pipes.push({
          rowIndex: idx + 2,
          section: row[1],
          local: row[4],
          trade: row[5],
          phase: row[7],
          unitid : row[3],
          desc: desc,
          qty: qty,
          uom: uom,
          diameter: diameter,
          isRiser: isRiser,
          originalRow: [...row]
        });
      }
    }
  });
  
  return pipes;
}

// ============= PROCESSAMENTO DE FIXADORES SELECIONADOS =============
function processarFixadoresSelecionados(selectedPipes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getAllConfigValues();
  const sourceSheet = ss.getSheetByName(config[CONSTANTS.CONFIG_KEYS.SOURCE_SHEET]);
  
  if (!sourceSheet || !selectedPipes || selectedPipes.length === 0) {
    return { success: false, message: 'Dados inválidos' };
  }
  
  // Ordenar por rowIndex DESC para inserir de baixo para cima (evita deslocamento)
  selectedPipes.sort((a, b) => b.rowIndex - a.rowIndex);
  
  let totalAdded = 0;
  const maxCol = sourceSheet.getLastColumn();
  
  selectedPipes.forEach(pipe => {
    const fixadorConfig = pipe.isRiser ? FIXADOR_CONFIG.RISER : FIXADOR_CONFIG.LOOP;
    const itemMap = pipe.isRiser ? fixadorConfig.clamps : fixadorConfig.hangs;
    const fixadorItem = itemMap[pipe.diameter];
    
    if (!fixadorItem) return;
    
    const qtyFixadores = arredondarInteligente(pipe.qty / fixadorConfig.interval);
    const linhasParaInserir = [];
    
    // Linha do fixador
    const linhaFixador = new Array(maxCol).fill('');
    linhaFixador[0] = pipe.originalRow[0]; // PROJECT
    linhaFixador[1] = pipe.originalRow[1]; // SECTION
    linhaFixador[2] = pipe.originalRow[2]; // QTT
    linhaFixador[3] = pipe.originalRow[3]; // UNIT ID
    linhaFixador[4] = pipe.originalRow[4]; // LOCAL
    linhaFixador[5] = 'FIX'; // TRADE
    linhaFixador[6] = pipe.originalRow[6]; // (vazia normalmente)
    linhaFixador[7] = pipe.originalRow[7]; // PHASE
    linhaFixador[8] = fixadorItem; // DESC
    linhaFixador[9] = `=ROUND(J${pipe.rowIndex}/${fixadorConfig.interval})`; // FÓRMULA QTY
    
    linhasParaInserir.push(linhaFixador);
    
    // Linhas dos materiais
    fixadorConfig.materials.forEach(mat => {
      const linhaMat = new Array(maxCol).fill('');
      linhaMat[0] = pipe.originalRow[0];
      linhaMat[1] = pipe.originalRow[1];
      linhaMat[2] = pipe.originalRow[2];
      linhaMat[3] = pipe.originalRow[3];
      linhaMat[4] = pipe.originalRow[4];
      linhaMat[5] = 'FIX';
      linhaMat[6] = pipe.originalRow[6];
      linhaMat[7] = pipe.originalRow[7];
      linhaMat[8] = mat.desc;
      linhaMat[9] = `=J${pipe.rowIndex + linhasParaInserir.length}*${mat.qtyFormula}`; // FÓRMULA
      
      linhasParaInserir.push(linhaMat);
    });
    
    // Inserir abaixo do pipe
    const insertRow = pipe.rowIndex + 1;
    sourceSheet.insertRowsAfter(pipe.rowIndex, linhasParaInserir.length);
    
    // Copiar formatação
    const formatoOrigem = sourceSheet.getRange(pipe.rowIndex, 1, 1, maxCol);
    formatoOrigem.copyFormatToRange(sourceSheet, 1, maxCol, insertRow, insertRow + linhasParaInserir.length - 1);
    
    // Inserir dados
    const rangeDestino = sourceSheet.getRange(insertRow, 1, linhasParaInserir.length, maxCol);
    
    rangeDestino.setValues(linhasParaInserir);
    
    totalAdded += linhasParaInserir.length;
  });
  
  return { success: true, added: totalAdded };
}

// ============= VERSÃO PARA RELATÓRIOS =============
function calcularFixadoresRelatorios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheets = getReportSheetNames();
  
  if (reportSheets.length === 0) {
    return { success: false, message: 'Nenhum relatório encontrado' };
  }
  
  let totalProcessed = 0;
  
  reportSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    const result = processarFixadoresSheet(sheet);
    totalProcessed += result.added;
  });
  
  return { success: true, processed: totalProcessed, sheets: reportSheets.length };
}

function processarFixadoresSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 8) return { added: 0 }; // Pula se só tem cabeçalho
  
  const dataStartRow = 8; // Após cabeçalho fixo de 6 linhas + 1 vazia + linha de headers
  const data = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 5).getValues();
  const linhasParaInserir = [];
  
  data.forEach((row, idx) => {
    const desc = String(row[1] || '');
    const qty = parseFloat(row[4]) || 0;
    
    if (desc.toUpperCase().includes('PIPE') && qty > 0) {
      const diameter = extrairDiametro(desc);
      if (!diameter) return;
      
      // Assume LOOP por padrão em relatórios (ajustar se necessário)
      const config = FIXADOR_CONFIG.LOOP;
      const fixadorItem = config.hangs[diameter];
      
      if (!fixadorItem) return;
      
      const actualRow = dataStartRow + idx;
      const qtyFixadores = arredondarInteligente(qty / config.interval);
      
      // Fixador
      linhasParaInserir.push([
        row[0], // ID
        fixadorItem,
        row[2], // UPC
        row[3], // UOM
        `=ARREDONDAR.PARA.CIMA(E${actualRow}/${config.interval};0)`
      ]);
      
      // Materiais
      config.materials.forEach(mat => {
        linhasParaInserir.push([
          '',
          mat.desc,
          '',
          'EA',
          `=E${dataStartRow + linhasParaInserir.length - 1}*${mat.qtyFormula}`
        ]);
      });
    }
  });
  
  if (linhasParaInserir.length > 0) {
    const insertRow = lastRow + 1;
    sheet.getRange(insertRow, 1, linhasParaInserir.length, 5).setValues(linhasParaInserir);
    
    // Aplica formatação de banding
    sheet.getRange(insertRow, 1, linhasParaInserir.length, 5)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  }
  
  return { added: linhasParaInserir.length };
}

function calcularFixadoresRelatoriosComFeedback() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Calculando fixadores em relatórios...', 'Aguarde', -1);
  const result = calcularFixadoresRelatorios();
  
  if (result.success) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ ${result.processed} linhas adicionadas em ${result.sheets} relatórios!`,
      'Sucesso',
      5
    );
  } else {
    SpreadsheetApp.getUi().alert('Erro', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
