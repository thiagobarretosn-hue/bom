/**
 * @OnlyCurrentDoc
 * SISTEMA UNIFICADO - RELATÓRIOS DINÂMICOS + FIXADORES
 * Versão 3.0 - Mapeamento Atualizado
 */

// ============================================================================
// MAPEAMENTO DE COLUNAS ATUALIZADO
// ============================================================================

const COLUMN_MAPPING = {
  SOURCE: {
    PROJECT_NAME: 0,  // A - PROJECT NAME
    SECTION: 1,       // B - SECTION
    QTT: 2,           // C - QTT
    UNIT_ID: 3,       // D - UNIT ID
    UNIT_TYPE: 4,     // E - UNIT TYPE
    LOCAL: 5,         // F - LOCAL
    TRADE: 6,         // G - TRADE
    PHASE: 7,         // H - PHASE
    FLOOR: 8,         // I - FLOOR
    DESC: 9,          // J - DESC
    QTY: 10,          // K - QTY
    UOM: 11,          // L - UOM
    UPC: 12,          // M - UPC
    UNIT_COST: 13,    // N - UNIT COST
    PROJECT: 14       // O - PROJECT
  }
};

// ============================================================================
// CONFIGURAÇÃO GLOBAL
// ============================================================================

const CONFIG = {
  SHEETS: { CONFIG: 'Config' },
  KEYS: {
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
    SORT_BY: 'CLASSIFICAR POR',
    SORT_ORDER: 'ORDEM'
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
  CACHE_TTL: 180,
  DELIMITER: '|||',
  FIXADORES: {
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
        { desc: 'NUT 3/8 IN. METAL', factor: 4 },
        { desc: 'FENDER WASHER 3/8 X 1-1/2', factor: 4 },
        { desc: 'ANCHOR DROP-IN 3/8 IN. X 3/4 IN. LONG HDI-P (W/ AUTO SET TOOL) [HILTI 409499]', factor: 2 },
        { desc: 'PLTD STEEL ALL THREAD ROD 3/8 IN. X 6 FT.', factor: 2 }
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
        { desc: 'NUT 3/8 IN. METAL', factor: 2 },
        { desc: 'FENDER WASHER 3/8 X 1-1/2', factor: 2 },
        { desc: 'ANCHOR DROP-IN 3/8 IN. X 3/4 IN. LONG HDI-P (W/ AUTO SET TOOL) [HILTI 409499]', factor: 1 },
        { desc: 'PLTD STEEL ALL THREAD ROD 3/8 IN. X 6 FT.', factor: 1 }
      ]
    }
  }
};

// ============================================================================
// UTILITÁRIOS
// ============================================================================

const Utils = {
  getColumnIndex: (colConfig) => {
    if (!colConfig || typeof colConfig !== 'string') return -1;
    return colConfig.charAt(0).toUpperCase().charCodeAt(0) - 64;
  },

  getColumnHeader: (colConfig) => {
    if (!colConfig || typeof colConfig !== 'string') return '';
    const parts = colConfig.split(' - ');
    return parts.length > 1 ? parts[1].trim() : '';
  },

  columnToLetter: (col) => {
    let result = '';
    while (col > 0) {
      col--;
      result = String.fromCharCode(65 + (col % 26)) + result;
      col = Math.floor(col / 26);
    }
    return result;
  },

  formatVersion: (input) => {
    if (!input) return '';
    const str = String(input).trim();
    if (!str) return '';
    const num = parseInt(str, 10);
    return (isNaN(num) || num < 1) ? str : (num < 10 ? `0${num}` : `${num}`);
  },

  sanitizeSheetName: (name) => {
    return String(name).replace(/[\/\\\?\*\[\]:]/g, '_').substring(0, 100);
  },

  extractDiameter: (desc) => {
    const patterns = [
      /PIPE\s+(\d+\/\d+)\s+IN/i,
      /PIPE\s+(\d+)\s+IN/i,
      /PIPE\s+(\d+-\d+\/\d+)\s+IN/i
    ];
    
    for (const pattern of patterns) {
      const match = desc.match(pattern);
      if (match) {
        const diam = match[1].replace(/\s+/g, '');
        const mapping = {
          '1/2': '1/2', '3/4': '3/4', '1': '1', '1-1/4': '1-1/4',
          '1-1/2': '1-1/2', '2': '2', '2-1/2': '2-1/2', '3': '3',
          '4': '4', '6': '6', '8': '8', '10': '10', '12': '12'
        };
        return mapping[diam] || diam;
      }
    }
    return null;
  },

  arredondarInteligente: (value) => {
    if (value <= 0) return 0;
    if (value < 1) return 1;
    const base = Math.floor(value);
    return (value - base) <= 0.3 ? base : base + 1;
  }
};

// ============================================================================
// CACHE
// ============================================================================

const CacheManager = {
  _cache: CacheService.getScriptCache(),
  
  get: (key) => {
    const cached = CacheManager._cache.get(key);
    return cached ? JSON.parse(cached) : null;
  },
  
  put: (key, value, ttl = CONFIG.CACHE_TTL) => {
    CacheManager._cache.put(key, JSON.stringify(value), ttl);
  },
  
  remove: (key) => {
    CacheManager._cache.remove(key);
  },
  
  invalidateAll: () => {
    CacheManager._cache.removeAll(['all_config_values', 'unique_values']);
  }
};

// ============================================================================
// CONFIGURAÇÃO
// ============================================================================

const ConfigService = {
  getAll: () => {
    const cached = CacheManager.get('all_config_values');
    if (cached) return cached;

    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CONFIG);
    if (!configSheet) return {};

    const values = configSheet.getRange("A1:B" + configSheet.getLastRow()).getDisplayValues();
    const config = {};
    
    values.forEach(row => {
      const key = String(row[0]).trim();
      if (key && !key.startsWith(' ')) config[key] = row[1];
    });

    CacheManager.put('all_config_values', config);
    return config;
  },

  get: (key, defaultValue = '') => {
    return ConfigService.getAll()[key] || defaultValue;
  }
};

// ============================================================================
// CRIAÇÃO DA ABA CONFIG
// ============================================================================

function ensureConfigExists() {
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CONFIG)) {
    forceCreateConfig();
  }
}

function forceCreateConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  if (configSheet) ss.deleteSheet(configSheet);
  
  configSheet = ss.insertSheet(CONFIG.SHEETS.CONFIG, 0);
  configSheet.clear();

  const configData = [
    ['CONFIGURAÇÃO', 'VALOR', 'DESCRIÇÃO'],
    ['', '', ''],
    [' 📊 AGRUPAMENTO', '', 'Defina as colunas para agrupar e os painéis serão atualizados automaticamente.'],
    [CONFIG.KEYS.SOURCE_SHEET, '', 'Aba que contém todos os dados brutos.'],
    [CONFIG.KEYS.GROUP_L1, '', 'Coluna principal de agrupamento (obrigatório).'],
    [CONFIG.KEYS.GROUP_L2, '', 'Subnível de agrupamento (opcional).'],
    [CONFIG.KEYS.GROUP_L3, '', 'Terceiro nível de agrupamento (opcional).'],
    ['', '', ''],
    [' 📋 DADOS BOMS', '', 'Mapeamento das colunas para a criação das listas de materiais (BOMs).'],
    [CONFIG.KEYS.COL_1, '', 'Coluna para o ID principal ou identificador do item.'],
    [CONFIG.KEYS.COL_2, 'J - DESC', 'Coluna para a descrição do item.'],
    [CONFIG.KEYS.COL_3, 'M - UPC', 'Coluna para o código de barras ou UPC.'],
    [CONFIG.KEYS.COL_4, 'L - UOM', 'Coluna para a unidade de medida (Unit of Measure).'],
    [CONFIG.KEYS.COL_5, 'O - PROJECT', 'Coluna com a quantidade a ser somada.'],
    ['', '', ''],
    [' 🏷️ CABEÇALHO', '', 'Informações que aparecerão no cabeçalho de cada relatório gerado.'],
    [CONFIG.KEYS.PROJECT, 'HG1 BE', 'Nome principal do projeto.'],
    [CONFIG.KEYS.BOM, 'RISERS JS', 'Nome da lista de materiais (Bill of Materials).'],
    [CONFIG.KEYS.KOJO_PREFIX, 'HG1.PLB.RGH.JS.BE.UNT.RISER', 'Prefixo fixo para o código KOJO.'],
    [CONFIG.KEYS.ENGINEER, 'WANDERSON', 'Nome do engenheiro responsável.'],
    [CONFIG.KEYS.VERSION, '', 'Versão do relatório (ex: 01, 02, 03...).'],
    ['', '', ''],
    [' ⚙️ OPÇÕES DE RELATÓRIO', '', 'Defina opções adicionais para a geração dos relatórios.'],
    [CONFIG.KEYS.SORT_BY, 'J - DESC', 'Coluna a ser usada para classificar os itens no relatório.'],
    [CONFIG.KEYS.SORT_ORDER, 'Ascendente (A-Z, 0-9)', 'A ordem de classificação (ascendente ou descendente).'],
    ['', '', ''],
    [' 💾 SALVAMENTO', '', 'Configurações para exportação e salvamento dos arquivos.'],
    [CONFIG.KEYS.DRIVE_FOLDER_ID, '', 'Cole o LINK COMPLETO da pasta ou apenas o ID.'],
    [CONFIG.KEYS.DRIVE_FOLDER_NAME, '', 'Nome da pasta a ser criada caso o link/ID não seja fornecido.'],
    [CONFIG.KEYS.PDF_PREFIX, 'HG1.PLB.RGH.JS.BE.UNT.RISER', 'Prefixo para os nomes dos arquivos PDF exportados.']
  ];

  configSheet.getRange(1, 1, configData.length, 3).setValues(configData);
  configSheet.getRange("A1:C" + configData.length).setFontFamily(CONFIG.COLORS.FONT_FAMILY).setVerticalAlignment('middle').setFontColor(CONFIG.COLORS.FONT_DARK);
  configSheet.setColumnWidth(1, 220).setColumnWidth(2, 300).setColumnWidth(3, 350);
  configSheet.getRange('A1:C1').merge().setValue('⚙️ PAINEL DE CONFIGURAÇÃO | RELATÓRIOS DINÂMICOS').setBackground(CONFIG.COLORS.HEADER_BG).setFontColor(CONFIG.COLORS.FONT_LIGHT).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  
  const sections = { 3: { endRow: 7 }, 9: { endRow: 14 }, 16: { endRow: 21 }, 23: { endRow: 26 }, 28: { endRow: 31 } };

  for (const startRow in sections) {
    const s = sections[startRow], start = parseInt(startRow);
    configSheet.getRange(start, 1, 1, 3).merge().setBackground(CONFIG.COLORS.SECTION_BG).setFontColor(CONFIG.COLORS.FONT_LIGHT).setFontSize(11).setFontWeight('bold').setHorizontalAlignment('left');
    configSheet.setRowHeight(start, 30);
    for (let r = start + 1; r <= s.endRow; r++) {
      configSheet.getRange(r, 2).setBackground(CONFIG.COLORS.INPUT_BG).setHorizontalAlignment('center');
      configSheet.getRange(r, 1).setFontWeight('500');
      configSheet.getRange(r, 3).setFontStyle('italic').setFontColor(CONFIG.COLORS.FONT_SUBTLE);
      configSheet.setRowHeight(r, 28);
    }
    configSheet.getRange(start, 1, s.endRow - start + 1, 3).setBorder(true, true, true, true, null, null, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }
  
  configSheet.getRange(21, 2).setNumberFormat('@STRING@');
  configSheet.setFrozenRows(1);
  
  configSheet.getRange('E1:M1').merge().setValue('PAINEL UNIFICADO DE AGRUPAMENTO E PRÉ-VISUALIZAÇÃO')
      .setBackground(CONFIG.COLORS.HEADER_BG).setFontColor(CONFIG.COLORS.FONT_LIGHT).setFontSize(14)
      .setFontWeight('bold').setHorizontalAlignment('center').setFontFamily(CONFIG.COLORS.FONT_FAMILY);
  configSheet.setRowHeight(1, 40);

  SpreadsheetApp.flush();
  Utilities.sleep(100);
  updateConfigDropdowns();
}

function updateConfigDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  if (!configSheet) return;
  
  const config = ConfigService.getAll();
  const sourceSheets = ss.getSheets().map(s => s.getName()).filter(n => n !== CONFIG.SHEETS.CONFIG);
  
  if (sourceSheets.length > 0) {
    configSheet.getRange('B4').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(sourceSheets, true).setAllowInvalid(false).build()
    );
  }

  const sourceSheet = ss.getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  if (!sourceSheet || sourceSheet.getLastColumn() === 0) return;
  
  const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const columnOptions = headers.map((h, i) => `${Utils.columnToLetter(i + 1)} - ${h || `Coluna ${Utils.columnToLetter(i + 1)}`}`);
  const columnRule = SpreadsheetApp.newDataValidation().requireValueInList(columnOptions, true).setAllowInvalid(false).build();
  
  [5, 6, 7, 10, 11, 12, 13, 14].forEach(row => configSheet.getRange(row, 2).setDataValidation(columnRule));
  configSheet.getRange(24, 2).setDataValidation(columnRule);

  const sortOrderOptions = ['Ascendente (A-Z, 0-9)', 'Descendente (Z-A, 9-0)'];
  configSheet.getRange(25, 2).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(sortOrderOptions, true).setAllowInvalid(false).build()
  );
}

// ============================================================================
// PAINEL DE AGRUPAMENTO
// ============================================================================

function getUniqueColumnValues(sheet, columnIndex) {
  const cache = CacheManager.get(`unique_${sheet.getSheetId()}_${columnIndex}`);
  if (cache) return cache;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const values = sheet.getRange(2, columnIndex, lastRow - 1).getValues().flat();
  const uniqueValues = [...new Set(values.filter(v => v))].sort((a, b) => 
    String(a).localeCompare(String(b), undefined, { numeric: true })
  );
  
  CacheManager.put(`unique_${sheet.getSheetId()}_${columnIndex}`, uniqueValues);
  return uniqueValues;
}

function updateSingleLevelPanel(level, config, sourceSheet) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CONFIG);
  const startCol = 5 + (level - 1) * 2;
  const levelId = `NÍVEL ${level}`;
  const groupConfigKey = `Agrupar por Nível ${level}`;
  const groupConfig = config[groupConfigKey];

  configSheet.getRange(3, startCol, configSheet.getMaxRows() - 2, 2).clear();
  configSheet.setColumnWidth(startCol, 150).setColumnWidth(startCol + 1, 50);
  const headerRange = configSheet.getRange(3, startCol, 1, 2).merge();
  
  if (sourceSheet && groupConfig && groupConfig.trim() !== '') {
    const colIndex = Utils.getColumnIndex(groupConfig);
    if (colIndex === -1) {
      headerRange.setValue(`${levelId}: ERRO`).setBackground(CONFIG.COLORS.PANEL_ERROR_BG).setFontColor(CONFIG.COLORS.FONT_LIGHT).setFontWeight('bold').setHorizontalAlignment('center');
      configSheet.getRange(4, startCol).setValue('Config. inválida.').setFontColor(CONFIG.COLORS.FONT_SUBTLE).setFontStyle('italic');
    } else {
      const colHeader = Utils.getColumnHeader(groupConfig);
      headerRange.setValue(`${levelId}: ${colHeader.toUpperCase()}`).setBackground(CONFIG.COLORS.SECTION_BG).setFontColor(CONFIG.COLORS.FONT_LIGHT).setFontWeight('bold').setHorizontalAlignment('center');
      
      const uniqueValues = getUniqueColumnValues(sourceSheet, colIndex);
      if (uniqueValues.length > 0) {
        const valuesFormatted = uniqueValues.map(v => [v === null || v === undefined || v === '' ? '' : v]);
        configSheet.getRange(4, startCol, valuesFormatted.length, 1).setValues(valuesFormatted);
        configSheet.getRange(4, startCol + 1, valuesFormatted.length, 1).insertCheckboxes();
      }
    }
  } else {
    headerRange.setValue(levelId).setBackground(CONFIG.COLORS.PANEL_EMPTY_BG).setFontColor(CONFIG.COLORS.FONT_LIGHT).setFontWeight('bold').setHorizontalAlignment('center');
    const message = sourceSheet ? '(Não configurado)' : '(Selecione Aba Origem)';
    configSheet.getRange(4, startCol).setValue(message).setFontColor(CONFIG.COLORS.FONT_SUBTLE).setFontStyle('italic');
  }

  const lastRowInPanel = configSheet.getRange(configSheet.getMaxRows(), startCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const effectiveLastRow = Math.max(3, lastRowInPanel);
  configSheet.getRange(3, startCol, effectiveLastRow - 2, 2).setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
}

function updateGroupingPanel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ConfigService.getAll();
  const sourceSheet = ss.getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  
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
        .map((v, index) => ({ value: v, checked: checkboxes[index] }))
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

// ============================================================================
// PRÉ-VISUALIZAÇÃO DE COMBINAÇÕES
// ============================================================================

function updatePreviewPanel(previewStartCol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  const internalDataCol = previewStartCol + 3;

  const combinations = getSelectedCombinationsFromPanel();
  const existingSuffixes = new Map();
  const lastPreviewRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  
  if (lastPreviewRow >= 4) {
    configSheet.getRange(4, previewStartCol, lastPreviewRow - 3, 3).getValues().forEach(row => {
      if (row[0]) existingSuffixes.set(row[0], row[2]);
    });
  }
  
  configSheet.getRange(3, previewStartCol, configSheet.getMaxRows() - 2, 4).clear();
  configSheet.getRange(3, previewStartCol, 1, 3).setValues([['COMBINAÇÃO GERADA', 'CRIAR?', 'SUFIXO KOJO FINAL']])
    .setBackground(CONFIG.COLORS.HEADER_BG).setFontColor(CONFIG.COLORS.FONT_LIGHT).setFontWeight('bold');
  
  if (combinations.length > 0) {
    const tableData = combinations.map(combo => {
      const displayCombo = combo.replace(/\|\|\|/g, '.');
      const kojoSuffix = existingSuffixes.get(combo) || displayCombo;
      return [displayCombo, true, kojoSuffix, combo];
    });
    configSheet.getRange(4, previewStartCol, tableData.length, 4).setValues(tableData);
    configSheet.getRange(4, previewStartCol + 1, tableData.length, 1).insertCheckboxes();
    configSheet.getRange(4, previewStartCol + 2, tableData.length, 1).setBackground(CONFIG.COLORS.INPUT_BG);
  }
  
  configSheet.setColumnWidth(previewStartCol, 250).setColumnWidth(previewStartCol + 1, 80).setColumnWidth(previewStartCol + 2, 250);
  configSheet.hideColumns(internalDataCol);

  const lastCombinationRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const effectiveLastRow = Math.max(3, lastCombinationRow);
  configSheet.getRange(3, previewStartCol, effectiveLastRow - 2, 3).setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  
  ss.toast('Pré-visualização atualizada!', 'Sucesso', 3);
}

function getSelectedCombinationsFromPanel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  const config = ConfigService.getAll();
  
  const sourceSheet = ss.getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  if (!sourceSheet) return [];
  
  const groupConfigs = [
    config[CONFIG.KEYS.GROUP_L1], 
    config[CONFIG.KEYS.GROUP_L2], 
    config[CONFIG.KEYS.GROUP_L3]
  ].filter(Boolean);
  
  if (groupConfigs.length === 0) return [];

  const groupIndices = groupConfigs.map(Utils.getColumnIndex);
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
      existingCombinations.add(combinationParts.join(CONFIG.DELIMITER));
    }
  });

  const finalCombinations = [...existingCombinations].filter(combo => {
    const parts = combo.split(CONFIG.DELIMITER);
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
    const aParts = a.split(CONFIG.DELIMITER);
    const bParts = b.split(CONFIG.DELIMITER);
    for (let i = 0; i < Math.min(aParts.length, bParts.length); i++) {
      const aNum = parseFloat(aParts[i]);
      const bNum = parseFloat(bParts[i]);
      
      if (!isNaN(aNum) && !isNaN(bNum)) {
        if (aNum !== bNum) return aNum - bNum;
      } else {
        const comparison = aParts[i].localeCompare(bParts[i], undefined, { numeric: true });
        if (comparison !== 0) return comparison;
      }
    }
    return aParts.length - bParts.length;
  });
}

// ============================================================================
// PROCESSAMENTO DE RELATÓRIOS
// ============================================================================

function runProcessing() {
  CacheManager.invalidateAll();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ConfigService.getAll();
  const configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  
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
  
  const sourceSheet = ss.getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  if (!sourceSheet) return { success: false, message: 'Aba de origem não encontrada.' };

  const groupConfigs = [config[CONFIG.KEYS.GROUP_L1], config[CONFIG.KEYS.GROUP_L2], config[CONFIG.KEYS.GROUP_L3]].filter(Boolean);
  const groupIndices = groupConfigs.map(Utils.getColumnIndex);
  
  const bomCols = {
    c1: Utils.getColumnIndex(config[CONFIG.KEYS.COL_1]),
    c2: Utils.getColumnIndex(config[CONFIG.KEYS.COL_2]),
    c3: Utils.getColumnIndex(config[CONFIG.KEYS.COL_3]),
    c4: Utils.getColumnIndex(config[CONFIG.KEYS.COL_4]),
    c5: Utils.getColumnIndex(config[CONFIG.KEYS.COL_5])
  };

  const sortColumnConfig = config[CONFIG.KEYS.SORT_BY] || 'J - DESC';
  const sortColumnIndex = Utils.getColumnIndex(sortColumnConfig) - 1;
  const sortOrder = config[CONFIG.KEYS.SORT_ORDER] === 'Descendente (Z-A, 9-0)' ? 'desc' : 'asc';
  
  if (sortColumnIndex < 0) return { success: false, message: `Coluna de classificação "${sortColumnConfig}" inválida.` };

  const dataMap = new Map();
  combinationsToProcess.forEach(item => dataMap.set(item.combination, []));

  const allData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
  
  for (const row of allData) {
    const rowCombination = groupIndices.map(index => {
      const value = row[index - 1];
      return value === null || value === undefined || value === '' ? '' : String(value).trim();
    }).join(CONFIG.DELIMITER);
    
    if (dataMap.has(rowCombination)) {
      dataMap.get(rowCombination).push([  
        row[bomCols.c1 - 1], row[bomCols.c2 - 1], row[bomCols.c3 - 1], 
        row[bomCols.c4 - 1], parseFloat(row[bomCols.c5 - 1]) || 0,
        row[sortColumnIndex]
      ]);
    }
  }
  
  let createdCount = 0;
  
  for (const item of combinationsToProcess) {
    const { displayCombination, combination, kojoSuffix } = item;
    const rawData = dataMap.get(combination);
    if (!rawData || rawData.length === 0) continue;
    
    const processedData = groupAndSumData(rawData, sortOrder);
    const sanitizedName = Utils.sanitizeSheetName(displayCombination);
    
    let targetSheet = ss.getSheetByName(sanitizedName) || ss.insertSheet(sanitizedName);
    targetSheet.clear();

    createAndFormatReport(targetSheet, combination, kojoSuffix, processedData);
    createdCount++;
  }
  
  return { success: true, created: createdCount };
}

function groupAndSumData(data, sortOrder) {
  const grouped = {};
  
  data.forEach(row => {
    const key = `${row[0]}|${row[1]}|${row[2]}|${row[3]}`;
    if (grouped[key]) {
      grouped[key][4] += row[4];
    } else {
      grouped[key] = [...row];
    }
  });

  const groupedData = Object.values(grouped);
  const direction = sortOrder === 'asc' ? 1 : -1;

  groupedData.sort((a, b) => {
    const valA = a[5];
    const valB = b[5];

    if (valA === null || valA === undefined || valA === '') return 1;
    if (valB === null || valB === undefined || valB === '') return -1;

    const aIsNum = !isNaN(parseFloat(valA)) && isFinite(valA);
    const bIsNum = !isNaN(parseFloat(valB)) && isFinite(valB);

    if (aIsNum && bIsNum) {
      return (parseFloat(valA) - parseFloat(valB)) * direction;
    }

    return String(valA).localeCompare(String(valB), undefined, { numeric: true }) * direction;
  });
  
  return groupedData.map(row => row.slice(0, 5));
}

function createAndFormatReport(sheet, combination, kojoSuffix, data) {
  const config = ConfigService.getAll();
  const reportConfig = { 
    project: config[CONFIG.KEYS.PROJECT], 
    bom: config[CONFIG.KEYS.BOM], 
    kojoPrefix: config[CONFIG.KEYS.KOJO_PREFIX], 
    engineer: config[CONFIG.KEYS.ENGINEER], 
    version: config[CONFIG.KEYS.VERSION]
  };
  
  const headers = { 
    h1: Utils.getColumnHeader(config[CONFIG.KEYS.COL_1]), 
    h2: Utils.getColumnHeader(config[CONFIG.KEYS.COL_2]), 
    h3: Utils.getColumnHeader(config[CONFIG.KEYS.COL_3]), 
    h4: Utils.getColumnHeader(config[CONFIG.KEYS.COL_4]), 
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
    ['PROJECT:', reportConfig.project],
    ['BOM:', reportConfig.bom],
    ['BOM KOJO:', bomKojoComplete],
    ['ENG.:', reportConfig.engineer],
    ['VERSION:', reportConfig.version],
    ['LAST UPDATE:', lastUpdate]
  ];
}

function formatHeader(sheet, headerLength) {
  sheet.getRange(1, 2, headerLength, 1).setNumberFormat('@STRING@');
  for (let r = 1; r <= headerLength; r++) {
    sheet.getRange(r, 2, 1, 4).merge();
  }
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
  } catch (e) {
    Logger.log('Erro ao proteger cabeçalho:', e.message);
  }
}

function clearOldReports() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirmação', 'Apagar TODAS as abas de relatório?', ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = ConfigService.getAll();
    const protectedSheets = [CONFIG.SHEETS.CONFIG, config[CONFIG.KEYS.SOURCE_SHEET]].filter(Boolean);
    
    let deletedCount = 0;
    ss.getSheets().forEach(sheet => {
      const sheetName = sheet.getName();
      if (!protectedSheets.includes(sheetName)) {
        ss.deleteSheet(sheet);
        deletedCount++;
      }
    });
    
    if (deletedCount > 0) {
      ui.alert('Limpeza Concluída', `${deletedCount} abas removidas.`, ui.ButtonSet.OK);
    } else {
      ui.alert('Limpeza', 'Nenhuma aba para remover.', ui.ButtonSet.OK);
    }
  }
}

// ============================================================================
// SISTEMA DE FIXADORES
// ============================================================================

function abrirSeletorFixadores() {
  const html = HtmlService.createHtmlOutputFromFile('FixadoresSidebar')
    .setTitle('🔧 Incluir Fixadores')
    .setWidth(750)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, '🔧 Incluir Fixadores');
}

function getPipesElegiveis() {
  const config = ConfigService.getAll();
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  
  if (!sourceSheet || sourceSheet.getLastRow() < 2) return [];
  
  const COL = COLUMN_MAPPING.SOURCE;
  const lastRow = sourceSheet.getLastRow();
  const data = sourceSheet.getRange(2, 1, lastRow - 1, 15).getValues();
  
  return data
    .filter(row => {
      const desc = String(row[COL.DESC] || '').trim().toUpperCase();
      const trade = String(row[COL.TRADE] || '').trim().toUpperCase();
      
      return (desc.includes('PIPE') && 
              (desc.includes('RISER') || desc.includes('COLGANTE')) &&
              (trade === 'WS' || trade === '10.01'));
    })
    .map(row => {
      const desc = String(row[COL.DESC] || '').trim();
      const section = String(row[COL.SECTION] || '').trim();
      const rowIndex = data.indexOf(row) + 2;
      
      const nextRow = rowIndex < lastRow ? sourceSheet.getRange(rowIndex + 1, COL.DESC + 1).getValue() : '';
      const jaTemFixador = String(nextRow).includes('LOOP HANG') || String(nextRow).includes('RISER CLAMP');
      
      return {
        row: rowIndex,
        section: section,
        trade: String(row[COL.TRADE] || '').trim(),
        floor: String(row[COL.FLOOR] || '').trim(),
        unitType: String(row[COL.UNIT_TYPE] || '').trim(),
        desc: desc,
        qty: parseFloat(row[COL.QTY]) || 0,
        uom: String(row[COL.UOM] || '').trim(),
        diameter: Utils.extractDiameter(desc),
        jaTemFixador: jaTemFixador
      };
    });
}

function processarFixadoresSelecionados(selectedPipes) {
  if (!selectedPipes || selectedPipes.length === 0) {
    return { success: false, message: 'Nenhum tubo selecionado' };
  }
  
  const config = ConfigService.getAll();
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  
  if (!sourceSheet) return { success: false, message: 'Aba de origem não encontrada' };
  
  const COL = COLUMN_MAPPING.SOURCE;
  let totalAdded = 0;
  
  selectedPipes.forEach(pipe => {
    const insertRow = pipe.row + 1;
    const isRiser = pipe.desc.toUpperCase().includes('RISER');
    const fixConfig = isRiser ? CONFIG.FIXADORES.RISER : CONFIG.FIXADORES.LOOP;
    
    const fixadorPrincipal = isRiser 
      ? fixConfig.clamps[pipe.diameter] 
      : fixConfig.hangs[pipe.diameter];
    
    if (!fixadorPrincipal) return;
    
    const qtdFixador = Utils.arredondarInteligente(pipe.qty / fixConfig.interval);
    if (qtdFixador === 0) return;
    
    const linhasParaInserir = [];
    
    // Linha do fixador principal
    const linhaFixador = new Array(15).fill('');
    linhaFixador[COL.SECTION] = pipe.section;
    linhaFixador[COL.QTT] = 1;
    linhaFixador[COL.UNIT_ID] = '';
    linhaFixador[COL.UNIT_TYPE] = '';
    linhaFixador[COL.LOCAL] = pipe.section;
    linhaFixador[COL.TRADE] = 'FIX';
    linhaFixador[COL.PHASE] = 'Job Site';
    linhaFixador[COL.FLOOR] = pipe.floor;
    linhaFixador[COL.DESC] = fixadorPrincipal;
    linhaFixador[COL.QTY] = `=RC[${COL.QTY - COL.QTT}]*${qtdFixador}`;
    linhaFixador[COL.UOM] = '#NAME?';
    
    linhasParaInserir.push(linhaFixador);
    
    // Materiais auxiliares
    fixConfig.materials.forEach(mat => {
      const linhaMaterial = new Array(15).fill('');
      linhaMaterial[COL.SECTION] = pipe.section;
      linhaMaterial[COL.QTT] = 1;
      linhaMaterial[COL.LOCAL] = pipe.section;
      linhaMaterial[COL.TRADE] = 'FIX';
      linhaMaterial[COL.PHASE] = 'Job Site';
      linhaMaterial[COL.FLOOR] = pipe.floor;
      linhaMaterial[COL.DESC] = mat.desc;
      linhaMaterial[COL.QTY] = `=RC[${COL.QTY - COL.QTT}]*${qtdFixador * mat.factor}`;
      linhaMaterial[COL.UOM] = '#NAME?';
      
      linhasParaInserir.push(linhaMaterial);
    });
    
    // Insere as linhas
    sourceSheet.insertRowsAfter(insertRow - 1, linhasParaInserir.length);
    sourceSheet.getRange(insertRow, 1, linhasParaInserir.length, 15).setValues(linhasParaInserir);
    
    // Aplica fórmulas R1C1
    linhasParaInserir.forEach((row, idx) => {
      const currentRow = insertRow + idx;
      const formulaK = row[COL.QTY];
      
      if (typeof formulaK === 'string' && formulaK.startsWith('=')) {
        sourceSheet.getRange(currentRow, COL.QTY + 1).setFormulaR1C1(formulaK);
      }
    });
    
    totalAdded += linhasParaInserir.length;
  });
  
  return { success: true, added: totalAdded };
}

// ============================================================================
// EXPORTAÇÃO PDF
// ============================================================================

function exportSelectedPDFs(selectedSheets) {
  if (!selectedSheets || selectedSheets.length === 0) {
    return { success: false, message: 'Nenhuma aba selecionada' };
  }
  
  const folder = getOrCreateOutputFolder();
  if (!folder) return { success: false, message: 'Pasta não configurada' };
  
  const config = ConfigService.getAll();
  const prefix = config[CONFIG.KEYS.PDF_PREFIX] || 'Relatorio';
  
  selectedSheets.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet) exportSheetToPdf(sheet, `${prefix}_${sheetName}`, folder);
  });
  
  return { success: true, exported: selectedSheets.length, folder: folder.getName() };
}

function getOrCreateOutputFolder() {
  const config = ConfigService.getAll();
  const folderInput = config[CONFIG.KEYS.DRIVE_FOLDER_ID];
  const folderName = config[CONFIG.KEYS.DRIVE_FOLDER_NAME];
  
  try {
    if (folderInput) {
      const match = folderInput.match(/folders\/([a-zA-Z0-9_-]+)/);
      const folderId = match ? match[1] : folderInput;
      return DriveApp.getFolderById(folderId);
    }
    
    if (folderName) {
      const folders = DriveApp.getFoldersByName(folderName);
      return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    }
    
    const defaultName = `${SpreadsheetApp.getActiveSpreadsheet().getName()} - PDFs`;
    const folders = DriveApp.getFoldersByName(defaultName);
    return folders.hasNext() ? folders.next() : DriveApp.createFolder(defaultName);
  } catch (error) {
    Logger.log(`Erro ao acessar pasta: ${error.message}`);
    return null;
  }
}

function exportSheetToPdf(sheet, pdfName, folder) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` +
    `gid=${sheet.getSheetId()}&format=pdf&size=A4&portrait=true&fitw=true&` +
    `sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;
  
  for (let attempt = 0, delay = 1000; attempt < 5; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() === 200) {
        folder.createFile(response.getBlob().setName(`${pdfName}.pdf`));
        return;
      }
      
      if (response.getResponseCode() === 429) {
        Utilities.sleep(delay + Math.floor(Math.random() * 500));
        delay *= 2;
        continue;
      }
      
      throw new Error(`Código HTTP ${response.getResponseCode()}`);
    } catch (error) {
      if (attempt === 4) throw error;
      Utilities.sleep(delay);
    }
  }
}

function getReportSheetNames() {
  const config = ConfigService.getAll();
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(s => s.getName())
    .filter(name => name !== CONFIG.SHEETS.CONFIG && name !== config[CONFIG.KEYS.SOURCE_SHEET]);
}

// ============================================================================
// TRIGGERS E UI
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔧 Relatórios Dinâmicos')
    .addItem('⚙️ Painel de Controle', 'openConfigSidebar')
    .addSeparator()
    .addItem('📊 Processar Relatórios', 'runProcessingWithFeedback')
    .addItem('🔧 Fixadores → Fonte', 'abrirSeletorFixadores')
    .addItem('📄 Exportar PDFs', 'exportPDFsWithFeedback')
    .addSeparator()
    .addItem('🗑️ Limpar Relatórios', 'clearOldReports')
    .addItem('🔄 Limpar Cache', 'forceRefreshCache')
    .addItem('🧪 Diagnóstico', 'testSystem')
    .addItem('🔧 Recriar Config', 'forceCreateConfig')
    .addToUi();
  
  ensureConfigExists();
}

function onEdit(e) {
  if (!e || !e.source) return;
  
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (sheetName === CONFIG.SHEETS.CONFIG) {
    CacheManager.invalidateAll();
    
    const range = e.range;
    const col = range.getColumn();
    const row = range.getRow();
    
    if (col === 2 && row === 21) {
      const formatted = Utils.formatVersion(range.getValue());
      if (formatted !== String(range.getValue())) {
        range.setValue(formatted);
      }
    }
    
    if (col === 2 && row >= 4 && row <= 7) {
      Utilities.sleep(200);
      const config = ConfigService.getAll();
      const sourceSheet = SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
      
      if (row === 4) {
        updateConfigDropdowns();
        updateGroupingPanel();
      } else {
        updateSingleLevelPanel(row - 4, config, sourceSheet);
      }
      
      updatePreviewPanel(11);
    }
    
    if ((col === 6 || col === 8 || col === 10) && row >= 4) {
      updatePreviewPanel(11);
    }
  } else {
    try {
      const config = ConfigService.getAll();
      if (sheetName === config[CONFIG.KEYS.SOURCE_SHEET]) {
        CacheManager.invalidateAll();
      }
    } catch (error) {
      // Config ainda não existe
    }
  }
}

function openConfigSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ConfigSidebar')
    .setTitle('Painel de Controle')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

function runProcessingWithFeedback() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Processando...', 'Aguarde', -1);
  const result = runProcessing();
  
  if (result.success) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ ${result.created} relatórios criados!`,
      'Sucesso',
      5
    );
  } else {
    SpreadsheetApp.getUi().alert('Erro', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
  
  return result;
}

function exportPDFsWithFeedback() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Exportando PDFs...', 'Aguarde', -1);
  const result = exportSelectedPDFs(getReportSheetNames());
  
  if (result.success) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ ${result.exported} PDFs exportados!`,
      'Sucesso',
      5
    );
  } else {
    SpreadsheetApp.getUi().alert('Erro', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function forceRefreshCache() {
  CacheManager.invalidateAll();
  updateGroupingPanel();
  updatePreviewPanel(11);
  SpreadsheetApp.getActiveSpreadsheet().toast('Cache limpo!', 'Sucesso', 3);
}

function testSystem() {
  try {
    const config = ConfigService.getAll();
    const sourceSheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
    
    const msg = [
      `Aba de Origem: ${config[CONFIG.KEYS.SOURCE_SHEET] || 'Não configurada'}`,
      `Linhas: ${sourceSheet ? sourceSheet.getLastRow() - 1 : 0}`,
      `Relatórios: ${getReportSheetNames().length}`,
      `Mapeamento: FLOOR=I(9), DESC=J(10), QTY=K(11)`
    ].join('\n');
    
    SpreadsheetApp.getUi().alert('🧪 Diagnóstico', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
