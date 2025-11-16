/**
 * @OnlyCurrentDoc
 * SISTEMA UNIFICADO DE RELATÓRIOS DINÂMICOS + FIXADORES
 * Versão 2.5 - Adicionada Função de REMOVER Fixadores
 *
 * NOVAS FUNCIONALIDADES:
 * 1. (FIXADORES) Nova função `removerFixadoresSelecionados` no backend.
 * 2. Esta função recebe os tubos processados, encontra as linhas 'FIX' 
 * abaixo deles e as apaga em ordem inversa (para segurança).
 * 3. (BOM) Mantém todas as correções anteriores (Classificação Avançada, 
 * colunas J-DESC, O-PROJECT, etc.).
 * 4. (UX) Mantém o `showModelessDialog` (não-modal) que você gostou.
 */

// ============================================================================
// CONFIGURAÇÃO GLOBAL
// ============================================================================

const CONFIG = {
  SHEETS: {
    CONFIG: 'Config'
  },
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
    CacheManager._cache.removeAll([
      'all_config_values',
      'unique_values'
    ]);
    // Limpa chaves de cache de coluna únicas
    const config = ConfigService.getAll();
    const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
    if (sourceSheet) {
        const sourceSheetId = sourceSheet.getSheetId();
        const groupConfigs = [
            config[CONFIG.KEYS.GROUP_L1],
            config[CONFIG.KEYS.GROUP_L2],
            config[CONFIG.KEYS.GROUP_L3]
        ].filter(Boolean);
        const cacheKeysToRemove = groupConfigs.map(conf => {
            const colIndex = Utils.getColumnIndex(conf);
            if (colIndex !== -1) {
                return `unique_${sourceSheetId}_${colIndex}`;
            }
            return null;
        }).filter(Boolean);
        if (cacheKeysToRemove.length > 0) {
            CacheManager._cache.removeAll(cacheKeysToRemove);
        }
    }
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
// CRIAÇÃO E ATUALIZAÇÃO DA ABA CONFIG
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
  
  // Valores padrão corrigidos para a nova estrutura de colunas
  const configData = [
    ['CONFIGURAÇÃO', 'VALOR', 'DESCRIÇÃO'],
    ['', '', ''],
    [' 📊 AGRUPAMENTO', '', 'Defina as colunas para agrupar'],
    [CONFIG.KEYS.SOURCE_SHEET, '', 'Aba com dados brutos'],
    [CONFIG.KEYS.GROUP_L1, '', 'Coluna principal de agrupamento'],
    [CONFIG.KEYS.GROUP_L2, '', 'Segundo nível (opcional)'],
    [CONFIG.KEYS.GROUP_L3, '', 'Terceiro nível (opcional)'],
    ['', '', ''],
    [' 📋 DADOS BOMS', '', 'Mapeamento das colunas'],
    [CONFIG.KEYS.COL_1, 'D - UNIT ID', 'ID ou identificador (ex: D - UNIT ID)'],
    [CONFIG.KEYS.COL_2, 'J - DESC', 'Descrição do item (ex: J - DESC)'],
    [CONFIG.KEYS.COL_3, 'M - UPC', 'Código de barras (ex: M - UPC)'],
    [CONFIG.KEYS.COL_4, 'L - UOM', 'Unidade de medida (ex: L - UOM)'],
    [CONFIG.KEYS.COL_5, 'O - PROJECT', 'Quantidade (ex: O - PROJECT)'],
    ['', '', ''],
    [' 🏷️ CABEÇALHO', '', 'Informações do relatório'],
    [CONFIG.KEYS.PROJECT, 'HG1 BE', 'Nome do projeto'],
    [CONFIG.KEYS.BOM, 'RISERS JS', 'Nome da BOM'],
    [CONFIG.KEYS.KOJO_PREFIX, 'HG1.PLB.RGH.JS.BE.UNT.RISER', 'Prefixo KOJO'],
    [CONFIG.KEYS.ENGINEER, 'WANDERSON', 'Engenheiro'],
    [CONFIG.KEYS.VERSION, '', 'Versão (ex: 01, 02)'],
    ['', '', ''],
    [' ⚙️ OPÇÕES', '', 'Classificação dos dados'],
    [CONFIG.KEYS.SORT_BY, 'J - DESC', 'Coluna da Aba Origem para classificar'],
    [CONFIG.KEYS.SORT_ORDER, 'Ascendente (A-Z, 0-9)', 'Ordem de classificação'],
    ['', '', ''],
    [' 💾 SALVAMENTO', '', 'Exportação de PDFs'],
    [CONFIG.KEYS.DRIVE_FOLDER_ID, '', 'Link ou ID da pasta'],
    [CONFIG.KEYS.DRIVE_FOLDER_NAME, '', 'Nome da pasta'],
    [CONFIG.KEYS.PDF_PREFIX, 'HG1.PLB.RGH.JS.BE.UNT.RISER', 'Prefixo dos PDFs']
  ];
  
  configSheet.getRange(1, 1, configData.length, 3).setValues(configData);
  
  // Formatação
  const s = CONFIG.COLORS;
  configSheet.getRange("A1:C" + configData.length)
    .setFontFamily(s.FONT_FAMILY)
    .setVerticalAlignment('middle')
    .setFontColor(s.FONT_DARK);
  
  configSheet.setColumnWidth(1, 220).setColumnWidth(2, 300).setColumnWidth(3, 350);
  configSheet.getRange('A1:C1').merge()
    .setValue('⚙️ PAINEL DE CONFIGURAÇÃO | RELATÓRIOS DINÂMICOS')
    .setBackground(s.HEADER_BG)
    .setFontColor(s.FONT_LIGHT)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  const sections = {
    3: { endRow: 7 },
    9: { endRow: 14 },
    16: { endRow: 21 },
    23: { endRow: 25 },
    27: { endRow: 30 }
  };
  
  for (const startRow in sections) {
    const start = parseInt(startRow);
    const endRow = sections[startRow].endRow;
    configSheet.getRange(start, 1, 1, 3).merge()
      .setBackground(s.SECTION_BG)
      .setFontColor(s.FONT_LIGHT)
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('left');
    configSheet.setRowHeight(start, 30);
    
    for (let r = start + 1; r <= endRow; r++) {
      configSheet.getRange(r, 2).setBackground(s.INPUT_BG).setHorizontalAlignment('center');
      configSheet.getRange(r, 1).setFontWeight('500');
      configSheet.getRange(r, 3).setFontStyle('italic').setFontColor(s.FONT_SUBTLE);
      configSheet.setRowHeight(r, 28);
    }
    
    configSheet.getRange(start, 1, endRow - start + 1, 3)
      .setBorder(true, true, true, true, null, null, s.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }
  
  configSheet.getRange(21, 2).setNumberFormat('@STRING@'); // Linha da Versão
  configSheet.setFrozenRows(1);
  
  // Painel de pré-visualização
  configSheet.getRange('E1:M1').merge()
    .setValue('PAINEL UNIFICADO DE AGRUPAMENTO E PRÉ-VISUALIZAÇÃO')
    .setBackground(s.HEADER_BG)
    .setFontColor(s.FONT_LIGHT)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontFamily(s.FONT_FAMILY);
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

  // Dropdown Aba Origem
  const sourceSheets = ss.getSheets()
    .map(s => s.getName())
    .filter(n => n !== CONFIG.SHEETS.CONFIG);
  if (sourceSheets.length > 0) {
    configSheet.getRange('B4').setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(sourceSheets, true)
        .setAllowInvalid(false)
        .build()
    );
  }

  const sourceSheet = ss.getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  if (!sourceSheet || sourceSheet.getLastColumn() === 0) return;
  
  // Dropdowns de colunas (com todas as opções)
  const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const columnOptions = headers.map((h, i) => 
    `${String.fromCharCode(65 + i)} - ${h || `Coluna ${String.fromCharCode(65 + i)}`}`
  );
  const columnRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(columnOptions, true)
    .setAllowInvalid(false)
    .build();
    
  // Linhas para aplicar a regra de coluna
  [5, 6, 7, 10, 11, 12, 13, 14].forEach(row => 
    configSheet.getRange(row, 2).setDataValidation(columnRule)
  );

  // Dropdown de classificação avançada
  configSheet.getRange(24, 2).setDataValidation(columnRule); // Linha 24 é 'CLASSIFICAR POR'

  // Dropdown de Ordem de Classificação
  const sortOrderOptions = ['Ascendente (A-Z, 0-9)', 'Descendente (Z-A, 9-0)'];
  configSheet.getRange(25, 2).setDataValidation( // Linha 25 é 'ORDEM'
    SpreadsheetApp.newDataValidation()
      .requireValueInList(sortOrderOptions, true)
      .setAllowInvalid(false)
      .build()
  );
}

// ============================================================================
// PAINÉIS DE AGRUPAMENTO
// ============================================================================

function updateGroupingPanel() {
  const config = ConfigService.getAll();
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  
  for (let level = 1; level <= 3; level++) {
    updateSingleLevelPanel(level, config, sourceSheet);
  }
}

function updateSingleLevelPanel(level, config, sourceSheet) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CONFIG);
  const startCol = 5 + (level - 1) * 2;
  const levelId = `NÍVEL ${level}`;
  const groupConfig = config[`Agrupar por Nível ${level}`];

  configSheet.getRange(3, startCol, configSheet.getMaxRows() - 2, 2).clear();
  configSheet.setColumnWidth(startCol, 150).setColumnWidth(startCol + 1, 50);
  
  const headerRange = configSheet.getRange(3, startCol, 1, 2).merge();
  if (sourceSheet && groupConfig && groupConfig.trim() !== '') {
    const colIndex = Utils.getColumnIndex(groupConfig);
    if (colIndex === -1) {
      headerRange.setValue(`${levelId}: ERRO`)
        .setBackground(CONFIG.COLORS.PANEL_ERROR_BG)
        .setFontColor(CONFIG.COLORS.FONT_LIGHT)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      configSheet.getRange(4, startCol)
        .setValue('Config. inválida.')
        .setFontColor(CONFIG.COLORS.FONT_SUBTLE)
        .setFontStyle('italic');
    } else {
      const colHeader = Utils.getColumnHeader(groupConfig);
      headerRange.setValue(`${levelId}: ${colHeader.toUpperCase()}`)
        .setBackground(CONFIG.COLORS.SECTION_BG)
        .setFontColor(CONFIG.COLORS.FONT_LIGHT)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      const uniqueValues = getUniqueColumnValues(sourceSheet, colIndex);
      
      if (uniqueValues.length > 0) {
        const valuesFormatted = uniqueValues.map(v => [v ?? '']);
        configSheet.getRange(4, startCol, valuesFormatted.length, 1).setValues(valuesFormatted);
        configSheet.getRange(4, startCol + 1, valuesFormatted.length, 1).insertCheckboxes();
      }
    }
  } else {
    headerRange.setValue(levelId)
      .setBackground(CONFIG.COLORS.PANEL_EMPTY_BG)
      .setFontColor(CONFIG.COLORS.FONT_LIGHT)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    const message = sourceSheet ? '(Não configurado)' : '(Selecione Aba Origem)';
    configSheet.getRange(4, startCol)
      .setValue(message)
      .setFontColor(CONFIG.COLORS.FONT_SUBTLE)
      .setFontStyle('italic');
  }

  const lastRowInPanel = configSheet.getRange(configSheet.getMaxRows(), startCol)
    .getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const effectiveLastRow = Math.max(3, lastRowInPanel);
  configSheet.getRange(3, startCol, effectiveLastRow - 2, 2)
    .setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
}

function getUniqueColumnValues(sheet, columnIndex) {
  const cacheKey = `unique_${sheet.getSheetId()}_${columnIndex}`;
  const cached = CacheManager.get(cacheKey);
  if (cached) return cached;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const values = sheet.getRange(2, columnIndex, lastRow - 1).getValues().flat();
  const uniqueValues = [...new Set(values.filter(v => v))]
    .sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
  
  CacheManager.put(cacheKey, uniqueValues);
  return uniqueValues;
}

function getPanelSelections(sheet) {
  const selections = {};
  for (let i = 0; i < 3; i++) {
    const startCol = 5 + (i * 2);
    const headerRange = sheet.getRange(3, startCol);
    if (headerRange.isBlank()) continue;

    const levelId = headerRange.getValue().split(':')[0].trim();
    const lastDataRow = sheet.getRange(sheet.getMaxRows(), startCol)
      .getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    
    if (lastDataRow >= 4) {
      const numItems = lastDataRow - 3;
      const values = sheet.getRange(4, startCol, numItems, 1).getValues().flat();
      const checkboxes = sheet.getRange(4, startCol + 1, numItems, 1).getValues().flat();
      const selectedValues = values
        .map((v, index) => ({ value: v, checked: checkboxes[index] }))
        .filter(item => item.checked === true)
        .map(item => item.value ?? '');
      selections[levelId] = new Set(selectedValues);
    } else {
      selections[levelId] = new Set();
    }
  }
  
  return selections;
}

// ============================================================================
// PRÉ-VISUALIZAÇÃO
// ============================================================================

function updatePreviewPanel(previewStartCol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  
  const combinations = getSelectedCombinations();
  const existingSuffixes = new Map();
  const lastPreviewRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol)
    .getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  
  if (lastPreviewRow >= 4) {
    configSheet.getRange(4, previewStartCol, lastPreviewRow - 3, 3).getValues()
      .forEach(row => {
        if (row[0]) existingSuffixes.set(row[0], row[2]);
      });
  }
  
  configSheet.getRange(3, previewStartCol, configSheet.getMaxRows() - 2, 4).clear();
  configSheet.getRange(3, previewStartCol, 1, 3)
    .setValues([['COMBINAÇÃO GERADA', 'CRIAR?', 'SUFIXO KOJO FINAL']])
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.FONT_LIGHT)
    .setFontWeight('bold');
  
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
  
  configSheet.setColumnWidth(previewStartCol, 250)
    .setColumnWidth(previewStartCol + 1, 80)
    .setColumnWidth(previewStartCol + 2, 250);
  configSheet.hideColumns(previewStartCol + 3);

  const lastCombinationRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol)
    .getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const effectiveLastRow = Math.max(3, lastCombinationRow);
  configSheet.getRange(3, previewStartCol, effectiveLastRow - 2, 3)
    .setBorder(true, true, true, true, true, true, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  ss.toast('Pré-visualização atualizada!', 'Sucesso', 3);
}

function getSelectedCombinations() {
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
  const activeSelectionLevels = Object.keys(panelSelections)
    .filter(level => panelSelections[level].size > 0);
  
  if (activeSelectionLevels.length === 0) return [];
  
  const allData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
  const existingCombinations = new Set();
  
  allData.forEach(row => {
    const combinationParts = groupIndices.map(index => {
      const value = row[index - 1];
      return value ?? '';
    });
    
    if (combinationParts.every(part => String(part).trim() !== '')) {
      existingCombinations.add(combinationParts.join(CONFIG.DELIMITER));
    }
  });
  
  const finalCombinations = [...existingCombinations].filter(combo => {
    const parts = combo.split(CONFIG.DELIMITER);
    return parts.every((part, i) => {
      const levelId = `NÍVEL ${i + 1}`;
      if (!panelSelections[levelId]) return false;
      
      return Array.from(panelSelections[levelId])
        .some(selected => String(selected).trim() === String(part).trim());
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
// PROCESSAMENTO DE RELATÓRIOS (BOM)
// ============================================================================

function runProcessing() {
  CacheManager.invalidateAll();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ConfigService.getAll();
  const configSheet = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  
  const previewStartCol = 11;
  const lastPreviewRow = configSheet.getRange(configSheet.getMaxRows(), previewStartCol)
    .getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  
  if (lastPreviewRow < 4) {
    return { success: false, message: 'Nenhuma combinação na pré-visualização' };
  }
  
  const previewData = configSheet.getRange(4, previewStartCol, lastPreviewRow - 3, 4).getValues();
  const combinationsToProcess = previewData
    .filter(row => row[1] === true)
    .map(row => ({
      displayCombination: String(row[0]),
      combination: String(row[3]),
      kojoSuffix: row[2]
    }));
  
  if (combinationsToProcess.length === 0) {
    return { success: false, message: 'Nenhum relatório selecionado' };
  }
  
  const sourceSheet = ss.getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  if (!sourceSheet) {
    return { success: false, message: 'Aba de origem não encontrada' };
  }

  const groupConfigs = [
    config[CONFIG.KEYS.GROUP_L1],
    config[CONFIG.KEYS.GROUP_L2],
    config[CONFIG.KEYS.GROUP_L3]
  ].filter(Boolean);
  const groupIndices = groupConfigs.map(Utils.getColumnIndex);
  
  const bomCols = {
    c1: Utils.getColumnIndex(config[CONFIG.KEYS.COL_1]),
    c2: Utils.getColumnIndex(config[CONFIG.KEYS.COL_2]),
    c3: Utils.getColumnIndex(config[CONFIG.KEYS.COL_3]),
    c4: Utils.getColumnIndex(config[CONFIG.KEYS.COL_4]),
    c5: Utils.getColumnIndex(config[CONFIG.KEYS.COL_5])
  };

  // Lógica de Classificação Avançada
  const sortColumnConfig = config[CONFIG.KEYS.SORT_BY] || config[CONFIG.KEYS.COL_2];
  const sortColumnIndex = Utils.getColumnIndex(sortColumnConfig) - 1; // Índice 0-based
  const sortOrder = config[CONFIG.KEYS.SORT_ORDER] === 'Descendente (Z-A, 9-0)' ? 'desc' : 'asc';
  
  if(sortColumnIndex < 0) {
    return { success: false, message: `Coluna de classificação "${sortColumnConfig}" inválida.` };
  }

  const dataMap = new Map();
  combinationsToProcess.forEach(item => dataMap.set(item.combination, []));
  const allData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
  
  for (const row of allData) {
    const rowCombination = groupIndices
      .map(index => String(row[index - 1] ?? '').trim())
      .join(CONFIG.DELIMITER);
    
    if (dataMap.has(rowCombination)) {
      // Adiciona o valor de classificação como um 6º elemento "oculto"
      dataMap.get(rowCombination).push([
        row[bomCols.c1 - 1],
        row[bomCols.c2 - 1],
        row[bomCols.c3 - 1],
        row[bomCols.c4 - 1],
        parseFloat(row[bomCols.c5 - 1]) || 0,
        row[sortColumnIndex] // Valor para classificação
      ]);
    }
  }
  
  let createdCount = 0;
  combinationsToProcess.forEach(item => {
    const { displayCombination, combination, kojoSuffix } = item;
    const rawData = dataMap.get(combination);
    if (!rawData || rawData.length === 0) return;
    
    const processedData = groupAndSumData(rawData, sortOrder);
    
    const sanitizedName = Utils.sanitizeSheetName(displayCombination);
    
    let targetSheet = ss.getSheetByName(sanitizedName);
    if (targetSheet) {
      targetSheet.clear();
    } else {
      targetSheet = ss.insertSheet(sanitizedName);
    }
    
    createAndFormatReport(targetSheet, combination, kojoSuffix, processedData);
    createdCount++;
  });
  
  return { success: true, created: createdCount };
}

/**
 * Agrupa, classifica pelo 6º elemento (oculto) e depois o remove.
 */
function groupAndSumData(data, sortOrder) {
  const grouped = {};
  
  data.forEach(row => {
    const key = `${row[0]}|${row[1]}|${row[2]}|${row[3]}`; 
    if (grouped[key]) {
      grouped[key][4] += row[4]; // Soma a quantidade
    } else {
      grouped[key] = [...row]; // Guarda a linha inteira (6 elementos)
    }
  });
  
  const groupedData = Object.values(grouped);
  const direction = sortOrder === 'asc' ? 1 : -1;

  // Ordena com base no 6º elemento (índice 5)
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
  
  // Remove o 6º elemento antes de retornar
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
  
  const lastUpdate = Utilities.formatDate(
    new Date(),
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    'MM/dd/yyyy'
  );
  const bomKojoComplete = `${reportConfig.kojoPrefix}.${kojoSuffix}`;
  
  const headerValues = [
    ['PROJECT:', reportConfig.project],
    ['BOM:', reportConfig.bom],
    ['BOM KOJO:', bomKojoComplete],
    ['ENG.:', reportConfig.engineer],
    ['VERSION:', reportConfig.version],
    ['LAST UPDATE:', lastUpdate]
  ];
  
  sheet.getRange(1, 1, headerValues.length, 2).setValues(headerValues);
  sheet.getRange(1, 2, headerValues.length, 1).setNumberFormat('@STRING@');
  
  for (let r = 1; r <= headerValues.length; r++) {
    sheet.getRange(r, 2, 1, 4).merge();
  }
  
  sheet.getRange(1, 1, headerValues.length, 5)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

  const dataStartRow = headerValues.length + 2;
  const finalData = [[headers.h1, headers.h2, headers.h3, headers.h4, headers.h5]].concat(data);
  
  sheet.getRange(dataStartRow, 1, finalData.length, 5).setValues(finalData);
  sheet.getRange(dataStartRow, 1, finalData.length, 5)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  sheet.getRange(dataStartRow, 1, 1, 5).setFontWeight('bold');
  
  sheet.setColumnWidth(1, 105)
    .setColumnWidth(2, 570)
    .setColumnWidth(3, 105)
    .setColumnWidth(4, 105)
    .setColumnWidth(5, 105);
  
  try {
    const protection = sheet.getRange(1, 1, dataStartRow - 1, 5).protect();
    protection.setDescription('Cabeçalho protegido');
    protection.removeEditors(protection.getEditors());
  } catch (e) {
    Logger.log(`Aviso: ${e.message}`);
  }
}

function clearOldReports() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Confirmação',
    'Apagar TODAS as abas de relatório?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ConfigService.getAll();
  const protectedSheets = [CONFIG.SHEETS.CONFIG, config[CONFIG.KEYS.SOURCE_SHEET]].filter(Boolean);
  
  let deletedCount = 0;
  ss.getSheets().forEach(sheet => {
    if (!protectedSheets.includes(sheet.getName())) {
      ss.deleteSheet(sheet);
      deletedCount++;
    }
  });
  ui.alert('Limpeza Concluída', `${deletedCount} abas removidas.`, ui.ButtonSet.OK);
}

// ============================================================================
// FIXADORES
// ============================================================================

/**
 * Abre um pop-up flutuante (não-modal)
 */
function abrirSeletorFixadores() {
  const html = HtmlService.createHtmlOutputFromFile('FixadoresSidebar')
    .setTitle('Seletor de Fixadores')
    .setWidth(900)
    .setHeight(800);
  
  SpreadsheetApp.getUi().showModelessDialog(html, 'Seletor de Fixadores');
}

/**
 * Deteta tubos processados (propriedade 'jaTemFixador')
 */
function getPipesElegiveis() {
  const config = ConfigService.getAll();
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  if (!sourceSheet) return [];
  
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) return [];
  
  const data = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();
  const pipes = [];
  
  data.forEach((row, idx) => {
    // Mapeamento (0-indexed): B=1, J=9, K=10, L=11, G=6
    const section = String(row[1] || ''); // Col B
    const desc = String(row[9] || '');    // Col J
    const qty = parseFloat(row[10]) || 0; // Col K
    const uom = String(row[11] || '');    // Col L
    
    if (validarTipoFixacao(section) && desc.toUpperCase().includes('PIPE') && qty > 0) {
      const diameter = Utils.extractDiameter(desc);
      const isRiser = section.toUpperCase().includes('RISER');
      const fixConfig = isRiser ? CONFIG.FIXADORES.RISER : CONFIG.FIXADORES.LOOP;
      const itemMap = isRiser ? fixConfig.clamps : fixConfig.hangs;
      
      if (diameter && itemMap[diameter]) {
        
        // --- Detecção de Fixador ---
        let jaTemFixador = false;
        if (idx + 1 < data.length) { 
          const nextRow = data[idx + 1];
          const nextRowTrade = String(nextRow[6] || '').toUpperCase(); // Col G (TRADE)
          if (nextRowTrade === 'FIX') {
            jaTemFixador = true;
          }
        }
        // --- Fim da Detecção ---

        pipes.push({
          rowIndex: idx + 2,
          section: row[1],  // B
          unitType: row[4], // E
          local: row[5],    // F
          trade: row[6],    // G
          phase: row[7],    // H
          floor: row[8],    // I
          unitid: row[3],   // D
          desc: desc,       // J
          qty: qty,         // K
          uom: uom,         // L
          diameter: diameter,
          isRiser: isRiser,
          originalRow: [...row],
          jaTemFixador: jaTemFixador // <-- Propriedade essencial
        });
      }
    }
  });
  
  return pipes;
}

function validarTipoFixacao(section) {
  const s = String(section).toUpperCase();
  return s.includes('RISER') || s.includes('COLGANTE');
}

function processarFixadoresSelecionados(selectedPipes) {
  const config = ConfigService.getAll();
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  
  if (!sourceSheet || !selectedPipes || selectedPipes.length === 0) {
    return { success: false, message: 'Dados inválidos' };
  }
  
  selectedPipes.sort((a, b) => b.rowIndex - a.rowIndex);
  
  let totalAdded = 0;
  const maxCol = sourceSheet.getLastColumn();
  
  selectedPipes.forEach(pipe => {
    const fixConfig = pipe.isRiser ? CONFIG.FIXADORES.RISER : CONFIG.FIXADORES.LOOP;
    const itemMap = pipe.isRiser ? fixConfig.clamps : fixConfig.hangs;
    const fixadorItem = itemMap[pipe.diameter];
    
    if (!fixadorItem) return;
    
    const formulaStartCol = 12; // Coluna L
    const numFormulaCols = 5; // L, M, N, O, P
    
    const formulasPipe = sourceSheet.getRange(pipe.rowIndex, formulaStartCol, 1, numFormulaCols).getFormulasR1C1()[0];
    const linhasParaInserir = [];
    const insertRow = pipe.rowIndex + 1;
    
    // Linha do fixador
    const linhaFixador = new Array(maxCol).fill('');
    linhaFixador[0] = pipe.originalRow[0]; // A
    linhaFixador[1] = pipe.originalRow[1]; // B
    linhaFixador[2] = pipe.originalRow[2]; // C
    linhaFixador[3] = pipe.originalRow[3]; // D
    linhaFixador[4] = pipe.originalRow[4]; // E
    linhaFixador[5] = pipe.originalRow[5]; // F
    linhaFixador[6] = 'FIX';               // G (Trade)
    linhaFixador[7] = pipe.originalRow[7]; // H
    linhaFixador[8] = pipe.originalRow[8]; // I
    linhaFixador[9] = fixadorItem;         // J (Desc)
    linhaFixador[10] = `=ROUNDUP(R[-1]C/${fixConfig.interval})`; // K (Qty)
    
    formulasPipe.forEach((formula, idx) => {
        const targetColIndex = (formulaStartCol - 1) + idx;
        if (targetColIndex < maxCol) {
            linhaFixador[targetColIndex] = formula || pipe.originalRow[targetColIndex];
        }
    });
    
    linhasParaInserir.push(linhaFixador);
    const fixadorRow = insertRow;
    
    // Materiais
    fixConfig.materials.forEach(mat => {
      const linhaMat = new Array(maxCol).fill('');
      linhaMat[0] = pipe.originalRow[0];
      linhaMat[1] = pipe.originalRow[1];
      linhaMat[2] = pipe.originalRow[2];
      linhaMat[3] = pipe.originalRow[3];
      linhaMat[4] = pipe.originalRow[4];
      linhaMat[5] = pipe.originalRow[5];
      linhaMat[6] = 'FIX';
      linhaMat[7] = pipe.originalRow[7];
      linhaMat[8] = pipe.originalRow[8];
      linhaMat[9] = mat.desc;
      linhaMat[10] = `=R${fixadorRow}C*${mat.factor}`;
      
      formulasPipe.forEach((formula, idx) => {
        const targetColIndex = (formulaStartCol - 1) + idx;
        if (targetColIndex < maxCol) {
            linhaMat[targetColIndex] = formula || pipe.originalRow[targetColIndex];
        }
      });
      
      linhasParaInserir.push(linhaMat);
    });
    
    sourceSheet.insertRowsAfter(pipe.rowIndex, linhasParaInserir.length);
    
    const formatoOrigem = sourceSheet.getRange(pipe.rowIndex, 1, 1, maxCol);
    formatoOrigem.copyFormatToRange(sourceSheet, 1, maxCol, insertRow, insertRow + linhasParaInserir.length - 1);
    
    const rangeDestino = sourceSheet.getRange(insertRow, 1, linhasParaInserir.length, maxCol);
    rangeDestino.setValues(linhasParaInserir);
    
    // Reaplica as fórmulas R1C1
    linhasParaInserir.forEach((row, idx) => {
      const currentRow = insertRow + idx;
      
      // Fórmula da Coluna K (Qty)
      const formulaK = row[10];
      if (typeof formulaK === 'string' && formulaK.startsWith('=')) {
        sourceSheet.getRange(currentRow, 11).setFormulaR1C1(formulaK);
      }
      
      // Fórmulas das colunas L em diante
      for (let col = formulaStartCol - 1; col < formulaStartCol - 1 + numFormulaCols; col++) {
        const formula = row[col];
         if (typeof formula === 'string' && formula.startsWith('=')) {
          sourceSheet.getRange(currentRow, col + 1).setFormulaR1C1(formula);
        }
      }
    });
    totalAdded += linhasParaInserir.length;
  });
  
  return { success: true, added: totalAdded };
}

/**
 * *** NOVA FUNÇÃO (V 2.5) ***
 * Remove as linhas de fixadores associadas a um tubo.
 */
function removerFixadoresSelecionados(selectedPipes) {
  const config = ConfigService.getAll();
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
  
  if (!sourceSheet || !selectedPipes || selectedPipes.length === 0) {
    return { success: false, message: 'Dados inválidos' };
  }

  // CRÍTICO: Ordenar por rowIndex DECRESCENTE para apagar de baixo para cima
  selectedPipes.sort((a, b) => b.rowIndex - a.rowIndex);
  
  let totalRemoved = 0;
  const allData = sourceSheet.getDataRange().getValues(); // Ler a folha toda 1 vez
  const tradeColIndex = 6; // Coluna G é o índice 6 (base 0)

  try {
    selectedPipes.forEach(pipe => {
      const rowIndex = pipe.rowIndex; // Esta é a linha do TUBO (base 1)
      let rowsToDelete = 0;
      
      // Começa a verificar a partir da linha *abaixo* do tubo
      // allData é base 0, então a linha *abaixo* (rowIndex + 1) está em allData[rowIndex]
      for (let i = rowIndex; i < allData.length; i++) {
        const rowData = allData[i];
        const trade = String(rowData[tradeColIndex] || '').toUpperCase();
        
        if (trade === 'FIX') {
          rowsToDelete++;
        } else {
          // Encontrou uma linha que não é 'FIX', para de contar
          break;
        }
      }
      
      if (rowsToDelete > 0) {
        // Apaga as linhas. A primeira linha a apagar é rowIndex + 1
        sourceSheet.deleteRows(rowIndex + 1, rowsToDelete);
        totalRemoved += rowsToDelete;
      }
    });
    
    return { success: true, removed: totalRemoved };

  } catch (e) {
    Logger.log(`Erro ao remover fixadores: ${e.message}`);
    return { success: false, message: `Erro ao apagar linhas: ${e.message}` };
  }
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
    
    // Formata a versão
    if (col === 2 && row === 21) { // Linha da Versão
      const formatted = Utils.formatVersion(range.getValue());
      if (formatted !== String(range.getValue())) {
        range.setValue(formatted);
      }
    }
    
    // Atualiza painéis de agrupamento
    if (col === 2 && row >= 4 && row <= 7) {
      Utilities.sleep(200);
      const config = ConfigService.getAll();
      const sourceSheet = SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
      
      if (row === 4) { // Mudou Aba Origem
        updateConfigDropdowns();
        updateGroupingPanel();
      } else { // Mudou Nível 1, 2 ou 3
        updateSingleLevelPanel(row - 4, config, sourceSheet);
      }
      
      updatePreviewPanel(11);
    }
    
    // Atualiza pré-visualização se um checkbox do painel for marcado
    if ((col === 6 || col === 8 || col === 10) && row >= 4) {
      updatePreviewPanel(11);
    }
  } else {
    // Se editar a aba de origem, limpa o cache
    try {
      const config = ConfigService.getAll();
      if (sheetName === config[CONFIG.KEYS.SOURCE_SHEET]) {
        CacheManager.invalidateAll();
      }
    } catch (error) {
      // Ignora erro
    }
  }
}

function openConfigSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ConfigSidebar')
    .setTitle('Painel de Controle')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Funções de Feedback (Wrapper)
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
  SpreadsheetApp.getActiveSpreadsheet().toast('Cache limpo e painéis atualizados!', 'Sucesso', 3);
}

function testSystem() {
  try {
    const config = ConfigService.getAll();
    const sourceSheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(config[CONFIG.KEYS.SOURCE_SHEET]);
    
    const msg = [
      `Versão do Script: 2.5 (Remoção Fixadores)`,
      `Aba de Origem: ${config[CONFIG.KEYS.SOURCE_SHEET] || 'Não configurada'}`,
      `Linhas na Origem: ${sourceSheet ? sourceSheet.getLastRow() - 1 : 0}`,
      `Relatórios Gerados: ${getReportSheetNames().length}`,
      `Coluna de Classificação: ${config[CONFIG.KEYS.SORT_BY] || 'Não configurada'}`
    ].join('\n');
    
    SpreadsheetApp.getUi().alert('🧪 Diagnóstico do Sistema', msg, SpreadsheetApp.getUi().ButtonSet.OK);
    
    return { 
      "Modo de Filtro": "Classificação Avançada", 
      "Tipologias Selecionadas": sourceSheet ? sourceSheet.getLastRow() - 1 : 0 
    };

  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro no Diagnóstico', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
