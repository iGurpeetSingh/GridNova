const CUSTOM_DICTIONARY_KEY = 'gridnova:custom-dictionary';
const REVIEW_OBJECT_ALT_KEY = (workbookId) => 'gridnova:review:alt-text:' + workbookId;
const SPELLING_IGNORED_KEY = (workbookId) => 'gridnova:review:ignored-words:' + workbookId;

export const COMMON_MISSPELLINGS = {
  accomodate: 'accommodate',
  acheive: 'achieve',
  adress: 'address',
  aggresive: 'aggressive',
  alchohol: 'alcohol',
  alot: 'a lot',
  amature: 'amateur',
  apparant: 'apparent',
  arguement: 'argument',
  beggining: 'beginning',
  beleive: 'believe',
  buisness: 'business',
  calender: 'calendar',
  catagory: 'category',
  comming: 'coming',
  definately: 'definitely',
  dependance: 'dependence',
  dimention: 'dimension',
  embarass: 'embarrass',
  enviroment: 'environment',
  existance: 'existence',
  familar: 'familiar',
  finaly: 'finally',
  foriegn: 'foreign',
  freind: 'friend',
  goverment: 'government',
  grammer: 'grammar',
  happend: 'happened',
  independant: 'independent',
  inteligence: 'intelligence',
  knowlege: 'knowledge',
  lenght: 'length',
  liason: 'liaison',
  maintenence: 'maintenance',
  neccessary: 'necessary',
  occassion: 'occasion',
  occured: 'occurred',
  oportunity: 'opportunity',
  persue: 'pursue',
  posession: 'possession',
  prefered: 'preferred',
  priviledge: 'privilege',
  pronounciation: 'pronunciation',
  recieve: 'receive',
  recomend: 'recommend',
  refered: 'referred',
  relevent: 'relevant',
  seperate: 'separate',
  succesful: 'successful',
  suprise: 'surprise',
  thier: 'their',
  untill: 'until',
  wierd: 'weird',
  writting: 'writing',
};

const BASE_DICTIONARY = new Set((
  'a about above across action active activity add address after again against all allow also always amount an and answer any app apple application are area around as ask at audit auto automatic average back bar base be because before below best between big blue body book both box branch break build business button by calculate calculation call can cancel change chart check cell cells center change changes choose clear close code color column comment comments common complete compute condition connect content copy count create current data date day decimal default delete depend dependency description design detail different display document done down draft draw each edit editor else empty end engine enter equal error event every example export expression extension external extra face field file filter find first fix fixed follow for formula formulas format from function future general get give go good grand green group guide header help here hide highlight history home horizontal how i icon if image import improve in include index input insert inside into issue it item its join jump just key keyboard label last layout left less let level library light like list live load local lock locked long look lookup lower main manual manage many map mark match material max mean measure merge message min minimum misspell mobile model more most move multiple must name named new next no note notes not now number object of off on one only open or order other our out output over own page pane parameter part party pass password paste path percentage people percent permission phone pie pin place plan point portal power prefer previous print process product project protect protection proof read real receive record redo reference refresh remove rename report resolve result review right role row rule run same scale scan schedule screen search section select selection separate set sheet sheets show side size sort source spell spelling standard start state status step stop style suggestion sum summary support switch table tab target text than that the them then there these they this those three through time title to today together top total trace true type under undo up update use user value verify version view visibility visible voice web week what where which while width with without work workbook write year yes zero zero-based'.split(/\s+/)
));

function safeRead(key, fallback) {
  try {
    const value = JSON.parse(window.localStorage.getItem(key) || 'null');
    return value == null ? fallback : value;
  } catch {
    return fallback;
  }
}

function safeWrite(key, value) {
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch {
    return false;
  }
}

export function getCustomDictionary() {
  return safeRead(CUSTOM_DICTIONARY_KEY, []);
}

export function addCustomDictionaryWord(word) {
  const normalized = String(word || '').trim().toLocaleLowerCase();
  if (!normalized) return getCustomDictionary();
  const next = [...new Set([...getCustomDictionary(), normalized])];
  safeWrite(CUSTOM_DICTIONARY_KEY, next);
  return next;
}

export function getIgnoredWords(workbookId) {
  return safeRead(SPELLING_IGNORED_KEY(workbookId), []);
}

export function saveIgnoredWords(workbookId, words) {
  safeWrite(SPELLING_IGNORED_KEY(workbookId), [...new Set(words.map((word) => String(word).toLocaleLowerCase()))]);
}

export function levenshteinDistance(left, right) {
  const a = String(left).toLocaleLowerCase();
  const b = String(right).toLocaleLowerCase();
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  let previous = Array.from({ length: b.length + 1 }, (_, index) => index);
  for (let i = 1; i <= a.length; i++) {
    const current = [i];
    for (let j = 1; j <= b.length; j++) {
      current[j] = Math.min(
        current[j - 1] + 1,
        previous[j] + 1,
        previous[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1),
      );
    }
    previous = current;
  }
  return previous[b.length];
}

export function suggestSpellings(word, limit = 5) {
  const normalized = String(word || '').toLocaleLowerCase();
  if (COMMON_MISSPELLINGS[normalized]) return [COMMON_MISSPELLINGS[normalized]];
  const dictionary = new Set([...BASE_DICTIONARY, ...getCustomDictionary()]);
  return [...dictionary]
    .map((candidate) => ({ candidate, distance: levenshteinDistance(normalized, candidate) }))
    .filter((item) => item.distance > 0 && item.distance <= Math.max(1, Math.floor(normalized.length / 3)))
    .sort((a, b) => a.distance - b.distance || a.candidate.localeCompare(b.candidate))
    .slice(0, limit)
    .map((item) => item.candidate);
}

function getCellValue(cell) {
  if (!cell || typeof cell !== 'object') return '';
  if (cell.v != null) return typeof cell.v === 'string' || typeof cell.v === 'number' ? String(cell.v) : '';
  if (cell.value != null) return String(cell.value);
  return '';
}

function cellIsFormula(cell) {
  return !!(cell && (cell.f || cell.formula || (typeof cell.v === 'string' && cell.v.startsWith('='))));
}

export function columnName(index) {
  let number = Number(index) + 1;
  let result = '';
  while (number > 0) {
    const remainder = (number - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    number = Math.floor((number - 1) / 26);
  }
  return result;
}

export function cellAddress(row, column) {
  return columnName(column) + (Number(row) + 1);
}

export function scanSheetForSpelling(snapshotSheet, activeAddress = 'A1', options = {}) {
  const cellData = snapshotSheet?.cellData || {};
  const customWords = new Set([...(options.customWords || getCustomDictionary())].map((word) => String(word).toLocaleLowerCase()));
  const ignored = new Set((options.ignoredWords || []).map((word) => String(word).toLocaleLowerCase()));
  const words = new Set([...BASE_DICTIONARY, ...customWords]);
  const entries = [];

  Object.entries(cellData).forEach(([rowText, rowCells]) => {
    const row = Number(rowText);
    if (!Number.isFinite(row) || !rowCells || typeof rowCells !== 'object') return;
    Object.entries(rowCells).forEach(([columnText, cell]) => {
      const column = Number(columnText);
      if (!Number.isFinite(column) || !cell || cellIsFormula(cell)) return;
      const value = getCellValue(cell);
      if (!value || !/[A-Za-z]/.test(value)) return;
      const address = cellAddress(row, column);
      const matches = [...value.matchAll(/[A-Za-z][A-Za-z']{1,}/g)];
      matches.forEach((match) => {
        const word = match[0];
        const normalized = word.toLocaleLowerCase();
        if (ignored.has(normalized) || words.has(normalized)) return;
        entries.push({
          address,
          row,
          column,
          value,
          word,
          start: match.index,
          end: match.index + word.length,
          suggestions: suggestSpellings(word),
        });
      });
    });
  });

  const active = String(activeAddress || 'A1').toUpperCase().match(/^([A-Z]+)(\d+)$/);
  let activeRow = 0;
  let activeColumn = 0;
  if (active) {
    for (const char of active[1]) activeColumn = activeColumn * 26 + char.charCodeAt(0) - 64;
    activeColumn -= 1;
    activeRow = Number(active[2]) - 1;
  }
  const activeIndex = entries.findIndex((item) =>
    item.row > activeRow || (item.row === activeRow && item.column >= activeColumn),
  );
  if (activeIndex > 0) entries.push(...entries.splice(0, activeIndex));
  return entries;
}

function rgbFromColor(value, fallback = '#000000') {
  if (!value) return fallback;
  if (typeof value === 'string') return value;
  if (typeof value === 'object') {
    if (value.rgb) return value.rgb;
    if (value.color) return value.color;
    if (value.hex) return value.hex;
  }
  return fallback;
}

function styleForCell(snapshot, cell) {
  const styles = snapshot?.styles || {};
  const styleKey = cell?.s ?? cell?.st ?? cell?.style;
  if (typeof styleKey === 'string' || typeof styleKey === 'number') return styles[styleKey] || {};
  return styleKey && typeof styleKey === 'object' ? styleKey : {};
}

function luminance(hex) {
  const value = String(hex || '#000000').replace('#', '');
  if (!/^[0-9a-f]{3}([0-9a-f]{3})?$/i.test(value)) return 0;
  const expanded = value.length === 3 ? value.split('').map((part) => part + part).join('') : value;
  const channels = [0, 2, 4].map((index) => parseInt(expanded.slice(index, index + 2), 16) / 255)
    .map((channel) => channel <= 0.04045 ? channel / 12.92 : Math.pow((channel + 0.055) / 1.055, 2.4));
  return channels[0] * 0.2126 + channels[1] * 0.7152 + channels[2] * 0.0722;
}

export function contrastRatio(foreground, background) {
  const first = luminance(foreground);
  const second = luminance(background);
  return (Math.max(first, second) + 0.05) / (Math.min(first, second) + 0.05);
}

export function readAltText(workbookId) {
  return safeRead(REVIEW_OBJECT_ALT_KEY(workbookId), {});
}

export function writeAltText(workbookId, objectId, text) {
  const current = readAltText(workbookId);
  const next = { ...current, [objectId]: String(text || '').trim() };
  safeWrite(REVIEW_OBJECT_ALT_KEY(workbookId), next);
  return next;
}

export function inspectWorkbookAccessibility(workbookFacade, snapshot, workbookId) {
  const issues = [];
  const altText = readAltText(workbookId);
  const sheets = workbookFacade?.getSheets?.() || [];
  const sheetSnapshots = snapshot?.sheets || {};

  sheets.forEach((sheet) => {
    const sheetId = sheet.getSheetId?.();
    const savedSheet = sheetSnapshots[sheetId] || {};
    const name = sheet.getSheetName?.() || savedSheet.name || 'Sheet';
    if (/^sheet\s*\d*$/i.test(name.trim())) {
      issues.push({
        id: 'sheet-name:' + sheetId,
        category: 'Default Sheet Names',
        title: 'Rename ' + name,
        description: 'Use a descriptive worksheet name.',
        sheetId,
        action: 'rename-sheet',
      });
    }

    (sheet.getMergedRanges?.() || sheet.getMergeData?.() || []).forEach((range) => {
      const address = range.getA1Notation?.() || '';
      issues.push({
        id: 'merge:' + sheetId + ':' + address,
        category: 'Merged Cells',
        title: 'Merged range ' + address,
        description: 'Merged cells can make navigation and screen-reader reading order harder.',
        sheetId,
        address,
        action: 'unmerge',
      });
    });

    const cellData = savedSheet.cellData || {};
    Object.entries(cellData).forEach(([rowText, rowCells]) => {
      Object.entries(rowCells || {}).forEach(([columnText, cell]) => {
        const style = styleForCell(snapshot, cell);
        const foreground = rgbFromColor(style?.cl?.rgb || style?.cl || style?.fontColor, '#202020');
        const background = rgbFromColor(style?.bg?.rgb || style?.bg || style?.backgroundColor, '#FFFFFF');
        if (!cell || (!style?.cl && !style?.fontColor) || (!style?.bg && !style?.backgroundColor)) return;
        const ratio = contrastRatio(foreground, background);
        if (ratio < 4.5) {
          const address = cellAddress(Number(rowText), Number(columnText));
          issues.push({
            id: 'contrast:' + sheetId + ':' + address,
            category: 'Hard-to-Read Text Contrast',
            title: 'Low contrast in ' + address,
            description: 'Text/background contrast is ' + ratio.toFixed(2) + ':1; WCAG AA text requires 4.5:1.',
            sheetId,
            address,
            ratio,
            action: 'fix-contrast',
          });
        }
      });
    });

    const objects = [
      ...(sheet.getCharts?.() || []).map((chart, index) => ({
        id: chart.getId?.() || chart.id || 'chart-' + index,
        label: chart.getTitle?.() || chart.title || 'Chart ' + (index + 1),
        type: 'Chart',
      })),
      ...(sheet.getAllFloatDoms?.() || []).map((object, index) => ({
        id: object.id || object.floatDomId || 'object-' + index,
        label: object.title || object.type || 'Drawing object ' + (index + 1),
        type: object.type || 'Drawing object',
      })),
    ];
    objects.forEach((object) => {
      if (!altText[object.id]) {
        issues.push({
          id: 'alt:' + sheetId + ':' + object.id,
          category: 'Missing Alt Text',
          title: object.label + ' has no Alt Text',
          description: 'Add a concise description so this object is accessible.',
          sheetId,
          objectId: object.id,
          action: 'alt-text',
        });
      }
    });
  });

  return issues;
}

function cellHasData(cell) {
  if (!cell || typeof cell !== 'object') return false;
  return cell.v !== undefined && cell.v !== null && String(cell.v) !== '' ||
    !!cell.f || !!cell.p || !!cell.formula;
}

export function calculateWorkbookStatistics(workbookFacade, snapshot) {
  const sheets = workbookFacade?.getSheets?.() || [];
  const active = workbookFacade?.getActiveSheet?.();
  const sheetSnapshots = snapshot?.sheets || {};
  const summarize = (sheet, isCurrent) => {
    const sheetId = sheet?.getSheetId?.();
    const saved = sheetSnapshots[sheetId] || {};
    const cellData = saved.cellData || {};
    let dataCells = 0;
    let formulas = 0;
    let endRow = -1;
    let endColumn = -1;
    Object.entries(cellData).forEach(([rowText, rowCells]) => {
      Object.entries(rowCells || {}).forEach(([columnText, cell]) => {
        if (cellHasData(cell)) {
          dataCells++;
          endRow = Math.max(endRow, Number(rowText));
          endColumn = Math.max(endColumn, Number(columnText));
        }
        if (cell?.f || cell?.formula || (typeof cell?.v === 'string' && cell.v.startsWith('='))) formulas++;
      });
    });
    const comments = (sheet?.getComments?.() || []).filter((comment) => comment.getIsRoot?.()).length;
    const notes = (sheet?.getNotes?.() || []).length;
    const tables = (sheet?.getSubTableInfos?.() || []).length;
    const charts = (sheet?.getCharts?.() || []).length;
    const drawings = (sheet?.getAllFloatDoms?.() || []).length;
    return {
      name: sheet?.getSheetName?.() || saved.name || 'Sheet',
      endOfSheet: endRow < 0 ? 'A1' : cellAddress(endRow, endColumn),
      cellsWithData: dataCells,
      formulas,
      comments,
      notes,
      tables,
      chartsAndObjects: charts + drawings,
    };
  };

  const current = active ? summarize(active, true) : summarize(sheets[0], true);
  const all = sheets.map((sheet) => summarize(sheet, false));
  const workbook = {
    totalSheets: sheets.length,
    totalCellsWithData: all.reduce((sum, item) => sum + item.cellsWithData, 0),
    totalFormulas: all.reduce((sum, item) => sum + item.formulas, 0),
    totalCommentsAndNotes: all.reduce((sum, item) => sum + item.comments + item.notes, 0),
    protectedSheets: sheets.filter((sheet) => {
      try { return !!sheet.getWorksheetPermission?.()?.isProtected?.(); } catch { return false; }
    }).length,
  };
  return { current, workbook };
}

export function readChangeHistory(workbookId) {
  return safeRead('gridnova:changes:' + workbookId, []);
}

export function writeChangeHistory(workbookId, entries) {
  safeWrite('gridnova:changes:' + workbookId, entries.slice(0, 500));
}

function snapshotCellValue(cell) {
  if (!cell) return '';
  if (cell.f) return '=' + String(cell.f).replace(/^=/, '');
  if (cell.v !== undefined && cell.v !== null) return String(cell.v);
  if (cell.p) return '[Rich text]';
  return '';
}

export function diffWorkbookSnapshots(previous, next, commandId = 'Workbook change') {
  if (!previous || !next) return [];
  const entries = [];
  const beforeSheets = previous.sheets || {};
  const afterSheets = next.sheets || {};
  const sheetIds = new Set([...Object.keys(beforeSheets), ...Object.keys(afterSheets)]);
  for (const sheetId of sheetIds) {
    const before = beforeSheets[sheetId];
    const after = afterSheets[sheetId];
    if (!before || !after) {
      entries.push({
        id: globalThis.crypto?.randomUUID?.() || String(Date.now() + entries.length),
        sheetId,
        sheetName: after?.name || before?.name || 'Worksheet',
        cell: '',
        action: commandId,
        oldValue: before ? 'Present' : 'Not present',
        newValue: after ? 'Present' : 'Removed',
        description: after ? 'Worksheet created' : 'Worksheet deleted',
        createdAt: Date.now(),
      });
      continue;
    }
    const beforeCells = before.cellData || {};
    const afterCells = after.cellData || {};
    const rowIds = new Set([...Object.keys(beforeCells), ...Object.keys(afterCells)]);
    for (const rowId of rowIds) {
      const oldRow = beforeCells[rowId] || {};
      const newRow = afterCells[rowId] || {};
      const columnIds = new Set([...Object.keys(oldRow), ...Object.keys(newRow)]);
      for (const columnId of columnIds) {
        const oldValue = snapshotCellValue(oldRow[columnId]);
        const newValue = snapshotCellValue(newRow[columnId]);
        if (oldValue === newValue) continue;
        const address = cellAddress(Number(rowId), Number(columnId));
        entries.push({
          id: globalThis.crypto?.randomUUID?.() || String(Date.now() + entries.length),
          sheetId,
          sheetName: after.name || before.name || 'Worksheet',
          cell: address,
          action: commandId,
          oldValue,
          newValue,
          canRevert: true,
          description: oldValue ? 'Changed from ' + oldValue + ' to ' + newValue : 'Entered ' + newValue,
          createdAt: Date.now(),
        });
      }
    }
    if (before.name !== after.name) {
      entries.push({
        id: globalThis.crypto?.randomUUID?.() || String(Date.now() + entries.length),
        sheetId,
        sheetName: after.name || before.name,
        cell: '',
        action: 'Rename sheet',
        oldValue: before.name,
        newValue: after.name,
        description: 'Renamed worksheet from ' + before.name + ' to ' + after.name,
        createdAt: Date.now(),
      });
    }
  }
  return entries.reverse();
}
