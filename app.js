// 엑셀 → 한글 표 변환 PoC
// - SheetJS로 xlsx 파싱
// - rhwp WASM(HwpDocument)으로 표 생성/병합/텍스트 삽입/.hwp 내보내기

import init, { HwpDocument, init_panic_hook } from './lib/rhwp/rhwp.js';

const $log = document.getElementById('log');
const $file = document.getElementById('file');
const $btnConvert = document.getElementById('btnConvert');
const $btnSelftest = document.getElementById('btnSelftest');
const $dlSlot = document.getElementById('dlSlot');
const $results = document.getElementById('results');

function log(msg, cls = '') {
  const line = document.createElement('div');
  if (cls) line.className = cls;
  line.textContent = msg;
  $log.appendChild(line);
  $log.scrollTop = $log.scrollHeight;
  console.log(msg);
}

function downloadBytes(bytes, filename) {
  const blob = new Blob([bytes], { type: 'application/x-hwp' });
  const url = URL.createObjectURL(blob);
  $dlSlot.innerHTML = '';
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.textContent = `↓ ${filename} 다운로드`;
  a.className = 'dl';
  $dlSlot.appendChild(a);
  // 자동 클릭으로 즉시 저장
  a.click();
}

function makeDownloadLink(bytes, filename) {
  const blob = new Blob([bytes], { type: 'application/x-hwp' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.textContent = `↓ ${filename} 다운로드`;
  a.className = 'dl';
  return a;
}

function resetResults() {
  if (!$results) return;
  $results.innerHTML = '';
  $results.hidden = true;
}

// ────────────────────────────────────────────────────────────────
// rhwp WASM 초기화
// ────────────────────────────────────────────────────────────────
async function ensureWasm() {
  if (ensureWasm._ready) return;
  log('WASM 로딩 중...');
  await init();
  try { init_panic_hook(); } catch (_) {}
  log('WASM 로딩 완료', 'ok');
  ensureWasm._ready = true;
}

function parseJsonResult(jsonStr, label) {
  let j;
  try { j = JSON.parse(jsonStr); }
  catch { throw new Error(`${label} 응답 파싱 실패: ${jsonStr}`); }
  if (j.ok === false || j.error) {
    throw new Error(`${label} 실패: ${jsonStr}`);
  }
  return j;
}

function excelColName(idx) {
  let n = idx + 1;
  let name = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    name = String.fromCharCode(65 + r) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function normalizeExcelCellText(value) {
  return String(value ?? '')
    .replace(/_x000D_/gi, '\n')
    .replace(/\r\n|\r/g, '\n')
    .replace(/\n{2,}/g, '\n')
    .trim();
}

function readSheetCellText(sheet, row, col) {
  const addr = XLSX.utils.encode_cell({ r: row, c: col });
  const cell = sheet[addr];
  if (!cell) return '';
  return normalizeExcelCellText(cell.w != null ? cell.w : (cell.v != null ? String(cell.v) : ''));
}

function extractColumnWidthsExcel(sheet, arrayBuffer, startCol, colCount, sheetIndex = 0) {
  const colsInfo = sheet['!cols'] || [];
  const colWidthsExcel = [];
  const colsInfoSamples = [];  // 진단용
  let nonDefaultCount = 0;
  for (let c = 0; c < colCount; c++) {
    const sourceCol = startCol + c;
    const info = colsInfo[sourceCol];
    let w = 10;
    if (info) {
      colsInfoSamples.push(`[${sourceCol}]${JSON.stringify(info)}`);
      if (typeof info.wch === 'number')       { w = info.wch;        nonDefaultCount++; }
      else if (typeof info.width === 'number'){ w = info.width;      nonDefaultCount++; }
      else if (typeof info.wpx === 'number')  { w = info.wpx / 7;    nonDefaultCount++; }
    } else {
      colsInfoSamples.push(`[${sourceCol}]undefined`);
    }
    colWidthsExcel.push(w);
  }

  // 정보가 충분치 않으면 worksheet XML을 우리가 직접 파싱 (확실한 폴백)
  if (nonDefaultCount < colCount) {
    const xmlFallback = parseColWidthsFromXlsxXml(sheet, arrayBuffer, sheetIndex);
    if (xmlFallback && xmlFallback.length > 0) {
      for (let c = 0; c < colCount; c++) {
        const sourceCol = startCol + c;
        if (xmlFallback[sourceCol] != null) colWidthsExcel[c] = xmlFallback[sourceCol];
      }
      nonDefaultCount = colCount;
    }
  }

  console.log('[debug] !cols samples:', colsInfoSamples);
  console.log('[debug] colWidthsExcel:', colWidthsExcel);
  return colWidthsExcel;
}

function isCurriculumOrganizationCandidate(sheetName, sheet) {
  const titleText = readSheetCellText(sheet, 1, 1); // B2
  const compact = normalizeCompactText(`${sheetName} ${titleText}`);
  return compact.includes('교육과정편제표');
}

function findLastDataRowInColumn(sheet, col, firstDataRow, maxRow) {
  for (let row = maxRow; row >= firstDataRow; row--) {
    if (readSheetCellText(sheet, row, col).length > 0) return row;
  }
  return -1;
}

function findLastDataRowInColumns(sheet, startCol, endCol, firstDataRow, maxRow) {
  let last = -1;
  for (let col = startCol; col <= endCol; col++) {
    last = Math.max(last, findLastDataRowInColumn(sheet, col, firstDataRow, maxRow));
  }
  return last;
}

function isCurriculumLandscapeSheetName(sheetName) {
  return normalizeCompactText(sheetName) === '교육과정편제표(가로)';
}

function readCurriculumOrganizationLandscapeTable(wb, arrayBuffer) {
  const tableStartRow = 5; // B6
  const titleRows = [0, 1]; // B1, B2
  const startCol = 1; // B
  const endCol = 29;  // AD
  const colCount = endCol - startCol + 1;

  for (let sheetIndex = 0; sheetIndex < wb.SheetNames.length; sheetIndex++) {
    const sheetName = wb.SheetNames[sheetIndex];
    if (!isCurriculumLandscapeSheetName(sheetName)) continue;

    const sheet = wb.Sheets[sheetName];
    if (!sheet || !sheet['!ref']) continue;

    const range = XLSX.utils.decode_range(sheet['!ref']);
    const lastDataRow = findLastDataRowInColumns(sheet, startCol, endCol, tableStartRow, range.e.r);
    if (lastDataRow < tableStartRow) continue;

    const rawRowCount = range.e.r - range.s.r + 1;
    const rawColCount = range.e.c - range.s.c + 1;
    const titleLines = titleRows
      .map(row => readSheetCellText(sheet, row, startCol))
      .filter(Boolean);
    const titleText = titleLines.join(' ') || '교육과정편제표(가로)';
    const tableDataRowCount = lastDataRow - tableStartRow + 1;
    const cells = [];

    for (let r = tableStartRow; r <= lastDataRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const text = readSheetCellText(sheet, r, c);
        if (text.length === 0) continue;
        cells.push({
          row: r - tableStartRow,
          col: c - startCol,
          text,
        });
      }
    }

    const merges = (sheet['!merges'] || [])
      .filter(m =>
        m.s.r >= tableStartRow && m.e.r <= lastDataRow &&
        m.s.c >= startCol && m.e.c <= endCol
      )
      .map(m => ({
        r1: m.s.r - tableStartRow,
        c1: m.s.c - startCol,
        r2: m.e.r - tableStartRow,
        c2: m.e.c - startCol,
      }));

    const colWidthsExcel = extractColumnWidthsExcel(sheet, arrayBuffer, startCol, colCount, sheetIndex);
    const borderByCell = parseExcelCellBordersFromXlsxXml(arrayBuffer, sheetIndex, {
      startRow: tableStartRow,
      endRow: lastDataRow,
      startCol,
      endCol,
    });
    const excelStyleByCell = parseExcelCellStylesFromXlsxXml(arrayBuffer, sheetIndex, {
      startRow: tableStartRow,
      endRow: lastDataRow,
      startCol,
      endCol,
    });
    const tableRangeLabel = `B6:AD${lastDataRow + 1}`;

    return {
      sheetName,
      format: 'curriculumOrganizationLandscape',
      formatLabel: '교육과정편제표(가로)',
      titleCell: 'B1/B2',
      titleLines,
      titleText,
      tableRangeLabel,
      rowCount: tableDataRowCount,
      colCount,
      rawRowCount,
      rawColCount,
      cells,
      merges,
      colWidthsExcel,
      borderByCell,
      excelStyleByCell,
      droppedColIdx: [],
      selectionRequirement: null,
      trim: { r1: tableStartRow, c1: startCol, r2: lastDataRow, c2: endCol },
      hasTitleRow: false,
      headerRowCount: 1,
      preserveZeroText: true,
      forceLandscape: true,
    };
  }

  return null;
}

function readCurriculumOrganizationTable(wb, arrayBuffer) {
  const tableStartRow = 4; // B5
  const titleRow = 1;      // B2
  const startCol = 1;      // B
  const endCol = 14;       // O
  const colCount = endCol - startCol + 1;

  for (let sheetIndex = 0; sheetIndex < wb.SheetNames.length; sheetIndex++) {
    const sheetName = wb.SheetNames[sheetIndex];
    if (isCurriculumLandscapeSheetName(sheetName)) continue;

    const sheet = wb.Sheets[sheetName];
    if (!sheet || !sheet['!ref']) continue;
    if (!isCurriculumOrganizationCandidate(sheetName, sheet)) continue;

    const range = XLSX.utils.decode_range(sheet['!ref']);
    const lastDataRow = findLastDataRowInColumn(sheet, startCol, tableStartRow, range.e.r);
    if (lastDataRow < tableStartRow) continue;

    const rawRowCount = range.e.r - range.s.r + 1;
    const rawColCount = range.e.c - range.s.c + 1;
    const titleText = readSheetCellText(sheet, titleRow, startCol) || '교육과정편제표';
    const tableDataRowCount = lastDataRow - tableStartRow + 1;
    const cells = [];

    for (let r = tableStartRow; r <= lastDataRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const text = readSheetCellText(sheet, r, c);
        if (text.length === 0) continue;
        cells.push({
          row: r - tableStartRow,
          col: c - startCol,
          text,
        });
      }
    }

    const merges = (sheet['!merges'] || [])
      .filter(m =>
        m.s.r >= tableStartRow && m.e.r <= lastDataRow &&
        m.s.c >= startCol && m.e.c <= endCol
      )
      .map(m => ({
        r1: m.s.r - tableStartRow,
        c1: m.s.c - startCol,
        r2: m.e.r - tableStartRow,
        c2: m.e.c - startCol,
      }));

    const colWidthsExcel = extractColumnWidthsExcel(sheet, arrayBuffer, startCol, colCount, sheetIndex);
    const excelStyleByCell = parseExcelCellStylesFromXlsxXml(arrayBuffer, sheetIndex, {
      startRow: tableStartRow,
      endRow: lastDataRow,
      startCol,
      endCol,
    });
    const tableRangeLabel = `B5:O${lastDataRow + 1}`;

    return {
      sheetName,
      format: 'curriculumOrganization',
      formatLabel: '교육과정편제표',
      titleCell: 'B2',
      titleText,
      tableRangeLabel,
      rowCount: tableDataRowCount,
      colCount,
      rawRowCount,
      rawColCount,
      cells,
      merges,
      colWidthsExcel,
      excelStyleByCell,
      droppedColIdx: [],
      selectionRequirement: null,
      trim: { r1: tableStartRow, c1: startCol, r2: lastDataRow, c2: endCol },
      hasTitleRow: false,
      headerRowCount: 2,
      preserveZeroText: true,
    };
  }

  return null;
}

// WASM/Rust에서 throw된 값은 Error가 아닐 수 있다. 강건하게 문자열화.
function errMsg(e) {
  if (e == null) return 'null';
  if (typeof e === 'string') return e;
  if (e.message) return e.message;
  if (typeof e.toString === 'function') {
    const s = e.toString();
    if (s && s !== '[object Object]') return s;
  }
  try { return JSON.stringify(e); } catch { return String(e); }
}

// createEmpty()만으론 sections/paragraphs가 비어 있어 createTable이 곧장 실패.
// createBlankDocument()로 내장 템플릿(blank2010.hwp)을 입혀 유효 골격을 만든다.
function makeBlankDoc() {
  const doc = HwpDocument.createEmpty();
  parseJsonResult(doc.createBlankDocument(), 'createBlankDocument');
  return doc;
}

// ────────────────────────────────────────────────────────────────
// 셀프 테스트: 3×3 표 만들고 첫 셀에 텍스트 + 한 번 병합
// ────────────────────────────────────────────────────────────────
async function selftest() {
  try {
    await ensureWasm();
    log('--- 셀프 테스트 시작 ---');

    const doc = makeBlankDoc();
    log('빈 문서 골격 준비 OK');

    const r = parseJsonResult(doc.createTable(0, 0, 0, 3, 3), 'createTable');
    const paraIdx = r.paraIdx ?? 0;
    const ctrlIdx = r.controlIdx ?? 0;
    log(`createTable(3×3) OK  paraIdx=${paraIdx} ctrlIdx=${ctrlIdx}`);

    // 셀(0,0)에 텍스트
    parseJsonResult(
      doc.insertTextInCell(0, paraIdx, ctrlIdx, 0, 0, 0, '안녕하세요'),
      'insertTextInCell(0,0)'
    );
    // 셀(1,1)에 텍스트
    parseJsonResult(
      doc.insertTextInCell(0, paraIdx, ctrlIdx, 1 * 3 + 1, 0, 0, '한글 셀'),
      'insertTextInCell(1,1)'
    );
    log('insertTextInCell 2회 OK');

    // 마지막에 (0,0)-(0,2) 병합 (1행 전체)
    parseJsonResult(
      doc.mergeTableCells(0, paraIdx, ctrlIdx, 0, 0, 0, 2),
      'mergeTableCells'
    );
    log('mergeTableCells(0,0)-(0,2) OK');

    const bytes = doc.exportHwp();
    log(`exportHwp() OK  ${bytes.length} bytes`, 'ok');
    downloadBytes(bytes, 'selftest.hwp');
    doc.free();
  } catch (e) {
    log(`[셀프 테스트 실패] ${errMsg(e)}`, 'err');
    console.error(e);
  }
}

// ────────────────────────────────────────────────────────────────
// xlsx 파싱 → 시트 데이터 추출
// ────────────────────────────────────────────────────────────────
function readXlsxTables(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const curriculumData = readCurriculumOrganizationTable(wb, arrayBuffer);
  const curriculumLandscapeData = readCurriculumOrganizationLandscapeTable(wb, arrayBuffer);
  const detectedTables = [curriculumData, curriculumLandscapeData].filter(Boolean);
  if (detectedTables.length > 0) return detectedTables;

  return [readGenericFirstSheetTable(wb, arrayBuffer)];
}

function readXlsxFirstSheet(arrayBuffer) {
  return readXlsxTables(arrayBuffer)[0];
}

function readGenericFirstSheetTable(wb, arrayBuffer) {
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  if (!sheet || !sheet['!ref']) throw new Error('첫 시트가 비어 있습니다.');

  const range = XLSX.utils.decode_range(sheet['!ref']);
  const rowOffset = range.s.r;       // 보통 0
  const colOffset = range.s.c;       // 보통 0
  const rawRowCount = range.e.r - range.s.r + 1;
  const rawColCount = range.e.c - range.s.c + 1;

  // 값 셀 추출 (.w 우선, 없으면 .v 문자열화)
  const cellsRaw = []; // {row, col, text}
  let minColWithData = Infinity;
  let minRowWithData = Infinity;
  let maxColWithData = -1;
  let maxRowWithData = -1;
  for (let r = 0; r < rawRowCount; r++) {
    for (let c = 0; c < rawColCount; c++) {
      const txt = readSheetCellText(sheet, r + rowOffset, c + colOffset);
      if (txt.length === 0) continue;
      cellsRaw.push({ row: r, col: c, text: txt });
      if (c < minColWithData) minColWithData = c;
      if (r < minRowWithData) minRowWithData = r;
      if (c > maxColWithData) maxColWithData = c;
      if (r > maxRowWithData) maxRowWithData = r;
    }
  }

  const rawMerges = (sheet['!merges'] || []).map(m => ({
    r1: m.s.r - rowOffset,
    c1: m.s.c - colOffset,
    r2: m.e.r - rowOffset,
    c2: m.e.c - colOffset,
  }));

  // 트리밍: 값이 있는 영역을 먼저 찾고, 그 영역과 닿는 병합 범위만 포함한다.
  // 기존 방식은 A1부터 마지막 데이터까지 전부 표로 만들어 앞쪽 빈 행/열이 크게 남았다.
  let trim = { r1: 0, c1: 0, r2: 0, c2: 0 };
  if (cellsRaw.length > 0) {
    trim = {
      r1: minRowWithData,
      c1: minColWithData,
      r2: maxRowWithData,
      c2: maxColWithData,
    };

    let expanded = true;
    while (expanded) {
      expanded = false;
      for (const m of rawMerges) {
        const overlaps =
          m.r1 <= trim.r2 && m.r2 >= trim.r1 &&
          m.c1 <= trim.c2 && m.c2 >= trim.c1;
        if (!overlaps) continue;

        const next = {
          r1: Math.min(trim.r1, m.r1),
          c1: Math.min(trim.c1, m.c1),
          r2: Math.max(trim.r2, m.r2),
          c2: Math.max(trim.c2, m.c2),
        };
        if (next.r1 !== trim.r1 || next.c1 !== trim.c1 || next.r2 !== trim.r2 || next.c2 !== trim.c2) {
          trim = next;
          expanded = true;
        }
      }
    }
  }

  const effRowCount = Math.max(trim.r2 - trim.r1 + 1, 1);
  const effColCount = Math.max(trim.c2 - trim.c1 + 1, 1);

  const cells = cellsRaw
    .filter(x => x.row >= trim.r1 && x.row <= trim.r2 && x.col >= trim.c1 && x.col <= trim.c2)
    .map(x => ({ row: x.row - trim.r1, col: x.col - trim.c1, text: x.text }));
  const merges = rawMerges
    .filter(m =>
      m.r1 >= trim.r1 && m.r2 <= trim.r2 &&
      m.c1 >= trim.c1 && m.c2 <= trim.c2
    )
    .map(m => ({
      r1: m.r1 - trim.r1,
      c1: m.c1 - trim.c1,
      r2: m.r2 - trim.r1,
      c2: m.c2 - trim.c1,
    }));

  const colWidthsExcel = extractColumnWidthsExcel(sheet, arrayBuffer, trim.c1 + colOffset, effColCount, 0);

  const selectionGroupMoved = moveSelectionRequirementIntoGroupColumn({
    rowCount: effRowCount,
    colCount: effColCount,
    cells,
    merges,
    colWidthsExcel,
  });

  // 헤더 행(0행)을 제외한 데이터 행이 모두 비어 있는 열은 제거 (예: 비고 열)
  const dropped = dropEmptyDataColumns(selectionGroupMoved);

  return {
    sheetName,
    rowCount: dropped.rowCount, colCount: dropped.colCount,
    rawRowCount, rawColCount,
    cells: dropped.cells, merges: dropped.merges, colWidthsExcel: dropped.colWidthsExcel,
    droppedColIdx: dropped.droppedColIdx,
    selectionRequirement: selectionGroupMoved.selectionRequirement,
    trim,
  };
}

// 헤더 외 데이터 행이 모두 빈 열을 제거. 병합·열폭 재매핑까지 처리.
function dropEmptyDataColumns({ rowCount, colCount, cells, merges, colWidthsExcel }) {
  const hasDataBelowHeader = new Array(colCount).fill(false);
  for (const cell of cells) {
    if (cell.row >= 1) hasDataBelowHeader[cell.col] = true;
  }
  const droppedColIdx = [];
  const oldToNew = new Array(colCount).fill(-1);
  const newWidths = [];
  for (let c = 0; c < colCount; c++) {
    if (!hasDataBelowHeader[c]) {
      droppedColIdx.push(c);
    } else {
      oldToNew[c] = newWidths.length;
      newWidths.push(colWidthsExcel[c]);
    }
  }
  const newColCount = newWidths.length;
  if (newColCount === 0) {
    return { rowCount, colCount, cells, merges, colWidthsExcel, droppedColIdx: [] };
  }
  if (droppedColIdx.length === 0) {
    return { rowCount, colCount, cells, merges, colWidthsExcel, droppedColIdx };
  }

  // 셀: 드롭된 열의 셀 제거 + col 인덱스 재매핑
  const newCells = [];
  for (const cell of cells) {
    const nc = oldToNew[cell.col];
    if (nc >= 0) newCells.push({ row: cell.row, col: nc, text: cell.text });
  }

  // 병합: 드롭된 열을 제외하고 범위 내 첫·끝 보존열 인덱스로 재계산.
  //       보존열이 1개 이하 + 단일 행이면 1x1이 되어 의미 없으므로 스킵.
  const newMerges = [];
  for (const m of merges) {
    let first = -1, last = -1;
    for (let c = m.c1; c <= m.c2; c++) {
      const nc = oldToNew[c];
      if (nc >= 0) {
        if (first < 0) first = nc;
        last = nc;
      }
    }
    if (first < 0) continue; // 병합 범위 전체가 드롭됨
    if (first === last && m.r1 === m.r2) continue; // 1×1
    newMerges.push({ r1: m.r1, c1: first, r2: m.r2, c2: last });
  }

  return {
    rowCount, colCount: newColCount,
    cells: newCells, merges: newMerges,
    colWidthsExcel: newWidths,
    droppedColIdx,
  };
}

function normalizeCompactText(value) {
  return String(value ?? '').replace(/\s+/g, '');
}

function isSelectionRequirementText(text) {
  const compact = normalizeCompactText(text);
  return (
    /학기별.*과목수.*최소선택/.test(compact) ||
    /최소선택[:：]?\d+.*최대선택[:：]?\d+/.test(compact)
  );
}

function formatSelectionRequirementText(text) {
  const compact = normalizeCompactText(text);
  const min = compact.match(/최소선택[:：]?(\d+)/)?.[1] || '';
  const max = compact.match(/최대선택[:：]?(\d+)/)?.[1] || '';
  if (min || max) {
    return `학기별 과목수: ${min ? `최소 ${min}` : ''}${min && max ? ' / ' : ''}${max ? `최대 ${max}` : ''}`;
  }
  return String(text).replace(/\s+/g, ' ').trim();
}

function findSelectionGroupColumn(cells) {
  const candidates = cells.filter(cell => cell.row <= 2);
  const hit = candidates.find(cell => /선택그룹명|선택그룹|그룹명|과목구분/.test(normalizeCompactText(cell.text)));
  return hit ? hit.col : -1;
}

function findMergeCoveringCell(merges, row, col) {
  return merges.find(m => m.r1 <= row && row <= m.r2 && m.c1 <= col && col <= m.c2) || null;
}

function isLikelySelectionGroupLabel(text) {
  const compact = normalizeCompactText(text);
  if (!compact || isSelectionRequirementText(compact)) return false;
  if (/^\d+(?:\.0+)?$/.test(compact)) return false;
  if (/^[-–—~]+$/.test(compact)) return false;
  return /학생선택|학교지정|선택그룹|그룹명|그룹|\d+학년\d+학기/.test(compact);
}

function findRequirementTargetCell(cell, cells, merges, fallbackCol) {
  if (fallbackCol >= 0) {
    const covering = findMergeCoveringCell(merges, cell.row, fallbackCol);
    if (covering) {
      const mergedStart = cells.find(candidate => candidate.row === covering.r1 && candidate.col === covering.c1);
      if (String(mergedStart?.text ?? '').trim()) return { row: covering.r1, col: covering.c1 };
    }
  }

  const sameRowCandidates = cells
    .filter(candidate =>
      candidate.row === cell.row &&
      candidate.col < cell.col &&
      isLikelySelectionGroupLabel(candidate.text)
    )
    .sort((a, b) => b.col - a.col);
  if (sameRowCandidates.length > 0) {
    return { row: sameRowCandidates[0].row, col: sameRowCandidates[0].col };
  }

  if (fallbackCol >= 0) {
    return { row: cell.row, col: fallbackCol };
  }

  return null;
}

function dropColumnsByIndex({ rowCount, colCount, cells, merges, colWidthsExcel }, colsToDrop) {
  const dropSet = new Set(colsToDrop);
  if (dropSet.size === 0) return { rowCount, colCount, cells, merges, colWidthsExcel };

  const oldToNew = new Array(colCount).fill(-1);
  const newWidths = [];
  for (let c = 0; c < colCount; c++) {
    if (dropSet.has(c)) continue;
    oldToNew[c] = newWidths.length;
    newWidths.push(colWidthsExcel[c]);
  }

  const newCells = cells
    .map(cell => ({ ...cell, col: oldToNew[cell.col] }))
    .filter(cell => cell.col >= 0);

  const newMerges = [];
  for (const m of merges) {
    let first = -1, last = -1;
    for (let c = m.c1; c <= m.c2; c++) {
      const nc = oldToNew[c];
      if (nc >= 0) {
        if (first < 0) first = nc;
        last = nc;
      }
    }
    if (first < 0) continue;
    if (first === last && m.r1 === m.r2) continue;
    newMerges.push({ r1: m.r1, c1: first, r2: m.r2, c2: last });
  }

  return {
    rowCount,
    colCount: newWidths.length,
    cells: newCells,
    merges: newMerges,
    colWidthsExcel: newWidths,
  };
}

function moveSelectionRequirementIntoGroupColumn({ rowCount, colCount, cells, merges, colWidthsExcel }) {
  const requirementCells = cells.filter(cell => isSelectionRequirementText(cell.text));
  const stats = {
    movedCount: 0,
    skippedCount: requirementCells.length,
    droppedColIdx: [],
    targetCol: -1,
  };
  if (requirementCells.length === 0) {
    return { rowCount, colCount, cells, merges, colWidthsExcel, selectionRequirement: stats };
  }

  const targetCol = findSelectionGroupColumn(cells);
  stats.targetCol = targetCol;

  const cellMap = new Map(cells.map(cell => [`${cell.row},${cell.col}`, { ...cell }]));
  const requirementStartKeys = new Set();
  const requirementCellsByKey = new Set(requirementCells.map(cell => `${cell.row},${cell.col}`));

  for (const cell of requirementCells) {
    const formatted = formatSelectionRequirementText(cell.text);
    const targetCell = findRequirementTargetCell(cell, cells, merges, targetCol);
    if (!targetCell) continue;
    const targetKey = `${targetCell.row},${targetCell.col}`;
    const target = cellMap.get(targetKey) || { row: targetCell.row, col: targetCell.col, text: '' };
    if (!target.text.includes(formatted)) {
      target.text = target.text.trim()
        ? `${target.text.trim()} (${formatted})`
        : formatted;
    }
    cellMap.set(targetKey, target);

    if (`${cell.row},${cell.col}` !== targetKey) {
      cellMap.delete(`${cell.row},${cell.col}`);
    }
    requirementStartKeys.add(`${cell.row},${cell.col}`);
    stats.movedCount++;
  }
  stats.skippedCount = requirementCells.length - stats.movedCount;

  const movedCells = Array.from(cellMap.values()).sort((a, b) => (a.row - b.row) || (a.col - b.col));
  const movedMerges = merges.filter(m => {
    if (requirementStartKeys.has(`${m.r1},${m.c1}`)) return false;
    for (const key of requirementCellsByKey) {
      const [row, col] = key.split(',').map(Number);
      if (m.r1 <= row && row <= m.r2 && m.c1 <= col && col <= m.c2) return false;
    }
    return true;
  });

  const colsToDrop = [];
  for (let col = 0; col < colCount; col++) {
    if (col === targetCol) continue;
    const colCells = movedCells.filter(cell => cell.col === col && String(cell.text).trim().length > 0);
    const originalRequirementInCol = requirementCells.some(cell => cell.col === col);
    const hasNonRequirementData = colCells.some(cell => cell.row > 0 && !isSelectionRequirementText(cell.text));
    if (originalRequirementInCol && !hasNonRequirementData) colsToDrop.push(col);
  }
  stats.droppedColIdx = colsToDrop;

  const dropped = dropColumnsByIndex({
    rowCount,
    colCount,
    cells: movedCells,
    merges: movedMerges,
    colWidthsExcel,
  }, colsToDrop);

  return { ...dropped, selectionRequirement: stats };
}

// SheetJS의 !cols 미흡 시 폴백: arrayBuffer에서 worksheet XML을 ZIP으로 다시 열어 <col> 태그를 직접 추출.
// xlsx는 ZIP 컨테이너. JSZip 없이 SheetJS 내부 유틸을 빌려 대상 시트 XML만 가져온다.
function parseColWidthsFromXlsxXml(sheet, arrayBuffer, sheetIndex = 0) {
  try {
    // SheetJS는 압축 해제된 파일들을 wb.Strings 등에 보관하지 않으므로, ZIP을 직접 다시 풀어야 한다.
    // 가벼운 ZIP 디코더 대신 SheetJS의 raw 모드를 이용한다: XLSX.read with bookFiles=true
    // → 다만 이 옵션을 위해선 다시 읽어야 함. 비용 작음 (수백KB).
    const wb = XLSX.read(arrayBuffer, { type: 'array', bookFiles: true });
    const files = wb.files || wb.Files || null;
    if (!files) return null;
    // 대상 시트의 XML 파일명 찾기
    const sheetXmlName = `xl/worksheets/sheet${sheetIndex + 1}.xml`;
    const sheetKey = Object.keys(files).find(k =>
      k.replace(/^\/+/, '').toLowerCase() === sheetXmlName.toLowerCase()
    );
    if (!sheetKey) return null;
    const xmlContent = files[sheetKey].content || files[sheetKey];
    const xml = (typeof xmlContent === 'string') ? xmlContent : new TextDecoder('utf-8').decode(xmlContent);
    // <col min="A" max="B" width="W" .../>
    const widths = [];
    const re = /<col\b[^>]*\bmin="(\d+)"[^>]*\bmax="(\d+)"[^>]*\bwidth="([^"]+)"/g;
    let m;
    while ((m = re.exec(xml)) !== null) {
      const min = parseInt(m[1], 10) - 1; // 0-based
      const max = parseInt(m[2], 10) - 1;
      const w = parseFloat(m[3]);
      if (!isFinite(w)) continue;
      for (let c = min; c <= max; c++) widths[c] = w;
    }
    return widths;
  } catch (e) {
    console.warn('[debug] parseColWidthsFromXlsxXml 실패:', e);
    return null;
  }
}

function parseXmlAttrs(attrText) {
  const attrs = {};
  const re = /([A-Za-z_:][\w:.-]*)="([^"]*)"/g;
  let m;
  while ((m = re.exec(attrText || '')) !== null) attrs[m[1]] = m[2];
  return attrs;
}

function readXlsxBookFiles(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', bookFiles: true });
  return wb.files || wb.Files || null;
}

function decodeXlsxFileContent(content) {
  if (content == null) return '';
  if (typeof content === 'string') return content;
  return new TextDecoder('utf-8').decode(content);
}

function readXlsxFileText(files, targetPath) {
  if (!files) return '';
  const normalizedTarget = targetPath.replace(/^\/+/, '').toLowerCase();
  const key = Object.keys(files).find(k => k.replace(/^\/+/, '').toLowerCase() === normalizedTarget);
  if (!key) return '';
  return decodeXlsxFileContent(files[key].content || files[key]);
}

function normalizeHexColor(value) {
  const clean = String(value ?? '').replace(/[^0-9A-Fa-f]/g, '').toUpperCase();
  if (clean.length >= 8) return `#${clean.slice(-6)}`;
  if (clean.length === 6) return `#${clean}`;
  return null;
}

const EXCEL_INDEXED_COLORS = [
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
  '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
  '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
  '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
  '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
];

function getFirstXmlTagBlock(xml, tagName) {
  const re = new RegExp(`<(?:[A-Za-z]+:)?${tagName}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z]+:)?${tagName}>`, 'i');
  return xml.match(re)?.[1] || '';
}

function parseThemeColorMap(themeXml) {
  const scheme = getFirstXmlTagBlock(themeXml || '', 'clrScheme');
  if (!scheme) return [];

  const names = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  return names.map(name => {
    const block = getFirstXmlTagBlock(scheme, name);
    if (!block) return null;

    const srgb = block.match(/<(?:[A-Za-z]+:)?srgbClr\b([^>]*)\/?>/i);
    if (srgb) return normalizeHexColor(parseXmlAttrs(srgb[1]).val);

    const sys = block.match(/<(?:[A-Za-z]+:)?sysClr\b([^>]*)\/?>/i);
    if (sys) return normalizeHexColor(parseXmlAttrs(sys[1]).lastClr);

    return null;
  });
}

function applyExcelTint(hex, tintValue) {
  const tint = Number(tintValue || 0);
  if (!hex || !isFinite(tint) || tint === 0) return hex;
  const clean = hex.replace('#', '');
  const channels = [0, 2, 4].map(offset => parseInt(clean.slice(offset, offset + 2), 16) || 0);
  const tinted = channels.map(channel => {
    const next = tint < 0
      ? channel * (1 + tint)
      : channel * (1 - tint) + 255 * tint;
    return Math.max(0, Math.min(255, Math.round(next)));
  });
  return `#${tinted.map(value => value.toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

function excelColorAttrsToHex(attrs, themeColors = []) {
  if (!attrs || attrs.auto === '1') return null;
  if (attrs.rgb) return normalizeHexColor(attrs.rgb);
  if (attrs.indexed != null) {
    const indexed = Number(attrs.indexed);
    return EXCEL_INDEXED_COLORS[indexed] || null;
  }
  if (attrs.theme != null) {
    const themeColor = themeColors[Number(attrs.theme)];
    return themeColor ? applyExcelTint(themeColor, attrs.tint) : null;
  }
  return null;
}

function parseExcelColorTag(block, tagName, themeColors) {
  const match = block.match(new RegExp(`<${tagName}\\b([^>]*)\\/?>`, 'i'));
  return match ? excelColorAttrsToHex(parseXmlAttrs(match[1]), themeColors) : null;
}

function parseStyleFillColors(stylesXml, themeColors) {
  const fillsSection = stylesXml.match(/<fills\b[^>]*>([\s\S]*?)<\/fills>/i)?.[1] || '';
  const fills = [];
  const fillRe = /<fill\b[^>]*\/>|<fill\b[^>]*>[\s\S]*?<\/fill>/gi;
  let fillMatch;
  while ((fillMatch = fillRe.exec(fillsSection)) !== null) {
    const block = fillMatch[0];
    const patternAttrs = parseXmlAttrs(block.match(/<patternFill\b([^>]*)/i)?.[1] || '');
    const patternType = String(patternAttrs.patternType || '').toLowerCase();
    if (patternType === 'none' || patternType === 'gray125') {
      fills.push(null);
      continue;
    }
    fills.push(
      parseExcelColorTag(block, 'fgColor', themeColors) ||
      parseExcelColorTag(block, 'bgColor', themeColors) ||
      null
    );
  }
  return fills;
}

function parseStyleFonts(stylesXml, themeColors) {
  const fontsSection = stylesXml.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i)?.[1] || '';
  const fonts = [];
  const fontRe = /<font\b[^>]*\/>|<font\b[^>]*>[\s\S]*?<\/font>/gi;
  let fontMatch;
  while ((fontMatch = fontRe.exec(fontsSection)) !== null) {
    const block = fontMatch[0];
    const color = parseExcelColorTag(block, 'color', themeColors);
    fonts.push({
      textColor: color,
      bold: /<b\b[^>]*\/?>/i.test(block),
    });
  }
  return fonts;
}

function parseStyleFormatMap(stylesXml, themeXml) {
  if (!stylesXml) return [];

  const themeColors = parseThemeColorMap(themeXml);
  const fills = parseStyleFillColors(stylesXml, themeColors);
  const fonts = parseStyleFonts(stylesXml, themeColors);
  const borders = parseStyleBorderMap(stylesXml, themeColors);

  const cellXfsSection = stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] || '';
  const styles = [];
  const xfRe = /<xf\b([^>]*)\/?>/gi;
  let xfMatch;
  while ((xfMatch = xfRe.exec(cellXfsSection)) !== null) {
    const attrs = parseXmlAttrs(xfMatch[1]);
    const fillId = Number(attrs.fillId || 0);
    const fontId = Number(attrs.fontId || 0);
    const borderId = Number(attrs.borderId || 0);
    const font = fonts[fontId] || {};
    const usesCustomFont = attrs.applyFont === '1' || fontId > 0;
    styles.push({
      fillColor: fills[fillId] || null,
      textColor: usesCustomFont ? (font.textColor || null) : null,
      bold: usesCustomFont && font.bold ? true : null,
      border: borders[borderId] || { left: null, right: null, top: null, bottom: null },
    });
  }

  return styles;
}

function isNonDefaultExcelStyle(style) {
  return !!(style?.fillColor || style?.textColor || style?.bold);
}

function parseExcelCellStylesFromXlsxXml(arrayBuffer, sheetIndex, { startRow, endRow, startCol, endCol }) {
  const styleByCell = new Map();
  try {
    const files = readXlsxBookFiles(arrayBuffer);
    if (!files) return styleByCell;

    const stylesXml = readXlsxFileText(files, 'xl/styles.xml');
    const themeXml = readXlsxFileText(files, 'xl/theme/theme1.xml');
    const styleFormats = parseStyleFormatMap(stylesXml, themeXml);
    if (styleFormats.length === 0) return styleByCell;

    const sheetXml = readXlsxFileText(files, `xl/worksheets/sheet${sheetIndex + 1}.xml`);
    if (!sheetXml) return styleByCell;

    const cellRe = /<c\b([^>]*)\/?>/gi;
    let cellMatch;
    while ((cellMatch = cellRe.exec(sheetXml)) !== null) {
      const attrs = parseXmlAttrs(cellMatch[1]);
      if (!attrs.r) continue;
      const addr = XLSX.utils.decode_cell(attrs.r);
      if (addr.r < startRow || addr.r > endRow || addr.c < startCol || addr.c > endCol) continue;

      const style = styleFormats[Number(attrs.s || 0)];
      if (!isNonDefaultExcelStyle(style)) continue;
      styleByCell.set(`${addr.r - startRow},${addr.c - startCol}`, {
        fillColor: style.fillColor,
        textColor: style.textColor,
        bold: style.bold,
      });
    }
  } catch (e) {
    console.warn('[debug] parseExcelCellStylesFromXlsxXml 실패:', e);
  }
  return styleByCell;
}

function parseBorderSide(borderBlock, side, themeColors) {
  const re = new RegExp(`<${side}\\b([^>]*)>([\\s\\S]*?)<\\/${side}>|<${side}\\b([^>]*)\\/?>`, 'i');
  const match = borderBlock.match(re);
  const attrs = parseXmlAttrs(match?.[1] || match?.[3] || '');
  const inner = match?.[2] || '';
  const visible = !!attrs.style && attrs.style !== 'none';
  const color = visible ? parseExcelColorTag(inner, 'color', themeColors) : null;
  return { visible, color };
}

function parseVisibleBorderSides(borderBlock, themeColors = []) {
  const sides = {};
  for (const side of ['left', 'right', 'top', 'bottom']) {
    sides[side] = parseBorderSide(borderBlock, side, themeColors);
  }
  return sides;
}

function parseStyleBorderMap(stylesXml, themeColors = []) {
  if (!stylesXml) return [];

  const bordersSection = stylesXml.match(/<borders\b[^>]*>([\s\S]*?)<\/borders>/i)?.[1] || '';
  const borders = [];
  const borderRe = /<border\b[^>]*\/>|<border\b[^>]*>[\s\S]*?<\/border>/gi;
  let borderMatch;
  while ((borderMatch = borderRe.exec(bordersSection)) !== null) {
    borders.push(parseVisibleBorderSides(borderMatch[0], themeColors));
  }

  const cellXfsSection = stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] || '';
  const styleBorders = [];
  const xfRe = /<xf\b([^>]*)\/?>/gi;
  let xfMatch;
  while ((xfMatch = xfRe.exec(cellXfsSection)) !== null) {
    const attrs = parseXmlAttrs(xfMatch[1]);
    const borderId = Number(attrs.borderId || 0);
    styleBorders.push(borders[borderId] || { left: null, right: null, top: null, bottom: null });
  }

  return styleBorders;
}

function isBorderSideVisible(side) {
  return typeof side === 'object' && side !== null ? !!side.visible : !!side;
}

function getBorderSideColor(side, fallbackColor) {
  return (typeof side === 'object' && side?.color) ? side.color : fallbackColor;
}

function hasAnyVisibleBorder(sides) {
  return !!(
    isBorderSideVisible(sides?.left) ||
    isBorderSideVisible(sides?.right) ||
    isBorderSideVisible(sides?.top) ||
    isBorderSideVisible(sides?.bottom)
  );
}

function parseExcelCellBordersFromXlsxXml(arrayBuffer, sheetIndex, { startRow, endRow, startCol, endCol }) {
  const borderByCell = new Map();
  try {
    const files = readXlsxBookFiles(arrayBuffer);
    if (!files) return borderByCell;

    const stylesXml = readXlsxFileText(files, 'xl/styles.xml');
    const themeXml = readXlsxFileText(files, 'xl/theme/theme1.xml');
    const styleBorders = parseStyleBorderMap(stylesXml, parseThemeColorMap(themeXml));
    if (styleBorders.length === 0) return borderByCell;

    const sheetXml = readXlsxFileText(files, `xl/worksheets/sheet${sheetIndex + 1}.xml`);
    if (!sheetXml) return borderByCell;

    const cellRe = /<c\b([^>]*)\/?>/gi;
    let cellMatch;
    while ((cellMatch = cellRe.exec(sheetXml)) !== null) {
      const attrs = parseXmlAttrs(cellMatch[1]);
      if (!attrs.r) continue;
      const addr = XLSX.utils.decode_cell(attrs.r);
      if (addr.r < startRow || addr.r > endRow || addr.c < startCol || addr.c > endCol) continue;

      const styleIdx = Number(attrs.s || 0);
      const border = styleBorders[styleIdx];
      if (!hasAnyVisibleBorder(border)) continue;

      borderByCell.set(`${addr.r - startRow},${addr.c - startCol}`, border);
    }
  } catch (e) {
    console.warn('[debug] parseExcelCellBordersFromXlsxXml 실패:', e);
  }
  return borderByCell;
}

const TABLE_THEME = {
  titleBg: '#18324A',
  headerBg: '#245C8F',
  headerBg2: '#285F91',
  headerBorder: '#E4EDF5',
  grid: '#D6E0EA',
  rowEven: '#FFFFFF',
  rowOdd: '#F7FAFC',
  indexEven: '#EAF2EF',
  indexOdd: '#E4EEEB',
  text: '#111827',
  indexText: '#243832',
  whiteText: '#FFFFFF',
};

const TABLE_FONT_FAMILY = '맑은 고딕';

const BORDER_COLOR_BY_FILL_COLOR = {
  [TABLE_THEME.titleBg]: TABLE_THEME.titleBg,
  [TABLE_THEME.headerBg]: TABLE_THEME.headerBorder,
  [TABLE_THEME.headerBg2]: '#E5EDF5',
  [TABLE_THEME.indexEven]: '#C9DCD6',
  [TABLE_THEME.indexOdd]: '#CADAD5',
  [TABLE_THEME.rowEven]: TABLE_THEME.grid,
  [TABLE_THEME.rowOdd]: '#D7E0EA',
};

const FILL_COLOR_BY_BORDER_COLOR = Object.fromEntries(
  Object.entries(BORDER_COLOR_BY_FILL_COLOR).map(([fill, border]) => [border.toUpperCase(), fill])
);

function hexToColorRefInt(hex) {
  const clean = String(hex).replace('#', '').trim();
  const r = parseInt(clean.slice(0, 2), 16) || 0;
  const g = parseInt(clean.slice(2, 4), 16) || 0;
  const b = parseInt(clean.slice(4, 6), 16) || 0;
  return ((b << 16) | (g << 8) | r) >>> 0;
}

function colorRefBytesToHex(bytes, offset) {
  const r = bytes[offset] ?? 0;
  const g = bytes[offset + 1] ?? 0;
  const b = bytes[offset + 2] ?? 0;
  return `#${[r, g, b].map(x => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

function readU32LE(bytes, offset) {
  return (
    bytes[offset] |
    (bytes[offset + 1] << 8) |
    (bytes[offset + 2] << 16) |
    (bytes[offset + 3] << 24)
  ) >>> 0;
}

function writeU32LE(bytes, offset, value) {
  bytes[offset] = value & 0xff;
  bytes[offset + 1] = (value >>> 8) & 0xff;
  bytes[offset + 2] = (value >>> 16) & 0xff;
  bytes[offset + 3] = (value >>> 24) & 0xff;
}

function writeU16LE(bytes, offset, value) {
  bytes[offset] = value & 0xff;
  bytes[offset + 1] = (value >>> 8) & 0xff;
}

function concatBytes(chunks) {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

function asUint8Array(value) {
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  return new Uint8Array(value);
}

async function transformRawDeflate(bytes, mode) {
  const cfbUtils = XLSX?.CFB?.utils;
  const sheetJsTransform = mode === 'inflate' ? cfbUtils?._inflateRaw : cfbUtils?._deflateRaw;
  if (sheetJsTransform) {
    return asUint8Array(sheetJsTransform(asUint8Array(bytes)));
  }

  const Ctor = mode === 'inflate' ? globalThis.DecompressionStream : globalThis.CompressionStream;
  if (!Ctor || typeof Blob === 'undefined' || typeof Response === 'undefined') {
    throw new Error('이 브라우저에서 HWP 압축 스트림 처리를 지원하지 않습니다.');
  }
  const stream = new Blob([bytes]).stream().pipeThrough(new Ctor('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function parseHwpRecords(bytes) {
  const records = [];
  let pos = 0;
  while (pos + 4 <= bytes.length) {
    const header = readU32LE(bytes, pos);
    pos += 4;
    const tag = header & 0x3ff;
    const level = (header >>> 10) & 0x3ff;
    let size = (header >>> 20) & 0xfff;
    if (size === 0xfff) {
      size = readU32LE(bytes, pos);
      pos += 4;
    }
    records.push({ tag, level, data: bytes.slice(pos, pos + size) });
    pos += size;
  }
  return records;
}

function buildHwpRecords(records) {
  const chunks = [];
  for (const record of records) {
    const size = record.data.length;
    if (size < 0xfff) {
      const header = (record.tag & 0x3ff) | ((record.level & 0x3ff) << 10) | (size << 20);
      const h = new Uint8Array(4);
      writeU32LE(h, 0, header);
      chunks.push(h, record.data);
    } else {
      const h = new Uint8Array(8);
      const header = (record.tag & 0x3ff) | ((record.level & 0x3ff) << 10) | (0xfff << 20);
      writeU32LE(h, 0, header);
      writeU32LE(h, 4, size);
      chunks.push(h, record.data);
    }
  }
  return concatBytes(chunks);
}

function makeSolidBorderFillData(baseData, fillHex) {
  const out = new Uint8Array(53);
  out.set(baseData.slice(0, Math.min(32, baseData.length)), 0);
  writeU32LE(out, 32, 1); // solid fill
  writeU32LE(out, 36, hexToColorRefInt(fillHex));
  writeU32LE(out, 40, 0);
  writeU32LE(out, 44, 0xffffffff);
  writeU32LE(out, 48, 1);
  out[52] = 0;
  return out;
}

function inferFillColorFromBorderFillData(data) {
  // BorderFill: attr(2) + 4 borders * (kind,width,colorref) + diagonal border.
  for (const offset of [4, 10, 16, 22]) {
    const fill = FILL_COLOR_BY_BORDER_COLOR[colorRefBytesToHex(data, offset)];
    if (fill) return fill;
  }
  return null;
}

async function applyHwpBorderFillColors(bytes, fillByBorderFillId) {
  if (!fillByBorderFillId || fillByBorderFillId.size === 0) return bytes;
  if (!XLSX?.CFB) {
    throw new Error('이 브라우저에서 HWP 색상 후처리를 지원하지 않습니다.');
  }

  const cfb = XLSX.CFB.read(bytes, { type: 'array' });
  const docInfo = XLSX.CFB.find(cfb, 'DocInfo') || cfb.FileIndex.find(file => file.name === 'DocInfo');
  if (!docInfo || !docInfo.content) throw new Error('DocInfo 스트림을 찾을 수 없습니다.');

  const inflated = await transformRawDeflate(new Uint8Array(docInfo.content), 'inflate');
  const records = parseHwpRecords(inflated);
  let borderFillId = 0;
  let patched = 0;
  for (const record of records) {
    if (record.tag !== 20) continue; // HWPTAG_BORDER_FILL
    borderFillId++;
    const fillHex = fillByBorderFillId.get(borderFillId) || inferFillColorFromBorderFillData(record.data);
    if (!fillHex) continue;
    record.data = makeSolidBorderFillData(record.data, fillHex);
    patched++;
  }

  if (patched === 0) return bytes;
  docInfo.content = await transformRawDeflate(buildHwpRecords(records), 'deflate');
  docInfo.size = docInfo.content.length;
  return new Uint8Array(XLSX.CFB.write(cfb, { type: 'array' }));
}

function patchCharShapeFontFaceIds(record, fontFaceId) {
  // HWPTAG_CHAR_SHAPE starts with 7 uint16 font face IDs
  // for language categories: Hangul, Latin, Hanja, Japanese, Other, Symbol, User.
  if (record.data.length < 14 || fontFaceId < 0 || fontFaceId > 0xffff) return false;
  for (let i = 0; i < 7; i++) {
    writeU16LE(record.data, i * 2, fontFaceId);
  }
  return true;
}

async function applyHwpCharShapeFont(bytes, fontFaceId) {
  if (fontFaceId == null || fontFaceId < 0) return { bytes, patched: 0 };
  if (!XLSX?.CFB) {
    throw new Error('이 브라우저에서 HWP 글꼴 후처리를 지원하지 않습니다.');
  }

  const cfb = XLSX.CFB.read(bytes, { type: 'array' });
  const docInfo = XLSX.CFB.find(cfb, 'DocInfo') || cfb.FileIndex.find(file => file.name === 'DocInfo');
  if (!docInfo || !docInfo.content) throw new Error('DocInfo 스트림을 찾을 수 없습니다.');

  const inflated = await transformRawDeflate(new Uint8Array(docInfo.content), 'inflate');
  const records = parseHwpRecords(inflated);
  let patched = 0;
  for (const record of records) {
    if (record.tag !== 21) continue; // HWPTAG_CHAR_SHAPE
    if (patchCharShapeFontFaceIds(record, fontFaceId)) patched++;
  }

  if (patched === 0) return { bytes, patched };
  docInfo.content = await transformRawDeflate(buildHwpRecords(records), 'deflate');
  docInfo.size = docInfo.content.length;
  return {
    bytes: new Uint8Array(XLSX.CFB.write(cfb, { type: 'array' })),
    patched,
  };
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, ch => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;',
  })[ch]);
}

function textToHtml(text) {
  return escapeHtml(text).replace(/\r\n|\r|\n/g, '<br>');
}

function styleAttr(style) {
  return Object.entries(style)
    .filter(([, value]) => value !== undefined && value !== null && value !== '')
    .map(([key, value]) => `${key}:${value}`)
    .join(';');
}

function computeCssColWidths(colWidthsExcel, totalPx = 660) {
  const count = colWidthsExcel.length;
  if (count === 0) return [];

  const minByCount = count >= 12 ? 30 : count >= 8 ? 40 : 52;
  const minPx = Math.max(24, Math.min(minByCount, Math.floor(totalPx / count) - 2));
  const baseTotal = minPx * count;
  if (baseTotal >= totalPx) {
    return new Array(count).fill(Math.max(24, Math.floor(totalPx / count)));
  }

  const weights = colWidthsExcel.map(w => Math.max(1, Number(w) || 1));
  const sum = weights.reduce((a, b) => a + b, 0) || 1;
  const remaining = totalPx - baseTotal;
  const widths = weights.map(w => Math.round(minPx + (w / sum * remaining)));
  const diff = totalPx - widths.reduce((a, b) => a + b, 0);
  widths[widths.length - 1] += diff;
  return widths;
}

function detectTitleRow(data) {
  if (data.rowCount < 2 || data.colCount < 2) return false;
  const textCells = data.cells.filter(cell => cell.row === 0 && cell.text.trim().length > 0);
  const wideTopMerge = data.merges.some(m =>
    m.r1 === 0 &&
    m.c1 === 0 &&
    (m.c2 - m.c1 + 1) >= Math.max(2, Math.ceil(data.colCount * 0.6))
  );
  return wideTopMerge && textCells.length <= 2;
}

function getHeaderRowCount(data, hasTitleRow) {
  if (typeof data.headerRowCount === 'number') return Math.max(1, data.headerRowCount);
  return hasTitleRow ? 2 : 1;
}

function buildStyledTableHtml(data, totalPx = 660) {
  const colWidthsPx = computeCssColWidths(data.colWidthsExcel, totalPx);
  const hasTitleRow = data.hasTitleRow ?? detectTitleRow(data);
  const headerRowCount = getHeaderRowCount(data, hasTitleRow);
  const { textByCell } = buildDisplayTextByCell(data, hasTitleRow, headerRowCount);
  const mergeStarts = new Map();
  const covered = new Set();

  for (const m of data.merges) {
    mergeStarts.set(`${m.r1},${m.c1}`, {
      rowSpan: m.r2 - m.r1 + 1,
      colSpan: m.c2 - m.c1 + 1,
      endRow: m.r2,
    });
    for (let r = m.r1; r <= m.r2; r++) {
      for (let c = m.c1; c <= m.c2; c++) {
        if (r !== m.r1 || c !== m.c1) covered.add(`${r},${c}`);
      }
    }
  }

  const rowsToRender = [];
  for (let r = 0; r < data.rowCount; r++) {
    let hasRenderableCell = false;
    for (let c = 0; c < data.colCount; c++) {
      if (!covered.has(`${r},${c}`)) {
        hasRenderableCell = true;
        break;
      }
    }
    if (hasRenderableCell) rowsToRender.push(r);
  }
  const renderedRowSpan = (startRow, endRow) =>
    rowsToRender.filter(row => row >= startRow && row <= endRow).length || 1;
  const cells = [];
  const html = [
    '<table style="border-collapse:collapse;table-layout:fixed;font-family:Malgun Gothic,Arial,sans-serif">',
  ];
  let cellIdx = 0;

  for (const r of rowsToRender) {
    html.push('<tr>');
    for (let c = 0; c < data.colCount; c++) {
      if (covered.has(`${r},${c}`)) continue;

      const span = mergeStarts.get(`${r},${c}`) || { rowSpan: 1, colSpan: 1 };
      const rowSpan = span.endRow == null ? 1 : renderedRowSpan(r, span.endRow);
      const text = textByCell.get(`${r},${c}`) || '';
      const isTitle = hasTitleRow && r === 0;
      const isHeader = !isTitle && r < headerRowCount;
      const isIndex = !isTitle && !isHeader && c === 0;
      const kind = isTitle ? 'title' : isHeader ? 'header' : isIndex ? 'index' : 'body';
      const fillColor = getCellFillColor(kind, r, c, data, text);
      const textFormat = getExcelTextFormat(data, r, c);
      const widthPx = colWidthsPx
        .slice(c, c + span.colSpan)
        .reduce((a, b) => a + b, 0) || 80 * span.colSpan;

      const style = {
        width: `${widthPx}px`,
        height: kind === 'title' ? '34px' : '28px',
        border: `1px solid ${kind === 'title' ? TABLE_THEME.titleBg : kind === 'header' ? TABLE_THEME.headerBorder : TABLE_THEME.grid}`,
        padding: kind === 'title' ? '7px 9px' : '5px 7px',
        'text-align': 'center',
        'vertical-align': 'middle',
        'word-break': 'keep-all',
        'overflow-wrap': 'break-word',
        background: fillColor,
        color: textFormat.textColor || (kind === 'index' ? TABLE_THEME.indexText : TABLE_THEME.text),
        'font-weight': textFormat.bold || kind === 'title' || kind === 'header' || kind === 'index' ? '700' : '400',
      };

      const tag = kind === 'body' || kind === 'index' ? 'td' : 'th';
      const attrs = [
        rowSpan > 1 ? `rowspan="${rowSpan}"` : '',
        span.colSpan > 1 ? `colspan="${span.colSpan}"` : '',
        `style="${styleAttr(style)}"`,
      ].filter(Boolean).join(' ');

      html.push(`<${tag} ${attrs}>${textToHtml(text)}</${tag}>`);
      cells.push({ cellIdx, row: r, col: c, text, kind });
      cellIdx++;
    }
    html.push('</tr>');
  }

  html.push('</table>');
  return { html: html.join(''), cells, hasTitleRow };
}

function getTableCellKind(row, col, hasTitleRow, headerRowCount) {
  const isTitle = hasTitleRow && row === 0;
  const isHeader = !isTitle && row < headerRowCount;
  const isIndex = !isTitle && !isHeader && col === 0;
  return isTitle ? 'title' : isHeader ? 'header' : isIndex ? 'index' : 'body';
}

function getExcelCellStyle(data, row, col) {
  return data?.excelStyleByCell?.get(`${row},${col}`) || null;
}

function parseHexRgb(hex) {
  const normalized = normalizeHexColor(hex);
  if (!normalized) return null;
  const clean = normalized.replace('#', '');
  return [0, 2, 4].map(offset => parseInt(clean.slice(offset, offset + 2), 16) || 0);
}

function rgbToHex(rgb) {
  return `#${rgb.map(value =>
    Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, '0')
  ).join('').toUpperCase()}`;
}

function mixHexColors(hexA, hexB, amountB = 0.1) {
  const a = parseHexRgb(hexA);
  const b = parseHexRgb(hexB);
  if (!a || !b) return normalizeHexColor(hexA) || normalizeHexColor(hexB) || TABLE_THEME.grid;
  const amountA = 1 - amountB;
  return rgbToHex(a.map((channel, idx) => channel * amountA + b[idx] * amountB));
}

function getCellFillColor(kind, row, col, data = null, text = '') {
  const excelFill = getExcelCellStyle(data, row, col)?.fillColor;
  const hasText = String(text ?? '').trim().length > 0;
  const excelBorder = data?.borderByCell?.get(`${row},${col}`);
  if (excelFill && (hasText || hasAnyVisibleBorder(excelBorder))) return excelFill;
  return TABLE_THEME.rowEven;
}

function getCellBorderColor(kind, fillHex) {
  // rhwp가 같은 테두리 스타일을 하나의 BorderFill로 공유하므로,
  // 배경색별로 눈에 거의 띄지 않는 테두리 차이를 줘서 채우기 레코드를 분리한다.
  if (BORDER_COLOR_BY_FILL_COLOR[fillHex]) return BORDER_COLOR_BY_FILL_COLOR[fillHex];
  if (kind === 'title') return mixHexColors(TABLE_THEME.titleBg, fillHex, 0.08);
  if (kind === 'header') return mixHexColors(TABLE_THEME.headerBorder, fillHex, 0.08);
  return mixHexColors(TABLE_THEME.grid, fillHex, 0.08);
}

function hwpBorderSide(visible, visibleColor, hiddenColor) {
  return {
    type: 1,
    width: 1,
    color: visible ? visibleColor : hiddenColor,
  };
}

const ADJACENT_BORDER_SIDE = {
  left: { dr: 0, dc: -1, opposite: 'right' },
  right: { dr: 0, dc: 1, opposite: 'left' },
  top: { dr: -1, dc: 0, opposite: 'bottom' },
  bottom: { dr: 1, dc: 0, opposite: 'top' },
};

function getAdjacentBorderSide(data, row, col, sideName) {
  const spec = ADJACENT_BORDER_SIDE[sideName];
  if (!spec) return null;
  const nextRow = row + spec.dr;
  const nextCol = col + spec.dc;
  if (nextRow < 0 || nextCol < 0 || nextRow >= data.rowCount || nextCol >= data.colCount) return null;
  return data.borderByCell?.get(`${nextRow},${nextCol}`)?.[spec.opposite] || null;
}

function resolveSharedBorderSide(data, row, col, sideName, fallbackColor, fillColor) {
  const border = data.borderByCell?.get(`${row},${col}`) || {};
  const ownSide = border[sideName];
  const adjacentSide = getAdjacentBorderSide(data, row, col, sideName);
  const visible = isBorderSideVisible(ownSide) || isBorderSideVisible(adjacentSide);
  const colorSource = isBorderSideVisible(ownSide) ? ownSide : adjacentSide;
  return {
    visible,
    color: mixHexColors(getBorderSideColor(colorSource, fallbackColor), fillColor, 0.04),
  };
}

function getCellBorderProps(data, key, kind, fillColor, borderColor) {
  if (data?.format !== 'curriculumOrganizationLandscape') {
    return {
      borderLeft: { type: 1, width: 1, color: borderColor },
      borderRight: { type: 1, width: 1, color: borderColor },
      borderTop: { type: 1, width: 1, color: borderColor },
      borderBottom: { type: 1, width: 1, color: borderColor },
    };
  }

  const [row, col] = key.split(',').map(Number);
  const hiddenColor = fillColor;
  const baseVisibleColor = kind === 'header' ? TABLE_THEME.headerBorder : TABLE_THEME.grid;
  const visibleColor = mixHexColors(baseVisibleColor, fillColor, 0.08);
  const left = resolveSharedBorderSide(data, row, col, 'left', visibleColor, fillColor);
  const right = resolveSharedBorderSide(data, row, col, 'right', visibleColor, fillColor);
  const top = resolveSharedBorderSide(data, row, col, 'top', visibleColor, fillColor);
  const bottom = resolveSharedBorderSide(data, row, col, 'bottom', visibleColor, fillColor);
  return {
    borderLeft: hwpBorderSide(left.visible, left.color, hiddenColor),
    borderRight: hwpBorderSide(right.visible, right.color, hiddenColor),
    borderTop: hwpBorderSide(top.visible, top.color, hiddenColor),
    borderBottom: hwpBorderSide(bottom.visible, bottom.color, hiddenColor),
  };
}

function getExcelTextFormat(data, row, col) {
  const style = getExcelCellStyle(data, row, col);
  const props = {};
  if (style?.textColor) props.textColor = style.textColor;
  if (style?.bold) props.bold = true;
  return props;
}

function hasTextFormatProps(props) {
  return props && Object.keys(props).length > 0;
}

function getLandscapeCellHeight(kind, paragraphCount) {
  const lines = Math.max(1, Number(paragraphCount) || 1);
  if (kind === 'title' || kind === 'header') {
    return 980 + ((lines - 1) * 320);
  }
  return 760 + ((lines - 1) * 360);
}

function getHeaderTextsByColumn(data, hasTitleRow, headerRowCount = getHeaderRowCount(data, hasTitleRow)) {
  const firstHeaderRow = hasTitleRow ? 1 : 0;
  const lastHeaderRow = Math.max(firstHeaderRow, headerRowCount - 1);
  const headers = new Array(data.colCount).fill('');
  for (const cell of data.cells) {
    if (cell.row >= firstHeaderRow && cell.row <= lastHeaderRow) {
      const text = cell.text.trim();
      if (text) headers[cell.col] = headers[cell.col] ? `${headers[cell.col]} ${text}` : text;
    }
  }
  return headers;
}

function getColumnRole(headerText) {
  const h = headerText.replace(/\s+/g, '');
  if (h === '학점') return 'credit';
  if (/기준학점|운영학점|이수단위|필수단위/.test(h)) return 'credit';
  if (/^\d+-\d+$/.test(h)) return 'term';
  if (/학기별|학점수|과목수|최소선택|학년.*학기/.test(h)) return 'term';
  if (/선택그룹|그룹명/.test(h)) return 'selectionGroup';
  if (/과목구분|구분/.test(h)) return 'category';
  if (/교과/.test(h)) return 'group';
  if (/과목유형|유형/.test(h)) return 'type';
  if (h === '과목' || /과목명/.test(h)) return 'subject';
  return 'default';
}

function computeCurriculumOrganizationColWidths(colCount, availableWidth) {
  // B:O 고정 구조: 구분, 과목, 과목-일반/정보, 기준/운영, 6개 학기, 이수/필수 단위.
  // 좁은 수치 열을 억지로 늘리지 않아 한글에서 한 페이지 안에 안정적으로 들어오게 한다.
  const HWPUNIT_PER_MM = 283;
  const widthsMm = [12, 15, 11, 28, 9, 9, 8, 8, 8, 8, 8, 8, 9, 9];
  let widths = widthsMm.slice(0, colCount).map(mm => Math.round(mm * HWPUNIT_PER_MM));

  if (colCount > widths.length) {
    widths = widths.concat(new Array(colCount - widths.length).fill(Math.round(8 * HWPUNIT_PER_MM)));
  }

  const total = widths.reduce((sum, width) => sum + width, 0);
  const scale = availableWidth / total;
  const scaled = widths.map(width => Math.floor(width * scale));
  scaled[scaled.length - 1] += availableWidth - scaled.reduce((sum, width) => sum + width, 0);
  return scaled;
}

function computeProportionalColWidths(colWidthsExcel, availableWidth, minWidth = 1300) {
  const count = colWidthsExcel.length;
  if (count === 0) return [];

  const minTotal = minWidth * count;
  if (minTotal >= availableWidth) {
    const width = Math.max(800, Math.floor(availableWidth / count));
    const widths = new Array(count).fill(width);
    widths[widths.length - 1] += availableWidth - widths.reduce((sum, item) => sum + item, 0);
    return widths;
  }

  const weights = colWidthsExcel.map(width => Math.max(1, Number(width) || 1));
  const weightTotal = weights.reduce((sum, width) => sum + width, 0) || 1;
  const extra = availableWidth - minTotal;
  const widths = weights.map(weight => Math.floor(minWidth + extra * (weight / weightTotal)));
  widths[widths.length - 1] += availableWidth - widths.reduce((sum, width) => sum + width, 0);
  return widths;
}

function computeHwpColWidths(data, hasTitleRow, availableWidth, headerRowCount = getHeaderRowCount(data, hasTitleRow)) {
  if (data.format === 'curriculumOrganizationLandscape') {
    return computeProportionalColWidths(data.colWidthsExcel, availableWidth, 1250);
  }

  if (data.format === 'curriculumOrganization') {
    return computeCurriculumOrganizationColWidths(data.colCount, availableWidth);
  }

  const headers = getHeaderTextsByColumn(data, hasTitleRow, headerRowCount);
  const specByRole = {
    category: { min: 7600, weight: 1.05 },
    selectionGroup: { min: 9200, weight: 1.25 },
    group:    { min: 8200, weight: 1.15 },
    type:     { min: 8200, weight: 1.05 },
    subject:  { min: 13000, weight: 1.75 },
    credit:   { min: 2600, weight: 0.03 },
    term:     { min: 2700, weight: 0.03 },
    default:  { min: 6500, weight: 0.9 },
  };

  const cols = headers.map(header => {
    const role = getColumnRole(header);
    return { header, role, ...specByRole[role] };
  });
  const minTotal = cols.reduce((sum, col) => sum + col.min, 0);
  if (minTotal >= availableWidth) {
    const scale = availableWidth / minTotal;
    return cols.map(col => Math.max(2500, Math.floor(col.min * scale)));
  }

  const weightTotal = cols.reduce((sum, col) => sum + col.weight, 0) || 1;
  const extra = availableWidth - minTotal;
  const widths = cols.map(col => Math.floor(col.min + extra * (col.weight / weightTotal)));
  widths[widths.length - 1] += availableWidth - widths.reduce((sum, width) => sum + width, 0);
  return widths;
}

function isZeroCreditText(text) {
  return /^0(?:\.0+)?$/.test(String(text).trim().replace(/,/g, ''));
}

function parseSelectionCreditLayout(text) {
  const match = String(text ?? '').match(/^\s*(택\s*\d+)\s+(\d+(?:\.\d+)?)\s*학점\s*$/);
  if (!match) return null;
  const choice = match[1].replace(/\s+/g, '');
  const credit = match[2];
  const unit = '학점';
  const paragraphs = [choice, credit, unit];
  return {
    text: paragraphs.join(''),
    paragraphs,
  };
}

function splitCellTextParagraphs(text) {
  return normalizeExcelCellText(text)
    .split('\n')
    .map(line => line.trim())
    .filter(Boolean);
}

function buildDisplayTextByCell(data, hasTitleRow, headerRowCount = getHeaderRowCount(data, hasTitleRow)) {
  const lastHeaderRow = Math.max(0, headerRowCount - 1);
  const roles = getHeaderTextsByColumn(data, hasTitleRow, headerRowCount).map(getColumnRole);
  const textByCell = new Map();
  const selectionCreditByCell = new Map();
  const paragraphLayoutByCell = new Map();
  let omittedZeroCount = 0;

  for (const cell of data.cells) {
    const key = `${cell.row},${cell.col}`;
    const selectionCredit = parseSelectionCreditLayout(cell.text);
    let text = selectionCredit ? selectionCredit.text : cell.text;
    const role = roles[cell.col];
    if (!data.preserveZeroText && cell.row > lastHeaderRow && (role === 'credit' || role === 'term') && isZeroCreditText(text)) {
      text = '';
      omittedZeroCount++;
    }
    const paragraphs = selectionCredit ? selectionCredit.paragraphs : splitCellTextParagraphs(text);
    text = paragraphs.join('\n');
    if (selectionCredit && text) {
      selectionCreditByCell.set(key, selectionCredit);
    } else if (paragraphs.length > 1) {
      paragraphLayoutByCell.set(key, { paragraphs });
    }
    textByCell.set(key, text);
  }

  return { textByCell, selectionCreditByCell, paragraphLayoutByCell, omittedZeroCount };
}

function createDocumentTitleStyle(doc, data = null) {
  const isLandscapeCurriculum = data?.format === 'curriculumOrganizationLandscape';
  const fontSize = isLandscapeCurriculum ? 700 : 1500;
  const lineSpacing = isLandscapeCurriculum ? 100 : 130;
  const styleId = doc.createStyle(JSON.stringify({
    name: 'Document Title',
    englishName: 'DocumentTitle',
    type: 0,
    nextStyleId: 0,
  }));
  doc.updateStyleShapes(
    styleId,
    JSON.stringify({
      bold: true,
      fontSize,
      fontFamily: TABLE_FONT_FAMILY,
      textColor: TABLE_THEME.text,
    }),
    JSON.stringify({ alignment: 'center', lineSpacing, spacingBefore: 0, spacingAfter: 0 })
  );
  return styleId;
}

function insertDocumentTitle(doc, titleText, data = null) {
  const isLandscapeCurriculum = data?.format === 'curriculumOrganizationLandscape';
  let titleLines = String(titleText ?? '')
    .split(/\r\n|\r|\n/)
    .map(line => line.trim())
    .filter(Boolean);
  if (isLandscapeCurriculum && titleLines.length > 1) {
    titleLines = [titleLines.join(' ')];
  }
  if (titleLines.length === 0) return 0;

  const fontSize = isLandscapeCurriculum ? 700 : 1500;
  const lineSpacing = isLandscapeCurriculum ? 100 : 130;
  const titleStyleId = createDocumentTitleStyle(doc, data);
  for (let i = 0; i < titleLines.length; i++) {
    const title = titleLines[i];
    parseJsonResult(doc.insertText(0, i, 0, title), `insertDocumentTitle(${i})`);
    parseJsonResult(doc.applyStyle(0, i, titleStyleId), `applyDocumentTitleStyle(${i})`);
    parseJsonResult(doc.applyParaFormat(0, i, JSON.stringify({
      alignment: 'center',
      lineSpacing,
      spacingBefore: 0,
      spacingAfter: 0,
    })), `applyDocumentTitlePara(${i})`);
    try {
      parseJsonResult(doc.applyCharFormat(0, i, 0, title.length, JSON.stringify({
        bold: true,
        fontSize,
        fontFamily: TABLE_FONT_FAMILY,
        textColor: TABLE_THEME.text,
      })), `applyDocumentTitleChar(${i})`);
    } catch (e) {
      console.warn('[debug] 제목 글자 서식 일부 적용 실패:', e);
    }
    parseJsonResult(doc.splitParagraph(0, i, title.length), `splitParagraphAfterTitle(${i})`);
  }
  return titleLines.length;
}

function insertCellParagraphs(doc, paraIdx, ctrlIdx, cellIdx, paragraphs, label) {
  if (!paragraphs || paragraphs.length === 0) return;
  parseJsonResult(
    doc.insertTextInCell(0, paraIdx, ctrlIdx, cellIdx, 0, 0, paragraphs[0]),
    `${label}:insertCellParagraph(0)`
  );
  for (let i = 1; i < paragraphs.length; i++) {
    parseJsonResult(
      doc.splitParagraphInCell(0, paraIdx, ctrlIdx, cellIdx, i - 1, paragraphs[i - 1].length),
      `${label}:splitParagraph(${i})`
    );
    parseJsonResult(
      doc.insertTextInCell(0, paraIdx, ctrlIdx, cellIdx, i, 0, paragraphs[i]),
      `${label}:insertCellParagraph(${i})`
    );
  }
}

function createTableTextStyles(doc, data) {
  const isCurriculum = data?.format === 'curriculumOrganization' || data?.format === 'curriculumOrganizationLandscape';
  const isLandscapeCurriculum = data?.format === 'curriculumOrganizationLandscape';
  const defs = {
    title:  { name: 'Minimal Title',  char: { bold: true, fontSize: isLandscapeCurriculum ? 700 : 1080, fontFamily: TABLE_FONT_FAMILY, textColor: TABLE_THEME.text } },
    header: { name: 'Minimal Header', char: { bold: true, fontSize: isLandscapeCurriculum ? 620 : isCurriculum ? 760 : 930, fontFamily: TABLE_FONT_FAMILY, textColor: TABLE_THEME.text } },
    index:  { name: 'Minimal Index',  char: { bold: true, fontSize: isLandscapeCurriculum ? 550 : isCurriculum ? 820 : 880, fontFamily: TABLE_FONT_FAMILY, textColor: TABLE_THEME.indexText } },
    body:   { name: 'Minimal Body',   char: { bold: false, fontSize: isLandscapeCurriculum ? 550 : isCurriculum ? 820 : 860, fontFamily: TABLE_FONT_FAMILY, textColor: TABLE_THEME.text } },
    selectionCredit: { name: 'Selection Credit', char: { bold: true, fontSize: isLandscapeCurriculum ? 550 : isCurriculum ? 860 : 900, fontFamily: TABLE_FONT_FAMILY, textColor: TABLE_THEME.text } },
  };

  const styleIds = {};
  for (const [kind, def] of Object.entries(defs)) {
    const styleId = doc.createStyle(JSON.stringify({
      name: def.name,
      englishName: def.name.replace(/\s+/g, ''),
      type: 0,
      nextStyleId: 0,
    }));
    doc.updateStyleShapes(styleId, JSON.stringify(def.char), '{}');
    styleIds[kind] = styleId;
  }
  return styleIds;
}

// ────────────────────────────────────────────────────────────────
// 본 변환: 감지한 양식별 HWP 생성 + 미리보기
// ────────────────────────────────────────────────────────────────
function logTableDataInfo(data) {
  log(`시트="${data.sheetName}"  원본 ${data.rawRowCount}×${data.rawColCount} → 표 ${data.rowCount}×${data.colCount}  값셀=${data.cells.length}  병합=${data.merges.length}`);
  if (data.formatLabel) {
    log(`${data.formatLabel} 양식 감지: 제목 ${data.titleCell}, 표 ${data.tableRangeLabel}`, 'ok');
  } else if (data.trim && (data.trim.r1 > 0 || data.trim.c1 > 0)) {
    log(`앞쪽 빈 행/열 제거: 시작 셀 ${excelColName(data.trim.c1)}${data.trim.r1 + 1}`, 'muted');
  }
  if (data.droppedColIdx && data.droppedColIdx.length > 0) {
    const baseCol = data.trim ? data.trim.c1 : 0;
    const labels = data.droppedColIdx.map(c => `${excelColName(baseCol + c)}(${baseCol + c})`).join(', ');
    log(`자동 제거된 빈 데이터 열: ${labels}`, 'muted');
  }
  if (data.selectionRequirement?.movedCount > 0) {
    const dropped = data.selectionRequirement.droppedColIdx.length > 0
      ? ` / 원래 안내 열 ${data.selectionRequirement.droppedColIdx.length}개 제거`
      : '';
    log(`선택그룹명으로 학기별 선택조건 이동: ${data.selectionRequirement.movedCount}건${dropped}`, 'muted');
  } else if (data.selectionRequirement?.skippedCount > 0) {
    log('학기별 선택조건 문구를 찾았지만 선택그룹명 열을 찾지 못해 이동하지 않았습니다.', 'muted');
  }
  if (data.excelStyleByCell?.size > 0) {
    log(`엑셀 셀 색상/글자색 감지: ${data.excelStyleByCell.size}칸`, 'muted');
  }
  log(`엑셀 열폭: [${data.colWidthsExcel.map(w => w.toFixed(1)).join(', ')}]`, 'muted');
}

function sanitizeFileNamePart(value) {
  return String(value ?? '변환')
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '') || '변환';
}

function outputFileNameForData(sourceName, data, index, total) {
  const base = sourceName.replace(/\.(xlsx|xls|xlsm)$/i, '');
  if (total === 1 && !data.formatLabel) return `${base}.hwp`;

  const suffix = data.format === 'curriculumOrganizationLandscape'
    ? '교육과정편제표_가로'
    : data.format === 'curriculumOrganization'
    ? '교육과정편제표'
    : sanitizeFileNamePart(data.sheetName || `양식_${index + 1}`);
  return `${base}_${suffix}.hwp`;
}

function renderConversionResults(results) {
  if (!$results) return;
  $results.innerHTML = '';
  $results.hidden = results.length === 0;
  $dlSlot.innerHTML = results.length > 0
    ? `<span class="muted">아래 미리보기에서 양식별로 다운로드하세요.</span>`
    : '';

  for (const result of results) {
    const { data, bytes, outName } = result;
    const card = document.createElement('section');
    card.className = 'result-card';

    const head = document.createElement('div');
    head.className = 'result-head';

    const title = document.createElement('div');
    title.className = 'result-title';
    const strong = document.createElement('strong');
    strong.textContent = data.formatLabel || data.sheetName || '변환 결과';
    const meta = document.createElement('span');
    meta.className = 'muted';
    meta.textContent = `${data.sheetName} · ${data.rowCount}행 × ${data.colCount}열 · ${(bytes.length / 1024).toFixed(1)}KB`;
    title.append(strong, meta);
    head.append(title, makeDownloadLink(bytes, outName));

    const preview = document.createElement('div');
    preview.className = 'preview-wrap';
    const previewWidth = data.format === 'curriculumOrganizationLandscape' ? 1800 : 980;
    preview.innerHTML = buildStyledTableHtml(data, previewWidth).html;

    card.append(head, preview);
    $results.appendChild(card);
  }
}

async function buildHwpForTableData(data) {
  logTableDataInfo(data);

  const doc = makeBlankDoc();
  try {
    log('빈 문서 골격 준비 OK');
    let tableFontFaceId = -1;
    try {
      tableFontFaceId = doc.findOrCreateFontId(TABLE_FONT_FAMILY);
      log(`문서 글꼴 등록: ${TABLE_FONT_FAMILY} (fontId=${tableFontFaceId})`, 'muted');
    } catch (e) {
      log(`문서 글꼴 등록 실패: ${errMsg(e)}`, 'err');
    }

    const HWPUNIT_PER_MM = 283;
    const PAGE_SHORT = 59528;
    const PAGE_LONG = 84186;
    const MIN_MARGIN = 10 * HWPUNIT_PER_MM;
    const useLandscape = !!data.forceLandscape || data.colCount >= 8;
    const pageDefWidth = PAGE_SHORT;
    const pageDefHeight = PAGE_LONG;
    const layoutPageWidth = useLandscape ? PAGE_LONG : PAGE_SHORT;
    parseJsonResult(doc.setPageDef(0, JSON.stringify({
      width: pageDefWidth,
      height: pageDefHeight,
      landscape: useLandscape,
      marginLeft:   MIN_MARGIN,
      marginRight:  MIN_MARGIN,
      marginTop:    MIN_MARGIN,
      marginBottom: MIN_MARGIN,
      marginHeader: 0,
      marginFooter: 0,
    })), 'setPageDef');
    log(`용지: ${useLandscape ? '가로' : '세로'} / 여백 사방 10mm`);

    const tableParaIdx = insertDocumentTitle(doc, data.titleText, data);
    if (data.titleText) {
      log(`문서 제목 삽입: "${data.titleText}"`, 'muted');
    }

    const r = parseJsonResult(
      doc.createTable(0, tableParaIdx, 0, data.rowCount, data.colCount),
      'createTable'
    );
    const paraIdx = r.paraIdx ?? 0;
    const ctrlIdx = r.controlIdx ?? 0;
    const hasTitleRow = data.hasTitleRow ?? detectTitleRow(data);
    const headerRowCount = getHeaderRowCount(data, hasTitleRow);
    const pageContentWidth = layoutPageWidth - (MIN_MARGIN * 2);
    const portraitSafeWidth = PAGE_SHORT - (MIN_MARGIN * 2) - 1200;
    const tableWidth = data.format === 'curriculumOrganizationLandscape'
      ? Math.max(52000, pageContentWidth - 500)
      : data.format === 'curriculumOrganization'
      ? Math.max(36000, Math.min(pageContentWidth - 2600, portraitSafeWidth - 500))
      : Math.max(24000, Math.min(pageContentWidth - 1200, portraitSafeWidth));
    const targetColWidths = computeHwpColWidths(data, hasTitleRow, tableWidth, headerRowCount);
    const { textByCell, selectionCreditByCell, paragraphLayoutByCell, omittedZeroCount } = buildDisplayTextByCell(data, hasTitleRow, headerRowCount);
    log(`createTable(${data.rowCount}×${data.colCount}) OK  paraIdx=${paraIdx} ctrlIdx=${ctrlIdx}`);
    log(`열폭 최적화: [${targetColWidths.map(w => (w / HWPUNIT_PER_MM).toFixed(1) + 'mm').join(', ')}]`, 'muted');
    if (omittedZeroCount > 0) {
      log(`0학점 표기 생략: ${omittedZeroCount}칸`, 'muted');
    }
    if (hasTitleRow) {
      log('상단 병합 제목 행 감지: 2행을 헤더 색상으로 처리', 'muted');
    } else if (headerRowCount > 1) {
      log(`다중 헤더 감지: 상단 ${headerRowCount}행을 헤더 색상으로 처리`, 'muted');
    }

    parseJsonResult(doc.setTableProperties(0, paraIdx, ctrlIdx, JSON.stringify({
      pageBreak: 2,
      repeatHeader: true,
    })), 'setTableProperties');
    log('표 속성: pageBreak=RowBreak, repeatHeader=true');

    let textFilled = 0, textErrors = 0;
    for (const { row, col } of data.cells) {
      const cellIdx = row * data.colCount + col;
      const key = `${row},${col}`;
      const text = textByCell.get(key) || '';
      if (text.length === 0) continue;
      try {
        const selectionCredit = selectionCreditByCell.get(key);
        const paragraphLayout = selectionCredit || paragraphLayoutByCell.get(key);
        if (paragraphLayout) {
          insertCellParagraphs(doc, paraIdx, ctrlIdx, cellIdx, paragraphLayout.paragraphs, `insertCellParagraphs(r=${row},c=${col})`);
        } else {
          parseJsonResult(
            doc.insertTextInCell(0, paraIdx, ctrlIdx, cellIdx, 0, 0, text),
            `insertTextInCell(r=${row},c=${col})`
          );
        }
        textFilled++;
      } catch (e) {
        textErrors++;
        if (textErrors <= 5) log(`  · ${errMsg(e)}`, 'err');
      }
    }
    log(`텍스트 삽입: 성공 ${textFilled} / 실패 ${textErrors}`);

    const isLandscapeCurriculum = data.format === 'curriculumOrganizationLandscape';
    const styleIdsByKind = createTableTextStyles(doc, data);
    log(`텍스트 스타일: ${isLandscapeCurriculum ? '가로 편제표 본문 5.5pt' : '본문 절제형'} 적용`, 'muted');

    const centerParaProps  = JSON.stringify({
      alignment: 'center',
      lineSpacing: isLandscapeCurriculum ? 86 : 105,
      spacingBefore: 0,
      spacingAfter: 0,
    });
    const headerParaProps  = JSON.stringify({
      alignment: 'center',
      lineSpacing: isLandscapeCurriculum ? 88 : 100,
      spacingBefore: 0,
      spacingAfter: 0,
    });
    const fillByBorderFillId = new Map();

    parseJsonResult(doc.beginBatch(), 'beginBatch(format)');
    let fmtStyleOk = 0, fmtParaOk = 0, fmtErr = 0;
    for (let row = 0; row < data.rowCount; row++) {
      for (let col = 0; col < data.colCount; col++) {
        const cellIdx = row * data.colCount + col;
        const key = `${row},${col}`;
        const text = textByCell.get(key) || '';
        const selectionCredit = selectionCreditByCell.get(key);
        const paragraphLayout = selectionCredit || paragraphLayoutByCell.get(key);
        const paragraphCount = paragraphLayout?.paragraphs?.length || 1;
        const kind = getTableCellKind(row, col, hasTitleRow, headerRowCount);
        try {
          try {
            const beforeCellProps = JSON.parse(doc.getCellProperties(0, paraIdx, ctrlIdx, cellIdx));
            const fillColor = getCellFillColor(kind, row, col, data, text);
            const borderColor = getCellBorderColor(kind, fillColor);
            const borderProps = getCellBorderProps(data, key, kind, fillColor, borderColor);
            parseJsonResult(
              doc.setCellProperties(0, paraIdx, ctrlIdx, cellIdx, JSON.stringify({
                width: targetColWidths[col] || beforeCellProps.width,
                height: isLandscapeCurriculum
                  ? getLandscapeCellHeight(kind, paragraphCount)
                  : Math.max(beforeCellProps.height || 0, kind === 'header' ? 1750 : 1280),
                isHeader: kind === 'title' || kind === 'header',
                verticalAlign: 1,
                paddingLeft: isLandscapeCurriculum ? 35 : kind === 'header' ? 120 : 220,
                paddingRight: isLandscapeCurriculum ? 35 : kind === 'header' ? 120 : 220,
                paddingTop: isLandscapeCurriculum ? 20 : kind === 'header' ? 80 : 140,
                paddingBottom: isLandscapeCurriculum ? 20 : kind === 'header' ? 80 : 140,
                ...borderProps,
              })),
              `setCellProperties(r=${row},c=${col})`
            );
            const afterCellProps = JSON.parse(doc.getCellProperties(0, paraIdx, ctrlIdx, cellIdx));
            const previousFill = fillByBorderFillId.get(afterCellProps.borderFillId);
            if (previousFill && previousFill !== fillColor) {
              console.warn('[debug] BorderFill 색상 충돌:', afterCellProps.borderFillId, previousFill, fillColor);
            }
            fillByBorderFillId.set(afterCellProps.borderFillId, fillColor);
          } catch (_) {}
          const paragraphTexts = paragraphLayout ? paragraphLayout.paragraphs : [text];
          for (let cellParaIdx = 0; cellParaIdx < paragraphTexts.length; cellParaIdx++) {
            const paragraphText = paragraphTexts[cellParaIdx] || '';
            if (paragraphText.length > 0) {
              const styleKey = selectionCredit && cellParaIdx === 1 ? 'selectionCredit' : kind;
              parseJsonResult(
                doc.applyCellStyle(0, paraIdx, ctrlIdx, cellIdx, cellParaIdx,
                  styleIdsByKind[styleKey] ?? styleIdsByKind.body),
                `applyCellStyle(r=${row},c=${col},p=${cellParaIdx})`
              );
              const excelTextFormat = getExcelTextFormat(data, row, col);
              if (hasTextFormatProps(excelTextFormat)) {
                parseJsonResult(
                  doc.applyCharFormatInCell(0, paraIdx, ctrlIdx, cellIdx, cellParaIdx, 0, paragraphText.length, JSON.stringify(excelTextFormat)),
                  `applyExcelTextFormat(r=${row},c=${col},p=${cellParaIdx})`
                );
              }
              fmtStyleOk++;
            }
            const paraProps = kind === 'header' ? headerParaProps : centerParaProps;
            parseJsonResult(
              doc.applyParaFormatInCell(0, paraIdx, ctrlIdx, cellIdx, cellParaIdx, paraProps),
              `applyParaFormatInCell(center,r=${row},c=${col},p=${cellParaIdx})`
            );
            fmtParaOk++;
          }
        } catch (e) {
          fmtErr++;
          if (fmtErr <= 5) log(`  · ${errMsg(e)}`, 'err');
        }
      }
    }
    parseJsonResult(doc.endBatch(), 'endBatch(format)');
    log(`서식 적용: 스타일 ${fmtStyleOk} / 문단 ${fmtParaOk} / 실패 ${fmtErr}`);
    log(`색상 처리: 셀 배경 ${fillByBorderFillId.size}개 추적${data.excelStyleByCell?.size ? ' / Excel 색상 우선 적용' : ''}`, 'muted');

    const sortedMerges = [...data.merges].sort((a, b) => {
      const areaA = (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1);
      const areaB = (b.r2 - b.r1 + 1) * (b.c2 - b.c1 + 1);
      return areaB - areaA;
    });

    let mergeOk = 0, mergeErr = 0;
    for (const m of sortedMerges) {
      try {
        parseJsonResult(
          doc.mergeTableCells(0, paraIdx, ctrlIdx, m.r1, m.c1, m.r2, m.c2),
          `mergeTableCells(${m.r1},${m.c1}-${m.r2},${m.c2})`
        );
        mergeOk++;
      } catch (e) {
        mergeErr++;
        if (mergeErr <= 5) log(`  · ${errMsg(e)}`, 'err');
      }
    }
    log(`병합 적용: 성공 ${mergeOk} / 실패 ${mergeErr}`);

    let bytes = doc.exportHwp();
    try {
      bytes = await applyHwpBorderFillColors(bytes, fillByBorderFillId);
      log(`색상 후처리: BorderFill ${fillByBorderFillId.size}개 반영`, 'ok');
    } catch (e) {
      log(`색상 후처리 실패: ${errMsg(e)}`, 'err');
    }
    try {
      const fontResult = await applyHwpCharShapeFont(bytes, tableFontFaceId);
      bytes = fontResult.bytes;
      log(`글꼴 후처리: CharShape ${fontResult.patched}개를 ${TABLE_FONT_FAMILY}(으)로 보정`, 'ok');
    } catch (e) {
      log(`글꼴 후처리 실패: ${errMsg(e)}`, 'err');
    }
    log(`exportHwp() OK  ${bytes.length} bytes`, 'ok');
    return bytes;
  } finally {
    doc.free();
  }
}

async function convertSelectedFile() {
  try {
    await ensureWasm();
    const f = $file.files[0];
    if (!f) { log('파일 미선택', 'err'); return; }
    $dlSlot.innerHTML = '';
    resetResults();
    log(`--- 변환 시작: ${f.name} ---`);

    const buf = await f.arrayBuffer();
    const tables = readXlsxTables(buf);
    log(`생성 대상 양식: ${tables.length}개`, tables.length > 1 ? 'ok' : 'muted');

    const results = [];
    for (let i = 0; i < tables.length; i++) {
      const data = tables[i];
      log(`--- [${i + 1}/${tables.length}] ${data.formatLabel || data.sheetName} ---`);
      const bytes = await buildHwpForTableData(data);
      const outName = outputFileNameForData(f.name, data, i, tables.length);
      results.push({ data, bytes, outName });
      log(`결과 준비: ${outName}`, 'ok');
    }

    renderConversionResults(results);
    log(`변환 완료: HWP ${results.length}개`, 'ok');
  } catch (e) {
    log(`[변환 실패] ${errMsg(e)}`, 'err');
    console.error(e);
  }
}

// ────────────────────────────────────────────────────────────────
// 이벤트 바인딩
// ────────────────────────────────────────────────────────────────
$file.addEventListener('change', () => {
  $btnConvert.disabled = !$file.files[0];
  $dlSlot.innerHTML = '';
  resetResults();
});
$btnSelftest.addEventListener('click', selftest);
$btnConvert.addEventListener('click', convertSelectedFile);

log('준비됨. 셀프 테스트로 동작 검증 후 xlsx/xls/xlsm 파일을 선택하세요.', 'muted');
