/**
 * 用最新的经销商名单重建「销售部业绩核对表」模板。
 *
 * 为什么要重建：模板里写死了 113 家（2026-05 名单）和合计行 116，
 * 名单一变（新增/改名/退出）就对不上——新商直接落在模板外被丢掉，
 * 合计行还固定引用 3:115，行数一动就错位。
 *
 * 做法：拿老模板当**格式母版**，只换 Sheet1 / Sheet1 (2) 的名单区行数，
 * 顺序按名单文件的顺序，合计行公式按实际行数重算。
 *
 * 用法：node build_template.js <名单xlsx> <名单sheet名> <老模板> <新模板>
 * 名单列：序号/区域/省份/城市/经销商（按表头名定位，不靠列号）
 */
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

const HEADER_ROW = 2;   // 第 1 行标题、第 2 行表头，数据从第 3 行起
const FIRST_DATA_ROW = HEADER_ROW + 1;
const COLS = { seq: 1, region: 2, province: 3, city: 4, dealer: 5 };
const FILL_COLS = [6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16];   // 数据列，模板里清空待填
const SUM_COLS = { 6: 'F', 7: 'G', 9: 'I', 10: 'J', 12: 'L', 13: 'M' };
const TOTAL_COLS = { 8: 'H', 14: 'N' };

function readDealerList(file, sheetName) {
  const wb = XLSX.readFile(file);
  const ws = wb.Sheets[sheetName] || wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error(`名单文件里找不到 sheet: ${sheetName}`);
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: '' });
  const header = (rows[0] || []).map(v => String(v).trim());
  const at = name => {
    const i = header.indexOf(name);
    if (i < 0) throw new Error(`名单表头缺少「${name}」，实际表头：${header.join(',')}`);
    return i;
  };
  const iSeq = header.indexOf('序号');
  const iRegion = at('区域'), iProv = at('省份'), iCity = at('城市'), iDealer = at('经销商');
  const list = [];
  for (const r of rows.slice(1)) {
    const dealer = String(r[iDealer] || '').trim();
    if (!dealer || dealer === '合计') continue;
    list.push([
      iSeq >= 0 && r[iSeq] !== '' ? r[iSeq] : list.length + 1,
      String(r[iRegion] || '').trim(),
      String(r[iProv] || '').trim(),
      String(r[iCity] || '').trim(),
      dealer,
    ]);
  }
  return list;
}

/** 从下往上找「合计」行号 */
function findTotalRow(ws) {
  let found = 0;
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    const v = String(row.getCell(1).value || '').trim();
    if (v === '合计') found = rowNumber;
  });
  if (!found) throw new Error(`sheet ${ws.name} 里找不到「合计」行`);
  return found;
}

async function rebuildSheet(ws, dealers) {
  const oldTotal = findTotalRow(ws);

  // 老模板里全是共享公式，删行/挪行后 ExcelJS 保存会报
  // "Shared Formula master must exist above and or left of clone"，
  // 而且模板只是格式母版（真正的公式由生成脚本重写），这里统一清掉。
  ws.eachRow({ includeEmpty: false }, row => {
    row.eachCell({ includeEmpty: false }, cell => {
      if (cell.type === ExcelJS.ValueType.Formula) cell.value = null;
    });
  });

  const oldCount = oldTotal - FIRST_DATA_ROW;      // 老数据行数
  const newCount = dealers.length;

  if (newCount < oldCount) {
    ws.spliceRows(FIRST_DATA_ROW + newCount, oldCount - newCount);
  } else if (newCount > oldCount) {
    // 复制最后一条数据行拿样式，补在合计行前面
    ws.duplicateRow(oldTotal - 1, newCount - oldCount, true);
  }

  const total = FIRST_DATA_ROW + newCount;

  // 换名单
  for (let i = 0; i < newCount; i++) {
    const row = ws.getRow(FIRST_DATA_ROW + i);
    row.getCell(COLS.seq).value = dealers[i][0];
    row.getCell(COLS.region).value = dealers[i][1];
    row.getCell(COLS.province).value = dealers[i][2];
    row.getCell(COLS.city).value = dealers[i][3];
    row.getCell(COLS.dealer).value = dealers[i][4];
    // 数据列清零（公式同列保留，H/N 由下面重设）
    for (const c of FILL_COLS) {
      if (c === 8 || c === 14) continue;
      const cell = row.getCell(c);
      cell.value = (c === 10 || c === 11) ? null : 0;   // 隐迹销量/新商标记留空
    }
    // 业绩合计 / 核对合计 用行内公式
    row.getCell(8).value = { formula: `SUM(F${FIRST_DATA_ROW + i}:G${FIRST_DATA_ROW + i})`, result: 0 };
    row.getCell(14).value = { formula: `SUM(L${FIRST_DATA_ROW + i}:M${FIRST_DATA_ROW + i})`, result: 0 };
    row.commit && row.commit();
  }

  // 合计行公式按新行数重算
  const tRow = ws.getRow(total);
  for (const [col, letter] of Object.entries(SUM_COLS)) {
    tRow.getCell(parseInt(col, 10)).value = {
      formula: `SUM(${letter}${FIRST_DATA_ROW}:${letter}${total - 1})`,
      result: 0,
    };
  }
  for (const [col, letter] of Object.entries(TOTAL_COLS)) {
    tRow.getCell(parseInt(col, 10)).value = {
      formula: `SUM(${letter === 'H' ? 'F' : 'L'}${total}:${letter === 'H' ? 'G' : 'M'}${total})`,
      result: 0,
    };
  }
  tRow.commit && tRow.commit();

  // 合计行的合并单元格跟着挪
  const merged = [];
  ws.model && ws.model.merges && ws.model.merges.forEach(m => merged.push(m));
  for (const ref of merged) {
    if (ref.startsWith('A') && ref.includes(':')) {
      ws.unMergeCells(ref);
    }
  }
  ws.mergeCells(`A${total}:E${total}`);

  // 删掉多余行后可能残留的空白行尾
  const lastRow = ws.rowCount;
  if (lastRow > total) ws.spliceRows(total + 1, lastRow - total);

  return { oldTotal, total, newCount };
}

async function main() {
  const [listFile, listSheet, oldTpl, outTpl] = process.argv.slice(2);
  if (!listFile || !oldTpl || !outTpl) {
    console.error('用法: node build_template.js <名单xlsx> <名单sheet> <老模板> <新模板>');
    process.exit(2);
  }
  const dealers = readDealerList(listFile, listSheet);
  console.log(`名单共 ${dealers.length} 家：${dealers.slice(0, 3).map(d => d[4]).join('、')} ...`);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(oldTpl);
  for (const name of ['Sheet1', 'Sheet1 (2)']) {
    const ws = wb.getWorksheet(name);
    if (!ws) throw new Error(`模板缺少 sheet: ${name}`);
    const r = await rebuildSheet(ws, dealers);
    console.log(`${name}: 数据行 ${r.oldTotal - 3} → ${r.newCount}，合计行移到第 ${r.total} 行`);
  }

  // 标题里的月份改成名单所在月份（名单 sheet 名叫「10月」时直接取用）
  const m = String(listSheet || '').match(/(\d+)\s*月/);
  if (m) {
    for (const name of ['Sheet1', 'Sheet1 (2)']) {
      const cell = wb.getWorksheet(name).getCell('A1');
      if (typeof cell.value === 'string') {
        cell.value = cell.value.replace(/\d+\s*月/, `${parseInt(m[1], 10)}月`);
      }
    }
  }

  await wb.xlsx.writeFile(outTpl);
  console.log(`已生成新模板: ${outTpl}`);
}

main().catch(err => {
  console.error('错误:', err);
  process.exit(1);
});
