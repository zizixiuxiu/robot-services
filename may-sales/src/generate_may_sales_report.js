const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// ==================== 配置 ====================
// 默认路径（本地调试用），Docker 中由调用方通过命令行参数传入
const DEFAULT_BASE_DIR = process.env.DEFAULT_BASE_DIR || 'D:/wechat/xwechat_files/wxid_0fh4oxng8dq212_f810/msg/file/2026-06';
const DEFAULT_ZHCX = path.join(DEFAULT_BASE_DIR, '综合查询5月(1).xls');
const DEFAULT_LIANSI = path.join(DEFAULT_BASE_DIR, '联思系统5月(1).xlsx');
const DEFAULT_SHEJIANG = path.join(DEFAULT_BASE_DIR, '奢匠下单统计5月.xlsx');
const DEFAULT_TEMPLATE = process.env.DEFAULT_TEMPLATE || 'D:/wechat/xwechat_files/wxid_0fh4oxng8dq212_f810/msg/file/2026-06/2026年5月销售部业绩核对表 - 副本.xlsx';
const DEFAULT_OUTPUT = path.join(DEFAULT_BASE_DIR, '2026年5月销售部业绩核对表.xlsx');

// 支持命令行参数：node generate_may_sales_report.js [zhcx] [liansi] [shejiang] [template] [output]
const args = process.argv.slice(2);
const SRC_ZHCX = args[0] || DEFAULT_ZHCX;
const SRC_LIANSI = args[1] || DEFAULT_LIANSI;
const SRC_SHEJIANG = args[2] || DEFAULT_SHEJIANG;
const TEMPLATE_FILE = args[3] || DEFAULT_TEMPLATE;
const OUTPUT_FILE = args[4] || DEFAULT_OUTPUT;
// =============================================

// 省份全称到简称的映射
const PROVINCE_SHORT_MAP = {
  '北京市': '北京', '天津市': '天津', '上海市': '上海', '重庆市': '重庆',
  '河北省': '河北', '山西省': '山西', '辽宁省': '辽宁', '吉林省': '吉林',
  '黑龙江省': '黑龙江', '江苏省': '江苏', '浙江省': '浙江', '安徽省': '安徽',
  '福建省': '福建', '江西省': '江西', '山东省': '山东', '河南省': '河南',
  '湖北省': '湖北', '湖南省': '湖南', '广东省': '广东', '海南省': '海南',
  '四川省': '四川', '贵州省': '贵州', '云南省': '云南', '陕西省': '陕西',
  '甘肃省': '甘肃', '青海省': '青海', '台湾省': '台湾',
  '内蒙古自治区': '内蒙古', '广西壮族自治区': '广西', '西藏自治区': '西藏',
  '宁夏回族自治区': '宁夏', '新疆维吾尔自治区': '新疆',
  '香港特别行政区': '香港', '澳门特别行政区': '澳门'
};

function normalizeProvince(province) {
  if (!province) return '';
  const s = String(province).trim();
  // 兼容模板里写的"内蒙"
  if (s === '内蒙') return '内蒙古';
  return PROVINCE_SHORT_MAP[s] || s;
}

function cleanDealer(name) {
  if (!name) return null;
  let s = String(name).trim();
  if (s.includes('.')) s = s.split('.')[0];
  if (['重庆直营店', '直营店.渝北', '直营店'].includes(s)) return '直营店';
  if (['国际电商部', '国际电商'].includes(s)) return '国际电商';
  // 去掉括号及内容，如 张鹏飞（奢匠）、郭亮(陕)
  // 省份区分交给 normalizeProvince + 经销商+省份 匹配键
  s = s.replace(/[（(].*?[)）]/g, '').trim();
  return s;
}

function extractYearMonthFromProdNo(prodNo) {
  if (!prodNo) return [null, null];
  const parts = String(prodNo).split('-');
  if (parts.length < 3) return [null, null];
  const first = parts[0];
  let yr = '';
  for (let i = first.length - 1; i >= 0; i--) {
    const c = first[i];
    if (/\d/.test(c)) yr = c + yr; else break;
  }
  if (yr.length < 2) return [null, null];
  const year = parseInt(yr.slice(-2), 10) + 2000;
  const month = parseInt(parts[1], 10);
  if (isNaN(month) || month < 1 || month > 12) return [year, null];
  return [year, month];
}

function readSheetRows(filePath, sheetName) {
  const wb = XLSX.readFile(filePath, { cellStyles: false, cellNF: false });
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet "${sheetName}" not found in ${filePath}`);
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const rows = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({r, c})];
      row.push(cell ? cell.v : undefined);
    }
    rows.push(row);
  }
  return rows;
}

function processShejiang(rows) {
  const results = [];
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    const month = row[1];
    const dealer = cleanDealer(row[5]);
    const province = normalizeProvince(row[3]);
    const city = row[4];
    const ot = row[6];
    const total = Number(row[61] || 0);
    const actual = Number(row[63] || 0);
    if (!month || !dealer) continue;
    const otMap = { '订单': '订单', '补单': '订单', '样品': '样品', '小样': '样品' };
    if (!otMap[ot]) continue;
    results.push({
      month: String(month).trim(),
      dealer,
      province,
      city,
      orderType: otMap[ot],
      total,
      actual,
      source: '线下'
    });
  }
  return results;
}

function processLiansi(rows) {
  const results = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const dealer = cleanDealer(row[2]);
    const province = normalizeProvince(row[3]);
    const city = row[4];
    const ot = row[19];
    const total = Number(row[7] || 0);
    const actual = Number(row[13] || 0);
    const cust = row[6];
    if (!dealer) continue;
    if (cust && String(cust).trim() === '色板') continue;
    const otMap = {
      '补单': '订单',
      '纯板式订单': '订单',
      '纯实木订单': '订单',
      '板式含油漆订单': '订单',
      '店样品单': '样品',
      '售后单': null
    };
    if (!otMap[ot]) continue;
    results.push({
      month: '5月',
      dealer,
      province,
      city,
      orderType: otMap[ot],
      total,
      actual,
      source: '联思'
    });
  }
  return results;
}

function processZhcx(rows, existingProdNos) {
  const results = [];
  for (let i = 2; i < rows.length; i++) {
    const row = rows[i];
    const prodNo = row[2];
    const dealer = cleanDealer(row[4]);
    const province = normalizeProvince(row[3]);
    const city = '';
    const cust = row[6];
    const sample = row[10];
    const prodType = row[11];
    const maker = row[13];
    const total = Number(row[24] || 0);
    const actual = Number(row[25] || 0);
    if (!prodNo || !dealer) continue;
    const prodNoStr = String(prodNo).trim();
    if (existingProdNos && existingProdNos.has(prodNoStr)) continue;
    if (['杨益琴', '陈明会'].includes(maker)) continue;
    if (prodType === '售后生产') continue;
    if (cust) {
      const cstr = String(cust).trim();
      if (cstr.includes('色卡') || cstr === '打色板') continue;
    }
    const [year, month] = extractYearMonthFromProdNo(prodNo);
    if (!year || year !== 2026 || !month) continue;
    const monthStr = `${month}月`;
    const ot = (sample === '是' || String(sample).includes('是')) ? '样品' : '订单';
    results.push({
      month: monthStr,
      dealer,
      province,
      city,
      orderType: ot,
      total,
      actual,
      source: '订单通'
    });
  }
  return results;
}

function aggregateByDealer(records) {
  const agg = {};
  for (const r of records) {
    const key = `${r.dealer}|${r.province}`;
    if (!agg[key]) {
      agg[key] = {
        dealer: r.dealer,
        province: r.province,
        正单: 0,
        样品: 0,
        业绩合计: 0,
        样品折后金额: 0,
        隐迹系列销量: 0
      };
    }
    if (r.orderType === '订单') {
      agg[key].正单 += r.total;
    } else if (r.orderType === '样品') {
      agg[key].样品 += r.total;
      agg[key].样品折后金额 += r.actual;
    }
  }
  for (const key of Object.keys(agg)) {
    agg[key].业绩合计 = agg[key].正单 + agg[key].样品;
  }
  return agg;
}

async function main() {
  console.log('读取数据源...');
  const shejiangRows = readSheetRows(SRC_SHEJIANG, '奢匠下单表');
  const liansiRows = readSheetRows(SRC_LIANSI, 'Sheet');
  const zhcxRows = readSheetRows(SRC_ZHCX, '综合查询');

  console.log('处理数据...');
  const existingProdNos = new Set();
  for (let i = 3; i < shejiangRows.length; i++) {
    const no = shejiangRows[i][2];
    if (no) existingProdNos.add(String(no).trim());
  }
  for (let i = 1; i < liansiRows.length; i++) {
    const no = liansiRows[i][1];
    if (no) existingProdNos.add(String(no).trim());
  }

  const records = [];
  records.push(...processShejiang(shejiangRows));
  records.push(...processLiansi(liansiRows));
  records.push(...processZhcx(zhcxRows, existingProdNos));

  console.log(`共提取 ${records.length} 条记录`);
  const agg = aggregateByDealer(records);

  const sourceSummary = {};
  for (const r of records) {
    if (!sourceSummary[r.source]) sourceSummary[r.source] = { count: 0, amount: 0 };
    sourceSummary[r.source].count++;
    sourceSummary[r.source].amount += r.actual;
  }
  console.log('按来源汇总:');
  for (const [src, s] of Object.entries(sourceSummary)) {
    console.log(`  ${src}: ${s.count}条 / ${s.amount.toFixed(2)}元`);
  }

  console.log('读取模板（ExcelJS，保留格式）...');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(TEMPLATE_FILE);

  const sourceSheet = workbook.getWorksheet('Sheet1');
  const targetSheet = workbook.getWorksheet('Sheet1 (2)');
  if (!sourceSheet) throw new Error('模板中找不到 Sheet "Sheet1"');
  if (!targetSheet) throw new Error('模板中找不到 Sheet "Sheet1 (2)"');

  // 读取模板经销商列表，使用 经销商+省份 作为匹配键
  const templateDealers = {};
  sourceSheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber < 3) return;
    const seq = row.getCell(1).value;
    const region = row.getCell(2).value;
    const province = normalizeProvince(row.getCell(3).value);
    const city = row.getCell(4).value;
    const dealer = row.getCell(5).value;
    if (dealer && dealer !== '合计') {
      const key = `${dealer}|${province}`;
      templateDealers[key] = { row: rowNumber, seq, region, province, city, dealer };
    }
  });
  console.log(`模板中共有 ${Object.keys(templateDealers).length} 个经销商（按名称+省份去重）`);

  // 按名称+省份匹配并填充
  let filledCount = 0;
  const unmatched = [];
  for (const [key, data] of Object.entries(agg)) {
    const tpl = templateDealers[key];
    if (!tpl) {
      unmatched.push(data);
      continue;
    }

    const isNew = data.正单 === 0 && data.样品 > 0;
    const row = targetSheet.getRow(tpl.row);

    row.getCell(1).value = tpl.seq;
    row.getCell(2).value = tpl.region;
    row.getCell(3).value = tpl.province;
    row.getCell(4).value = tpl.city;
    row.getCell(5).value = tpl.dealer;
    row.getCell(6).value = round2(data.正单);
    row.getCell(7).value = round2(data.样品);
    row.getCell(8).value = round2(data.业绩合计);
    row.getCell(9).value = round2(data.样品折后金额);
    row.getCell(10).value = round2(data.隐迹系列销量);
    if (isNew) {
      row.getCell(11).value = '新商';
    }
    filledCount++;
  }

  if (unmatched.length > 0) {
    console.log('\n数据源中有但模板中无的经销商（按名称+省份）：');
    unmatched.forEach(d => {
      console.log(`  ${d.dealer} / ${d.province}：正单=${d.正单.toFixed(2)} 样品=${d.样品.toFixed(2)}`);
    });
  }

  // 合计行：用普通 SUM 公式替换原模板中可能损坏的共享公式
  const totalRow = targetSheet.getRow(116);
  const sums = { 6: 0, 7: 0, 9: 0, 10: 0 };
  for (const key of Object.keys(agg)) {
    const data = agg[key];
    sums[6] += data.正单;
    sums[7] += data.样品;
    sums[9] += data.样品折后金额;
    sums[10] += data.隐迹系列销量;
  }
  sums[8] = sums[6] + sums[7];
  const mainCols = { 6: 'F', 7: 'G', 8: 'H', 9: 'I', 10: 'J' };
  for (const [col, val] of Object.entries(sums)) {
    const c = parseInt(col);
    const letter = mainCols[c];
    let formula;
    if (c === 8) {
      formula = 'F116+G116';
    } else {
      formula = `SUM(${letter}3:${letter}115)`;
    }
    totalRow.getCell(c).value = { formula, result: round2(val) };
  }
  // 核对列 L-P 也重置为普通公式，避免共享公式损坏
  const checkCols = { 12: 'L', 13: 'M', 14: 'N', 15: 'O', 16: 'P' };
  for (const [col, letter] of Object.entries(checkCols)) {
    let formula;
    if (parseInt(col) === 14) {
      formula = 'L116+M116';
    } else {
      formula = `SUM(${letter}3:${letter}115)`;
    }
    totalRow.getCell(parseInt(col)).value = { formula, result: 0 };
  }

  // 保存
  try {
    if (fs.existsSync(OUTPUT_FILE)) {
      fs.chmodSync(OUTPUT_FILE, 0o666);
    }
  } catch (e) {
    console.warn('无法修改目标文件权限:', e.message);
  }
  await workbook.xlsx.writeFile(OUTPUT_FILE);
  console.log(`\n完成：已填充 ${filledCount} 个经销商，保存到 ${OUTPUT_FILE}`);
  if (unmatched.length > 0) {
    console.log(`有 ${unmatched.length} 个经销商未匹配到模板行`);
  }
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

main().catch(err => {
  console.error('错误:', err);
  process.exit(1);
});

