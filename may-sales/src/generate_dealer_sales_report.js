// -*- coding: utf-8 -*-
/**
 * 奢匠各经销商销量情况 月报生成（dealer-sales 服务配套脚本）
 *
 * 流程（移植自 D:/1/dealer_sales_pipeline_monthly.js）：
 *   步骤1：从 奢匠明细表 / 联思系统 / 综合查询 三源提取 → 汇总表_提取结果_X月.xlsx（SheetJS）
 *   步骤2：当月数据写入累积表「销量累计_2026.xlsx」（月度数据 sheet，每月递增持久化）
 *   步骤3：主表「奢匠专卖店26年」1-12 月列全部由累积表驱动填充（ExcelJS 保格式/保公式），
 *          输出附「2026月度数据」sheet → 2026年奢匠各经销商销量情况（X月）.xlsx
 *
 * 用法：
 *   node generate_dealer_sales_report.js <月份> <奢匠文件> <联思文件> <综合查询文件> <模板文件> <输出目录> <累积表文件>
 */
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// ==================== 参数 ====================
const args = process.argv.slice(2);
if (args.length < 7) {
  console.log('用法：node generate_dealer_sales_report.js <月份> <奢匠文件> <联思文件> <综合查询文件> <模板文件> <输出目录> <累积表文件>');
  process.exit(1);
}
const TARGET_MONTH_NUM = parseInt(args[0], 10);
if (isNaN(TARGET_MONTH_NUM) || TARGET_MONTH_NUM < 1 || TARGET_MONTH_NUM > 12) {
  console.error('错误：月份必须是 1-12 的数字');
  process.exit(1);
}
const TARGET_MONTH_STR = `${TARGET_MONTH_NUM}月`;
const SHEJIANG_FILE = args[1];
const LIANSI_FILE = args[2];
const ZHCX_FILE = args[3];
const TEMPLATE_FILE = args[4];
const OUTPUT_DIR = args[5];
const ACC_FILE = args[6];

const OUTPUT_EXTRACT = path.join(OUTPUT_DIR, `汇总表_提取结果_${TARGET_MONTH_NUM}月.xlsx`);
const OUTPUT_MAPPED = path.join(OUTPUT_DIR, `2026年奢匠各经销商销量情况（${TARGET_MONTH_NUM}月）.xlsx`);
// ===================================================

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
  if (s === '内蒙') return '内蒙古';
  return PROVINCE_SHORT_MAP[s] || s;
}

function normalizeCity(city) {
  if (!city) return '';
  let s = String(city).trim().replace(/\s+/g, '');

  // 去掉开头的省份前缀（保留直辖市本身，如"重庆"）
  const provinces = ['河北','山西','辽宁','吉林','黑龙江','江苏','浙江','安徽','福建','江西','山东','河南','湖北','湖南','广东','海南','四川','贵州','云南','陕西','甘肃','青海','台湾','内蒙古','广西','西藏','宁夏','新疆'];
  for (const p of provinces) {
    if (s.startsWith(p) && s.length > p.length) {
      s = s.slice(p.length);
      break;
    }
  }

  // 去掉末尾行政区划后缀（市/区/县/镇/乡）
  s = s.replace(/[市区县镇乡]$/, '');

  // 处理 "九江市共青城市" → "共青城"：取最后一个"市"之后的内容
  const lastCityIdx = s.lastIndexOf('市');
  if (lastCityIdx > 0 && lastCityIdx < s.length - 1) {
    const tail = s.slice(lastCityIdx + 1);
    if (tail) s = tail;
  }

  // 再次去掉末尾后缀
  s = s.replace(/[市区县镇乡]$/, '');

  // 常见城市别名映射（数据源城市名与模板不统一时统一为标准名）
  const CITY_ALIAS_MAP = {
    '张家港': '苏州市',
    '张家港市': '苏州市',
    '扬州江都区': '扬州市',
    '扬州江都': '扬州市',
    '江都': '扬州市',
    '太原尖草坪店': '太原尖草坪',
    '尖草坪': '太原尖草坪',
    '临洮': '临洮县',
    '固原市原州区': '固原',
    '原州': '固原',
    '固原市': '固原',
    '鄂尔多斯市': '鄂尔多斯',
    '鄂尔多斯准格尔旗薛家湾': '鄂尔多斯',
    '天水市': '天水',
    '麦积': '天水',
    '甘肃省天水市麦积区': '天水',
    '吕梁市': '吕梁',
  };
  if (CITY_ALIAS_MAP[s]) s = CITY_ALIAS_MAP[s];

  return s;
}

function cleanDealer(name) {
  if (!name) return '';
  let s = String(name).trim();
  if (s.includes('.')) s = s.split('.')[0];
  if (['重庆直营店', '直营店.渝北', '直营店'].includes(s)) return '直营店';
  if (['国际电商部', '国际电商'].includes(s)) return '国际电商';
  return s.replace(/[（(].*?[)）]/g, '').trim();
}

function extractDealerCity(rawName, rawCity) {
  if (!rawName) return { dealer: '', city: '' };
  const nameStr = String(rawName).trim();
  let dealer = nameStr;
  let city = rawCity ? String(rawCity).trim() : '';
  if (nameStr.includes('.')) {
    const parts = nameStr.split('.');
    dealer = parts[0];
    if (!city) city = parts[1];
  }
  dealer = cleanDealer(dealer);
  return { dealer, city: normalizeCity(city) };
}

function readWorkbook(filePath) {
  return XLSX.readFile(filePath, { cellStyles: false, cellNF: false });
}

function readSheetRows(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet "${sheetName}" not found`);
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

function findSheetByKeyword(wb, keywords) {
  const names = wb.SheetNames;
  for (const kw of keywords) {
    const found = names.find(n => n.includes(kw));
    if (found) return found;
  }
  return null;
}

function findHeaderIndex(headerRow, keywords) {
  for (let i = 0; i < headerRow.length; i++) {
    const h = String(headerRow[i] || '').trim();
    for (const kw of keywords) {
      if (h.includes(kw)) return i;
    }
  }
  return -1;
}

function extractYyMm(prodNo) {
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

function processShejiang(wb) {
  const results = [];
  const sheetName = findSheetByKeyword(wb, ['奢匠下单表', '下单表']);
  if (!sheetName) throw new Error(`奢匠明细文件中找不到"奢匠下单表"Sheet（现有: ${wb.SheetNames.join('、')}）`);
  const rows = readSheetRows(wb, sheetName);
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    const month = String(row[1] || '').trim();
    const { dealer, city } = extractDealerCity(row[5], row[4]);
    const province = normalizeProvince(row[3]);
    const ot = row[6];
    // 用总金额（col 22）
    const actual = Number(row[22] || 0);
    if (!dealer) continue;
    if (month !== TARGET_MONTH_STR && month !== String(TARGET_MONTH_NUM)) continue;
    const otMap = { '订单': '订单', '补单': '订单', '样品': '样品', '小样': '样品' };
    if (!otMap[ot]) continue;
    results.push({ 月份: TARGET_MONTH_STR, 经销商: dealer, 省份: province, 城市: city, 订单类型: otMap[ot], 实际销售额: actual, 来源: '线下' });
  }
  return results;
}

function processDirect(wb) {
  const results = [];
  const sheetName = findSheetByKeyword(wb, ['直营店+电商下单表', '直营店+电商', '直营店']);
  if (!sheetName) {
    console.log(`提示: 找不到"直营店+电商下单表"Sheet（现有: ${wb.SheetNames.join('、')}），直营店数据按 0 处理`);
    return results;
  }
  const rows = readSheetRows(wb, sheetName);
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    const month = String(row[1] || '').trim();
    const { dealer, city } = extractDealerCity(row[5], row[4]);
    const province = normalizeProvince(row[3]);
    const ot = row[6];
    // 用总金额（col 21）
    const actual = Number(row[21] || 0);
    if (!dealer) continue;
    if (month !== TARGET_MONTH_STR && month !== String(TARGET_MONTH_NUM)) continue;
    const otMap = { '订单': '订单', '补单': '订单', '样品': '样品', '小样': '样品' };
    if (!otMap[ot]) continue;
    results.push({ 月份: TARGET_MONTH_STR, 经销商: dealer, 省份: province, 城市: city, 订单类型: otMap[ot], 实际销售额: actual, 来源: '线下' });
  }
  return results;
}

function processLiansi(wb) {
  const results = [];
  const sheetName = findSheetByKeyword(wb, ['联思', 'Sheet']);
  if (!sheetName) throw new Error('联思文件中找不到包含"联思"或"Sheet"的 Sheet');
  console.log(`联思文件使用 Sheet: ${sheetName}`);
  const rows = readSheetRows(wb, sheetName);
  if (rows.length < 2) return results;

  // 通过表头自动识别列索引
  const headerRow = rows[0];
  const colDealer = findHeaderIndex(headerRow, ['经销商名称', '经销商']);
  const colProvince = findHeaderIndex(headerRow, ['省']);
  const colCity = findHeaderIndex(headerRow, ['市', '城市']);
  const colOrderType = findHeaderIndex(headerRow, ['订单类型']);
  let colTotal = findHeaderIndex(headerRow, ['订单金额', '订单总金额']);
  if (colTotal < 0) colTotal = findHeaderIndex(headerRow, ['原价总金额']);  // 7月起联思导出列名改为原价总金额
  const colActual = findHeaderIndex(headerRow, ['应收金额', '折后金额']);
  const colCustomer = findHeaderIndex(headerRow, ['客户']);
  const colMonth = findHeaderIndex(headerRow, ['月份']);

  if (colDealer < 0 || colOrderType < 0) {
    throw new Error('联思文件表头识别失败，找不到经销商或订单类型列');
  }
  console.log(`联思列映射：经销商=${colDealer}, 省=${colProvince}, 订单类型=${colOrderType}, 金额=${colTotal}/${colActual}`);

  const startRow = 1; // 表头下一行
  for (let i = startRow; i < rows.length; i++) {
    const row = rows[i];
    const { dealer, city } = extractDealerCity(row[colDealer], colCity >= 0 ? row[colCity] : '');
    const province = colProvince >= 0 ? normalizeProvince(row[colProvince]) : '';
    const ot = row[colOrderType];
    const total = colTotal >= 0 ? Number(row[colTotal] || 0) : 0;
    // 联思统一用订单金额（与现有模板对齐）；如需要折后金额，可把下一行改回 row[colActual]
    const actual = total;
    const cust = colCustomer >= 0 ? row[colCustomer] : '';
    const month = colMonth >= 0 ? String(row[colMonth] || '').trim() : '';

    if (!dealer) continue;
    if (cust && String(cust).trim() === '色板') continue;
    if (colMonth >= 0 && month && month !== TARGET_MONTH_STR && month !== String(TARGET_MONTH_NUM)) continue;

    const otMap = {
      '补单': '订单', '纯板式订单': '订单', '纯实木订单': '订单',
      '板式含油漆订单': '订单', '店样品单': '样品', '售后单': null
    };
    if (!otMap[ot]) continue;
    results.push({ 月份: TARGET_MONTH_STR, 经销商: dealer, 省份: province, 城市: city, 订单类型: otMap[ot], 实际销售额: actual, 来源: '联思' });
  }
  return results;
}

function processZhcx(wb, existingProdNos) {
  const results = [];
  // 优先精确匹配 "综合查询" sheet；找不到再按关键字模糊匹配
  let sheetName = wb.SheetNames.find(n => n === '综合查询');
  if (!sheetName) sheetName = findSheetByKeyword(wb, ['综合查询']);
  if (!sheetName) throw new Error('综合查询文件中找不到包含"综合查询"的 Sheet');
  console.log(`综合查询文件使用 Sheet: ${sheetName}`);
  const rows = readSheetRows(wb, sheetName);
  if (rows.length < 2) return results;

  // 自动查找表头行：包含"生产编号"和"经销商"的行
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    if (findHeaderIndex(rows[i], ['生产编号']) >= 0 && findHeaderIndex(rows[i], ['经销商']) >= 0) {
      headerRowIdx = i;
      break;
    }
  }
  if (headerRowIdx < 0) {
    throw new Error('综合查询文件表头识别失败');
  }
  const headerRow = rows[headerRowIdx];
  const colProdNo = findHeaderIndex(headerRow, ['生产编号']);
  const colDealer = findHeaderIndex(headerRow, ['经销商']);
  const colProvince = findHeaderIndex(headerRow, ['省份']);
  const colCustomer = findHeaderIndex(headerRow, ['终端客户']);
  const colSample = findHeaderIndex(headerRow, ['样品']);
  const colProdType = findHeaderIndex(headerRow, ['生产类型']);
  const colMaker = findHeaderIndex(headerRow, ['制单人']);
  const colTotal = findHeaderIndex(headerRow, ['总金额']);
  const colPrice = findHeaderIndex(headerRow, ['原价']);

  console.log(`综合查询列映射：生产编号=${colProdNo}, 经销商=${colDealer}, 总金额=${colTotal}`);

  const startRow = headerRowIdx + 1;
  for (let i = startRow; i < rows.length; i++) {
    const row = rows[i];
    const prodNo = row[colProdNo];
    const { dealer, city } = extractDealerCity(row[colDealer], '');
    const province = colProvince >= 0 ? normalizeProvince(row[colProvince]) : '';
    const cust = colCustomer >= 0 ? row[colCustomer] : '';
    const sample = colSample >= 0 ? row[colSample] : '';
    const prodType = colProdType >= 0 ? row[colProdType] : '';
    const maker = colMaker >= 0 ? row[colMaker] : '';
    // 综合查询按「原价」取金额；如没有原价列再回退到总金额
    const actual = colPrice >= 0 ? Number(row[colPrice] || 0) : (colTotal >= 0 ? Number(row[colTotal] || 0) : 0);

    if (!prodNo || !dealer) continue;
    const prodNoStr = String(prodNo).trim();
    if (existingProdNos && existingProdNos.has(prodNoStr)) continue;
    if (['杨益琴', '陈明会'].includes(maker)) continue;
    if (prodType === '售后生产') continue;
    if (cust) {
      const cstr = String(cust).trim();
      if (cstr.includes('色卡') || cstr === '打色板') continue;
    }
    const [year, month] = extractYyMm(prodNo);
    if (!year || year !== 2026 || month !== TARGET_MONTH_NUM) continue;
    const ot = (sample === '是' || String(sample).includes('是')) ? '样品' : '订单';
    results.push({ 月份: TARGET_MONTH_STR, 经销商: dealer, 省份: province, 城市: city, 订单类型: ot, 实际销售额: actual, 来源: '订单通' });
  }
  return results;
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

function createExtractWorkbook(records) {
  const summaryData = records.map(r => ({
    '月份': r.月份,
    '经销商': r.经销商,
    '省份': r.省份,
    '城市': r.城市 || '',
    '订单类型': r.订单类型,
    '实际销售额': round2(r.实际销售额),
    '来源': r.来源
  }));

  const pivotAgg = {};
  for (const r of records) {
    if (!pivotAgg[r.经销商]) pivotAgg[r.经销商] = { '经销商': r.经销商, [TARGET_MONTH_STR]: 0, '2026年金额': 0 };
    pivotAgg[r.经销商][TARGET_MONTH_STR] += r.实际销售额;
    pivotAgg[r.经销商]['2026年金额'] += r.实际销售额;
  }
  const pivotData = Object.values(pivotAgg)
    .map(d => ({ '经销商': d.经销商, [TARGET_MONTH_STR]: round2(d[TARGET_MONTH_STR]), '2026年金额': round2(d['2026年金额']) }))
    .sort((a, b) => b['2026年金额'] - a['2026年金额']);

  const sourceAgg = {};
  for (const r of records) {
    const key = `${r.来源}|${r.月份}`;
    if (!sourceAgg[key]) sourceAgg[key] = { '来源': r.来源, '月份': r.月份, '实际销售额': 0 };
    sourceAgg[key]['实际销售额'] += r.实际销售额;
  }
  const sourceData = Object.values(sourceAgg).map(d => ({ ...d, '实际销售额': round2(d['实际销售额']) }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryData), '汇总数据');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pivotData), '透视表');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sourceData), '按来源月份汇总');
  return wb;
}

// ==================== 步骤2：ExcelJS 映射到模板（保格式/保公式） ====================

/** 展开共享公式（避免 ExcelJS 保存时报 Shared Formula 错误），与 generate_may_sales_report.js 一致 */
function expandSharedFormulas(wb) {
  wb.eachSheet((sheet) => {
    const masters = {};
    sheet.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        if (cell.type === ExcelJS.ValueType.Formula && cell.value && typeof cell.value === 'object') {
          const v = cell.value;
          if (v.formula && v.ref) {
            masters[cell.row] = v.formula;
          }
        }
      });
    });
    sheet.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        if (cell.type === ExcelJS.ValueType.Formula && cell.value && typeof cell.value === 'object') {
          const v = cell.value;
          if (v.master && masters[v.master]) {
            cell.value = { formula: masters[v.master], result: v.result };
          } else if (v.formula && v.ref) {
            cell.value = { formula: v.formula, result: v.result };
          }
        }
      });
    });
  });
}

/** 读取单元格的"显示值"（公式单元格取 result） */
function cellText(cell) {
  if (!cell || cell.value == null) return '';
  const v = cell.value;
  if (typeof v === 'object') {
    if (v.result != null) return String(v.result);
    if (v.richText) return v.richText.map(t => t.text).join('');
    if (v.text) return String(v.text);
    return '';
  }
  return String(v);
}

// ==================== 步骤2：累积表（每月递增，持久化） ====================
const ACC_SHEET_NAME = '月度数据';

/** 从销量模板建立经销商城市归集映射。
 *  key: 经销商|省份  -> 城市（省份非空）
 *  key: 经销商|*     -> 城市（省份为空，仅按经销商+城市归集）
 * 保证同一经销商的数据落到模板所在城市，避免被数据源不同城市写法拆成多行。 */
function buildTemplateMap(templatePath) {
  const wb = XLSX.readFile(templatePath);
  const ws = wb.Sheets['奢匠专卖店26年'];
  if (!ws) return new Map();
  const rows = readSheetRows(wb, '奢匠专卖店26年');
  const map = new Map();
  for (let i = 2; i < rows.length; i++) {
    const row = rows[i];
    const dealer = cleanDealer(String(row[3] || ''));
    const province = normalizeProvince(String(row[1] || ''));
    const city = normalizeCity(String(row[2] || ''));
    if (!dealer) continue;
    if (province) {
      const key = `${dealer}|${province}`;
      if (!map.has(key)) map.set(key, city);
    } else if (city) {
      const key = `${dealer}|*`;
      if (!map.has(key)) map.set(key, city);
    }
  }
  return map;
}

function loadAccumulator(accPath) {
  if (!fs.existsSync(accPath)) throw new Error(`累积表不存在: ${accPath}`);
  const wb = XLSX.readFile(accPath);
  const sheetName = wb.SheetNames.find(n => n.includes(ACC_SHEET_NAME)) || wb.SheetNames[0];
  const rows = readSheetRows(wb, sheetName);
  const acc = [];
  for (let i = 1; i < rows.length; i++) {  // 第 1 行是表头
    const row = rows[i];
    const dealer = String(row[0] || '').trim();
    if (!dealer) continue;
    const months = {};
    for (let m = 1; m <= 12; m++) {
      const raw = row[2 + m];
      if (raw != null && raw !== '' && !isNaN(Number(raw))) months[m] = Number(raw);
    }
    acc.push({ dealer, province: String(row[1] || '').trim(), city: String(row[2] || '').trim(), months, used: false });
  }
  return acc;
}

function accKey(dealer, province, city) {
  return `${dealer}|${province}|${normalizeCity(city)}`;
}

/** 匹配链：精确(经销商+省份+城市) → 经销商+省份 → 经销商+城市（模板省份为空时） */
function findAccRow(acc, dealer, province, city) {
  const k = accKey(dealer, province, city);
  let hit = acc.find(r => accKey(r.dealer, r.province, r.city) === k);
  if (!hit) hit = acc.find(r => r.dealer === dealer && r.province === province);
  if (!hit && !province && city) {
    const nc = normalizeCity(city);
    const candidates = acc.filter(r => r.dealer === dealer && normalizeCity(r.city) === nc);
    // 优先匹配当月有数据的行，避免历史旧行抢占模板行导致新增重复行
    hit = candidates.find(r => r.months[TARGET_MONTH_NUM] != null) || candidates[0];
  }
  return hit || null;
}

/** 把当月正单数据写入累积表（当月列先清零再按提取结果整体重写，幂等）并保存 */
function updateAccumulator(acc, records, templateMap) {
  const pivotAgg = {};
  for (const r of records) {
    if (r.订单类型 !== '订单') continue;  // 样品仅汇总到提取表供核对
    // 优先以模板中该经销商对应的城市为准，避免同一经销商因数据源城市写法不同被拆成多行
    let city = r.城市 || '';
    if (templateMap) {
      const cityByProvince = templateMap.get(`${r.经销商}|${r.省份}`);
      const cityByDealer = templateMap.get(`${r.经销商}|*`);
      city = cityByProvince || cityByDealer || city;
    }
    const key = `${r.经销商}|${r.省份}|${city}`;
    if (!pivotAgg[key]) pivotAgg[key] = { dealer: r.经销商, province: r.省份, city, amount: 0 };
    pivotAgg[key].amount += r.实际销售额;
  }
  // 当月列整体清零：本月数据以本次提取为准，不残留旧值
  for (const row of acc) delete row.months[TARGET_MONTH_NUM];
  // 先按匹配到的累积行聚合（多个 pivot 键可能命中同一行，需累加而非覆盖）
  const rowAmounts = new Map(); // accRow -> amount
  const newRows = [];
  for (const data of Object.values(pivotAgg)) {
    const val = round2(data.amount);
    if (val === 0) continue;
    // 严格按 经销商+省份+城市 精确匹配，避免同名不同地经销商被合并
    const k = accKey(data.dealer, data.province, data.city);
    let row = acc.find(r => accKey(r.dealer, r.province, r.city) === k);

    // 若模板中有该经销商+省份但累积表尚无对应行，按 dealer+省份 查找已有行并合并（同时会带入模板城市）
    if (!row) {
      row = acc.find(r => r.dealer === data.dealer && r.province === data.province);
    }

    if (row) {
      rowAmounts.set(row, round2((rowAmounts.get(row) || 0) + val));
    } else {
      newRows.push({ dealer: data.dealer, province: data.province, city: data.city || '', amount: val });
    }
  }
  for (const [row, val] of rowAmounts) {
    row.months[TARGET_MONTH_NUM] = val;
  }
  for (const nr of newRows) {
    acc.push({ dealer: nr.dealer, province: nr.province, city: nr.city || '', months: { [TARGET_MONTH_NUM]: nr.amount }, used: false });
  }
  console.log(`累积表更新: 当月 ${TARGET_MONTH_STR} 写入${rowAmounts.size}个经销商, 新增${newRows.length}个`);

  const out = XLSX.utils.book_new();
  const aoa = [['经销商', '省份', '城市', ...Array.from({ length: 12 }, (_, i) => `${i + 1}月`)]];
  for (const r of acc) {
    aoa.push([r.dealer, r.province, r.city, ...Array.from({ length: 12 }, (_, i) => r.months[i + 1] != null ? r.months[i + 1] : null)]);
  }
  XLSX.utils.book_append_sheet(out, XLSX.utils.aoa_to_sheet(aoa), ACC_SHEET_NAME);
  XLSX.writeFile(out, ACC_FILE);
  console.log(`累积表已保存: ${ACC_FILE} (共${acc.length}行)`);
}

// ==================== 步骤3：主表 1-12 月全部由累积表驱动填充 ====================
async function mapToTemplate(acc) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(TEMPLATE_FILE);
  expandSharedFormulas(wb);

  const ws = wb.getWorksheet('奢匠专卖店26年');
  if (!ws) throw new Error('模板中找不到 Sheet "奢匠专卖店26年"');

  // 找到 "2025年金额" 列（表头在第 2 行），其后 1-12 月为目标填充区域
  let yearBaseCol = -1;
  const headerRow = ws.getRow(2);
  for (let c = 1; c <= ws.columnCount; c++) {
    if (cellText(headerRow.getCell(c)).includes('2025年金额')) {
      yearBaseCol = c;
      break;
    }
  }
  if (yearBaseCol < 0) {
    throw new Error('模板中找不到 "2025年金额" 列，无法定位目标月份列');
  }
  const yearStartCol = yearBaseCol + 1;   // CN (1月)
  const yearEndCol = yearBaseCol + 12;    // CY (12月)
  const yearTotalCol = yearBaseCol + 13;  // CZ 2026金额（SUM 公式）
  const colLetter = (c) => {
    let s = '';
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  };
  console.log(`2026 月度列: ${colLetter(yearStartCol)}~${colLetter(yearEndCol)} (基于2025年金额列 ${colLetter(yearBaseCol)})`);

  // 读取模板经销商（列：A 序号 / B 省份 / C 地区城市 / D 客户姓名），自第 3 行起
  const templateAllRows = [];
  let lastDealerRow = 2;
  for (let r = 3; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const dealer = cleanDealer(cellText(row.getCell(4)));
    if (dealer) {
      templateAllRows.push(r);
      lastDealerRow = r;
    }
  }

  // 逐行匹配累积表，填充 1-12 月（无数据的单元格清空，避免残留旧公式/旧值）
  let filled = 0, cleared = 0;
  for (const r of templateAllRows) {
    const row = ws.getRow(r);
    const dealer = cleanDealer(cellText(row.getCell(4)));
    const province = normalizeProvince(cellText(row.getCell(2)));
    const city = cellText(row.getCell(3)).trim();
    const accRow = findAccRow(acc, dealer, province, city);
    if (accRow) accRow.used = true;
    for (let m = 1; m <= 12; m++) {
      const cell = row.getCell(yearBaseCol + m);
      const v = accRow ? accRow.months[m] : null;
      if (v != null) {
        cell.value = v;
        filled++;
      } else {
        if (cell.value != null) cleared++;
        cell.value = null;
      }
    }
  }

  // 累积表里有、模板里没有的经销商 → 表尾新增行（带全年各月值 + CZ 年合计公式）
  // 只新增当月有数据的行，避免历史月份有值但本月无值的旧行也被输出
  const newRows = acc.filter(r => !r.used && r.months[TARGET_MONTH_NUM] != null);
  if (newRows.length > 0) {
    const startRow = lastDealerRow + 1;
    const baseSeq = templateAllRows.length;
    for (let i = 0; i < newRows.length; i++) {
      const r = startRow + i;
      const nr = newRows[i];
      const row = ws.getRow(r);
      row.getCell(1).value = baseSeq + i + 1;
      row.getCell(2).value = nr.province;
      row.getCell(3).value = nr.city;
      row.getCell(4).value = nr.dealer;
      for (let m = 1; m <= 12; m++) {
        if (nr.months[m] != null) row.getCell(yearBaseCol + m).value = nr.months[m];
      }
      row.getCell(yearTotalCol).value = { formula: `SUM(${colLetter(yearStartCol)}${r}:${colLetter(yearEndCol)}${r})` };
      row.commit && row.commit();
    }
  }
  console.log(`主表填充完成: 写入${filled}格, 清空${cleared}格, 新增${newRows.length}行`);

  // 输出附「2026月度数据」sheet（累积表快照 + 合计公式）
  const sh = wb.addWorksheet('2026月度数据');
  const head = sh.addRow(['经销商', '省份', '城市', ...Array.from({ length: 12 }, (_, i) => `${i + 1}月`), '合计']);
  head.font = { bold: true };
  for (const r of acc) {
    const row = sh.addRow([r.dealer, r.province, r.city,
      ...Array.from({ length: 12 }, (_, i) => r.months[i + 1] != null ? r.months[i + 1] : null)]);
    row.getCell(16).value = { formula: `SUM(D${row.number}:O${row.number})` };
  }
  sh.getColumn(1).width = 20;
  sh.getColumn(2).width = 10;
  sh.getColumn(3).width = 12;

  console.log(`保存: ${OUTPUT_MAPPED}`);
  await wb.xlsx.writeFile(OUTPUT_MAPPED);
}

// ==================== 主流程 ====================
async function main() {
  console.log(`目标月份：${TARGET_MONTH_STR}`);
  console.log('步骤1: 提取数据');

  const shejiangWb = readWorkbook(SHEJIANG_FILE);
  const liansiWb = readWorkbook(LIANSI_FILE);
  const zhcxWb = readWorkbook(ZHCX_FILE);

  const records = [];
  records.push(...processShejiang(shejiangWb));
  records.push(...processDirect(shejiangWb));
  records.push(...processLiansi(liansiWb));
  // 线下和综合查询不去重，都计算
  records.push(...processZhcx(zhcxWb, null));

  console.log(`提取完成: ${records.length}条`);
  const sourceSummary = {};
  for (const r of records) {
    if (!sourceSummary[r.来源]) sourceSummary[r.来源] = { count: 0, amount: 0 };
    sourceSummary[r.来源].count++;
    sourceSummary[r.来源].amount += r.实际销售额;
  }
  for (const [src, s] of Object.entries(sourceSummary)) {
    console.log(`  ${src}: ${s.count}条 / ${s.amount.toFixed(2)}元`);
  }

  console.log(`保存: ${OUTPUT_EXTRACT}`);
  const extractWb = createExtractWorkbook(records);
  XLSX.writeFile(extractWb, OUTPUT_EXTRACT);

  console.log('\n步骤2: 更新累积表');
  const templateMap = buildTemplateMap(TEMPLATE_FILE);
  console.log(`模板经销商映射: ${templateMap.size}个`);
  const acc = loadAccumulator(ACC_FILE);
  updateAccumulator(acc, records, templateMap);

  console.log('\n步骤3: 按累积表填充模板');
  await mapToTemplate(acc);

  console.log('\n完成！');
}

main().catch(e => {
  console.error('执行失败:', e.message);
  console.error(e.stack);
  process.exit(1);
});
