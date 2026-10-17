import http from 'node:http';
import path from 'node:path';
import fs from 'node:fs/promises';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { dateFromExcelSerial, dateToISO, monthKey, monthLabel, calculateMonth, normalizeState } from './src/performance/calc.js';

const root = path.resolve('.');
const publicDir = path.join(root, 'out');
const port = Number(process.env.PORT || 4173);

const headers = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
};
const json = (response, status, payload) => {
  response.writeHead(status, { ...headers, 'Content-Type': 'application/json; charset=utf-8' });
  response.end(JSON.stringify(payload));
};

const SPREADSHEML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

function columnToNumber(letters) {
  return letters.split('').reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0);
}

function parseCellRef(ref) {
  const match = /^([A-Z]+)(\d+)$/.exec(ref || '');
  if (!match) return null;
  return { column: columnToNumber(match[1]), row: Number(match[2]) };
}

// 部分导出工具会产生互相重叠的合并区域，exceljs 遇到重叠会直接抛错，这里只保留互不冲突的那些
function dedupeOverlappingMerges(xml) {
  const ranges = [];
  const pattern = /<(?:x:)?mergeCell\b[^>]*ref="([A-Z]+\d+(?::[A-Z]+\d+)?)"[^>]*\/?>/g;
  let match = pattern.exec(xml);
  while (match) {
    const [startRef, endRef = match[1]] = match[1].split(':');
    const start = parseCellRef(startRef);
    const end = parseCellRef(endRef);
    if (start && end) {
      ranges.push({
        raw: match[1],
        minColumn: Math.min(start.column, end.column),
        maxColumn: Math.max(start.column, end.column),
        minRow: Math.min(start.row, end.row),
        maxRow: Math.max(start.row, end.row),
      });
    }
    match = pattern.exec(xml);
  }
  if (ranges.length < 2) return xml;

  const kept = [];
  for (const range of ranges) {
    const overlaps = kept.some((other) => range.minColumn <= other.maxColumn && range.maxColumn >= other.minColumn && range.minRow <= other.maxRow && range.maxRow >= other.minRow);
    if (!overlaps) kept.push(range);
  }
  if (kept.length === ranges.length) return xml;

  const keptRefs = new Set(kept.map((range) => range.raw));
  const dropRefs = new Set(ranges.filter((range) => !keptRefs.has(range.raw)).map((range) => range.raw));
  return xml.replace(pattern, (full, ref) => (dropRefs.has(ref) ? '' : full));
}

async function normalizeWorkbookXml(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  let changed = false;
  const targets = Object.keys(zip.files).filter(
    (name) => name === 'xl/workbook.xml' || name === 'xl/sharedStrings.xml' || name === 'xl/styles.xml' || /^xl\/worksheets\/sheet\d+\.xml$/.test(name),
  );
  for (const partName of targets) {
    const entry = zip.file(partName);
    if (!entry) continue;
    const xml = await entry.async('string');
    if (!/<x:[A-Za-z]/.test(xml) && !/<mergeCell\b/.test(xml)) continue;
    let normalized = xml;
    if (/<x:[A-Za-z]/.test(normalized)) {
      normalized = normalized
        .replace(/<x:([A-Za-z])/g, '<$1')
        .replace(/<\/x:([A-Za-z])/g, '</$1')
        .replace(new RegExp(`xmlns:x="${SPREADSHEML_NS}"`, 'g'), `xmlns="${SPREADSHEML_NS}"`)
        .replace(/xmlns:x=/g, 'xmlns=');
    }
    normalized = dedupeOverlappingMerges(normalized);
    zip.file(partName, normalized);
    changed = true;
  }
  if (!changed) return bytes;
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function text(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') {
    if (value.text !== undefined) return String(value.text);
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text).join('');
    if (value.result !== undefined) return String(value.result);
    if (value.hyperlink && value.text) return String(value.text);
  }
  return String(value);
}

function cellValue(sheet, address) {
  const cell = sheet.getCell(address);
  const value = cell.value;
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return dateToISO(value);
  if (typeof value === 'object' && !Array.isArray(value) && value.result !== undefined) return value.result;
  return value;
}

function dateValue(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) return dateToISO(value);
  if (typeof value === 'number') return dateFromExcelSerial(value);
  if (typeof value === 'object' && value.result instanceof Date) return dateToISO(value.result);
  if (typeof value === 'object' && value.result !== undefined) return dateValue(value.result);
  return dateToISO(text(value));
}

function rowValues(sheet, startRow, startCol, endRow, endCol) {
  const rows = [];
  for (let r = startRow; r <= endRow; r += 1) {
    const row = [];
    for (let c = startCol; c <= endCol; c += 1) row.push(sheet.getRow(r).getCell(c).value ?? '');
    rows.push(row);
  }
  return rows;
}

function colLetterToNumber(letters) {
  return letters.split('').reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0);
}

function columnLetters(number) {
  let result = '';
  let value = number;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

async function importWorkbook(bytes) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await normalizeWorkbookXml(bytes));
  const state = normalizeState({});
  const warnings = [];

  const profileSheet = workbook.getWorksheet('基础信息');
  if (profileSheet) {
    // 按标签文字在整行内寻找对应的值，兼容不同模板把标签放在左侧或右侧的差异
    const labelledValue = (labels) => {
      for (let row = 1; row <= Math.min(profileSheet.rowCount || 10, 12); row += 1) {
        for (let column = 1; column <= Math.min(profileSheet.columnCount || 10, 12); column += 1) {
          const label = text(cellValue(profileSheet, `${columnLetters(column)}${row}`)).replace(/\s/g, '');
          if (label && labels.some((candidate) => label.includes(candidate))) {
            const next = cellValue(profileSheet, `${columnLetters(column + 1)}${row}`);
            if (next !== null && next !== undefined && text(next) !== '') return next;
          }
        }
      }
      return null;
    };
    state.profile.name = text(labelledValue(['姓名', '员工姓名']) ?? cellValue(profileSheet, 'B3'));
    state.profile.department = text(labelledValue(['部门', '所属部门']) ?? cellValue(profileSheet, 'D3'));
    state.profile.manager = text(labelledValue(['直属负责人', '直属上级', '上级']) ?? cellValue(profileSheet, 'B4'));
    state.profile.position = text(labelledValue(['岗位', '职位', '职务']) ?? cellValue(profileSheet, 'F3') ?? cellValue(profileSheet, 'E3'));
    const year = labelledValue(['考核年度', '年度']);
    state.profile.year = Number(year) || Number(cellValue(profileSheet, 'D4')) || state.profile.year;
    state.profile.activationDate = dateValue(
      labelledValue(['考核启用日期', '启用日期', '入职日期', '入职', '起算日期', '激活日期']) ?? cellValue(profileSheet, 'F4') ?? cellValue(profileSheet, 'E4'),
    );
  } else warnings.push('未找到基础信息工作表');

  const parameterSheet = workbook.getWorksheet('绩效参数');
  if (parameterSheet) {
    const params = [
      ['standardDayHours', ['标准日工时', '日标准工时']],
      ['baseRewardPerHour', ['基础奖励', '基础分']],
      ['expansionFactor', ['拓展工作', '拓展系数']],
      ['overtimeFactor', ['加班工作', '加班系数']],
      ['monthlyInnovationMax', ['创新加分上限', '积极创新']],
      ['efficiencyMin', ['效率系数下限', '效率下限']],
      ['efficiencyMax', ['效率系数上限', '效率上限']],
      ['plannedOverageFactor', ['计划内超额', '计划内']],
    ];
    // 优先按参数名精确查找，找不到时回退到按行号读取，保证旧版模板也能导入
    const byLabel = (candidates) => {
      for (let row = 1; row <= Math.min(parameterSheet.rowCount || 40, 60); row += 1) {
        for (let column = 1; column <= Math.min(parameterSheet.columnCount || 6, 6); column += 1) {
          const label = text(cellValue(parameterSheet, `${columnLetters(column)}${row}`)).replace(/\s/g, '');
          if (label && candidates.some((candidate) => label.includes(candidate))) {
            const next = cellValue(parameterSheet, `${columnLetters(column + 1)}${row}`);
            if (Number.isFinite(Number(next))) return Number(next);
          }
        }
      }
      return null;
    };
    params.forEach(([key, candidates], index) => {
      const value = byLabel(candidates) ?? Number(cellValue(parameterSheet, `B${index + 4}`));
      if (Number.isFinite(value)) state.settings[key] = value;
    });
  } else warnings.push('未找到绩效参数工作表，使用默认参数');

  const dropdownSheet = workbook.getWorksheet('下拉配置');
  if (dropdownSheet) {
    const categories = rowValues(dropdownSheet, 2, 1, 100, 1).flat().map((value) => text(value).trim()).filter(Boolean);
    if (categories.length) state.settings.taskCategories = [...new Set(categories)];
  }

  const monthSheets = workbook.worksheets.filter((sheet) => /^\d{1,2}月任务明细$/.test(sheet.name));
  const archiveTasksSheet = workbook.getWorksheet('任务明细');
  const taskSheets = monthSheets.length
    ? monthSheets.map((sheet) => ({ sheet, start: 7, endColumn: 18 }))
    : archiveTasksSheet
      ? [{ sheet: archiveTasksSheet, start: 2, endColumn: 20 }]
      : [];
  if (!taskSheets.length) warnings.push('未找到任务明细工作表');

  const TASK_FIELD_RULES = [
  ['date', ['日期']],
  ['content', ['今日工作内容', '工作内容', '任务内容', '工作事项']],
  ['category', ['任务分类', '工作分类', '分类']],
  ['planType', ['工作类型', '计划类型']],
  ['collaborator', ['协作人', '协同人']],
  ['estimatedHours', ['预计工时']],
  ['actualHours', ['实际工时']],
  ['completionPct', ['交付完成', '完成占比']],
  ['collaborationPct', ['协作完成']],
  ['innovation', ['创新分']],
  ['selfScore', ['本人评分', '自评分']],
  ['reviewerScore', ['负责人评分', '主管评分']],
  ['nextPlan', ['明日计划', '下一步计划']],
  ['blocker', ['阻塞', '问题', '备注']],
  ['breakthrough', ['突破', '创新成果']],
];

// 任务表的表头行位置在不同版本模板里不一致（6 / 7 / 2），列顺序也不完全相同，这里按表头文字定位
function detectTaskHeader(sheet) {
  const limit = Math.min(sheet.rowCount || 20, 20);
  for (let row = 1; row <= limit; row += 1) {
    const headers = rowValues(sheet, row, 1, row, Math.min(sheet.columnCount || 20, 60))[0].map((value) => text(value).replace(/\s/g, ''));
    const mapping = {};
    let matched = 0;
    for (const [field, candidates] of TASK_FIELD_RULES) {
      const index = headers.findIndex((header) => header && candidates.some((candidate) => header.includes(candidate)));
      if (index >= 0) {
        mapping[field] = index + 1;
        matched += 1;
      }
    }
    if (matched >= 4 && mapping.date && mapping.content) return { headerRow: row, mapping };
  }
  return null;
}

for (const { sheet, start, endColumn } of taskSheets) {
    const detected = detectTaskHeader(sheet);
    if (!detected) {
      warnings.push(`${sheet.name} 未找到可识别的表头行，已跳过`);
      continue;
    }
    const { headerRow, mapping } = detected;
    const lastColumn = Math.min(sheet.columnCount || endColumn, 60);
    let imported = 0;
    for (let row = headerRow + 1; row <= Math.min(sheet.rowCount || headerRow + 300, headerRow + 400); row += 1) {
      const read = (field) => (mapping[field] ? cellValue(sheet, `${columnLetters(mapping[field])}${row}`) : '');
      const task = { id: crypto.randomUUID() };
      for (const [field] of TASK_FIELD_RULES) {
        const value = read(field);
        task[field] = field === 'date' ? dateValue(value) : value ?? '';
      }
      if (!task.date || !task.content) continue;
      state.tasks.push(task);
      imported += 1;
    }
    if (imported === 0) warnings.push(`${sheet.name} 未导入有效任务`);
  }

  const poolSheet = workbook.getWorksheet('任务池');
  if (poolSheet) {
    for (const row of rowValues(poolSheet, 2, 1, 300, 11)) {
      if (!row[0] && !row[1]) continue;
      state.workItems.push({
        id: text(row[0]) || crypto.randomUUID(),
        title: text(row[1]),
        description: text(row[2]),
        category: text(row[3]) || state.settings.taskCategories[0],
        planType: text(row[4]) || '单人工作',
        priority: text(row[5]) || '普通',
        dueDate: dateValue(row[6]),
        status: text(row[7]) || 'todo',
        reviewer: text(row[8]),
        acceptanceStatus: text(row[9]),
        performanceTaskId: text(row[10]),
      });
    }
  }

  const plansSheet = workbook.getWorksheet('计划');
  if (plansSheet) {
    for (const row of rowValues(plansSheet, 2, 1, 500, 7)) {
      if (!row[0] || !row[1]) continue;
      const done = row[6] === true || ['是', '1', 'true'].includes(text(row[6]).toLowerCase());
      const dailyBucket = row[0] === '今日计划' ? 'today' : row[0] === '明日计划' ? 'tomorrow' : '';
      if (dailyBucket) {
        const date = dateValue(row[1]);
        if (!date || !text(row[4])) continue;
        const plan = state.dailyPlans[date] || (state.dailyPlans[date] = { today: [], tomorrow: [] });
        plan[dailyBucket].push({ id: text(row[4]), text: text(row[5]), done });
        continue;
      }
      const kind = row[0] === '周计划' ? 'weekly' : row[0] === '月计划' ? 'monthly' : '';
      if (!kind) continue;
      const plan = state.plans[kind][text(row[1])] || (state.plans[kind][text(row[1])] = { title: text(row[2]), notes: text(row[3]), items: [] });
      if (text(row[4])) plan.items.push({ id: text(row[4]), text: text(row[5]), done });
    }
  }

  const monthlySheet = workbook.getWorksheet('月度绩效');
  if (monthlySheet) {
    for (let row = 4; row <= 15; row += 1) {
      const label = text(cellValue(monthlySheet, `A${row}`));
      const match = label.match(/(\d{1,2})月/);
      if (!match) continue;
      const key = monthKey(state.profile.year, Number(match[1]));
      state.monthlyReviews[key] = {
        collaboration: cellValue(monthlySheet, `E${row}`) ?? '',
        loyalty: cellValue(monthlySheet, `F${row}`) ?? '',
        discipline: cellValue(monthlySheet, `G${row}`) ?? '',
        learning: cellValue(monthlySheet, `H${row}`) ?? '',
        innovationBonus: cellValue(monthlySheet, `L${row}`) ?? '',
        note: text(cellValue(monthlySheet, `P${row}`)),
      };
    }
  }

  return { state, warnings, sheets: taskSheets.map(({ sheet }) => sheet.name) };
}

function styleHeader(sheet, rowCount) {
  const header = sheet.getRow(1);
  header.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF172B4D' } };
  header.alignment = { vertical: 'middle', horizontal: 'center' };
  header.height = 22;
  for (let c = 1; c <= rowCount; c += 1) sheet.getColumn(c).width = Math.max(12, Math.min(40, String(header.getCell(c).value ?? '').length * 2 + 8));
}

async function exportWorkbook(rawState) {
  const state = normalizeState(rawState);
  const workbook = new ExcelJS.Workbook();

  const overview = workbook.addWorksheet('基础信息');
  // 导入端固定读取 B3 / D3 / E3 / B4 / D4 / E4：值严格落在这些格子上，标签写在紧邻的左侧格
  const setProfileField = (labelColumn, valueColumn, row, label, value) => {
    overview.getCell(`${labelColumn}${row}`).value = label;
    overview.getCell(`${labelColumn}${row}`).font = { bold: true };
    overview.getCell(`${valueColumn}${row}`).value = value;
  };
  setProfileField('A', 'B', 3, '姓名', state.profile.name);
  setProfileField('C', 'D', 3, '部门', state.profile.department);
  overview.getCell('E3').value = state.profile.position;
  setProfileField('A', 'B', 4, '直属负责人', state.profile.manager);
  setProfileField('C', 'D', 4, '考核年度', state.profile.year);
  overview.getCell('E4').value = state.profile.activationDate;
  overview.getCell('A1').value = '绩效考核基础信息';
  overview.mergeCells('A1:E1');
  const titleCell = overview.getCell('A1');
  titleCell.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF172B4D' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
  overview.getRow(1).height = 30;
  [16, 22, 14, 22, 20].forEach((width, index) => { overview.getColumn(index + 1).width = width; });
  overview.getCell('A6').value = '任务总数';
  overview.getCell('B6').value = state.tasks.length;
  overview.getCell('A7').value = '导出时间';
  overview.getCell('B7').value = dateToISO(new Date());
  overview.getCell('A8').value = '说明';
  overview.getCell('B8').value = '本文件由绩效考核工作台导出，可直接回传导入以继续编辑';
  overview.getCell('A6').font = { bold: true };
  overview.getCell('A7').font = { bold: true };
  overview.getCell('A8').font = { bold: true };

  const tasksSheet = workbook.addWorksheet('任务明细');
  tasksSheet.addRow(['日期', '工作内容', '任务分类', '计划类型', '协作人', '预计工时', '实际工时', '交付完成%', '协作完成%', '创新分', '自评分', '负责人评分', '明日计划', '阻塞问题', '重大突破', '任务得分', '综合加权有效工时', '当日绩效加权平均分', '当日奖励分', '校验']);
  const taskRows = state.tasks.map((task) => {
    const month = calculateMonth(Number(task.date.slice(0, 4)), Number(task.date.slice(5, 7)), state);
    const day = month.days.find((entry) => entry.date === task.date);
    const calculated = day?.tasks?.find((entry) => entry.id === task.id) || task;
    return [task.date, task.content, task.category, task.planType, task.collaborator, task.estimatedHours, task.actualHours, task.completionPct, task.collaborationPct, task.innovation, task.selfScore, task.reviewerScore, task.nextPlan, task.blocker, task.breakthrough, calculated.taskScore ?? '', calculated.weightedEffectiveHours ?? '', day?.performance ?? '', day?.dailyReward ?? '', calculated.validationErrors?.join('；') || '正常'];
  });
  if (taskRows.length) tasksSheet.addRows(taskRows);
  styleHeader(tasksSheet, 20);
  tasksSheet.getColumn(2).width = 40;

  const daysSheet = workbook.addWorksheet('每日汇总');
  daysSheet.addRow(['日期', '标准工时', '任务数', '计划内有效工时', '拓展有效工时', '加班有效工时', '总有效工时', '绩效换算工时', '当日绩效', '奖励换算工时', '当日奖励', '实际工时', '状态', '校验']);
  const allDays = [];
  for (let month = 1; month <= 12; month += 1) allDays.push(...calculateMonth(state.profile.year, month, state).days);
  if (allDays.length) daysSheet.addRows(allDays.map((day) => [day.date, day.standardHours, day.taskCount, day.plannedHours, day.expansionHours, day.overtimeHours, day.totalHours, day.performanceHours, day.performance ?? '', day.rewardHours, day.dailyReward, day.actualHours, day.status, day.validationErrors.join('；')]));
  styleHeader(daysSheet, 14);

  const monthSheet = workbook.addWorksheet('月度绩效');
  monthSheet.addRow(['月份', '已填写天数', '平均每日绩效', '月度奖励总分', '协同能力', '忠诚度', '纪律性', '学习力', '综合素养', '日常表现折算', '素养折算', '积极创新加分', '月度最终绩效', '绩效兑现比例', '绩效等级']);
  const monthRows = Array.from({ length: 12 }, (_, index) => calculateMonth(state.profile.year, index + 1, state)).map((month) => [monthLabel(month.key), month.scoredDays, month.averageDailyPerformance ?? '', month.monthlyReward, month.review.collaboration ?? '', month.review.loyalty ?? '', month.review.discipline ?? '', month.review.learning ?? '', month.competencyScore ?? '', month.dailyWeighted ?? '', month.competencyWeighted ?? '', month.innovationBonus, month.finalPerformance ?? '', month.payoutRatio ?? '', month.grade]);
  monthSheet.addRows(monthRows);
  styleHeader(monthSheet, 15);

  const settingsSheet = workbook.addWorksheet('绩效参数');
  const settingRows = [
    ['标准日工时', state.settings.standardDayHours],
    ['基础奖励分/小时', state.settings.baseRewardPerHour],
    ['拓展工作奖励系数', state.settings.expansionFactor],
    ['加班工作奖励系数', state.settings.overtimeFactor],
    ['月度积极创新加分上限', state.settings.monthlyInnovationMax],
    ['效率系数下限', state.settings.efficiencyMin],
    ['效率系数上限', state.settings.efficiencyMax],
    ['计划内超额奖励系数', state.settings.plannedOverageFactor],
  ];
  settingsSheet.getCell('A1').value = '参数';
  settingsSheet.getCell('B1').value = '当前值';
  // 导入端固定从 B4 开始读取，参数行必须精确落在第 4 至 11 行
  settingRows.forEach(([label, value], index) => {
    settingsSheet.getCell(`A${index + 4}`).value = label;
    settingsSheet.getCell(`B${index + 4}`).value = value;
  });
  styleHeader(settingsSheet, 2);
  settingsSheet.getCell('A13').value = '日常表现权重';
  settingsSheet.getCell('B13').value = state.settings.dailyPerformanceWeight;
  settingsSheet.getCell('A14').value = '综合素养权重';
  settingsSheet.getCell('B14').value = state.settings.competencyWeight;

  const poolSheet = workbook.addWorksheet('任务池');
  poolSheet.addRow(['ID', '任务名称', '说明', '工作分类', '计划类型', '优先级', '截止日期', '状态', '验收人', '验收状态', '绩效任务ID']);
  if (state.workItems.length) poolSheet.addRows(state.workItems.map((item) => [item.id, item.title, item.description, item.category, item.planType, item.priority, item.dueDate, item.status, item.reviewer, item.acceptanceStatus, item.performanceTaskId]));
  styleHeader(poolSheet, 11);

  const plansSheet = workbook.addWorksheet('计划');
  plansSheet.addRow(['计划类型', '计划周期', '主题', '说明', '事项ID', '计划事项', '已完成']);
  const planRows = [];
  for (const [kind, label] of [['weekly', '周计划'], ['monthly', '月计划']]) {
    for (const [key, plan] of Object.entries(state.plans[kind] || {})) {
      const items = plan.items || [];
      if (!items.length) planRows.push([label, key, plan.title || '', plan.notes || '', '', '', '']);
      for (const item of items) planRows.push([label, key, plan.title || '', plan.notes || '', item.id, item.text, item.done ? '是' : '否']);
    }
  }
  for (const [date, plan] of Object.entries(state.dailyPlans || {})) {
    for (const [bucket, label] of [['today', '今日计划'], ['tomorrow', '明日计划']]) {
      for (const item of plan[bucket] || []) planRows.push([label, date, '', '', item.id, item.text, item.done ? '是' : '否']);
    }
  }
  if (planRows.length) plansSheet.addRows(planRows);
  styleHeader(plansSheet, 7);

  for (const sheet of [overview, tasksSheet, daysSheet, monthSheet, settingsSheet, poolSheet, plansSheet]) {
    sheet.views = [{ showGridLines: false }];
  }

  return workbook.xlsx.writeBuffer();
}

async function serveStatic(request, response) {
  let pathname = new URL(request.url, `http://${request.headers.host}`).pathname;
  if (pathname === '/') pathname = '/index.html';
  const filePath = path.resolve(publicDir, `.${pathname}`);
  if (!filePath.startsWith(publicDir)) return response.writeHead(403).end();
  try {
    const data = await fs.readFile(filePath);
    const ext = path.extname(filePath);
    const contentType = {
      '.html': 'text/html; charset=utf-8',
      '.js': 'text/javascript; charset=utf-8',
      '.css': 'text/css; charset=utf-8',
      '.json': 'application/json',
      '.svg': 'image/svg+xml',
      '.png': 'image/png',
      '.jpg': 'image/jpeg',
      '.woff2': 'font/woff2',
    }[ext] || 'application/octet-stream';
    response.writeHead(200, { ...headers, 'Content-Type': contentType });
    response.end(data);
  } catch {
    response.writeHead(404, headers).end('Not found');
  }
}

function readBody(request) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    request.on('data', (chunk) => chunks.push(chunk));
    request.on('end', () => resolve(Buffer.concat(chunks)));
    request.on('error', reject);
  });
}

const server = http.createServer(async (request, response) => {
  if (request.method === 'OPTIONS') return response.writeHead(204, headers).end();
  try {
    if (request.method === 'POST' && request.url === '/api/import') {
      const body = await readBody(request);
      return json(response, 200, await importWorkbook(body));
    }
    if (request.method === 'POST' && request.url === '/api/export') {
      const body = await readBody(request);
      const buffer = await exportWorkbook(JSON.parse(body.toString('utf8')));
      response.writeHead(200, {
        ...headers,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="performance-review.xlsx"',
      });
      return response.end(Buffer.from(buffer));
    }
  } catch (error) {
    return json(response, 400, { error: String(error?.message || error) });
  }
  return serveStatic(request, response);
});

server.listen(port, () => console.log(`绩效考核工作台已启动: http://localhost:${port}`));