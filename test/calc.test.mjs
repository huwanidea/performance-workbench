import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { DEFAULT_STATE, calculateDay, calculateMonth, normalizeState, validateTask } from '../src/performance/calc.js';
import { dateFromExcelSerial, dateToISO, monthKey, monthLabel } from '../src/performance/calc.js';

const base = normalizeState({ profile: { year: 2026, activationDate: '2026-01-01' } });
const task = (overrides = {}) => ({
  id: 't',
  date: '2026-08-03',
  content: '测试任务',
  category: '需求开发',
  planType: '单人工作',
  collaborator: '',
  estimatedHours: 7 + 20 / 60,
  actualHours: 7 + 20 / 60,
  completionPct: 100,
  collaborationPct: 0,
  innovation: 0,
  selfScore: 100,
  reviewerScore: 100,
  nextPlan: '',
  blocker: '',
  breakthrough: '',
  ...overrides,
});

let state = normalizeState({ ...base, tasks: [task()] });
let day = calculateDay('2026-08-03', state.tasks, state);
assert.equal(Math.round(day.performance), 100);
assert.equal(day.dailyReward, 0);

state = normalizeState({ ...base, tasks: [task({ estimatedHours: 10, actualHours: 10 })] });
day = calculateDay('2026-08-03', state.tasks, state);
assert.equal(Math.round(day.performance), 100);
assert.ok(day.dailyReward > 0);

state = normalizeState({ ...base, tasks: [task({ estimatedHours: 5, actualHours: 5, completionPct: 100 })] });
day = calculateDay('2026-08-03', state.tasks, state);
assert.ok(day.performance < 100);
assert.equal(day.dailyReward, 0);

state = normalizeState({ ...base, tasks: [task({ planType: '协同工作', collaborator: '', collaborationPct: 30 })] });
assert.ok(validateTask(state.tasks[0]).some((message) => message.includes('协同工作')));

state = normalizeState({
  ...base,
  tasks: [task()],
  monthlyReviews: { '2026-08': { collaboration: 100, loyalty: 100, discipline: 100, learning: 100, innovationBonus: 0 } },
});
const month = calculateMonth(2026, 8, state);
assert.equal(month.scoredDays, 1);
assert.equal(month.finalPerformance, 100);
assert.equal(month.payoutRatio, 1);

state = normalizeState({ ...base, calendarOverrides: { '2026-08-03': { kind: 'holiday' } }, tasks: [task()] });
day = calculateDay('2026-08-03', state.tasks, state);
assert.equal(day.standardHours, 0);
assert.equal(day.status, '非考核日有工作');

state = normalizeState({
  ...base,
  dailyPlans: { '2026-08-03': { today: [{ id: 'd1', text: '整理需求', done: true }], tomorrow: [{ id: 'd2', text: '开始开发' }] } },
});
assert.deepEqual(state.dailyPlans['2026-08-03'].today[0], { id: 'd1', text: '整理需求', done: true });
assert.deepEqual(state.dailyPlans['2026-08-03'].tomorrow[0], { id: 'd2', text: '开始开发', done: false });

state = normalizeState({
  ...base,
  tasks: [task({ order: 20, body: '# 记录', subtasks: [{ id: 's1', text: '拆分页面', done: true }], nextPlan: '继续回归', nextPlanOrder: 10, nextPlanDone: true, sourceDailyItemId: 'daily-1', frozen: true, nextPlanFrozen: true })],
  dailyPlans: { '2026-08-03': { today: [{ id: 'd1', text: '整理需求', done: false, order: 30, frozen: true }], tomorrow: [] } },
});
assert.equal(state.tasks[0].order, 20);
assert.equal(state.tasks[0].nextPlanOrder, 10);
assert.equal(state.tasks[0].nextPlanDone, true);
assert.equal(state.tasks[0].sourceDailyItemId, 'daily-1');
assert.equal(state.tasks[0].body, '# 记录');
assert.deepEqual(state.tasks[0].subtasks, [{ id: 's1', text: '拆分页面', done: true }]);
assert.equal(state.tasks[0].frozen, true);
assert.equal(state.tasks[0].nextPlanFrozen, true);
assert.equal(state.dailyPlans['2026-08-03'].today[0].order, 30);
assert.equal(state.dailyPlans['2026-08-03'].today[0].frozen, true);

assert.equal(monthKey(2026, 8), '2026-08');
assert.equal(monthLabel('2026-08'), '2026年8月');
assert.equal(dateFromExcelSerial(46237), '2026-08-03');
assert.equal(dateToISO('2026-08-03'), '2026-08-03');

const excelDate = new Date(Date.UTC(2026, 7, 3));
assert.equal(dateToISO(excelDate), '2026-08-03');

const roundTrip = normalizeState({});
assert.equal(roundTrip.tasks.length, 0);
assert.ok(DEFAULT_STATE.settings.standardDayHours > 0);

const xlsxRoundTrip = new ExcelJS.Workbook();
const sheet = xlsxRoundTrip.addWorksheet('基础信息');
sheet.getCell('B3').value = '测试用户';
sheet.getCell('D3').value = '综合管理部';
sheet.getCell('F4').value = excelDate;
const buffer = await xlsxRoundTrip.xlsx.writeBuffer();
const reloaded = new ExcelJS.Workbook();
await reloaded.xlsx.load(buffer);
const reloadedSheet = reloaded.getWorksheet('基础信息');
assert.equal(reloadedSheet.getCell('B3').value, '测试用户');
assert.equal(dateToISO(reloadedSheet.getCell('F4').value), '2026-08-03');

console.log('calc tests passed');