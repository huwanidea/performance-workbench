import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import path from 'node:path';
import fs from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const port = 4399;

const server = spawn(process.execPath, ['server.mjs'], {
  cwd: root,
  env: { ...process.env, PORT: String(port) },
  stdio: ['ignore', 'pipe', 'pipe'],
});

let serverOutput = '';
server.stdout.on('data', (chunk) => { serverOutput += chunk.toString(); });
server.stderr.on('data', (chunk) => { serverOutput += chunk.toString(); });

async function waitForServer() {
  for (let attempt = 0; attempt < 60; attempt += 1) {
    try {
      const response = await fetch(`http://127.0.0.1:${port}/api/import`, { method: 'POST', body: Buffer.alloc(0) });
      if (response.status) return;
    } catch {
      await new Promise((resolve) => setTimeout(resolve, 250));
    }
  }
  throw new Error(`服务未在预期时间内启动。输出：\n${serverOutput}`);
}

try {
  await waitForServer();

  const fixture = new ExcelJS.Workbook();
  const profile = fixture.addWorksheet('基础信息');
  profile.getCell('A3').value = '姓名';
  profile.getCell('B3').value = '张三';
  profile.getCell('C3').value = '部门';
  profile.getCell('D3').value = '综合管理部';
  profile.getCell('E3').value = '岗位';
  profile.getCell('F3').value = '项目综合管理';
  profile.getCell('A4').value = '直属负责人';
  profile.getCell('B4').value = '李经理';
  profile.getCell('C4').value = '考核年度';
  profile.getCell('D4').value = 2026;
  profile.getCell('E4').value = '考核启用日期';
  profile.getCell('F4').value = new Date(Date.UTC(2026, 0, 1));

  const params = fixture.addWorksheet('绩效参数');
  params.getCell('A4').value = '标准日工时';
  params.getCell('B4').value = 8;
  params.getCell('A5').value = '基础奖励分/小时';
  params.getCell('B5').value = 1.5;

  const dropdown = fixture.addWorksheet('下拉配置');
  dropdown.getCell('A2').value = '需求开发';
  dropdown.getCell('A3').value = 'Bug 修复';

  const tasks = fixture.addWorksheet('任务明细');
  tasks.getRow(1).values = ['日期', '今日工作内容', '任务分类', '工作类型', '协作人', '预计工时 (h)', '本人当日实际工时 (h)', '当日交付完成占比 (%)', '其中协作完成占比 (%)', '创新分 (0-100)', '本人评分', '负责人评分', '明日计划', '阻塞问题', '重大突破'];
  tasks.getRow(2).values = [new Date(Date.UTC(2026, 7, 3)), '编写月度总结', '需求开发', '单人工作', '', 7.5, 7.5, 100, 0, 0, 100, 100, '', '', ''];

  const pool = fixture.addWorksheet('任务池');
  pool.getCell('A2').value = 'w1';
  pool.getCell('B2').value = '搭建同步机制';
  pool.getCell('C2').value = '支持每日计划回写任务池';
  pool.getCell('D2').value = '需求开发';

  const plans = fixture.addWorksheet('计划');
  plans.getCell('A2').value = '今日计划';
  plans.getCell('B2').value = new Date(Date.UTC(2026, 7, 3));
  plans.getCell('E2').value = 'p1';
  plans.getCell('F2').value = '整理需求文档';
  plans.getCell('G2').value = '是';
  plans.getCell('A3').value = '周计划';
  plans.getCell('B3').value = '2026-W32';
  plans.getCell('C3').value = '第32周冲刺';
  plans.getCell('E3').value = 'p2';
  plans.getCell('F3').value = '联调同步池';
  plans.getCell('G3').value = '否';

  const monthly = fixture.addWorksheet('月度绩效');
  monthly.getCell('A4').value = '8月';
  monthly.getCell('E4').value = 90;
  monthly.getCell('F4').value = 95;
  monthly.getCell('G4').value = 92;
  monthly.getCell('H4').value = 88;

  const fixtureBuffer = Buffer.from(await fixture.xlsx.writeBuffer());

  const importResponse = await fetch(`http://127.0.0.1:${port}/api/import`, { method: 'POST', body: fixtureBuffer });
  assert.equal(importResponse.status, 200);
  const imported = await importResponse.json();

  assert.equal(imported.state.profile.name, '张三');
  assert.equal(imported.state.profile.department, '综合管理部');
  assert.equal(imported.state.profile.position, '项目综合管理');
  assert.equal(imported.state.profile.manager, '李经理');
  assert.equal(imported.state.profile.year, 2026);
  assert.equal(imported.state.profile.activationDate, '2026-01-01');
  assert.equal(imported.state.settings.standardDayHours, 8);
  assert.deepEqual(imported.state.settings.taskCategories, ['需求开发', 'Bug 修复']);
  assert.equal(imported.state.tasks.length, 1);
  assert.equal(imported.state.tasks[0].content, '编写月度总结');
  assert.equal(imported.state.tasks[0].date, '2026-08-03');
  assert.equal(imported.state.workItems.length, 1);
  assert.equal(imported.state.workItems[0].title, '搭建同步机制');
  assert.equal(imported.state.dailyPlans['2026-08-03'].today[0].text, '整理需求文档');
  assert.equal(state1Done(imported), true);
  assert.equal(imported.state.plans.weekly['2026-W32'].title, '第32周冲刺');
  assert.equal(imported.state.monthlyReviews['2026-08'].collaboration, 90);
  assert.equal(imported.state.monthlyReviews['2026-08'].discipline, 92);

  function state1Done(payload) {
    return payload.state.dailyPlans['2026-08-03'].today[0].done;
  }

  const exportResponse = await fetch(`http://127.0.0.1:${port}/api/export`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(imported.state),
  });
  assert.equal(exportResponse.status, 200);
  assert.match(exportResponse.headers.get('content-type'), /spreadsheetml/);
  const exportBuffer = Buffer.from(await exportResponse.arrayBuffer());
  assert.ok(exportBuffer.length > 5000, '导出文件体积异常');

  const exported = new ExcelJS.Workbook();
  await exported.xlsx.load(exportBuffer);
  for (const name of ['基础信息', '任务明细', '每日汇总', '月度绩效', '绩效参数', '任务池', '计划']) {
    assert.ok(exported.getWorksheet(name), `导出缺少工作表：${name}`);
  }
  const exportedTasks = exported.getWorksheet('任务明细');
  assert.equal(exportedTasks.getCell('A2').value, '2026-08-03');
  assert.equal(exportedTasks.getCell('B2').value, '编写月度总结');

  const reimportResponse = await fetch(`http://127.0.0.1:${port}/api/import`, { method: 'POST', body: exportBuffer });
  assert.equal(reimportResponse.status, 200);
  const reimported = await reimportResponse.json();
  assert.equal(reimported.state.profile.name, '张三');
  assert.equal(reimported.state.profile.department, '综合管理部');
  assert.equal(reimported.state.profile.position, '项目综合管理');
  assert.equal(reimported.state.profile.manager, '李经理');
  assert.equal(reimported.state.profile.year, 2026);
  assert.equal(reimported.state.profile.activationDate, '2026-01-01');
  assert.equal(reimported.state.settings.standardDayHours, 8);
  assert.equal(reimported.state.tasks.length, 1);
  assert.equal(reimported.state.tasks[0].content, '编写月度总结');
  assert.equal(reimported.state.tasks[0].date, '2026-08-03');
  assert.equal(reimported.state.workItems.length, 1);
  assert.equal(reimported.state.workItems[0].title, '搭建同步机制');
  assert.equal(reimported.state.dailyPlans['2026-08-03'].today[0].text, '整理需求文档');
  assert.equal(reimported.state.monthlyReviews['2026-08'].collaboration, 90);

  const badResponse = await fetch(`http://127.0.0.1:${port}/api/import`, { method: 'POST', body: Buffer.from('not a workbook') });
  assert.equal(badResponse.status, 400);

  const realTemplatePath = process.env.REAL_TEMPLATE_PATH;
  if (realTemplatePath) {
    const realBuffer = await fs.readFile(realTemplatePath);
    const realResponse = await fetch(`http://127.0.0.1:${port}/api/import`, { method: 'POST', body: realBuffer });
    assert.equal(realResponse.status, 200, '真实模板导入失败');
    const realPayload = await realResponse.json();
    console.log(`真实模板导入成功：姓名=${JSON.stringify(realPayload.state.profile.name)} 任务数=${realPayload.state.tasks.length} 警告=${realPayload.warnings.length}`);
    assert.ok(realPayload.state.tasks.length > 0, '真实模板未导入任何任务');
    console.log('real template import passed');
  }

  console.log('excel round-trip tests passed');
} finally {
  server.kill();
}