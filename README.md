# 绩效考核工作台 · Performance Workbench

> 把「每日做了什么」自动折算成「这个月绩效多少、拿多少奖励」的一体化工具。
> 任务记录、当日绩效核算、月度综合素养评分、Excel 归档导入导出，全部在浏览器里完成。

[![CI](https://github.com/huwanidea/performance-workbench/actions/workflows/ci.yml/badge.svg)](https://github.com/huwanidea/performance-workbench/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Node](https://img.shields.io/badge/node-%3E%3D20-339933?logo=node.js)](https://nodejs.org/)

---

## 这个工具解决什么问题

绩效考核最烦的不是填表，而是**算账**：

- 每天的任务工时怎么折算成当日绩效分？
- 月底了，这个月一共上了多少有效工时，绩效工资怎么算？
- 协同、忠诚度、纪律性、学习力这些主观项怎么和日常表现加权？
- 换模板、换参数后，历史数据怎么对齐？

这个工作台把这些规则写进了代码：**改参数立刻全局重算**，不需要在 Excel 里拖公式。

## 核心能力

| 模块 | 说明 |
| --- | --- |
| 每日工作台 | 今日任务 / 明日计划 / 周计划 / 月计划四象限布局，拖拽调整顺序，面板可折叠可冻结 |
| 绩效自动核算 | 按计划类型（单人 / 协同 / 拓展 / 加班）分别计算有效工时，输出当日绩效分与奖励分 |
| 月度绩效 | 综合素养四维（协同 / 忠诚 / 纪律 / 学习）加权，叠加创新加分，给出最终绩效与兑现比例 |
| 绩效看板 | 全年 365 天逐日趋势，识别异常日期（缺卡、工时异常、协同缺协作人） |
| 任务池 | 需求 / 任务分离，任务可挂验收人，支持验收状态流转 |
| 计划体系 | 日 / 周 / 月计划三层结构，今日计划可一键同步到任务池 |
| Excel 归档 | 导入既有绩效模板，或导出归档文件**再原样导回继续编辑** |
| 多用户工作空间 | 本地多套独立数据空间，切换互不干扰 |
| 中英双语 | 基于 i18next，界面语言可切换 |

## 技术栈

- **前端**：React 19 · TypeScript 5.8 · Vite 8 · Tailwind CSS 3
- **服务端**：Node.js 原生 `http`（零框架依赖）· ExcelJS
- **质量**：ESLint（typescript-eslint 严格模式）· `tsc --noEmit` 类型检查 · Node 内置测试运行器
- **数据**：浏览器 `localStorage` 持久化，无需数据库

## 快速开始

环境要求：Node.js 20 或更高版本。

```bash
git clone https://github.com/huwanidea/performance-workbench.git
cd performance-workbench
npm install

# 只跑前端开发服务器
npm run dev

# 构建 + 启动完整服务（含 Excel 导入导出接口）
npm run serve
```

启动后访问 <http://localhost:4173>。

## 常用命令

| 命令 | 作用 |
| --- | --- |
| `npm run dev` | 启动 Vite 开发服务器 |
| `npm run build` | 构建到 `out/` 目录 |
| `npm run serve` | 构建并启动带 Excel 接口的完整服务 |
| `npm run server` | 仅启动 Excel 服务（需先build） |
| `npm test` | 运行全部测试（计分引擎 + Excel 往返） |
| `npm run type-check` | TypeScript 类型检查 |
| `npm run lint` | ESLint 检查 |

## 架构说明

```
src/
├─ performance/
│  ├─ calc.js          ← 计分引擎：纯函数，无副作用，测试覆盖率的重点
│  ├─ app.js           ← 应用主逻辑与视图渲染
│  ├─ theme.css        ← 设计令牌与主题变量
│  └─ ...              ← 各功能模块的样式与逻辑
├─ pages/home/page.tsx ← 页面入口与 DOM 骨架
├─ router/             ← 路由配置
└─ i18n/               ← 国际化配置

server.mjs              ← Excel 导入导出 HTTP 服务
test/
├─ calc.test.mjs        ← 计分引擎单元测试
└─ excel.test.mjs       ← Excel 导入导出端到端测试（含往返验证）
```

**计分引擎与视图是分离的**：`calc.js` 全部导出为纯函数，不碰 DOM、不读`localStorage`，因此可以被直接单元测试。

## 几个值得一提的实现细节

### Excel 模板兼容性

导入逻辑不是硬编码单元格位置，而是**按表头文字动态定位字段**。原因是不同版本的绩效模板列顺序和表头行位置都不同（实测同一项目内就存在表头在第 1 行和第 6 行两种布局）。

此外还处理了两类真实模板里会出现的问题：

- **命名空间前缀**：部分工具导出的 xlsx 内部使用 `x:` 命名空间前缀（`<x:workbook>` 而非 `<workbook>`），标准解析器会直接失败，导入前会先做规范化。
- **重叠合并单元格**：部分导出工具会产生互相重叠的合并区域，会让解析器抛错，导入时会剔除冲突的合并声明。

### 归档可回读

导出的归档文件可以被**原样导回**继续编辑——导出布局与导入契约严格对齐，并附带自动化往返测试守住这条链路。

## 测试

```bash
npm test
```

- `test/calc.test.mjs`：覆盖计分引擎的绩效折算、工时校验、综合素养加权、日历覆盖等逻辑，含日期与 Excel 序列号互转。
- `test/excel.test.mjs`：真实启动 HTTP 服务，跑通「导入 → 导出 → 再导入」完整往返，并验证字段值、异常处理。

## 隐私

所有数据保存在浏览器 `localStorage`，不会上传到任何服务器。Excel 导入导出全程在本地服务与浏览器之间完成。

## License

[MIT](LICENSE)