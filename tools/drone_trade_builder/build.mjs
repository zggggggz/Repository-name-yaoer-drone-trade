import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/lm/Documents/幺儿/outputs/drone_trade_startup";
const outputPath = `${outputDir}/无人机外贸启动总表.xlsx`;
const wb = Workbook.create();

const color = {
  navy: "#123047",
  teal: "#0E7C7B",
  ink: "#1F2933",
  header: "#D9E8EF",
  line: "#C8D3DA",
  light: "#EAF4F8",
  pale: "#F7FBFC",
  good: "#E7F5E8",
  warn: "#FFF4DE",
  bad: "#FCE8E1",
};

function sh(name) {
  const s = wb.worksheets.add(name);
  s.showGridLines = false;
  return s;
}

function title(s, text, note, end = "J") {
  s.getRange("A1").values = [[text]];
  s.getRange(`A1:${end}1`).format.fill = { color: color.navy };
  s.getRange(`A1:${end}1`).format.font = { color: color.navy, bold: true, size: 18 };
  s.getRange(`A1:${end}1`).format.rowHeightPx = 34;
  s.getRange("A2").values = [[note]];
  s.getRange(`A2:${end}2`).format.fill = { color: color.light };
  s.getRange(`A2:${end}2`).format.font = { color: color.ink, italic: true, size: 10 };
  s.getRange(`A2:${end}2`).format.rowHeightPx = 28;
}

function table(s, range, name) {
  const t = s.tables.add(range, true, name);
  t.style = "TableStyleMedium2";
  t.showFilterButton = true;
  s.getRange(range).format.borders = { preset: "all", style: "thin", color: color.line };
  s.getRange(range).getRow(0).format.fill = { color: color.header };
  s.getRange(range).getRow(0).format.font = { bold: true, color: color.ink };
  s.getRange(range).format.wrapText = true;
  return t;
}

function widths(s, spec) {
  for (const [col, px] of spec) s.getRange(`${col}:${col}`).format.columnWidthPx = px;
}

function statusRules(r, col, first) {
  r.conditionalFormats.addCustom(`=$${col}${first}="已完成"`, { fill: { color: color.good }, font: { color: "#2E7D32" } });
  r.conditionalFormats.addCustom(`=$${col}${first}="进行中"`, { fill: { color: color.warn }, font: { color: "#B7791F" } });
  r.conditionalFormats.addCustom(`=$${col}${first}="风险"`, { fill: { color: color.bad }, font: { color: "#C2410C", bold: true } });
}

const dash = sh("总览Dashboard");
const products = sh("产品线SKU");
const markets = sh("目标市场");
const competitors = sh("竞品分析");
const leads = sh("潜在客户");
const quotes = sh("报价&利润");
const suppliers = sh("供应商&物流");
const content = sh("内容素材");
const tasks = sh("90天任务");
const lists = sh("下拉选项");

lists.getRange("A1:H8").values = [
  ["状态", "产品类型", "客户类型", "来源", "跟进阶段", "优先级", "付款方式", "渠道"],
  ["未开始", "无人机配件", "经销商", "Alibaba", "新线索", "高", "T/T", "阿里国际站"],
  ["进行中", "农用无人机", "终端农场/合作社", "Google", "已联系", "中", "PayPal", "Google独立站"],
  ["已完成", "FPV", "系统集成商", "LinkedIn", "已报价", "低", "信用卡", "LinkedIn"],
  ["风险", "工业无人机", "政府/公共机构", "Facebook", "样品/测试", "", "L/C", "Facebook"],
  ["暂停", "整机", "维修商", "展会/社群", "谈判中", "", "Alibaba Trade Assurance", "YouTube"],
  ["", "", "品牌商/OEM", "转介绍", "已成交", "", "", "TikTok"],
  ["", "", "", "其他", "沉默/丢失", "", "", "邮件开发"],
];
table(lists, "A1:H8", "ListsTable");
widths(lists, [["A", 90], ["B", 110], ["C", 140], ["D", 110], ["E", 100], ["F", 80], ["G", 150], ["H", 120]]);

title(products, "产品线SKU", "定义主推产品、参数、认证、目标价和资料准备状态。", "N");
products.getRange("A3:N7").values = [
  ["产品ID", "产品类型", "产品名称", "核心卖点", "关键参数", "目标客户", "目标市场", "样品价USD", "批量价USD", "MOQ", "认证/合规", "资料状态", "负责人", "备注"],
  ["P-001", "无人机配件", "碳纤维螺旋桨", "轻量、耐冲击、适配多轴机型", "18-30 inch / carbon fiber", "维修商、改装商、经销商", "美国、巴西、德国", 12, 8.5, "50 pairs", "材质说明、HS Code", "进行中", "待定", "优先整理高清图和适配清单"],
  ["P-002", "农用无人机", "30L 农用喷洒无人机", "大载荷、折叠机臂、RTK 可选", "30L tank / 8-12 ha/hour", "农场、经销商、合作社", "巴西、墨西哥、东南亚", 4200, 3800, "1 set", "电池运输、CE/FCC视市场", "未开始", "待定", "第1个月重点验证需求"],
  ["P-003", "FPV", "5寸 FPV 穿越机套装", "即飞套装、竞速/航拍入门", "5 inch / 6S / ELRS optional", "FPV店铺、俱乐部、个人玩家", "美国、法国、澳大利亚", 260, 225, "5 sets", "无线电频段合规", "未开始", "待定", "需确认当地频段限制"],
  ["P-004", "工业无人机", "巡检无人机平台", "长航时、可挂载云台相机", "45-60 min / payload 1-2 kg", "电力、安防、测绘集成商", "中东、欧洲、拉美", 6800, 6200, "1 set", "CE、加密链路说明", "未开始", "待定", "高客单价，需案例资料"],
];
table(products, "A3:N7", "ProductsTable");
products.getRange("H4:I203").setNumberFormat("$#,##0.00");
products.getRange("B4:B203").dataValidation = { rule: { type: "list", formula1: "下拉选项!$B$2:$B$6" } };
products.getRange("L4:L203").dataValidation = { rule: { type: "list", formula1: "下拉选项!$A$2:$A$6" } };
statusRules(products.getRange("A4:N203"), "L", 4);
widths(products, [["A", 80], ["B", 110], ["C", 170], ["D", 220], ["E", 180], ["F", 170], ["G", 160], ["H", 95], ["I", 95], ["J", 90], ["K", 160], ["L", 90], ["M", 80], ["N", 220]]);
products.freezePanes.freezeRows(3);

title(markets, "目标市场", "按国家评估需求、监管、物流和进入优先级。", "M");
markets.getRange("A3:M8").values = [
  ["国家/地区", "主推产品", "买家画像", "需求信号", "价格带USD", "认证/监管", "物流难点", "竞争强度", "付款习惯", "进入优先级", "本月动作", "状态", "备注"],
  ["巴西", "农用无人机、配件", "大型农场、农机经销商", "农业喷洒需求强，地块大", "3,500-8,000", "ANATEL/电池运输需确认", "清关和关税较复杂", "高", "T/T 分阶段", "高", "收集10个经销商名单", "进行中", "适合农用机首批验证"],
  ["美国", "FPV、配件、工业机", "FPV店铺、维修商、集成商", "消费级/工业级需求都活跃", "50-7,000", "FCC、FAA相关要求", "电池运输要求严格", "高", "信用卡/PayPal/T/T", "高", "Google关键词和竞品页收集", "进行中", "适合配件和FPV切入"],
  ["墨西哥", "农用无人机、配件", "农资商、农业服务商", "价格敏感但需求增长", "2,800-6,000", "进口资质待确认", "可从美国/中国直发", "中", "T/T", "中", "找西语关键词和客户", "未开始", ""],
  ["德国", "配件、工业无人机", "工业集成商、经销商", "重视质量、认证、售后", "80-9,000", "CE、RoHS、电池合规", "要求文件完整", "高", "T/T/银行转账", "中", "准备英文/德文技术资料", "未开始", ""],
  ["阿联酋", "工业无人机、整机", "安防、能源、测绘公司", "高客单价项目型采购", "5,000-20,000", "进口许可/飞行许可待确认", "项目周期较长", "中", "T/T/L/C", "中", "LinkedIn搜集系统集成商", "未开始", ""],
];
table(markets, "A3:M8", "MarketsTable");
markets.getRange("J4:J203").dataValidation = { rule: { type: "list", formula1: "下拉选项!$F$2:$F$4" } };
markets.getRange("L4:L203").dataValidation = { rule: { type: "list", formula1: "下拉选项!$A$2:$A$6" } };
statusRules(markets.getRange("A4:M203"), "L", 4);
widths(markets, [["A", 100], ["B", 150], ["C", 190], ["D", 220], ["E", 115], ["F", 180], ["G", 170], ["H", 90], ["I", 130], ["J", 90], ["K", 210], ["L", 90], ["M", 180]]);
markets.freezePanes.freezeRows(3);

title(competitors, "竞品分析", "记录竞争对手价格、卖点、弱点和我们的差异化打法。", "L");
competitors.getRange("A3:L6").values = [
  ["平台", "竞品/店铺名", "国家", "产品类型", "产品/链接关键词", "标价USD", "MOQ", "核心卖点", "弱点/机会", "我们应对", "优先级", "备注"],
  ["Alibaba", "待补充", "中国", "农用无人机", "30L agricultural drone", 3999, "1 set", "价格低、参数齐", "同质化严重，案例少", "突出售后、配件包、交期", "高", "第1周补齐10家"],
  ["Google", "待补充", "美国", "FPV", "5 inch FPV drone kit", 299, "1 set", "套装清晰、零售友好", "批发政策不明显", "做B2B批量价和可定制", "中", ""],
  ["Made-in-China", "待补充", "中国", "工业无人机", "inspection drone platform", 6500, "1 set", "工业场景描述完整", "认证资料不透明", "准备案例PDF和参数对照", "中", ""],
];
table(competitors, "A3:L6", "CompetitorsTable");
competitors.getRange("F4:F203").setNumberFormat("$#,##0.00");
widths(competitors, [["A", 105], ["B", 130], ["C", 90], ["D", 110], ["E", 190], ["F", 95], ["G", 75], ["H", 190], ["I", 200], ["J", 200], ["K", 85], ["L", 170]]);
competitors.freezePanes.freezeRows(3);

title(leads, "潜在客户", "所有客户线索统一沉淀，跟进阶段和下次动作必须持续更新。", "O");
leads.getRange("A3:O5").values = [
  ["客户ID", "公司/联系人", "国家", "客户类型", "需求产品", "来源", "联系方式", "首次联系日期", "跟进阶段", "预算USD", "痛点/需求", "下一步动作", "下次跟进", "优先级", "备注"],
  ["L-001", "示例：Brazil Agri Dealer", "巴西", "经销商", "30L 农用无人机", "LinkedIn", "WhatsApp/Email 待补充", new Date("2026-06-12"), "新线索", 5000, "想比较喷洒效率和备件成本", "发送产品参数和报价", new Date("2026-06-14"), "高", "示例数据，可替换"],
  ["L-002", "示例：US FPV Store", "美国", "经销商", "FPV套装、螺旋桨", "Google", "Email 待补充", new Date("2026-06-13"), "已联系", 1200, "需要稳定供货和小批量测试", "确认MOQ和样品价", new Date("2026-06-16"), "中", "示例数据，可替换"],
];
table(leads, "A3:O5", "LeadsTable");
leads.getRange("H4:H203").setNumberFormat("yyyy-mm-dd");
leads.getRange("M4:M203").setNumberFormat("yyyy-mm-dd");
leads.getRange("J4:J203").setNumberFormat("$#,##0");
leads.getRange("D4:D203").dataValidation = { rule: { type: "list", formula1: "下拉选项!$C$2:$C$7" } };
leads.getRange("F4:F203").dataValidation = { rule: { type: "list", formula1: "下拉选项!$D$2:$D$8" } };
leads.getRange("I4:I203").dataValidation = { rule: { type: "list", formula1: "下拉选项!$E$2:$E$8" } };
leads.getRange("N4:N203").dataValidation = { rule: { type: "list", formula1: "下拉选项!$F$2:$F$4" } };
leads.getRange("A4:O203").conditionalFormats.addCustom("=$M4<TODAY()", { fill: { color: color.bad }, font: { color: "#C2410C" } });
widths(leads, [["A", 75], ["B", 175], ["C", 90], ["D", 130], ["E", 150], ["F", 100], ["G", 175], ["H", 110], ["I", 90], ["J", 95], ["K", 220], ["L", 180], ["M", 110], ["N", 75], ["O", 180]]);
leads.freezePanes.freezeRows(3);

title(quotes, "报价&利润", "录入成本、运费和报价后自动计算总报价和毛利率。", "N");
quotes.getRange("A3:N5").values = [
  ["报价ID", "日期", "客户ID", "产品ID", "产品名称", "数量", "币种", "采购成本/件", "包装物流/件", "其他成本/件", "报价/件", "总报价", "毛利率", "报价状态"],
  ["Q-001", new Date("2026-06-14"), "L-001", "P-002", "30L 农用喷洒无人机", 1, "USD", 3100, 280, 120, 4200, "", "", "草稿"],
  ["Q-002", new Date("2026-06-15"), "L-002", "P-001", "碳纤维螺旋桨", 100, "USD", 5.8, 0.7, 0.2, 9.5, "", "", "草稿"],
];
quotes.getRange("L4:L203").formulasR1C1 = [["=IF(RC[-6]=\"\",\"\",RC[-6]*RC[-1])"]];
quotes.getRange("M4:M203").formulasR1C1 = [["=IF(RC[-2]=\"\",\"\",(RC[-2]-RC[-5]-RC[-4]-RC[-3])/RC[-2])"]];
table(quotes, "A3:N5", "QuotesTable");
quotes.getRange("B4:B203").setNumberFormat("yyyy-mm-dd");
quotes.getRange("H4:L203").setNumberFormat("$#,##0.00");
quotes.getRange("M4:M203").setNumberFormat("0.0%");
quotes.getRange("M4:M203").conditionalFormats.add("cellIs", { operator: "lessThan", formula: 0.18, format: { fill: { color: color.bad }, font: { color: "#C2410C", bold: true } } });
quotes.getRange("M4:M203").conditionalFormats.add("cellIs", { operator: "greaterThanOrEqual", formula: 0.3, format: { fill: { color: color.good }, font: { color: "#2E7D32", bold: true } } });
widths(quotes, [["A", 80], ["B", 105], ["C", 80], ["D", 80], ["E", 175], ["F", 70], ["G", 70], ["H", 110], ["I", 110], ["J", 110], ["K", 95], ["L", 95], ["M", 75], ["N", 90]]);
quotes.freezePanes.freezeRows(3);

title(suppliers, "供应商&物流", "沉淀供应商、配套件、物流方式和风险，避免报价时临时找资料。", "N");
suppliers.getRange("A3:N5").values = [
  ["供应商ID", "供应商名称", "类型", "可供产品", "联系人", "联系方式", "城市", "交期", "最小起订", "付款条件", "质量风险", "物流/报关备注", "评级", "状态"],
  ["S-001", "待补充：农用机厂家", "整机/配件", "30L/50L 农用无人机、喷头、电池", "待补充", "待补充", "深圳/广州", "7-15天", "1 set", "30%订金/70%发货前", "需确认质保和备件价格", "电池需单独确认运输方案", "A", "进行中"],
  ["S-002", "待补充：FPV配件工厂", "配件", "电机、桨叶、机架、飞控", "待补充", "待补充", "深圳", "3-10天", "50 pcs", "全款/账期待谈", "批次一致性需抽检", "小件可快递，大电池限制多", "B", "未开始"],
];
table(suppliers, "A3:N5", "SuppliersTable");
suppliers.getRange("N4:N203").dataValidation = { rule: { type: "list", formula1: "下拉选项!$A$2:$A$6" } };
statusRules(suppliers.getRange("A4:N203"), "N", 4);
widths(suppliers, [["A", 85], ["B", 165], ["C", 100], ["D", 220], ["E", 90], ["F", 130], ["G", 90], ["H", 80], ["I", 90], ["J", 155], ["K", 190], ["L", 220], ["M", 65], ["N", 90]]);
suppliers.freezePanes.freezeRows(3);

title(content, "内容素材", "把产品页、社媒、视频、FAQ和资料包统一排期。", "M");
content.getRange("A3:M6").values = [
  ["素材ID", "渠道", "产品类型", "内容主题", "内容形式", "关键词/标签", "目标客户", "负责人", "截止日期", "状态", "素材链接/文件", "复用场景", "备注"],
  ["C-001", "阿里国际站", "农用无人机", "30L 农用喷洒无人机详情页", "产品页文案", "agricultural drone, spraying drone", "经销商/农场", "待定", new Date("2026-06-20"), "未开始", "", "Alibaba/独立站", "含标题、五点卖点、FAQ"],
  ["C-002", "Google独立站", "无人机配件", "无人机螺旋桨选型指南", "SEO文章", "drone propeller, carbon fiber propeller", "维修商/DIY客户", "待定", new Date("2026-06-25"), "未开始", "", "SEO/LinkedIn", ""],
  ["C-003", "YouTube", "工业无人机", "巡检无人机应用场景视频", "视频脚本", "inspection drone, industrial UAV", "系统集成商", "待定", new Date("2026-07-05"), "未开始", "", "YouTube/客户跟进", "需要实拍素材"],
];
table(content, "A3:M6", "ContentTable");
content.getRange("B4:B203").dataValidation = { rule: { type: "list", formula1: "下拉选项!$H$2:$H$8" } };
content.getRange("C4:C203").dataValidation = { rule: { type: "list", formula1: "下拉选项!$B$2:$B$6" } };
content.getRange("J4:J203").dataValidation = { rule: { type: "list", formula1: "下拉选项!$A$2:$A$6" } };
content.getRange("I4:I203").setNumberFormat("yyyy-mm-dd");
statusRules(content.getRange("A4:M203"), "J", 4);
widths(content, [["A", 75], ["B", 110], ["C", 110], ["D", 230], ["E", 110], ["F", 200], ["G", 130], ["H", 80], ["I", 105], ["J", 90], ["K", 165], ["L", 145], ["M", 180]]);
content.freezePanes.freezeRows(3);

title(tasks, "90天任务", "第1个月跑市场，第2个月开渠道，第3个月搭流程。", "L");
tasks.getRange("A3:L16").values = [
  ["任务ID", "阶段", "周次", "模块", "任务", "产出物", "负责人", "开始日期", "截止日期", "状态", "优先级", "备注"],
  ["T-001", "第1个月：跑通市场", "W1", "产品线", "整理主推SKU和参数", "产品线SKU表、产品参数", "父亲大人/幺儿", new Date("2026-06-10"), new Date("2026-06-14"), "进行中", "高", ""],
  ["T-002", "第1个月：跑通市场", "W1", "竞品", "收集阿里/Google竞品各10个", "竞品分析表", "幺儿", new Date("2026-06-10"), new Date("2026-06-16"), "未开始", "高", ""],
  ["T-003", "第1个月：跑通市场", "W2", "客户开发", "建立首批50个潜在客户名单", "潜在客户表", "父亲大人/幺儿", new Date("2026-06-17"), new Date("2026-06-23"), "未开始", "高", ""],
  ["T-004", "第1个月：跑通市场", "W3", "报价", "完成3类产品报价模板", "报价&利润表", "幺儿", new Date("2026-06-24"), new Date("2026-06-30"), "未开始", "高", ""],
  ["T-005", "第1个月：跑通市场", "W4", "复盘", "判断主推国家和主推产品", "月度复盘结论", "父亲大人/幺儿", new Date("2026-07-01"), new Date("2026-07-09"), "未开始", "高", ""],
  ["T-006", "第2个月：开通渠道", "W5", "阿里国际站", "准备店铺首页和产品详情页资料", "产品标题/详情/FAQ", "幺儿", new Date("2026-07-10"), new Date("2026-07-16"), "未开始", "高", ""],
  ["T-007", "第2个月：开通渠道", "W6", "独立站", "规划网站栏目和SEO关键词", "站点结构、关键词表", "幺儿", new Date("2026-07-17"), new Date("2026-07-23"), "未开始", "中", ""],
  ["T-008", "第2个月：开通渠道", "W7", "社媒", "建立LinkedIn/Facebook/YouTube内容节奏", "内容日历", "父亲大人/幺儿", new Date("2026-07-24"), new Date("2026-07-30"), "未开始", "中", ""],
  ["T-009", "第2个月：开通渠道", "W8", "询盘", "搭建询盘回复和客户分级模板", "询盘SOP", "幺儿", new Date("2026-07-31"), new Date("2026-08-08"), "未开始", "高", ""],
  ["T-010", "第3个月：成熟流程", "W9", "外贸流程", "完成询盘-报价-付款-发货SOP", "外贸流程SOP", "幺儿", new Date("2026-08-09"), new Date("2026-08-16"), "未开始", "高", ""],
  ["T-011", "第3个月：成熟流程", "W10", "供应链", "确认核心供应商、物流、付款方式", "供应商&物流表", "父亲大人/幺儿", new Date("2026-08-17"), new Date("2026-08-23"), "未开始", "高", ""],
  ["T-012", "第3个月：成熟流程", "W11", "订单管理", "建立订单、合同、发票和售后模板", "订单管理模板包", "幺儿", new Date("2026-08-24"), new Date("2026-08-30"), "未开始", "中", ""],
  ["T-013", "第3个月：成熟流程", "W12", "公司化", "形成每周数据复盘机制", "周报模板和指标看板", "父亲大人/幺儿", new Date("2026-08-31"), new Date("2026-09-07"), "未开始", "高", ""],
];
table(tasks, "A3:L16", "TasksTable");
tasks.getRange("H4:I203").setNumberFormat("yyyy-mm-dd");
tasks.getRange("J4:J203").dataValidation = { rule: { type: "list", formula1: "下拉选项!$A$2:$A$6" } };
tasks.getRange("K4:K203").dataValidation = { rule: { type: "list", formula1: "下拉选项!$F$2:$F$4" } };
statusRules(tasks.getRange("A4:L203"), "J", 4);
widths(tasks, [["A", 75], ["B", 150], ["C", 65], ["D", 90], ["E", 235], ["F", 160], ["G", 115], ["H", 105], ["I", 105], ["J", 90], ["K", 75], ["L", 170]]);
tasks.freezePanes.freezeRows(3);

title(dash, "无人机外贸启动总表", "3个月作战台：跑市场、开渠道、搭流程。", "J");
dash.getRange("A4:J4").merge();
dash.getRange("A4").values = [["核心指标"]];
dash.getRange("A4:J4").format.fill = { color: color.teal };
dash.getRange("A4:J4").format.font = { color: "#FFFFFF", bold: true };
dash.getRange("A5:F6").values = [
  ["产品SKU数", "", "高优先级市场", "", "潜在客户数", ""],
  ["已报价数", "", "任务完成率", "", "平均毛利率", ""],
];
dash.getRange("B5").formulas = [["=COUNTA('产品线SKU'!$A$4:$A$203)"]];
dash.getRange("D5").formulas = [["=COUNTIF('目标市场'!$J$4:$J$203,\"高\")"]];
dash.getRange("F5").formulas = [["=COUNTA('潜在客户'!$A$4:$A$203)"]];
dash.getRange("B6").formulas = [["=COUNTIF('潜在客户'!$I$4:$I$203,\"已报价\")+COUNTIF('潜在客户'!$I$4:$I$203,\"谈判中\")+COUNTIF('潜在客户'!$I$4:$I$203,\"已成交\")"]];
dash.getRange("D6").formulas = [["=IFERROR(COUNTIF('90天任务'!$J$4:$J$203,\"已完成\")/COUNTA('90天任务'!$A$4:$A$203),0)"]];
dash.getRange("F6").formulas = [["=IFERROR(AVERAGE('报价&利润'!$M$4:$M$203),0)"]];
dash.getRange("A5:F6").format.fill = { color: color.pale };
dash.getRange("A5:F6").format.borders = { preset: "all", style: "thin", color: color.line };
dash.getRange("D6:F6").setNumberFormat("0.0%");

dash.getRange("A8:E11").values = [
  ["三个月里程碑", "目标", "关键动作", "负责人", "状态"],
  ["第1个月", "跑通市场", "产品线、市场、竞品、首批客户、报价模板", "父亲大人/幺儿", "进行中"],
  ["第2个月", "开通渠道", "阿里国际站、Google独立站、社媒内容矩阵", "父亲大人/幺儿", "未开始"],
  ["第3个月", "成熟流程", "询盘、报价、付款、发货、售后、复盘SOP", "父亲大人/幺儿", "未开始"],
];
table(dash, "A8:E11", "MilestoneTable");
statusRules(dash.getRange("A9:E11"), "E", 9);

dash.getRange("G8:J15").values = [
  ["模块", "待做", "进行中", "已完成"],
  ["产品线", "", "", ""],
  ["目标市场", "", "", ""],
  ["潜在客户", "", "", ""],
  ["内容素材", "", "", ""],
  ["90天任务", "", "", ""],
  ["报价", "", "", ""],
  ["供应商", "", "", ""],
];
dash.getRange("H9:J9").formulas = [["=COUNTIF('产品线SKU'!$L$4:$L$203,\"未开始\")", "=COUNTIF('产品线SKU'!$L$4:$L$203,\"进行中\")", "=COUNTIF('产品线SKU'!$L$4:$L$203,\"已完成\")"]];
dash.getRange("H10:J10").formulas = [["=COUNTIF('目标市场'!$L$4:$L$203,\"未开始\")", "=COUNTIF('目标市场'!$L$4:$L$203,\"进行中\")", "=COUNTIF('目标市场'!$L$4:$L$203,\"已完成\")"]];
dash.getRange("H11:J11").formulas = [["=COUNTIF('潜在客户'!$I$4:$I$203,\"新线索\")+COUNTIF('潜在客户'!$I$4:$I$203,\"已联系\")", "=COUNTIF('潜在客户'!$I$4:$I$203,\"已报价\")+COUNTIF('潜在客户'!$I$4:$I$203,\"样品/测试\")+COUNTIF('潜在客户'!$I$4:$I$203,\"谈判中\")", "=COUNTIF('潜在客户'!$I$4:$I$203,\"已成交\")"]];
dash.getRange("H12:J12").formulas = [["=COUNTIF('内容素材'!$J$4:$J$203,\"未开始\")", "=COUNTIF('内容素材'!$J$4:$J$203,\"进行中\")", "=COUNTIF('内容素材'!$J$4:$J$203,\"已完成\")"]];
dash.getRange("H13:J13").formulas = [["=COUNTIF('90天任务'!$J$4:$J$203,\"未开始\")", "=COUNTIF('90天任务'!$J$4:$J$203,\"进行中\")", "=COUNTIF('90天任务'!$J$4:$J$203,\"已完成\")"]];
dash.getRange("H14:J14").formulas = [["=COUNTIF('报价&利润'!$N$4:$N$203,\"草稿\")", "=COUNTIF('报价&利润'!$N$4:$N$203,\"已发送\")+COUNTIF('报价&利润'!$N$4:$N$203,\"谈判中\")", "=COUNTIF('报价&利润'!$N$4:$N$203,\"成交\")"]];
dash.getRange("H15:J15").formulas = [["=COUNTIF('供应商&物流'!$N$4:$N$203,\"未开始\")", "=COUNTIF('供应商&物流'!$N$4:$N$203,\"进行中\")", "=COUNTIF('供应商&物流'!$N$4:$N$203,\"已完成\")"]];
table(dash, "G8:J15", "ProgressTable");

dash.getRange("A14:E20").values = [
  ["本周必须推进", "来源表", "截止日期", "状态", "备注"],
  ["整理主推SKU和参数", "90天任务 T-001", new Date("2026-06-14"), "进行中", "农用机、FPV、配件参数补齐"],
  ["收集阿里/Google竞品各10个", "90天任务 T-002", new Date("2026-06-16"), "未开始", "记录价格、MOQ、卖点、弱点"],
  ["建立首批50个潜在客户名单", "90天任务 T-003", new Date("2026-06-23"), "未开始", "国家优先：巴西、美国、墨西哥"],
  ["确认2-3个核心供应商", "供应商&物流", new Date("2026-06-20"), "未开始", "交期、质保、配件价、物流方式"],
  ["准备第一版英文公司介绍", "内容素材", new Date("2026-06-22"), "未开始", "用于客户开发和平台资料"],
  ["建立报价模板", "报价&利润", new Date("2026-06-30"), "未开始", "含EXW/FOB、样品价、批量价"],
];
table(dash, "A14:E20", "ThisWeekTable");
dash.getRange("C15:C20").setNumberFormat("yyyy-mm-dd");
statusRules(dash.getRange("A15:E20"), "D", 15);

dash.getRange("G18:J23").values = [
  ["使用顺序", "", "", ""],
  ["1. 先补产品线SKU，明确主推产品和基础报价。", "", "", ""],
  ["2. 再补目标市场和竞品分析，判断优先国家。", "", "", ""],
  ["3. 每天新增潜在客户并更新跟进阶段。", "", "", ""],
  ["4. 有询盘就进报价&利润，低毛利及时调整。", "", "", ""],
  ["5. 每周看总览页，决定下周重点。", "", "", ""],
];
dash.getRange("G18:J18").merge();
dash.getRange("G19:J23").merge(true);
dash.getRange("G18:J23").format.fill = { color: color.warn };
dash.getRange("G18:J23").format.borders = { preset: "all", style: "thin", color: color.line };
widths(dash, [["A", 135], ["B", 110], ["C", 120], ["D", 105], ["E", 180], ["F", 100], ["G", 115], ["H", 75], ["I", 75], ["J", 75]]);
dash.getRange("A1:J23").format.wrapText = true;
dash.freezePanes.freezeRows(2);

await fs.mkdir(outputDir, { recursive: true });
const check = await wb.inspect({ kind: "table", range: "总览Dashboard!A1:J23", include: "values,formulas", tableMaxRows: 24, tableMaxCols: 10 });
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 300 }, summary: "formula error scan" });
await fs.writeFile(`${outputDir}/verification.txt`, `${check.ndjson}\n${errors.ndjson}`);

for (const sheetName of ["总览Dashboard", "产品线SKU", "目标市场", "潜在客户", "报价&利润", "90天任务"]) {
  const png = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${sheetName}.png`, new Uint8Array(await png.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);
console.log(outputPath);
