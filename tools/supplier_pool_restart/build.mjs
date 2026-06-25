import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/lm/Documents/幺儿/outputs/supplier_pool_restart";
const outputPath = `${outputDir}/无人机供应商开发池_重新收集版.xlsx`;
const wb = Workbook.create();

const c = {
  navy: "#123047",
  teal: "#0E7C7B",
  header: "#D9E8EF",
  border: "#C8D3DA",
  pale: "#F7FBFC",
  good: "#E7F5E8",
  warn: "#FFF4DE",
  bad: "#FCE8E1",
  ink: "#1F2933",
};

function sheet(name) {
  const s = wb.worksheets.add(name);
  s.showGridLines = false;
  return s;
}

function title(s, text, subtitle, endCol) {
  s.getRange("A1").values = [[text]];
  s.getRange(`A1:${endCol}1`).format.fill = { color: c.navy };
  s.getRange(`A1:${endCol}1`).format.font = { color: "#FFFFFF", bold: true, size: 16 };
  s.getRange(`A1:${endCol}1`).format.rowHeightPx = 32;
  s.getRange("A2").values = [[subtitle]];
  s.getRange(`A2:${endCol}2`).format.fill = { color: c.pale };
  s.getRange(`A2:${endCol}2`).format.font = { color: c.ink, italic: true, size: 10 };
  s.getRange(`A2:${endCol}2`).format.rowHeightPx = 30;
}

function table(s, range, name) {
  const t = s.tables.add(range, true, name);
  t.style = "TableStyleMedium2";
  t.showFilterButton = true;
  s.getRange(range).format.borders = { preset: "all", style: "thin", color: c.border };
  s.getRange(range).getRow(0).format.fill = { color: c.header };
  s.getRange(range).getRow(0).format.font = { bold: true, color: c.ink };
  s.getRange(range).format.wrapText = true;
  return t;
}

function widths(s, list) {
  for (const [col, px] of list) s.getRange(`${col}:${col}`).format.columnWidthPx = px;
}

const pool = sheet("供应商池");
const verify = sheet("验厂问题");
const message = sheet("询价话术");
const source = sheet("来源说明");

title(pool, "无人机供应商开发池 - 重新收集版", "原则：中文公司名优先；中文主体或地址未能公开核验时必须标注待核验，后续联系供应商拿营业执照确认。", "V");

const rows = [
  ["S-001", "无人机配件", "T-MOTOR / Tiger Motor", "中国", "无人机电机、螺旋桨、电调、动力系统", "https://uav-en.tmotor.com/", "https://uav-en.tmotor.com/Company/AboutUs/", "官网有 UAV 动力系统品牌与公司介绍，产品线垂直，非杂货铺型平台。", "品牌方/制造商", "高", "优先", "电机、桨叶、动力套装", "中高端，适合质量背书", "官网询盘/邮件", "要阶梯价、适配表、OEM/中性包装、质保规则", "价格可能偏高", "可先作为配件线标杆供应商", "未联系", "第一批联系"],
  ["S-002", "无人机配件", "Grepow / Tattu", "中国", "无人机电池、FPV电池、工业电池包", "https://www.grepow.com/", "https://www.grepow.com/page/about-grepow.html", "公开资料显示其为电池制造商，Tattu 为无人机电池品牌，适合电池供应链。", "品牌方/制造商", "高", "优先", "FPV电池、农用/工业电池包", "中高端，物流合规要求高", "官网询盘", "要 UN38.3、MSDS、空运/海运方案、循环寿命", "电池运输和认证复杂", "电池必须做合规资料包", "未联系", "第一批联系"],
  ["S-003", "无人机配件", "Hobbywing / XRotor", "中国", "电调、无刷电机、无人机动力系统", "https://www.hobbywing.com/", "https://www.hobbywing.com/en/about-us.html", "品牌长期做电调和动力系统，XRotor 系列与无人机配件高度相关。", "品牌方/制造商", "高", "备选", "ESC、电机、动力系统", "中高端", "官网/渠道", "要外贸批发政策、授权限制、售后换新规则", "品牌渠道可能有限制", "适合做高品质配件补充", "未联系", "第二批联系"],

  ["S-004", "FPV", "iFlight", "中国", "FPV整机、机架、电机、飞控、电调、图传套装", "https://iflight-rc.com/", "https://shop.iflight.com/pages/about-us", "FPV品牌产品线完整，有整机与零配件体系，适合直接做 FPV 供应商候选。", "品牌方/制造商", "高", "优先", "5寸/7寸FPV整机、配件", "中高端，海外认知度高", "官网/经销合作", "要批发价、热销SKU、经销政策、备件包", "价格透明，利润需谈", "FPV主线优先", "未联系", "第一批联系"],
  ["S-005", "FPV", "GEPRC", "中国", "FPV整机、机架、飞控、电调、配件", "https://geprc.com/", "https://geprc.com/about-us/", "公开品牌和产品线完整，适合 FPV 整机与配件线。", "品牌方/制造商", "高", "优先", "FPV整机、机架、飞控", "中端到中高端", "官网联系", "要渠道价、保修政策、畅销型号、配件库存", "售后换件流程需确认", "适合美国/欧洲FPV客户", "未联系", "第一批联系"],
  ["S-006", "FPV", "BETAFPV", "中国", "微型FPV、TinyWhoop、遥控器、飞控、电机、电池", "https://betafpv.com/", "https://betafpv.com/pages/about-us", "品牌专注微型 FPV，产品线垂直，适合入门/室内 FPV 和教育套装。", "品牌方/制造商", "中高", "备选", "微型FPV、遥控器、配件", "中端，复购型", "官网联系", "要批量价、教育套装、热销SKU、质保", "客单价较低", "适合铺配件和小套装", "未联系", "第二批联系"],
  ["S-007", "FPV", "CaddxFPV / Walksnail", "中国", "FPV摄像头、数字图传、影像模块", "https://caddxfpv.com/", "https://caddxfpv.com/pages/about-us", "影像和图传产品线垂直，适合作为 FPV 高价值配件供应商。", "品牌方/制造商", "中高", "备选", "摄像头、图传系统", "中高端", "官网联系", "要图传系统价格、授权限制、售后换新", "兼容性和渠道限制需确认", "适合做FPV配件利润点", "未联系", "第二批联系"],

  ["S-008", "农用无人机", "XAG / 极飞", "中国", "农业无人机、农业机器人、智慧农业系统", "https://www.xa.com/", "https://www.xa.com/en/about", "农业无人机头部品牌，研发与产品体系强，适合高端农用机和行业标杆。", "品牌方/制造商", "高", "标杆/谨慎", "农用整机、农业系统", "高端", "官网渠道", "要海外代理政策、巴西适配、配件和质保政策", "合作门槛可能较高", "可作为高端标杆，不一定最适合贴牌", "未联系", "先了解政策"],
  ["S-009", "农用无人机", "DJI Agriculture", "中国", "农业无人机、Agras 系列、农业解决方案", "https://ag.dji.com/", "https://ag.dji.com/", "DJI 农业线成熟，工厂属性明确，但通常走授权渠道体系。", "品牌方/制造商", "高", "标杆/谨慎", "Agras农用无人机、配件", "高端，渠道管控强", "官网/授权渠道", "要授权经销政策、区域限制、售后条款", "不一定能直接供货", "适合作为竞品和高端参考", "未联系", "先了解政策"],
  ["S-010", "农用无人机", "JIYI UAV / 极翼", "中国", "农业飞控、喷洒系统、雷达、农用无人机控制方案", "https://www.jiyiuav.com/", "https://www.jiyiuav.com/en/about.html", "产品偏农用无人机核心控制系统，适合农用机配件和方案供应链。", "品牌方/制造商", "中高", "优先", "农用飞控、雷达、喷洒系统", "中端", "官网联系", "要整套控制系统价格、兼容清单、英文资料", "整机供应能力需确认", "适合农用配件和维修件", "未联系", "第一批联系"],

  ["S-011", "工业无人机", "JOUAV / 纵横股份", "中国", "垂直起降固定翼、测绘、巡检、应急无人机", "https://www.jouav.com/", "https://www.jouav.com/about-us/", "工业无人机制造商属性强，产品偏测绘/巡检/应急项目型。", "品牌方/制造商", "高", "标杆/谨慎", "VTOL、测绘、巡检", "高端项目型", "官网联系", "要渠道合作、项目案例、报价区间、样机政策", "客单高、周期长", "适合工业项目背书", "未联系", "先了解政策"],
  ["S-012", "工业无人机", "MMC UAV", "中国", "工业无人机、巡检/安防平台、特种无人机", "https://www.mmcuav.cn/", "https://www.mmcuav.cn/", "官网展示工业无人机产品与解决方案，适合工业整机候选。", "品牌方/制造商", "中高", "优先", "工业整机、巡检平台", "中高端", "官网联系", "要代理价、案例、交期、售后、英文资料包", "需确认当前主推型号", "适合中东/拉美工业客户", "未联系", "第一批联系"],
  ["S-013", "工业无人机", "Autel Robotics / 道通智能", "中国", "行业无人机、热成像、巡检、测绘平台", "https://www.autelrobotics.com/", "https://www.autelrobotics.com/about-us/", "成熟无人机品牌，行业产品线完善，工厂和品牌属性强。", "品牌方/制造商", "高", "备选", "工业无人机、热成像", "高端，授权体系强", "官网/授权渠道", "要授权政策、批发价、区域限制、售后规则", "渠道限制可能较强", "适合高信任客户", "未联系", "第二批联系"],
  ["S-014", "工业无人机", "ViewPro", "中国", "云台相机、热成像、激光测距、无人机载荷", "https://www.viewprouav.com/", "https://www.viewprouav.com/about-us/", "专注 UAV payload，适合作为工业无人机高毛利挂载供应商。", "品牌方/制造商", "中高", "优先", "云台、热成像、载荷", "中高端", "官网联系", "要兼容清单、SDK、项目价、质保政策", "需确认平台兼容性", "工业配件利润点", "未联系", "第一批联系"],
];

const companyInfo = {
  "S-001": ["南昌三瑞智能科技有限公司（待营业执照核验）", "官网未披露详细工厂地址；公开资料指向中国南昌/国内供应链，需联系确认", "待核验"],
  "S-002": ["深圳市格瑞普电池有限公司（待营业执照核验）", "深圳市（详细工厂地址需联系确认）", "待核验"],
  "S-003": ["深圳市好盈科技股份有限公司（待营业执照核验）", "深圳市（详细工厂地址需联系确认）", "待核验"],
  "S-004": ["惠州市翼飞智能科技有限公司（待营业执照核验）", "惠州/深圳供应链区域；详细工厂地址需联系确认", "待核验"],
  "S-005": ["深圳市格普科技有限公司（待营业执照核验）", "深圳市（详细工厂地址需联系确认）", "待核验"],
  "S-006": ["中文主体待核验（BETAFPV）", "深圳市（详细工厂地址需联系确认）", "待核验"],
  "S-007": ["中文主体待核验（CaddxFPV / Walksnail）", "深圳市（详细工厂地址需联系确认）", "待核验"],
  "S-008": ["广州极飞科技股份有限公司（待营业执照核验）", "广州市（详细工厂/办公地址需联系确认）", "待核验"],
  "S-009": ["深圳市大疆创新科技有限公司 / 大疆农业（授权供货主体待确认）", "深圳市；农业线供货/授权主体和工厂地址需联系确认", "待核验"],
  "S-010": ["上海极翼智能科技有限公司（待营业执照核验）", "中国（详细工厂/办公地址需联系确认）", "待核验"],
  "S-011": ["成都纵横自动化技术股份有限公司（待营业执照核验）", "成都市（详细工厂/办公地址需联系确认）", "待核验"],
  "S-012": ["深圳市科卫泰实业发展有限公司（待营业执照核验）", "深圳市（详细工厂/办公地址需联系确认）", "待核验"],
  "S-013": ["深圳市道通智能航空技术股份有限公司（待授权主体确认）", "深圳市（详细工厂/办公地址需联系确认）", "待核验"],
  "S-014": ["深圳市视普泰科技有限公司（待营业执照核验）", "深圳市（详细工厂/办公地址需联系确认）", "待核验"],
};

const enrichedRows = rows.map((row) => {
  const [companyName, factoryAddress, addressStatus] = companyInfo[row[0]] ?? ["待核验", "待核验", "待核验"];
  return [row[0], row[1], row[2], companyName, factoryAddress, addressStatus, ...row.slice(3)];
});

pool.getRange("A3:V17").values = [
  ["编号", "品类", "品牌/供应商简称", "供应商中文公司名", "工厂/办公地址", "地址核验状态", "国家", "主营产品", "官网", "证据链接", "公开工厂判断依据", "属性判断", "置信度", "优先级", "适合采购", "价格带判断", "联系渠道", "必须询问", "主要风险", "使用建议", "状态", "批次"],
  ...enrichedRows,
];
table(pool, "A3:V17", "SupplierPoolRestartTable");
pool.freezePanes.freezeRows(3);
pool.getRange("M4:N103").conditionalFormats.addCustom('=$N4="优先"', { fill: { color: c.good }, font: { color: c.ink, bold: true } });
pool.getRange("M4:N103").conditionalFormats.addCustom('=$N4="标杆/谨慎"', { fill: { color: c.warn }, font: { color: c.ink, bold: true } });
pool.getRange("M4:M103").conditionalFormats.addCustom('=$M4="高"', { font: { color: c.ink, bold: true } });
widths(pool, [["A", 70], ["B", 100], ["C", 160], ["D", 230], ["E", 280], ["F", 110], ["G", 70], ["H", 230], ["I", 220], ["J", 240], ["K", 300], ["L", 125], ["M", 75], ["N", 95], ["O", 190], ["P", 140], ["Q", 120], ["R", 260], ["S", 190], ["T", 220], ["U", 90], ["V", 100]]);

title(verify, "验厂和反二道贩子问题", "联系供应商时必须问。答不上来、拒绝提供证据、只催付款的，直接降级。", "G");
verify.getRange("A3:G13").values = [
  ["序号", "问题", "合格回答", "危险信号", "要拿到的证据", "适用品类", "备注"],
  [1, "你们是工厂、品牌方、还是代理/贸易商？", "明确说明身份并能提供公司资料", "含糊其辞，只说都有货", "营业执照、官网、公司抬头报价", "全部", "第一问就要问清楚"],
  [2, "能否提供工厂/仓库/装配线视频？", "可提供当天视频或安排视频会议", "只发宣传图，不给实时视频", "视频、照片、视频会议", "全部", "最重要"],
  [3, "是否支持 OEM 或中性包装？", "说明支持范围、MOQ、费用", "只会说可以但不给细节", "包装图、打样费用、MOQ", "配件/FPV/农用", "做外贸很关键"],
  [4, "样品和批量交期分别多久？", "明确天数和库存状态", "只说很快、有货", "样品交期、批量交期", "全部", ""],
  [5, "请给阶梯报价：样品/10/50/100。", "能给结构化报价和有效期", "只口头报价或价格乱跳", "正式报价单", "全部", "用于毛利测算"],
  [6, "质保政策是什么？易损件价格表有吗？", "能给质保期限、范围、配件价", "坏了再说、没有配件表", "质保文件、配件清单", "全部", "售后决定复购"],
  [7, "是否有英文资料包？", "图片、视频、参数、说明书齐全", "只给中文截图或低清图", "英文目录、图片、视频、说明书", "全部", "上架平台必需"],
  [8, "电池/带电产品有哪些认证？", "UN38.3、MSDS、运输方案明确", "不了解运输要求", "UN38.3、MSDS、物流方案", "电池/整机", "一票否决项"],
  [9, "是否能提供海外客户案例或出货记录？", "可提供国家/应用场景/案例资料", "没有任何案例", "案例PDF、发货记录可打码", "农用/工业", "提高客户信任"],
  [10, "能否签采购合同并开正规发票？", "可以，主体一致", "要求私人收款", "合同、发票、收款账户", "全部", "资金安全"],
];
table(verify, "A3:G13", "VerifyQuestionsTable");
verify.freezePanes.freezeRows(3);
widths(verify, [["A", 70], ["B", 250], ["C", 240], ["D", 230], ["E", 210], ["F", 120], ["G", 160]]);

title(message, "首次询价话术", "父亲大人复制后发给供应商即可；如果对方回复完整，就进入报价和样品评估。", "F");
message.getRange("A3:F8").values = [
  ["场景", "中文话术", "English Message", "必须拿到", "判断重点", "备注"],
  ["首次联系", "你好，我们是做无人机外贸的，目前在筛选长期供应商。请问贵司是工厂/品牌方，还是代理/贸易商？我们主要关注无人机配件、FPV、农用无人机和工业无人机。请发公司介绍、主推产品目录、阶梯报价、MOQ、交期和质保政策。", "Hello, we are building a long-term supplier base for drone export business. Are you a factory/brand manufacturer, or a trading company? We focus on drone parts, FPV drones, agricultural drones, and industrial UAVs. Please share your company profile, main product catalog, tier pricing, MOQ, lead time, and warranty policy.", "公司介绍、目录、报价、MOQ、交期、质保", "是否直接说明身份", "所有供应商通用"],
  ["工厂确认", "为了进入核心供应商名单，我们需要确认生产/装配能力。请提供营业执照、工厂或仓库视频、生产/装配区照片、质检流程，并说明是否可以视频会议看厂。", "To qualify you as a core supplier, we need to verify production or assembly capability. Please provide business license, factory or warehouse video, production/assembly photos, QC process, and confirm whether a video factory tour is available.", "营业执照、视频、质检流程", "防止二道贩子", "不配合则降级"],
  ["报价要求", "请按样品、10套、50套、100套分别报价，注明 EXW 和 FOB 价格、包装尺寸重量、付款方式、报价有效期。电池或带电产品请附 UN38.3/MSDS 和可用物流方案。", "Please quote sample, 10 units, 50 units, and 100 units separately. Include EXW and FOB prices, package size/weight, payment terms, and quote validity. For batteries or products with batteries, please attach UN38.3/MSDS and available shipping solutions.", "阶梯价、包装、付款、认证", "能否结构化报价", "用于利润测算"],
  ["资料要求", "我们需要用于阿里国际站和独立站上架的英文资料：高清图片、短视频、参数表、说明书、认证、FAQ、易损件清单和配件价格表。", "We need English materials for Alibaba and our independent website: HD photos, short videos, spec sheets, manuals, certificates, FAQ, wearing parts list, and spare parts price list.", "英文资料包", "是否支持外贸上架", "资料差会影响获客"],
  ["样品推进", "如果资料和价格合适，我们会先采购样品测试。请说明样品价格、样品交期、发货方式、售后处理，以及后续批量订单是否可抵扣样品差价。", "If documents and pricing are suitable, we will purchase samples first. Please confirm sample price, sample lead time, shipping method, after-sales process, and whether future bulk orders can offset sample cost differences.", "样品政策、物流、售后", "能否低风险启动", "进入样品池"],
];
table(message, "A3:F8", "MessageRestartTable");
message.freezePanes.freezeRows(3);
widths(message, [["A", 95], ["B", 380], ["C", 380], ["D", 180], ["E", 170], ["F", 150]]);

title(source, "来源说明", "本表为第一轮公开资料初筛，不等同于最终验厂。下一步必须由父亲大人联系供应商拿证据。", "F");
source.getRange("A3:F9").values = [
  ["说明", "内容", "风险", "下一步", "负责人", "状态"],
  ["来源类型", "优先使用供应商官网、公司介绍页、品牌/产品官网。", "官网也可能是营销信息", "联系时索要营业执照和视频验厂", "父亲大人", "待执行"],
  ["工厂判断", "本表的“品牌方/制造商”是公开资料初判。", "不能替代实地验厂或第三方审厂", "用验厂问题表逐项确认", "父亲大人/幺儿", "待执行"],
  ["二道贩子识别", "拒绝提供公司资料、工厂视频、正式报价、质保政策的供应商要降级。", "低价诱惑可能导致售后崩盘", "先样品后批量", "父亲大人", "待执行"],
  ["第一批策略", "优先联系标记为“优先”的供应商，标杆/谨慎类用于了解政策和做高端参考。", "大品牌可能有授权限制", "先拿价格和渠道规则", "父亲大人", "待执行"],
  ["资料用途", "拿到的图片、视频、参数和报价会进入产品目录、阿里国际站、独立站和客户报价体系。", "资料不完整无法做平台", "建立资料包文件夹", "幺儿", "待执行"],
  ["成功标准", "每个品类筛出2家核心供应商+1家备用供应商。", "只靠一家供应商风险高", "完成A/B/C评级", "父亲大人/幺儿", "待执行"],
];
table(source, "A3:F9", "SourceNotesTable");
source.freezePanes.freezeRows(3);
widths(source, [["A", 120], ["B", 330], ["C", 260], ["D", 220], ["E", 120], ["F", 90]]);

await fs.mkdir(outputDir, { recursive: true });
const check = await wb.inspect({ kind: "table", range: "供应商池!A3:V17", include: "values", tableMaxRows: 20, tableMaxCols: 22 });
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "error scan" });
await fs.writeFile(`${outputDir}/verification.txt`, `${check.ndjson}\n${errors.ndjson}`);
for (const sheetName of ["供应商池", "验厂问题", "询价话术", "来源说明"]) {
  const png = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${sheetName}.png`, new Uint8Array(await png.arrayBuffer()));
}
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);
console.log(outputPath);
