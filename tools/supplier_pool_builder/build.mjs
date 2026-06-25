import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/lm/Documents/幺儿/outputs/supplier_pool";
const outputPath = `${outputDir}/无人机供应商开发池_第一版.xlsx`;
const wb = Workbook.create();

const colors = {
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

function addSheet(name) {
  const sheet = wb.worksheets.add(name);
  sheet.showGridLines = false;
  return sheet;
}

function setTitle(sheet, title, subtitle, endCol) {
  sheet.getRange("A1").values = [[title]];
  sheet.getRange(`A1:${endCol}1`).format.fill = { color: colors.navy };
  sheet.getRange(`A1:${endCol}1`).format.font = { color: "#FFFFFF", bold: true, size: 16 };
  sheet.getRange(`A1:${endCol}1`).format.rowHeightPx = 32;
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange(`A2:${endCol}2`).format.fill = { color: colors.pale };
  sheet.getRange(`A2:${endCol}2`).format.font = { color: colors.ink, italic: true, size: 10 };
  sheet.getRange(`A2:${endCol}2`).format.rowHeightPx = 28;
}

function styleTable(sheet, range, name) {
  const table = sheet.tables.add(range, true, name);
  table.style = "TableStyleMedium2";
  table.showFilterButton = true;
  sheet.getRange(range).format.borders = { preset: "all", style: "thin", color: colors.border };
  sheet.getRange(range).getRow(0).format.fill = { color: colors.header };
  sheet.getRange(range).getRow(0).format.font = { bold: true, color: colors.ink };
  sheet.getRange(range).format.wrapText = true;
  return table;
}

function widths(sheet, spec) {
  for (const [col, px] of spec) sheet.getRange(`${col}:${col}`).format.columnWidthPx = px;
}

const pool = addSheet("供应商池");
const criteria = addSheet("筛选标准");
const message = addSheet("询价话术");
const next = addSheet("下一步动作");

setTitle(pool, "无人机供应商开发池 - 第一版", "初筛目标：优先选择有研发/制造/生产能力的公司，排除纯二道贩子；所有候选仍需通过营业执照、工厂视频、报价和样品二次确认。", "R");

const suppliers = [
  ["S-001", "无人机配件", "T-MOTOR / Tiger Motor", "中国", "南昌/全球销售", "无人机电机、螺旋桨、电调、动力系统", "https://uav-en.tmotor.com/", "https://uav-en.tmotor.com/Company/AboutUs/", "官网品牌线清晰，长期专注 UAV 动力系统，适合作为电机/桨叶/动力套装供应商", "强工厂/品牌方", "高", "官网联系/询盘", "重点要批量价、代理政策、适配清单", "电机、桨叶、ESC", "中高", "优先", "询问是否支持中性包装、OEM标、海外售后配件包", "父亲大人可优先联系"],
  ["S-002", "无人机配件", "Grepow / Tattu", "中国", "深圳", "无人机/FPV/工业电池、电池包", "https://www.grepow.com/", "https://www.grepow.com/page/about-grepow.html", "公开资料显示为电池制造商，Tattu 是其无人机电池品牌，适合电池供应链", "强工厂/品牌方", "高", "官网联系/Alibaba", "电池运输、认证和最小起订需重点确认", "无人机电池、FPV电池", "中高", "优先", "必须要 UN38.3/MSDS/空运海运方案", "适合做配件和整机备件包"],
  ["S-003", "无人机配件", "Hobbywing / XRotor", "中国", "深圳", "无人机电调、动力系统、无刷电机", "https://www.hobbywing.com/", "https://www.hobbywing.com/en/about-us.html", "品牌产品线覆盖无人机动力和电调，适合 FPV/工业/农用动力配件", "强工厂/品牌方", "中高", "官网联系/渠道询价", "品牌强但价格可能不低，需确认批发门槛", "ESC、电机、动力系统", "中高", "备选", "询问经销价、最小订单、是否允许外贸转售", "适合做高品质配件"],
  ["S-004", "无人机配件", "CUAV", "中国", "广州/深圳周边", "飞控、RTK、自动驾驶模块、无人机电子系统", "https://www.cuav.net/", "https://www.cuav.net/en/about/", "官网产品以飞控和导航电子为主，适合作为工业/农用无人机核心电子配件供应商", "强工厂/品牌方", "中高", "官网联系", "需确认技术支持响应和英文资料完整度", "飞控、RTK、数传", "中", "备选", "重点要技术文档、SDK/兼容清单", "适合工业机配套"],

  ["S-005", "FPV", "iFlight", "中国", "惠州/深圳供应链", "FPV整机、机架、电机、飞控、电调、图传套装", "https://iflight-rc.com/", "https://shop.iflight.com/pages/about-us", "FPV品牌方，产品线覆盖整机和零配件，适合做 FPV 套装与配件", "强工厂/品牌方", "高", "官网/经销渠道", "品牌货价格透明，利润空间需谈渠道价", "FPV整机、套装、零件", "中高", "优先", "询问批发折扣、海外经销政策、备件包", "FPV主线优先候选"],
  ["S-006", "FPV", "GEPRC", "中国", "深圳", "FPV整机、机架、飞控、电调、图传、配件", "https://geprc.com/", "https://geprc.com/about-us/", "公开品牌和产品线完整，FPV圈认知度高，可作为 FPV 整机/配件供应商", "强工厂/品牌方", "高", "官网联系", "需确认渠道价和售后换件流程", "FPV整机、机架、飞控", "中", "优先", "询问热销型号、批发价、是否可定制", "适合美国/欧洲 FPV 客户"],
  ["S-007", "FPV", "BETAFPV", "中国", "深圳", "微型 FPV、TinyWhoop、遥控器、飞控、电机、电池", "https://betafpv.com/", "https://betafpv.com/pages/about-us", "品牌专注微型 FPV 产品，适合做入门/室内 FPV 和教育套装", "强工厂/品牌方", "中高", "官网联系", "客单价较低但复购高，需控制库存型号", "微型FPV、遥控器、配件", "中", "备选", "询问热销套装、批量价、渠道保护", "适合小B客户"],
  ["S-008", "FPV", "CaddxFPV / Walksnail", "中国", "深圳", "FPV摄像头、数字图传系统、影像模块", "https://caddxfpv.com/", "https://caddxfpv.com/pages/about-us", "影像/图传类品牌，适合补齐 FPV 高价值配件线", "强工厂/品牌方", "中高", "官网联系", "需确认授权和售后换新政策", "FPV摄像头、图传", "中", "备选", "询问经销价、保修规则、热门 SKU", "适合做配件增利"],

  ["S-009", "农用无人机", "XAG / 极飞", "中国", "广州", "农业无人机、农业机器人、智慧农业系统", "https://www.xa.com/", "https://www.xa.com/en/about", "农业无人机头部品牌，研发和产品体系强，适合作为标杆和高端供应链参考", "强工厂/品牌方", "高", "官网渠道/海外合作", "品牌强但合作门槛可能高，未必适合贴牌", "农用整机、农业系统", "高", "标杆/谨慎", "询问海外代理、渠道采购、配件供应政策", "可作为竞品/供应链标杆"],
  ["S-010", "农用无人机", "EAVISION / 易瓦特农业视觉", "中国", "苏州/全球农业市场", "农业喷洒无人机、果树/山地场景无人机", "https://www.eavisionag.com/", "https://www.eavisionag.com/about.html", "专注农业无人机和视觉/避障场景，适合差异化农用机供应商候选", "强工厂/品牌方", "中高", "官网联系", "需确认外贸合作政策和价格带", "农用喷洒无人机", "中高", "优先", "询问30L/40L型号、巴西适配、配件价", "差异化卖点较强"],
  ["S-011", "农用无人机", "EFT / EFT UAV", "中国", "农业无人机供应链", "农业无人机平台、喷洒系统、播撒系统、配件", "https://www.eftuav.com/", "https://www.eftuav.com/about-us/", "农用无人机平台和配件供应商属性明显，可能更适合外贸/OEM合作", "疑似强工厂/需确认", "中高", "官网/Alibaba询价", "需要核验工厂视频、营业执照、产能", "农用平台、喷洒配件", "中", "优先", "重点问OEM、散件、SKD、配件包", "很适合先询价"],
  ["S-012", "农用无人机", "JIYI UAV / 极翼", "中国", "农业无人机控制系统供应链", "农业飞控、喷洒系统、雷达、农用无人机解决方案", "https://www.jiyiuav.com/", "https://www.jiyiuav.com/en/about.html", "产品偏农用飞控和整机控制系统，适合作为农用核心配件/方案供应商", "强工厂/品牌方", "中高", "官网联系", "整机能力和外贸价格需二次确认", "农用飞控、雷达、喷洒系统", "中", "备选", "询问控制系统套件价、英文资料、适配品牌", "适合做农用配件线"],

  ["S-013", "工业无人机", "JOUAV / 纵横股份", "中国", "成都", "垂直起降固定翼、测绘、巡检、应急工业无人机", "https://www.jouav.com/", "https://www.jouav.com/about-us/", "上市/知名工业无人机制造商，工业场景资料完整，适合高端项目型供应链参考", "强工厂/品牌方", "高", "官网联系/渠道合作", "合作门槛和价格高，适合项目单而非低价铺货", "测绘、巡检、VTOL", "高", "标杆/谨慎", "询问渠道合作、项目资料、样机政策", "适合工业项目背书"],
  ["S-014", "工业无人机", "MMC UAV", "中国", "深圳", "工业无人机、氢燃料无人机、巡检/安防平台", "https://www.mmcuav.com/", "https://www.mmcuav.com/about-us/", "工业无人机品牌方，产品偏专业场景，适合工业整机和项目型客户", "强工厂/品牌方", "中高", "官网联系", "需确认当前产品线、交期和售后政策", "工业整机、巡检平台", "中高", "优先", "询问代理价、项目支持、案例资料", "适合中东/拉美工业客户"],
  ["S-015", "工业无人机", "Autel Robotics / 道通智能", "中国", "深圳", "行业无人机、热成像、巡检、测绘平台", "https://www.autelrobotics.com/", "https://www.autelrobotics.com/about-us/", "品牌和产品成熟，行业机产品资料完善，适合高端工业无人机供应链", "强工厂/品牌方", "高", "官网/授权渠道", "授权体系严格，需确认是否能做转售", "工业无人机、热成像", "高", "备选", "询问授权、批发价、区域限制", "适合高信任客户"],
  ["S-016", "工业无人机", "ViewPro", "中国", "深圳", "无人机吊舱、云台相机、热成像、激光测距载荷", "https://www.viewprouav.com/", "https://www.viewprouav.com/about-us/", "专注 UAV payload，适合工业无人机高毛利配件/挂载供应商", "强工厂/品牌方", "中高", "官网联系", "需确认不同飞控/平台兼容性", "云台、热成像、载荷", "中高", "优先", "询问兼容清单、SDK、项目价", "工业配件利润点"],
];

pool.getRange("A3:R19").values = [
  ["编号", "品类", "供应商", "国家", "城市/区域", "主营产品", "官网", "证据链接", "工厂判断依据", "工厂属性判断", "置信度", "联系渠道", "关键风险", "适合采购", "价格带判断", "优先级", "下一步询问重点", "备注"],
  ...suppliers,
];
styleTable(pool, "A3:R19", "SupplierPoolTable");
pool.getRange("A4:R103").conditionalFormats.addCustom('=$K4="高"', { fill: { color: colors.good }, font: { color: "#2E7D32" } });
pool.getRange("A4:R103").conditionalFormats.addCustom('=$J4="疑似强工厂/需确认"', { fill: { color: colors.warn }, font: { color: "#B7791F" } });
pool.getRange("A4:R103").conditionalFormats.addCustom('=$P4="优先"', { font: { bold: true, color: colors.teal } });
pool.freezePanes.freezeRows(3);
widths(pool, [["A", 70], ["B", 100], ["C", 150], ["D", 70], ["E", 110], ["F", 230], ["G", 210], ["H", 220], ["I", 300], ["J", 130], ["K", 75], ["L", 120], ["M", 230], ["N", 160], ["O", 90], ["P", 80], ["Q", 240], ["R", 180]]);

setTitle(criteria, "供应商筛选标准", "用于判断对方是工厂、品牌方，还是二道贸易商。联系前先看公开资料，联系后必须拿证据。", "H");
criteria.getRange("A3:H14").values = [
  ["维度", "强工厂/品牌方信号", "贸易商/二道贩子风险信号", "我们要的证据", "权重", "通过标准", "备注", "执行人"],
  ["公司身份", "官网有研发、制造、生产、品牌历史、产品线", "只展示杂乱 SKU，无研发生产信息", "营业执照、工厂视频、公司抬头报价", 20, "能提供完整公司资料", "先看官网和工商信息", "幺儿/父亲大人"],
  ["产品线", "产品系列完整，有参数、手册、配件体系", "什么都卖，型号混乱，图片来自不同品牌", "产品目录、参数表、说明书", 15, "至少有清晰主线产品", "越垂直越好", "幺儿"],
  ["价格", "能给阶梯价、EXW/FOB、样品价", "只会口头报价，价格波动大", "报价单、有效期、MOQ", 15, "价格结构清楚", "便于外贸报价", "父亲大人"],
  ["交期", "能说明库存、生产周期、备件周期", "只说有货但不能给交期依据", "交期表、库存说明", 10, "样品和批量交期明确", "", "父亲大人"],
  ["资料", "有英文图册、图片、视频、认证资料", "只给低清图片，资料不愿给", "英文资料包、高清图、视频", 15, "能支持平台上架", "阿里国际站必须要", "幺儿"],
  ["质保", "质保条款明确，能提供配件价", "售后只说坏了再看", "质保政策、易损件价格", 10, "质保可写进报价", "", "父亲大人"],
  ["OEM/中性包装", "支持贴牌、中性包装、定制", "只能原品牌零售出货", "包装图、贴牌规则", 10, "至少支持中性包装", "品牌货可例外", "父亲大人"],
  ["沟通响应", "24小时内清楚回复", "拖延、答非所问、只催付款", "聊天记录、报价响应", 5, "能持续配合", "", "父亲大人"],
  ["一票否决", "无", "拒绝提供公司资料、拒绝视频验厂、报价明显异常低", "视频会议/验厂/第三方认证", "一票否决", "不进入核心供应商", "低价陷阱要避开", "父亲大人"],
  ["最终分级", "A=核心；B=备选；C=观察；D=淘汰", "无", "评分表", "合计100", "A≥80，B≥65", "先筛每类2主1备", "幺儿"],
  ["第一阶段目标", "每个品类至少2家核心+1家备用", "无", "完整供应商池", "无", "可报价、可发样、可售后", "这是外贸底盘", "父亲大人/幺儿"],
];
styleTable(criteria, "A3:H14", "CriteriaTable");
criteria.getRange("E4:E11").setNumberFormat("0");
criteria.freezePanes.freezeRows(3);
widths(criteria, [["A", 120], ["B", 260], ["C", 250], ["D", 200], ["E", 80], ["F", 150], ["G", 170], ["H", 110]]);

setTitle(message, "供应商询价话术", "复制后按品类微调。目标不是客气聊天，而是一次性拿到能报价、能上架、能判断是否工厂的资料。", "F");
message.getRange("A3:F7").values = [
  ["场景", "中文话术", "英文话术", "必须索取资料", "判断点", "备注"],
  ["首次联系", "你好，我们是做无人机外贸的，目前在建立长期供应链。想了解你们是否是工厂或品牌方，是否支持外贸批发/OEM/中性包装？请发公司介绍、主推产品目录、阶梯报价、MOQ、交期和质保政策。", "Hello, we are building a long-term supply chain for drone export business. Are you a factory or brand manufacturer? Do you support wholesale, OEM, or neutral packaging? Please share your company profile, main product catalog, tier pricing, MOQ, lead time, and warranty policy.", "公司介绍、目录、报价、MOQ、交期、质保", "是否直接回答工厂/品牌方，资料是否完整", "适用于所有品类"],
  ["确认工厂", "为了筛选核心供应商，我们需要确认生产能力。请提供营业执照、工厂/仓库视频、生产线或装配区照片、质检流程，以及是否可以视频会议看厂。", "To qualify core suppliers, we need to verify production capability. Please provide business license, factory/warehouse video, production or assembly photos, QC process, and whether a video factory tour is available.", "营业执照、工厂视频、质检流程", "不愿提供则降级", "防二道贩子"],
  ["要报价", "请按样品、10套/50套/100套分别报价，注明EXW和FOB价格、包装尺寸重量、付款方式、报价有效期。电池类请附UN38.3/MSDS。", "Please quote sample, 10 units, 50 units, and 100 units separately. Include EXW and FOB price, package size/weight, payment terms, and quote validity. For batteries, please attach UN38.3/MSDS.", "阶梯价、包装、付款、认证", "是否能给结构化报价", "用于报价&利润表"],
  ["要资料", "我们需要用于阿里国际站和独立站上架的英文资料：高清图片、短视频、参数表、说明书、认证、FAQ、易损件清单和配件价格。", "We need English materials for Alibaba and our independent website: HD photos, short videos, spec sheet, user manual, certificates, FAQ, wearing parts list, and spare parts price list.", "图片、视频、参数、说明书、认证", "是否支持我们上架", "资料差会影响获客"],
  ["样品确认", "如果价格和资料合适，我们会先采购样品测试。请说明样品价格、发货方式、打样周期、售后处理和后续批量价是否可抵扣样品差价。", "If pricing and documents are suitable, we will purchase samples for testing. Please confirm sample price, shipping method, sample lead time, after-sales process, and whether bulk orders can offset sample price differences.", "样品政策、物流、售后", "能否支持低风险启动", "下一步进入样品池"],
];
styleTable(message, "A3:F7", "MessageTable");
message.freezePanes.freezeRows(3);
widths(message, [["A", 100], ["B", 360], ["C", 360], ["D", 180], ["E", 170], ["F", 150]]);

setTitle(next, "下一步动作", "这张表不是终点。下一步要把候选供应商变成可报价、可发样、可售后的核心供应链。", "G");
next.getRange("A3:G12").values = [
  ["序号", "动作", "对象", "产出物", "优先级", "负责人", "备注"],
  [1, "每个品类先联系优先级为“优先”的供应商", "S-001/S-002/S-005/S-006/S-010/S-011/S-014/S-016", "第一轮报价和资料包", "高", "父亲大人", "先用中文/英文话术"],
  [2, "要求工厂证明材料", "全部候选", "营业执照、工厂视频、生产/仓库照片", "高", "父亲大人", "拒绝提供则降级"],
  [3, "补齐产品资料", "配件/FPV/农用/工业", "参数表、高清图、视频、认证", "高", "幺儿", "后续用于阿里国际站"],
  [4, "拿阶梯报价", "每个品类至少3家", "EXW/FOB/样品价/批量价", "高", "父亲大人", "用于毛利测算"],
  [5, "建立样品测试清单", "每类2-3个SKU", "样品预算和测试计划", "中", "父亲大人/幺儿", "先小额验证"],
  [6, "筛出核心供应商", "每个品类", "2家核心+1家备选", "高", "父亲大人/幺儿", "达到A/B评级才进入报价体系"],
  [7, "整理可售产品目录", "核心供应商", "英文产品目录第一版", "高", "幺儿", "客户开发前必须完成"],
  [8, "同步到启动总表", "供应商&物流页", "正式供应链台账", "中", "幺儿", "后续可合并进总表"],
  [9, "风险复核", "价格异常低/资料不足供应商", "淘汰或观察名单", "中", "父亲大人", "宁缺毋滥"],
];
styleTable(next, "A3:G12", "NextStepsTable");
next.freezePanes.freezeRows(3);
widths(next, [["A", 70], ["B", 230], ["C", 240], ["D", 220], ["E", 80], ["F", 120], ["G", 220]]);

await fs.mkdir(outputDir, { recursive: true });
const overview = await wb.inspect({ kind: "table", range: "供应商池!A3:R19", include: "values", tableMaxRows: 20, tableMaxCols: 18 });
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula error scan" });
await fs.writeFile(`${outputDir}/verification.txt`, `${overview.ndjson}\n${errors.ndjson}`);

for (const sheetName of ["供应商池", "筛选标准", "询价话术", "下一步动作"]) {
  const png = await wb.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${sheetName}.png`, new Uint8Array(await png.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);
console.log(outputPath);
