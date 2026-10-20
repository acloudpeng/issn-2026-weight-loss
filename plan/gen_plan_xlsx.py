#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""根据 ISSN 2026《减重饮食立场声明》生成减肥计划表 + 记录表（xlsx）"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ---------- 样式 ----------
H_FILL = PatternFill("solid", fgColor="1F6F4A")
H_FONT = Font(color="FFFFFF", bold=True, size=11)
SUB_FILL = PatternFill("solid", fgColor="E8F3ED")
INPUT_FILL = PatternFill("solid", fgColor="FFF9E6")
CALC_FILL = PatternFill("solid", fgColor="EAF3FB")
THIN = Side(style="thin", color="C9D6CE")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
TITLE_FONT = Font(bold=True, size=14, color="1F6F4A")


def style_header(ws, row=1, ncol=None):
    ncol = ncol or ws.max_column
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = H_FILL
        cell.font = H_FONT
        cell.alignment = CENTER
        cell.border = BORDER


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def title(ws, text, sub=None, span=8):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    c = ws.cell(row=1, column=1, value=text)
    c.font = TITLE_FONT
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 26
    if sub:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
        c2 = ws.cell(row=2, column=1, value=sub)
        c2.font = Font(size=9.5, color="5B6B62")
        c2.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[2].height = 30


wb = Workbook()

# ============================================================
# Sheet 1 · 说明
# ============================================================
ws = wb.active
ws.title = "00-说明"
title(ws, "减重计划与记录表（依据 ISSN 2026 立场声明）",
      "数据来源：Aragon AA et al. International Society of Sports Nutrition position stand: diets for body weight reduction. "
      "J Int Soc Sports Nutr. 2026;23(1):2740266. 在线发表 2026-10-13. DOI 10.1080/15502783.2026.2740266 · PMID 42867412 · CC BY",
      span=6)
rows = [
    ("", ""),
    ("表格结构", "① 01-12周计划表：逐周目标与任务　② 02-体重围度记录：每周固定时间填一次　③ 03-饮食记录：每日一行，算总热量与蛋白　"
                 "④ 04-训练记录：主要动作的重量/组数/次数　⑤ 05-参考值：论文给出的蛋白区间、饮食法要点、平台期规则　⑥ 06-计算器：填 4 个空自动出结果"),
    ("核心结论（论文）", "决定减重质量的是三件事：总能量缺口、蛋白质摄入量、抗阻训练。"
                        "生酮/低脂/轻断食之间的差别，在热量与蛋白匹配后基本消失。"),
    ("关键蛋白区间", "一般人群 1.2–1.6 g/kg；减脂+抗阻训练 1.6–2.2 g/kg；精瘦且有训练基础者可 1.9–3.2 g/kg（按去脂体重 2.5–4.2 g/kg）。"),
    ("减重速度参考", "每周 0.5%–1% 体重为稳妥区间。VLED（<800 kcal）虽快（1.5–2.5 kg/周），但属高风险方案，需医学监督。"),
    ("平台期规则", "平台定义：12 周内体重下降不足 5%。应对顺序：先核对摄入记录（自报普遍低报 25%–50%）→ 提高活动量与训练 → "
                   "插入 7 天饮食休息（每减 3–5 kg）或 1–3 个月维持期（每减 5–10 kg）。"),
    ("怎么用这张表", "1) 先在「06-计算器」填入性别/体重/身高/年龄/活动量，得到目标热量与蛋白克数；"
                     "2) 把结果抄到「01-12周计划表」的对应周；3) 每天在「03-饮食记录」记一行；每周固定一天在「02-体重围度记录」记一行。"),
    ("黄色单元格", "需要你自己填写。蓝色单元格是公式自动计算，不要覆盖。"),
    ("重要提醒", "本表是依据论文结论整理的执行模板，不是医疗建议。极低能量饮食、有慢性病、进食障碍史、孕期哺乳期人群请先咨询医生或注册营养师。"),
    ("色标说明", "输入区=浅黄　计算区=浅蓝　表头=深绿"),
]
r = 4
for k, v in rows:
    if not k:
        r += 1
        continue
    a = ws.cell(row=r, column=1, value=k)
    a.font = Font(bold=True)
    a.alignment = WRAP
    a.fill = SUB_FILL
    a.border = BORDER
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    b = ws.cell(row=r, column=2, value=v)
    b.alignment = WRAP
    b.border = BORDER
    ws.row_dimensions[r].height = 44 if len(v) < 90 else 60
    r += 1
set_widths(ws, [16, 30, 24, 18, 18, 18])

# ============================================================
# Sheet 2 · 12 周计划表
# ============================================================
ws = wb.create_sheet("01-12周计划表")
title(ws, "12 周减重计划表", "以「A 稳妥型（缺口 20%–25%）」为模板；B 运动型把缺口调到 15%、蛋白提到 1.8–2.2 g/kg；C 激进型须医学监督，勿自行套用。", span=8)
hdr = ["周次", "阶段", "能量（相对维持热量）", "蛋白目标", "抗阻训练", "有氧 / 步数", "本周任务与监测", "备注"]
ws.append([])
ws.append(hdr)
style_header(ws, row=3, ncol=len(hdr))

plan = [
    ("第 0 周", "打底", "维持热量（不设缺口）", "1.2 g/kg", "2 次", "步数 6000–8000",
     "测基线：晨起空腹体重、腰围、臀围、正侧背照片；用 3 天饮食记录估算真实摄入", "论文提醒：自报摄入普遍低报 25%–50%"),
    ("第 1–2 周", "适应", "−15%", "1.4 g/kg，分 ≥3 餐", "3 次（全身复合动作）", "有氧 2×25 分钟",
     "只改两件事：蛋白达标 + 抗阻到位；体重取 7 日均值，不看单日波动", "速度 0.5%–1%/周"),
    ("第 3–4 周", "主缺口", "−20%", "1.6 g/kg", "3 次（渐进加重）", "有氧 2–3 次；步数 8000–10000",
     "查速度：过快→加 100–150 kcal 碳水；完全不动→先查摄入记录，不要先砍热量", ""),
    ("第 5–6 周", "加深", "−22%～−25%", "1.6–1.8 g/kg", "3–4 次", "有氧 3 次",
     "腰围比体重更重要；若力量明显下降或持续疲劳 → 缺口过大或蛋白不足", "注意睡眠与饥饿感"),
    ("第 7 周", "饮食休息（diet break）", "回到维持热量，约 7 天", "保持 1.6 g/kg", "照常", "有氧减半",
     "论文推荐：每减 3–5 kg 安排一次约 7 天的有意识自由进食，缓解心理疲劳", "不是放纵，专注饥饱信号"),
    ("第 8–10 周", "第二轮缺口", "−20%～−25%", "1.6–2.0（训练量大可 2.2）", "3–4 次", "有氧 2–3 次",
     "易进入平台期：用腰围/照片/衣物松紧/训练重量综合判断", "平台定义：12 周内<5%"),
    ("第 11–12 周", "收口 + 稳定", "缺口收窄至 −10%", "1.6 g/kg", "4 次，维持强度", "有氧 2 次",
     "进入维持期：论文建议每减 5–10 kg 安排 1–3 个月稳定期，把体重钉住再决定下一轮", ""),
]
for row in plan:
    ws.append(list(row))
for r_ in range(4, 4 + len(plan)):
    for c_ in range(1, len(hdr) + 1):
        cell = ws.cell(row=r_, column=c_)
        cell.alignment = WRAP
        cell.border = BORDER
        if c_ in (3, 4):
            cell.fill = INPUT_FILL
    ws.row_dimensions[r_].height = 52
set_widths(ws, [11, 18, 22, 20, 20, 20, 46, 24])
ws.freeze_panes = "A4"

# ============================================================
# Sheet 3 · 体重围度记录
# ============================================================
ws = wb.create_sheet("02-体重围度记录")
title(ws, "体重与围度记录表", "每周固定一天、同样条件（晨起空腹、排便后、同一体重秤）测一次；体重列建议填 7 日均值。", span=11)
hdr = ["周次", "日期", "体重 7 日均值(kg)", "较上周变化(kg)", "较起始变化(kg)", "周降幅(%体重)",
       "腰围(cm)", "臀围(cm)", "体脂率(%)", "日均步数", "主观饥饿感(1-5)", "备注（睡眠/训练/状态）"]
ws.append([])
ws.append(hdr)
style_header(ws, row=3, ncol=len(hdr))
start = 4
for i in range(13):  # 第0周 ~ 第12周
    r_ = start + i
    label = f"第{i}周" if i else "第0周(基线)"
    ws.cell(row=r_, column=1, value=label).alignment = CENTER
    for c_ in range(2, len(hdr) + 1):
        ws.cell(row=r_, column=c_).fill = INPUT_FILL
    if i >= 1:
        ws.cell(row=r_, column=4, value=f'=IF(OR(C{r_}="",C{r_-1}=""),"",ROUND(C{r_}-C{r_-1},2))')
        ws.cell(row=r_, column=5, value=f'=IF(OR(C{r_}="",$C${start}=""),"",ROUND(C{r_}-$C${start},2))')
        ws.cell(row=r_, column=6, value=f'=IF(OR(C{r_}="",C{r_-1}=""),"",ROUND((C{r_-1}-C{r_})/C{r_-1}*100,2))')
    for c_ in range(1, len(hdr) + 1):
        cell = ws.cell(row=r_, column=c_)
        cell.border = BORDER
        cell.alignment = CENTER if c_ != len(hdr) else WRAP
        if c_ in (4, 5, 6):
            cell.fill = CALC_FILL
r_ = start + 13
ws.cell(row=r_, column=1, value="合计").font = Font(bold=True)
ws.cell(row=r_, column=5, value=f'=IF(OR($C${start}="",C{start+12}=""),"",ROUND(C{start+12}-$C${start},2))').font = Font(bold=True)
ws.cell(row=r_, column=6, value=f'=IF(OR($C${start}="",C{start+12}=""),"",ROUND(($C${start}-C{start+12})/$C${start}*100,2))').font = Font(bold=True)
ws.cell(row=r_, column=7, value="↑ 总减重(kg)")
ws.cell(row=r_, column=8, value="↑ 总降幅(%)")
set_widths(ws, [13, 12, 18, 15, 15, 14, 10, 10, 11, 11, 14, 30])
ws.freeze_panes = "C4"

# ============================================================
# Sheet 4 · 每日饮食记录
# ============================================================
ws = wb.create_sheet("03-饮食记录")
title(ws, "每日饮食记录（31 行，可复制扩展）",
      "蛋白目标见「06-计算器」结果；K 列自动按你在 J 列填的蛋白克数汇总，H 列填热量。蛋白优先：每餐 30–40 g。", span=11)
hdr = ["日期", "体重(kg)", "早餐", "加餐", "午餐", "加餐2", "晚餐", "总热量(kcal)", "蛋白(g)",
       "蛋白是否达标", "训练", "饥饿感(1-5)", "备注"]
ws.append([])
ws.append(hdr)
style_header(ws, row=3, ncol=len(hdr))
for i in range(31):
    r_ = 4 + i
    for c_ in range(1, len(hdr) + 1):
        cell = ws.cell(row=r_, column=c_)
        cell.border = BORDER
        cell.alignment = WRAP
        if c_ not in (10,):
            cell.fill = INPUT_FILL
    ws.cell(row=r_, column=10, value=f'=IF(I{r_}="","",IFERROR(IF(I{r_}>=\'06-计算器\'!$B$14,"✅ 达标","⚠ 不足"),""))').fill = CALC_FILL
    ws.row_dimensions[r_].height = 30
dv = DataValidation(type="list", formula1='"✅,❌"', allow_blank=True)
ws.add_data_validation(dv)
dv.add(f"K4:K34")
set_widths(ws, [12, 10, 22, 18, 22, 18, 22, 12, 9, 12, 16, 11, 22])
ws.freeze_panes = "B4"

# ============================================================
# Sheet 5 · 训练记录
# ============================================================
ws = wb.create_sheet("04-训练记录")
title(ws, "训练记录表（抗阻优先）",
      "论文立场：抗阻训练 + 足量蛋白是减脂期保住瘦体重的必需项。记录主要动作的重量×次数，保证渐进超负荷。", span=9)
ws.append([])
ws.append(["日期", "训练类型", "动作", "重量(kg)", "组数", "次数/组", "主观强度(RPE 1-10)", "有氧时长(分钟)", "备注"])
style_header(ws, row=3, ncol=9)
for i in range(40):
    r_ = 4 + i
    for c_ in range(1, 10):
        cell = ws.cell(row=r_, column=c_)
        cell.border = BORDER
        cell.alignment = CENTER
        cell.fill = INPUT_FILL
dv2 = DataValidation(type="list", formula1='"下肢主导,上肢推,上肢拉,全身复合,有氧,休息,其他"', allow_blank=True)
ws.add_data_validation(dv2)
dv2.add("B4:B43")
set_widths(ws, [13, 14, 24, 11, 8, 11, 16, 15, 26])
ws.freeze_panes = "A4"

# ============================================================
# Sheet 6 · 参考值
# ============================================================
ws = wb.create_sheet("05-参考值")
title(ws, "论文关键参考值（可直接对照执行）", span=5)
blocks = [
    ("蛋白摄入阶梯（论文立场 6/7）", [
        ("人群 / 目标", "每日蛋白 (g/kg 体重)", "每日蛋白 (g/kg 去脂体重)", "说明"),
        ("长期沿用的 RDA", "0.8", "—", "满足久坐成人基本需求，但对保瘦体重很可能不足"),
        ("一般人群", "1.2–1.6", "—", "支持运动、肌肉骨骼、心血管代谢健康；2025–2030 美国膳食指南首次写入"),
        ("减脂 + 抗阻训练", "1.6–2.2", "—", "可优化肌肉蛋白合成、减少瘦体重流失"),
        ("精瘦 + 有训练基础", "1.9–3.2", "2.5–4.2", "热量缺口下最大化瘦体重保留"),
        ("追求同时增肌减脂（重组）", "—", "2.6–3.5", "提高重组概率与幅度，但重组是有限过程，越精瘦越难"),
    ]),
    ("减重幅度与健康收益", [
        ("减重幅度", "可获得的收益", "", ""),
        ("2%–5%", "改善血糖控制、降低甘油三酯", "", ""),
        ("5%–10%", "改善血压、减轻关节疼痛、提升活动能力、增加肠道菌群多样性、减轻抑郁症状、降低心血管死亡风险", "", ""),
        ("约 15%", "已有研究报告 2 型糖尿病缓解", "", ""),
    ]),
    ("平台期与恢复期规则（论文立场 13）", [
        ("规则", "数值", "", ""),
        ("平台期定义", "12 周内体重下降不足 5%", "", ""),
        ("典型出现时间", "饮食法开始后 8–12 周", "", ""),
        ("饮食休息 diet break", "每减 3–5 kg，安排约 7 天的有意识自由进食", "", ""),
        ("维持期 maintenance", "每减 5–10 kg，安排约 1–3 个月或更长的体重稳定期", "", ""),
    ]),
    ("各饮食法一句话结论", [
        ("方案", "结论要点", "", ""),
        ("VLED (<800 kcal)", "短期最快（10–15% 体重），长期无优势，瘦体重与骨量流失风险高，需医疗监督", "", ""),
        ("LED (800–1200 kcal)", "减重幅度大但烈度低，长期结局与温和方案类似", "", ""),
        ("间歇禁食 (ADF/5:2/TRE)", "能量匹配时与每日限能量无临床意义差别（6 个月仅多减 0.94 kg）", "", ""),
        ("生酮", "短期有效，但瘦体重掉得更多、LDL 常升高；蛋白与热量匹配后无内在优势", "", ""),
        ("低碳 / 低脂 / 极低脂", "长期效果相近，选能坚持的；低碳 12–24 个月后优势消失", "", ""),
        ("高蛋白", "提高饱腹感与食物热效应，减脂同时保肌，健康人群在宽范围内安全", "", ""),
    ]),
]
r_ = 3
for name, table in blocks:
    ws.merge_cells(start_row=r_, start_column=1, end_row=r_, end_column=5)
    c = ws.cell(row=r_, column=1, value=name)
    c.font = Font(bold=True, color="1F6F4A", size=11.5)
    c.fill = SUB_FILL
    c.border = BORDER
    r_ += 1
    for j, row in enumerate(table):
        for k, val in enumerate(row):
            cell = ws.cell(row=r_, column=k + 1, value=val)
            cell.border = BORDER
            cell.alignment = WRAP
            if j == 0:
                cell.fill = H_FILL
                cell.font = H_FONT
                cell.alignment = CENTER
        r_ += 1
    r_ += 1
set_widths(ws, [26, 34, 20, 40, 10])

# ============================================================
# Sheet 7 · 计算器
# ============================================================
ws = wb.create_sheet("06-计算器")
title(ws, "热量与蛋白计算器", "黄色格填自己数据；Mifflin-St Jeor 公式估算 BMR，蛋白区间取自论文立场 6/7。结果为起点估算，需按 2–4 周实际变化校准。", span=3)
ws["A4"] = "性别（男/女）"; ws["B4"] = "男"
ws["A5"] = "体重 (kg)"; ws["B5"] = 75
ws["A6"] = "身高 (cm)"; ws["B6"] = 175
ws["A7"] = "年龄"; ws["B7"] = 30
ws["A8"] = "活动水平"; ws["B8"] = "轻度(1.375)"
ws["A9"] = "路径"; ws["B9"] = "A 稳妥型(−20%)"

ws["A11"] = "—— 计算结果 ——"; ws["A11"].font = Font(bold=True, color="1F6F4A")
ws["A12"] = "基础代谢 BMR (kcal)"; ws["B12"] = '=IF(B4="男",10*B5+6.25*B6-5*B7+5,10*B5+6.25*B6-5*B7-161)'
ws["A13"] = "维持热量 TDEE (kcal)"
act = 'IF(LEFT(B8,2)="久坐",1.2,IF(LEFT(B8,2)="轻度",1.375,IF(LEFT(B8,2)="中度",1.55,1.725)))'
ws["B13"] = f'=ROUND(B12*{act},0)'
ws["A14"] = "每日蛋白目标 (g)"
ws["B14"] = '=IF(LEFT(B9,1)="A",ROUND(B5*1.6,0),ROUND(B5*2.2,0))'
ws["A15"] = "每日目标热量 (kcal)"
ws["B15"] = '=ROUND(MAX(B13*IF(LEFT(B9,1)="A",0.8,0.85),B12*1.1),0)'
ws["A16"] = "预计减重速度 (kg/周)"
ws["B16"] = '=ROUND((B13-B15)*7/7700,2)'
ws["A17"] = "蛋白下限参考 (g)"; ws["B17"] = '=ROUND(B5*1.2,0)'
ws["A18"] = "蛋白上限参考 (g)"; ws["B18"] = '=ROUND(B5*3.2,0)'

for r_ in range(4, 10):
    ws.cell(row=r_, column=1).font = Font(bold=True)
    ws.cell(row=r_, column=2).fill = INPUT_FILL
    ws.cell(row=r_, column=2).border = BORDER
for r_ in range(11, 19):
    ws.cell(row=r_, column=1).font = Font(bold=True)
    ws.cell(row=r_, column=2).fill = CALC_FILL
    ws.cell(row=r_, column=2).border = BORDER
dv3 = DataValidation(type="list", formula1='"男,女"', allow_blank=False)
ws.add_data_validation(dv3); dv3.add("B4")
dv4 = DataValidation(type="list", formula1='"久坐(1.2),轻度(1.375),中度(1.55),高度(1.725)"', allow_blank=False)
ws.add_data_validation(dv4); dv4.add("B8")
dv5 = DataValidation(type="list", formula1='"A 稳妥型(−20%),B 运动型(−15%)"', allow_blank=False)
ws.add_data_validation(dv5); dv5.add("B9")
ws["A20"] = "提示：若把路径改成 B 运动型，蛋白目标自动提到 2.2 g/kg、缺口收窄到 15%。"
ws["A20"].font = Font(size=9.5, color="5B6B62")
ws.merge_cells("A20:C20")
set_widths(ws, [26, 22, 40])

wb.save("/root/repos/issn-2026-weight-loss/plan/减重计划与记录表.xlsx")
print("已生成: /root/repos/issn-2026-weight-loss/plan/减重计划与记录表.xlsx")
print("工作表:", wb.sheetnames)
