#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据管理员培训课件 · 28 页 PPTX 构建脚本（专业培训风）
=========================================================
用法：python build_pptx.py
输入：lesson_plan.md（同目录，逐字稿级教案）
输出：data_steward_course.pptx / notes_per_page.json

纪律：
- 所有页面正文要点取自 lesson_plan.md 对应 page_id；讲解词原样进 notes，不进正文。
- 示例数据全部虚构，不含客户名/真实企业名/信用代码/手机号。
- 中文避头尾：每个 a:pPr 设 eaLnBrk="1" hangingPunct="1"。
- 不使用标题左侧竖线 / 块状衬底 / emoji / 正中文字底纹。
"""
import json
import os
import re

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

BASE = os.path.dirname(os.path.abspath(__file__))
LESSON = os.path.join(BASE, "lesson_plan.md")
OUT_PPTX = os.path.join(BASE, "data_steward_course.pptx")
OUT_JSON = os.path.join(BASE, "notes_per_page.json")

# ---- 色板（专业培训风）----------------------------------------------------
NAVY = (0x1F, 0x3A, 0x5F)    # 主色 深蓝
SLATE = (0x4A, 0x6F, 0xA5)   # 辅色 青灰
ORANGE = (0xC8, 0x79, 0x1A)  # 强调 暖橙，仅作关键标记（不承载文本）
BG = (0xF7, 0xF9, 0xFC)      # 背景 近白
TEXT = (0x26, 0x34, 0x45)    # 正文 深灰
WHITE = (0xFF, 0xFF, 0xFF)
BORDER = (0xC4, 0xD1, 0xE0)
LIGHT = (0xEC, 0xF1, 0xF7)
FONT = "Microsoft YaHei"

W, H = 13.333, 7.5
M = 0.62
CW = W - 2 * M
BODY_TOP = 1.78

PAGES = []


def page(pid, title, chapter, subtitle=""):
    def deco(fn):
        PAGES.append((pid, title, chapter, subtitle, fn))
        return fn
    return deco


def rgb(c):
    return RGBColor(*c)


# ---- 中文避头尾 + 字体 ---------------------------------------------------
def _ppr(p, align=PP_ALIGN.LEFT, line=1.18, before=0, after=0):
    pPr = p._p.get_or_add_pPr()
    pPr.set("eaLnBrk", "1")
    pPr.set("hangingPunct", "1")
    p.alignment = align
    p.line_spacing = line
    p.space_before = Pt(before)
    p.space_after = Pt(after)
    return pPr


def _font(run, size, color, bold=False):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = rgb(color)
    f.name = FONT
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", FONT)


def text(slide, x, y, w, h, content, size=16, color=TEXT, bold=False,
         align=PP_ALIGN.LEFT, line=1.18, anchor=MSO_ANCHOR.TOP,
         ml=0, mt=0, mr=0, mb=0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left, tf.margin_right = Inches(ml), Inches(mr)
    tf.margin_top, tf.margin_bottom = Inches(mt), Inches(mb)
    tf.vertical_anchor = anchor
    for i, ln in enumerate(str(content).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = ln
        _ppr(p, align, line)
        for r in p.runs:
            _font(r, size, color, bold)
    return box


def rect(slide, x, y, w, h, fill=None, line=None, lw=0.75,
         shape=MSO_SHAPE.RECTANGLE, radius=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(lw)
    if radius is not None:
        try:
            s.adjustments[0] = radius
        except Exception:
            pass
    return s


def hline(slide, x, y, w, color=SLATE, t=0.018):
    return rect(slide, x, y, w, t, fill=color)


def arrow(slide, x, y, w, h, direction="right", color=SLATE):
    shp = MSO_SHAPE.RIGHT_ARROW if direction == "right" else MSO_SHAPE.DOWN_ARROW
    return rect(slide, x, y, w, h, fill=color, shape=shp)


def card(slide, x, y, w, h, title, body, tsize=15.5, bsize=13.5,
         accent=False, fill=WHITE, line=BORDER, tcolor=NAVY, bcolor=TEXT):
    rect(slide, x, y, w, h, fill=fill, line=line, lw=0.75,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    if accent:
        rect(slide, x + w - 0.22, y + 0.13, 0.10, 0.10, fill=ORANGE)
    text(slide, x + 0.15, y + 0.10, w - 0.30, 0.34, title, tsize, tcolor, True, line=1.05)
    if body:
        text(slide, x + 0.15, y + 0.47, w - 0.30, max(0.32, h - 0.56), body, bsize, bcolor, line=1.20)


def band(slide, x, y, w, h, content, size=15, color=NAVY, fill=LIGHT,
         align=PP_ALIGN.LEFT, bold=False):
    rect(slide, x, y, w, h, fill=fill, line=None,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    text(slide, x + 0.18, y + 0.06, w - 0.36, h - 0.12, content, size, color, bold,
         align=align, line=1.18, anchor=MSO_ANCHOR.MIDDLE)


def _cell_border(cell, color=BORDER, width_pt=0.6):
    tcPr = cell._tc.get_or_add_tcPr()
    edges = ["a:lnL", "a:lnR", "a:lnT", "a:lnB"]
    for edge in edges:
        old = tcPr.find(qn(edge))
        if old is not None:
            tcPr.remove(old)
    for edge in reversed(edges):
        ln = tcPr.makeelement(qn(edge), {"w": str(int(Pt(width_pt)))})
        sf = ln.makeelement(qn("a:solidFill"), {})
        clr = sf.makeelement(qn("a:srgbClr"), {"val": "%02X%02X%02X" % color})
        sf.append(clr)
        ln.append(sf)
        tcPr.insert(0, ln)


def table(slide, x, y, w, h, data, widths, fsize=12.5, hsize=None, first_col_bold=False):
    rows, cols = len(data), len(data[0])
    shp = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(h))
    tbl = shp.table
    for i, cw in enumerate(widths):
        tbl.columns[i].width = Inches(cw)
    rh = h / rows
    for i in range(rows):
        tbl.rows[i].height = Inches(rh)
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(NAVY if r == 0 else (WHITE if r % 2 else LIGHT))
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.025)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            _cell_border(cell)
            for p in cell.text_frame.paragraphs:
                _ppr(p, PP_ALIGN.LEFT, 1.05)
                for run in p.runs:
                    _font(run, (hsize or fsize) if r == 0 else fsize,
                          WHITE if r == 0 else TEXT,
                          bold=(r == 0 or (first_col_bold and c == 0)))
    return shp


def frame(slide, pid, title, chapter, subtitle=""):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(BG)
    text(slide, M, 0.26, 7.6, 0.32, "数据管理员培训 · " + chapter, 10, SLATE)
    text(slide, W - M - 4.6, 0.26, 4.6, 0.32, "主数据治理 · 培训课件",
         10, SLATE, align=PP_ALIGN.RIGHT)
    rect(slide, M, 0.64, 0.22, 0.06, fill=ORANGE)
    text(slide, M, 0.78, CW, 0.52, title, 29, NAVY, True, line=1.02)
    if subtitle:
        text(slide, M, 1.32, CW, 0.32, subtitle, 12.5, SLATE, line=1.05)
    text(slide, M, 6.78, 9.0, 0.32, "数据管理员培训 · 内部培训材料 · 示例数据均为虚构",
         10, SLATE)
    text(slide, W - M - 2.2, 6.78, 2.2, 0.32, pid + " / 28",
         10, NAVY, True, align=PP_ALIGN.RIGHT)


# ============================ 28 页正文 ==================================
@page("P01", "这门课解决什么问题", "开场", "90 分钟 · 28 页 · 8 章节 · 面向会员企业数据管理员")
def p01(s):
    text(s, M, 1.86, CW, 0.62, "数据管理员培训", 34, NAVY, True, line=1.02)
    text(s, M, 2.55, CW, 0.36, "不讲抽象概念，只讲你坐在数据菜单前每天真正要做的判断和操作。",
         17, SLATE, line=1.1)
    rect(s, M, 3.04, 0.16, 0.16, fill=ORANGE)
    band(s, M, 3.28, CW, 0.78,
         "六入口：导入中心 · 查重合并 · 字段字典 · 数据质量报告 · 导出审批 · 操作审计",
         16.5, NAVY, LIGHT, bold=True)
    chips = ["导入中心", "查重合并", "字段字典", "数据质量报告", "导出审批", "操作审计"]
    for i, name in enumerate(chips):
        cx = M + (i % 3) * 4.12
        cy = 4.28 + (i // 3) * 0.68
        rect(s, cx, cy, 3.85, 0.56, fill=WHITE, line=BORDER,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
        text(s, cx, cy + 0.04, 3.85, 0.46, name, 16, NAVY, True,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, M, 5.78, CW, 0.32,
         "学完你能独立处理一批主数据的导入、查重、合并与质量核查。", 14, NAVY)
    text(s, M, 6.16, CW, 0.32,
         "所有示例数据均为虚构：企业名、信用代码、手机号都是编出来的，仅供练手，不涉及任何真实信息。",
         12.5, SLATE)


@page("P02", "数据菜单六入口与课程地图", "开场", "先把主线记住，每章都会提醒你站在哪一段")
def p02(s):
    row1 = [("字段字典", "知道什么能填、怎么填"), ("导入中心", "把数据搬进来"),
            ("查重合并", "处理重名与同码记录"), ("数据质量报告", "完整率 / 重复率 / 过期率")]
    for i, (t, d) in enumerate(row1):
        cx = 0.72 + i * 3.05
        card(s, cx, 1.95, 2.65, 0.86, t, d, tsize=15, bsize=12)
        if i < 3:
            arrow(s, cx + 2.70, 2.25, 0.26, 0.24)
    card(s, 0.72, 3.00, 4.30, 0.86, "第八章 · AI 能力边界", "独立专题：建议进复核队列，不直写、不外发、不算金额",
         tsize=15, bsize=12, accent=True)
    card(s, 6.82, 3.00, 2.65, 0.86, "导出审批", "涉及对外给数据时走这里", tsize=15, bsize=12)
    arrow(s, 9.53, 3.30, 0.26, 0.24)
    card(s, 9.87, 3.00, 2.65, 0.86, "操作审计", "所有关键动作都留痕", tsize=15, bsize=12)
    text(s, M, 4.12, CW, 0.32, "每一步的输出交给谁", 14, NAVY, True)
    hand = ["字段字典 → 模板口径", "导入中心 → 校验报告", "查重合并 → 合并结果",
            "质量报告 → 治理动作", "导出审批 → 受控文件", "操作审计 → 可追溯链路"]
    for i, h in enumerate(hand):
        cx = M + (i % 3) * 4.12
        cy = 4.52 + (i // 3) * 0.46
        text(s, cx, cy, 3.85, 0.38, "· " + h, 14, TEXT)
    band(s, M, 5.62, CW, 0.78, "每进入一章，先确认我们站在主线的哪一段、这一步的输出会交给谁。",
         16, NAVY, LIGHT)


@page("P03", "六类主数据对象全景", "第一章", "企业是根，联系人与供需长在根上，商机从对接中产生")
def p03(s):
    items = [
        ("会员企业", "平台里的组织主体，是根。\n法人 / 组织主体，属 Party 主数据。"),
        ("联系人", "挂在企业下面，承担业务联络的自然人。"),
        ("能力卡", "描述一家企业能供给什么：产品 / 服务 / 资源能力。"),
        ("需求卡", "描述它想要什么：采购 / 合作 / 资源需求。"),
        ("商机", "能力卡与需求卡对接后产生的机会。"),
        ("活动参与记录", "记下谁参加了哪场活动，补充互动轨迹。"),
    ]
    for i, (t, d) in enumerate(items):
        cx = M + (i % 3) * 4.12
        cy = 1.92 + (i // 3) * 2.10
        card(s, cx, cy, 3.85, 1.92, t, d, tsize=16, bsize=14)
    band(s, M, 6.14, CW, 0.44,
         "后面所有导入、查重、质量核查，本质上都是在这六个对象上做文章。",
         14.5, NAVY, LIGHT)


@page("P04", "对象归属与唯一可信源", "第一章", "每个对象说清四件事：谁维护 / 必填什么 / 从哪来 / 被谁用")
def p04(s):
    band(s, M, 1.86, CW, 0.62,
         "两处来源维护同一对象：指定唯一可信源（SoR），另一处只读引用；冲突以 SoR 为准，并在对象卡上写明。",
         14, NAVY, LIGHT)
    data = [
        ["对象", "归属角色 owner", "必填字段（要点）", "上游来源", "下游用途"],
        ["会员企业", "会员运营", "ent_id / ent_name_c / uscc_c（可空但唯一）/ ent_status_c（枚举）",
         "入驻申请 / CRM / 批量导入", "CRM客户 / 商机归属 / 资源匹配 / 驾驶舱"],
        ["联系人", "会员运营 / CRM管理员", "contact_id / ent_id / contact_name_c / mobile_c（敏感）",
         "企业提交 / 名片录入 / 导入", "通知 / 商机跟进 / 账号关联"],
        ["能力卡", "内容运营", "cap_id / ent_id / cap_name_c / cap_status_c",
         "企业自报 / 运营审核", "资源匹配 / 商机推荐 / 官网展示"],
        ["需求卡", "资源对接", "demand_id / ent_id / demand_name_c / demand_status_c",
         "企业发布 / 运营录入", "商机生成 / 资源匹配 / 看板"],
        ["商机", "招商业务员", "opp_id / ent_id / opp_stage_c / owner_user_id（金额以分存整数）",
         "需求/能力匹配 / 业务员录入", "销售工作台 / 驾驶舱 / 漏斗统计"],
        ["活动参与记录", "活动运营", "part_id / ent_id、contact_id / event_id / join_status_c",
         "小程序报名签到 / 官网", "活跃度分析 / 精准运营 / 驾驶舱"],
    ]
    table(s, M, 2.58, CW, 3.55, data, [1.35, 1.55, 3.60, 2.05, 3.54], fsize=11.5, hsize=12,
          first_col_bold=True)
    text(s, M, 6.20, CW, 0.32,
         "软删除行不计入质量指标，但保留在操作审计中可查。", 12.5, SLATE)


@page("P05", "六类对象卡逐项解读（演练）", "第一章", "演练：把虚构字段卡贴到正确的对象下面")
def p05(s):
    band(s, M, 1.86, CW, 0.50,
         "每张卡写清 owner / 必填字段 / 上游来源 / 下游用途，教师对照知识卡核对后点评。",
         14, NAVY, LIGHT)
    cards = [
        ("会员企业", "owner：会员运营\n必填：ent_id / ent_name_c / uscc_c / ent_status_c\n上游：入驻申请 / CRM / 导入\n下游：CRM客户 / 商机归属 / 驾驶舱"),
        ("联系人", "owner：会员运营 / CRM管理员\n必填：contact_id / ent_id / contact_name_c / mobile_c（敏感）\n上游：企业提交 / 名片录入 / 导入\n下游：通知 / 商机跟进 / 账号关联"),
        ("能力卡", "owner：内容运营\n必填：cap_id / ent_id / cap_name_c / cap_status_c\n上游：企业自报 / 运营审核\n下游：资源匹配 / 商机推荐 / 官网展示"),
        ("需求卡", "owner：资源对接\n必填：demand_id / ent_id / demand_name_c / demand_status_c\n上游：企业发布 / 运营录入\n下游：商机生成 / 资源匹配 / 看板"),
        ("商机", "owner：招商业务员\n必填：opp_id / ent_id / opp_stage_c / owner_user_id\n上游：需求/能力匹配 / 业务员录入\n下游：销售工作台 / 驾驶舱 / 漏斗统计"),
        ("活动参与记录", "owner：活动运营\n必填：part_id / ent_id、contact_id / event_id / join_status_c\n上游：小程序报名签到 / 官网\n下游：活跃度分析 / 精准运营 / 驾驶舱"),
    ]
    for i, (t, d) in enumerate(cards):
        cx = M + (i % 3) * 4.12
        cy = 2.48 + (i // 3) * 1.92
        card(s, cx, cy, 3.85, 1.80, t, d, tsize=15, bsize=12)
    band(s, M, 6.20, CW, 0.44,
         "边界：被软删除的对象行不计入质量指标，但变更痕迹仍留在审计里，查得到、说得清。",
         12.5, NAVY, LIGHT)


@page("P06", "字段字典的结构与读法", "第二章", "字典是口径源，模板按字典生成")
def p06(s):
    card(s, M, 1.92, 6.35, 4.45, "字段字典应逐项登记的元数据（13 项）",
         "物理字段名 / 中文显示名 · 业务定义 · 数据类型与长度精度 · 是否必填（可空性）\n"
         "默认值 · 格式校验规则 · 枚举值域 · 唯一性约束\n"
         "敏感级别与脱敏规则 · 来源系统 / SoR · 数据责任人\n"
         "版本与生效日 · 废弃标记",
         tsize=16, bsize=14.5)
    card(s, 7.20, 1.92, 5.51, 2.15, "读字典三步",
         "1. 翻到对应对象那一页\n2. 确认哪些必填、哪些选填、类型对不对\n3. 按字典口径录入——口头“大概填一下”不算数",
         tsize=16, bsize=14.5, accent=True)
    card(s, 7.20, 4.22, 5.51, 1.35, "字典与模板的关系",
         "字典是口径源；模板按字典生成；字典变更须同步模板。",
         tsize=15, bsize=14)
    band(s, 7.20, 5.72, 5.51, 0.65, "先查字典，再动手录。", 17, NAVY, LIGHT, bold=True)


@page("P07", "必填规则与校验（演练）", "第二章", "必填缺失即拦截，与其余字段是否齐全无关")
def p07(s):
    data = [
        ["约束", "示例（虚构字段）", "含义"],
        ["必填", "ent_name_c NOT NULL 且 TRIM(ent_name_c) <> ''", "既非空也非空串；空串与 NULL 要分别校验"],
        ["格式", "mobile_c 须匹配 11 位数字格式正则", "只验形态不验真实性（不写真实号码）"],
        ["枚举", "ent_status_c ∈ {ACTIVE, INACTIVE, PENDING}", "值域受控，变更须版本化"],
        ["唯一", "UNIQUE(uscc_c)，但允许 NULL", "同一信用代码不重复入库；空码不触发唯一"],
    ]
    table(s, M, 1.92, CW, 2.55, data, [1.30, 5.00, 5.79], fsize=13, hsize=13,
          first_col_bold=True)
    card(s, M, 4.66, CW, 1.45, "演练：5 条虚构记录，判断放行还是拦截",
         "判断标准只有一条：必填缺了就拦，选填空着不拦；字段类型也要对得上（金额以「分」为单位存整数）。\n"
         "拦截不是刁难，而是保护中台——今天放进去一条缺字段的脏数据，明天查重和质量报告都会跟着出错。",
         tsize=15, bsize=13.5)
    band(s, M, 6.16, CW, 0.44, "字典改了，模板要跟着改；校验口径以字典为准。", 12.5, NAVY, LIGHT)


@page("P08", "批量导入四步流程与模板", "第三章", "导入不是一次性赌博，而是一来一回的对话")
def p08(s):
    steps = [
        ("1 下载最新模板", "列结构由字段字典派生\n不留合并单元格 / 公式\nUTF-8 编码"),
        ("2 按模板填数据", "金额以「分」为单位存整数\n日期格式统一"),
        ("3 上传看校验报告", "逐行成功 / 失败\n失败原因以错误码呈现"),
        ("4 修正失败行再传", "部分成功还是全量回滚须事先约定\n建议行级事务 + 失败行隔离"),
    ]
    for i, (t, d) in enumerate(steps):
        cx = 0.62 + i * 3.11
        card(s, cx, 2.00, 2.75, 1.90, t, d, tsize=14.5, bsize=12.5, accent=(i == 3))
        if i < 3:
            arrow(s, cx + 2.80, 2.78, 0.24, 0.24)
    band(s, M, 4.22, CW, 0.92,
         "模板定义列结构 · 校验报告逐行反馈通过或失败及原因 · 失败行须修正后重新上传。",
         16, NAVY, LIGHT)
    band(s, M, 5.38, CW, 0.92,
         "跳过校验报告强制导入，会把脏数据直接写进主数据；导入是写操作，必须留审计。",
         15, TEXT, LIGHT)


@page("P09", "校验报告与失败行识别", "第三章", "别只看失败总数，要定位到具体行与具体原因")
def p09(s):
    data = [
        ["行号", "状态", "失败原因（错误码）"],
        ["0007", "失败", "必填字段为空（ent_name_c）"],
        ["0012", "失败", "类型错误（金额非「分」整数）"],
        ["0018", "失败", "格式不合法（mobile_c 非 11 位数字形态）"],
        ["0023", "成功", "—"],
        ["0031", "成功", "—"],
        ["0044", "失败", "枚举值不在值域（ent_status_c）"],
    ]
    table(s, M, 1.95, 7.00, 3.45, data, [1.20, 1.20, 4.60], fsize=12.5, hsize=13)
    card(s, 7.85, 1.95, 4.86, 3.45, "读报告三步",
         "1. 别只看“有 4 行失败”这种总数\n"
         "2. 定位到具体是哪几行、为什么失败\n"
         "3. 是必填空了、类型不对，还是编码不合法\n"
         "4. 把失败行号圈出来，这就是下一步要动刀的地方",
         tsize=15.5, bsize=13.5, accent=True)
    band(s, M, 5.62, CW, 0.78,
         "部分成功还是全量回滚须事先约定，建议行级事务 + 失败行隔离；失败是反馈，不是判决。",
         15, NAVY, LIGHT)


@page("P10", "失败行修正与重传（演练）", "第三章", "只修正并重传失败行，禁止重传全量")
def p10(s):
    data = [
        ["行号", "失败原因", "诊断", "修正动作"],
        ["0007", "必填为空", "补 ent_name_c 值", "填入去空格后非空串的中文名"],
        ["0012", "类型错误", "金额改为「分」整数", "按字典把金额存为以分为单位的整数"],
        ["0018", "格式不合法", "按字典纠正形态", "改为 11 位数字形态（只验形态不验真伪）"],
    ]
    table(s, M, 1.95, 7.00, 2.70, data, [0.95, 1.55, 1.85, 2.65], fsize=12, hsize=12.5)
    card(s, 7.85, 1.95, 4.86, 2.70, "重传红线",
         "只重传失败的 3 行，不把整批一百行重新传一遍。\n"
         "按业务键 / 批次号幂等重传；禁止重传全量，否则已成功行二次入库变重复。\n"
         "导入是写操作，必须留审计。",
         tsize=15, bsize=12.5, accent=True)
    band(s, M, 4.90, CW, 0.80,
         "跳过校验报告强制导入，会把脏数据直接写进主数据。", 15, TEXT, LIGHT)
    band(s, M, 5.88, CW, 0.56,
         "改完再传，再看报告，直到全部通过。", 16, NAVY, LIGHT, bold=True)


@page("P11", "查重两类判定总览", "第四章", "同码即确定，同名规范化即疑似")
def p11(s):
    card(s, M, 2.00, 5.85, 3.25, "确定重复",
         "判定依据：统一社会信用代码相同\n"
         "处理：系统可直接列为合并候选\n"
         "关键：身份以代码为准\n"
         "边界：代码相同但名称冲突，仍属确定重复，加备注即可（见 P12）",
         tsize=17, bsize=14.5)
    card(s, 6.86, 2.00, 5.85, 3.25, "疑似重复",
         "判定依据：企业名规范化后相同\n"
         "处理：必须人工复核，绝不自动合并\n"
         "关键：“像”不等于“是”\n"
         "边界：存在未决字段则挂起，不合并",
         tsize=17, bsize=14.5, accent=True)
    band(s, M, 5.52, CW, 0.88,
         "确定可自动候选，疑似必须人工过目；后面逐类拆开练，最后用一张决策图固定判断顺序。",
         16, NAVY, LIGHT, bold=True)


@page("P12", "确定重复：统一社会信用代码相同", "第四章", "身份以代码为准，名字差异是备注事项")
def p12(s):
    card(s, M, 1.95, 5.85, 1.85, "记录 A（虚构示例）",
         "uscc_c：91FICT00000000001X（虚构18位）\n企业名：星海科技（简称）", tsize=15, bsize=14)
    card(s, 6.86, 1.95, 5.85, 1.85, "记录 B（虚构示例）",
         "uscc_c：91FICT00000000001X（虚构18位）\n企业名：星海科技有限公司（全称）", tsize=15, bsize=14)
    text(s, M, 3.92, CW, 0.34, "两条记录代码完全相同", 14, SLATE,
         align=PP_ALIGN.CENTER)
    band(s, M, 4.34, CW, 0.88,
         "代码完全相同 → 判定为确定重复 → 系统直接生成合并候选。", 17, NAVY, LIGHT, bold=True)
    card(s, M, 5.40, CW, 1.05, "边界情形",
         "同码不同名（简称 vs 全称）不改变类别：名字冲突只需加备注说明，不能降级成疑似；空 / 无效信用代码转名称分支。",
         tsize=14, bsize=13)


@page("P13", "疑似重复：企业名规范化后相同", "第四章", "规范化同名只是“像”，不能证明是同一主体")
def p13(s):
    card(s, M, 1.95, 6.35, 2.55, "名称规范化六步",
         "去法律后缀词表 → 去括号及内容 → 全 / 半角统一 →\n"
         "繁 / 简统一 → 去首尾空格标点 → 大小写统一\n"
         "保留原文字段，用于回溯。",
         tsize=15.5, bsize=13.5)
    card(s, 7.20, 1.95, 5.51, 2.55, "复核闸门",
         "规范化后相同只说明它们像，不能证明是同一家（分公司 / 连锁可能同名），必须进入人工复核队列。\n"
         "若存在某个字段悬而未决、无法判断取哪个值：挂起（hold），暂不合并。",
         tsize=15.5, bsize=13.5, accent=True)
    band(s, M, 4.72, CW, 0.78,
         "词表与是否引入相似度阈值需按本题标定，不臆造数值阈值。", 15, TEXT, LIGHT)
    band(s, M, 5.66, CW, 0.78,
         "疑似重复：规范化同名 → 人工复核 → 决定合并或挂起，宁可慢一步，也不要错合并。",
         15, NAVY, LIGHT)


@page("P14", "查重决策流程图与裁决", "第四章", "起点 → 同码？ → 规范化同名？ → 人工复核 → 合并 / 挂起")
def p14(s):
    # 第一行
    rect(s, 0.62, 1.88, 1.35, 0.62, fill=WHITE, line=SLATE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    text(s, 0.62, 1.88, 1.35, 0.62, "两条记录", 12.5, NAVY, True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    arrow(s, 2.02, 2.07, 0.28, 0.24)
    rect(s, 2.36, 1.84, 2.25, 0.70, fill=LIGHT, line=SLATE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    text(s, 2.36, 1.84, 2.25, 0.70, "Q1 同码？", 13, NAVY, True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    arrow(s, 4.66, 2.07, 0.28, 0.24)
    rect(s, 5.00, 1.84, 2.25, 0.70, fill=WHITE, line=NAVY,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    text(s, 5.00, 1.84, 2.25, 0.70, "确定重复\n合并候选", 12.5, NAVY, True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 4.66, 1.62, 0.60, 0.20, "是", 11, ORANGE, True, align=PP_ALIGN.CENTER)
    text(s, 3.20, 2.58, 0.60, 0.20, "否", 11, SLATE, True, align=PP_ALIGN.CENTER)
    arrow(s, 3.35, 2.58, 0.22, 0.24, direction="down")
    # 第二行
    rect(s, 2.36, 2.94, 2.25, 0.70, fill=LIGHT, line=SLATE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    text(s, 2.36, 2.94, 2.25, 0.70, "Q2 规范化同名？", 12.5, NAVY, True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    arrow(s, 4.66, 3.17, 0.28, 0.24)
    rect(s, 5.00, 2.94, 2.05, 0.70, fill=WHITE, line=NAVY,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    text(s, 5.00, 2.94, 2.05, 0.70, "疑似重复\n人工复核", 12, NAVY, True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    arrow(s, 7.10, 3.17, 0.26, 0.24)
    rect(s, 7.42, 2.94, 1.95, 0.70, fill=LIGHT, line=SLATE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    text(s, 7.42, 2.94, 1.95, 0.70, "Q3 未决字段？", 12.5, NAVY, True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    arrow(s, 9.42, 3.17, 0.26, 0.24)
    rect(s, 9.74, 2.94, 2.20, 0.70, fill=WHITE, line=ORANGE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    text(s, 9.74, 2.94, 2.20, 0.70, "挂起 hold\n不合并", 12.5, NAVY, True,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 第三行：Q2 否 -> 不重复
    text(s, 4.66, 3.70, 0.60, 0.20, "否", 11, SLATE, True, align=PP_ALIGN.CENTER)
    arrow(s, 3.35, 3.68, 0.22, 0.24, direction="down")
    rect(s, 5.00, 3.94, 2.25, 0.58, fill=WHITE, line=SLATE,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    text(s, 5.00, 3.94, 2.25, 0.58, "不重复，各自保留", 12, NAVY,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 裁决优先级
    text(s, M, 4.72, CW, 0.32, "裁决优先级", 14, NAVY, True)
    text(s, M, 5.08, CW, 1.35,
         "· 统一社会信用代码相同 → 确定重复（自动合并候选）\n"
         "· 空 / 无效信用代码 → 转名称分支\n"
         "· 企业名规范化后相同 → 疑似重复 → 人工复核，不得自动合并\n"
         "· 疑似但字段未决 → 挂起 hold，不合并，待人工补齐再定",
         14, TEXT, line=1.28)


@page("P15", "合并工作表与逐字段取值", "第四章", "保留主记录，逐字段挑值并写明理由")
def p15(s):
    data = [
        ["字段名", "主记录原值", "从记录值", "选定值", "选择理由（字典 / 来源 / 置信度）"],
        ["ent_name_c", "星海科技（简称）", "星海科技有限公司", "星海科技有限公司", "全称更完整，依字段字典中文名口径"],
        ["mobile_c", "空", "****5678", "****5678", "从记录非空；保留打码形态，明文须另行审批"],
        ["opp_stage_c", "PENDING", "ACTIVE", "ACTIVE", "来源记录更新鲜，且值在枚举值域内"],
    ]
    table(s, M, 1.95, CW, 2.90, data, [1.85, 2.20, 2.20, 2.20, 3.64], fsize=12, hsize=12.5,
          first_col_bold=True)
    card(s, M, 5.00, CW, 1.15, "合并动作要领",
         "保留主记录 ID；从记录标记 MERGED / 软删；逐字段选择取值并写明理由（依据字段字典、来源记录、置信度）；"
         "同一事务完成下游外键重指向——只改字段不重指向会造成数据分裂。",
         tsize=14, bsize=13)
    band(s, M, 6.19, CW, 0.44,
         "理由就是 30 天内被质疑时的证据，复核和撤销都有据可依。", 12.5, NAVY, LIGHT)


@page("P16", "30 天撤销、超窗审批与三演练", "第四章", "窗口起点 = 合并提交时刻")
def p16(s):
    card(s, M, 1.95, 5.85, 2.55, "30 天时间轴",
         "T0 合并提交 → 窗口内（≤30 天）可撤销重来\n"
         "撤销须回滚外键重指向与下游影响\n"
         "超 30 天：下游已扩散，改为审批流程，不得自行回滚",
         tsize=15.5, bsize=13.5)
    card(s, 6.86, 1.95, 5.85, 2.55, "三演练答案路径",
         "1. 同码不同名 → 确定重复 + 备注\n"
         "2. 规范化同名且有未决字段 → 疑似重复并挂起\n"
         "3. 合并后第 31 天发现错值 → 超窗须审批",
         tsize=15.5, bsize=13.5, accent=True)
    band(s, M, 4.72, CW, 0.88,
         "撤销须回滚外键重指向与下游影响；超 30 天因下游已扩散，改为审批流程而非直接撤销。",
         15, TEXT, LIGHT)
    band(s, M, 5.78, CW, 0.62,
         "对照答案路径核对自己每一步的判断依据，而不是只记结论。", 16, NAVY, LIGHT, bold=True)


@page("P17", "数据质量三率总览", "第五章", "三个数，三个不同维度，不能互相替代")
def p17(s):
    gauges = [
        ("完整率  92%", "关心该填的字段填没填全", "必填字段非空记录数 ÷ 活跃记录数"),
        ("重复率  3%", "关心有没有同一条记录存了两份", "命中确定 + 疑似规则记录数 ÷ 活跃数"),
        ("过期率  15%", "关心信息是不是太久没更新", "超过 N 天未更新记录数 ÷ 活跃数"),
    ]
    for i, (big, q, f) in enumerate(gauges):
        cx = M + i * 4.12
        rect(s, cx, 2.00, 3.85, 2.62, fill=WHITE, line=BORDER,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
        rect(s, cx + 3.55, 2.12, 0.10, 0.10, fill=ORANGE)
        text(s, cx + 0.15, 2.18, 3.55, 0.50, big, 24, NAVY, True,
             align=PP_ALIGN.CENTER)
        text(s, cx + 0.15, 2.82, 3.55, 0.70, q, 15, TEXT,
             align=PP_ALIGN.CENTER, line=1.2)
        text(s, cx + 0.15, 3.62, 3.55, 0.80, f, 12.5, SLATE,
             align=PP_ALIGN.CENTER, line=1.2)
    text(s, M, 4.78, CW, 0.32, "以上为虚构示例：活跃记录数 1000（分母不含软删行）", 12, SLATE,
         align=PP_ALIGN.CENTER)
    band(s, M, 5.20, CW, 0.80,
         "完整率 100% 只说明字段齐全，不代表没有重复，也不代表信息不过期。",
         16, NAVY, LIGHT)
    band(s, M, 6.06, CW, 0.46,
         "任何一项异常都指向不同的治理动作，看报告要把三个数一起看。", 13.5, TEXT, LIGHT)


@page("P18", "三率口径与分母", "第五章", "同样是 90%，分母不同含义天差地别")
def p18(s):
    data = [
        ["指标", "计算口径", "虚构示例（分母 1000）"],
        ["完整率", "必填字段非空记录数 ÷ 活跃记录数（记录级；亦可字段级，口径须统一）", "920 ÷ 1000 = 92%"],
        ["重复率", "命中确定 + 疑似规则记录数 ÷ 活跃记录数", "30 ÷ 1000 = 3%"],
        ["过期率", "超过 N 天未更新（updated_at）记录数 ÷ 活跃数（N 需标定）", "150 ÷ 1000 = 15%"],
    ]
    table(s, M, 1.95, CW, 2.45, data, [1.40, 7.25, 3.44], fsize=12.5, hsize=13,
          first_col_bold=True)
    card(s, M, 4.60, CW, 1.25, "四条口径纪律",
         "分母统一为 active（未软删）记录数；软删行不计入质量指标但留在审计；"
         "分母口径不一致会让三率失真；三率之和无业务意义。",
         tsize=14, bsize=13)
    band(s, M, 6.00, CW, 0.50,
         "阈值只作监控告警线，触发后由人核查而非自动改主数据。", 14, NAVY, LIGHT)


@page("P19", "读一份质量报告（演练）", "第五章", "报告不是给你打分，而是给你派活")
def p19(s):
    data = [
        ["对象（虚构）", "完整率", "重复率", "过期率", "异常提示 → 治理动作"],
        ["会员企业", "98%", "18%", "5%", "重复率明显偏高 → 回查重合并流程"],
        ["联系人", "92%", "3%", "22%", "过期率偏高 → 更新联系人 / 资质信息"],
        ["能力卡", "95%", "2%", "8%", "正常，持续监控"],
        ["需求卡", "90%", "4%", "12%", "关注过期率"],
        ["商机", "96%", "6%", "9%", "正常，持续监控"],
    ]
    table(s, M, 1.95, CW, 2.65, data, [2.30, 1.55, 1.55, 1.55, 5.14], fsize=12.5, hsize=13,
          first_col_bold=True)
    band(s, M, 4.82, CW, 0.82,
         "读报告顺序：先看整体三率 → 下钻到对象 → 锁定到具体行。", 16, NAVY, LIGHT, bold=True)
    band(s, M, 5.82, CW, 0.62,
         "每一个异常数字背后都对应一个具体动作；三率维度不同，齐全不等于不重复、不过期。",
         14.5, TEXT, LIGHT)


@page("P20", "权限角色与最小必要", "第六章", "查看不等于导出，导出须单独审批")
def p20(s):
    data = [
        ["角色", "负责对象 / 入口"],
        ["会员运营", "会员企业、联系人"],
        ["内容运营", "能力卡"],
        ["资源对接", "需求卡"],
        ["招商业务员", "商机"],
        ["活动运营", "活动参与记录"],
        ["数据管理员", "导入 / 查重 / 合并 / 质量 / 导出审批 / 审计 六入口"],
        ["说明", "明文查看敏感字段需更高权限且留痕"],
    ]
    table(s, M, 1.95, 7.20, 3.65, data, [2.00, 5.20], fsize=12.5, hsize=13,
          first_col_bold=True)
    card(s, 8.05, 1.95, 4.66, 3.65, "查看 ≠ 导出",
         "界面上看得到手机号 = 查看权限；\n"
         "把它变成一份文件带出去 = 导出权限，必须单独审批。\n\n"
         "最小必要：看什么取决于岗位需要，不是想看什么。数据管理员负责维护数据，不等于拥有全部数据的导出权。",
         tsize=16, bsize=13.5, accent=True)
    band(s, M, 5.78, CW, 0.62,
         "守住这条边界，数据才不会从“内部可见”滑向“外部不可控”。", 15, NAVY, LIGHT)


@page("P21", "手机号默认打码", "第六章", "默认打码让泄露没有可乘之机")
def p21(s):
    card(s, M, 1.95, 5.85, 2.60, "打码前后对照（虚构示例）",
         "完整形态（仅示意，界面不展示）：1X0 XXXX 5678\n"
         "界面显示（一律）：****5678\n"
         "规则：保留后 4 位，前段显示 ****（尾号虚构）；明文仅审批后可见。",
         tsize=15.5, bsize=14)
    card(s, 6.86, 1.95, 5.85, 2.60, "三条边界",
         "明文查看需更高权限，并留审计。\n"
         "打码后仍可能被关联重识别，不可当作彻底匿名。\n"
         "默认打码：你不需要记得去打码，系统已经帮你做了。",
         tsize=15.5, bsize=13.5, accent=True)
    band(s, M, 4.78, CW, 0.92,
         "确实需要完整号码时，走的是申请和审批，而不是绕过显示规则。", 16, NAVY, LIGHT, bold=True)
    band(s, M, 5.88, CW, 0.52,
         "打码规则统一执行，具体规则以知识卡 / 字段字典登记为准。", 13.5, TEXT, LIGHT)


@page("P22", "导出审批与文件水印", "第六章", "默认打码保护查看 · 审批控制导出 · 水印保障追溯")
def p22(s):
    card(s, M, 1.95, 5.85, 2.70, "导出审批单",
         "申请人 · 审批人 · 审批单号 · 数据范围 · 用途\n"
         "未审批一律拒绝。\n"
         "导出是数据离开系统的最后一道闸，获批后才能生成文件。",
         tsize=15.5, bsize=13.5)
    card(s, 6.86, 1.95, 5.85, 2.70, "文件水印字段",
         "导出人工号 · 导出时间（UTC）· 审批单号\n"
         "数据范围（org_id / site_id）· 文件序列号\n"
         "用途 · 页码；一旦外泄可追溯来源。",
         tsize=15.5, bsize=13.5, accent=True)
    band(s, M, 4.90, CW, 0.60,
         "默认打码保护查看 → 审批控制导出 → 水印保障追溯，三道闸门缺一不可。",
         16, NAVY, LIGHT, bold=True)
    band(s, M, 5.68, CW, 0.72,
         "水印不替代访问控制；导出文件一旦落地，每一次流转都应在操作审计里留痕。",
         14.5, TEXT, LIGHT)


@page("P23", "操作审计追踪的字段与链路", "第七章", "谁 · 何时 · 做了什么 · 结果状态")
def p23(s):
    card(s, M, 1.95, 3.85, 3.05, "审计回答四个问题",
         "谁操作 · 什么时间 · 做了什么动作 · 结果状态如何。\n"
         "每一次关键动作——导入、合并、撤销、导出、权限变更——都会留下一条记录。",
         tsize=15, bsize=13)
    card(s, 4.74, 1.95, 3.85, 3.05, "必须留痕的操作",
         "创建 / 修改 / 软删 / 恢复、批量导入、查重判定、合并与撤销、导出及审批、权限变更、敏感字段明文查看、AI 建议采纳驳回。",
         tsize=15, bsize=12.5)
    card(s, 8.86, 1.95, 3.85, 3.05, "审计记录字段",
         "审计ID、操作人、时间UTC、操作类型、对象类型+ID、变更前后值、来源IP、关联审批单、结果状态；append-only，不可 UPDATE / DELETE。",
         tsize=15, bsize=12.5, accent=True)
    band(s, M, 5.22, CW, 0.90,
         "审计的态度是中立的：只客观记录发生了什么，不替你判断对错。像飞机黑匣子，平时不显眼，关键时刻离不开。",
         15, NAVY, LIGHT)
    text(s, M, 6.24, CW, 0.32,
         "软删除行不计入质量指标，但变更仍留在审计中，可追溯。", 12.5, SLATE)


@page("P24", "审计复盘演练与边界", "第七章", "沿审计链路还原一次导出的来龙去脉")
def p24(s):
    steps = ["提交申请\n谁 / 何时", "审批\n审批人 / 单号", "导出对象", "水印字段", "审计留痕"]
    for i, t in enumerate(steps):
        cx = 0.72 + i * 2.42
        rect(s, cx, 2.12, 2.20, 1.00, fill=WHITE, line=SLATE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
        text(s, cx, 2.12, 2.20, 1.00, t, 12.5, NAVY, True,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line=1.15)
        if i < 4:
            arrow(s, cx + 2.22, 2.50, 0.18, 0.22)
    card(s, M, 3.62, 5.85, 1.50, "复盘要点",
         "谁在什么时间提交、审批人是谁、导出了哪个对象、文件带的水印是什么。每一步都能在审计里找到对应痕迹，这就是可追溯。",
         tsize=15, bsize=13)
    card(s, 6.86, 3.62, 5.85, 1.50, "边界回收",
         "软删除记录不计入质量指标，但变更仍在审计查得到；合并即使撤销，原始合并动作也留痕不抹除。",
         tsize=15, bsize=13, accent=True)
    band(s, M, 5.40, CW, 0.90,
         "审计给了数据管理员一层保护：只要按流程操作、有据可查，就不怕被质疑。",
         16, NAVY, LIGHT, bold=True)


@page("P25", "AI 能做什么、不能做什么", "第八章", "AI 输出是“建议”，不是“决定”")
def p25(s):
    card(s, M, 1.95, 3.85, 2.60, "可做",
         "发现线索：提示哪两条记录可能重复、哪个字段可能填错。\n"
         "输出定位是“建议”，供人参考。", tsize=16, bsize=13.5)
    card(s, 4.74, 1.95, 3.85, 2.60, "不可做",
         "不能直接写入系统；不得对外发送；不计算金额。\n"
         "金额涉及精确口径和责任归属，由系统按规则计算。",
         tsize=16, bsize=13.5, accent=True)
    card(s, 8.86, 1.95, 3.85, 2.60, "须复核",
         "建议必须进入复核队列，人工确认后才执行。\n"
         "合并决策权在人，确定性规则不因 AI 介入而改变。",
         tsize=16, bsize=13.5)
    card(s, M, 4.76, CW, 1.20, "三条红线",
         "红线 1：清洗建议必须进复核队列，不得直接写入　·　红线 2：不得对外发送　·　红线 3：不计算金额（金额一律系统按分存整数）",
         tsize=14, bsize=13.5)
    band(s, M, 6.06, CW, 0.46, "AI 是你的助手，不是你的替身。", 15, NAVY, LIGHT, bold=True)


@page("P26", "AI 复核队列与三条红线", "第八章", "三个场景，三种越界")
def p26(s):
    flow = ["AI 建议", "复核队列", "逐条确认", "正常写入"]
    for i, t in enumerate(flow):
        cx = 0.62 + i * 3.11
        rect(s, cx, 1.98, 2.75, 0.76, fill=WHITE, line=SLATE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
        text(s, cx, 1.98, 2.75, 0.76, t, 14, NAVY, True,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        if i < 3:
            arrow(s, cx + 2.80, 2.24, 0.24, 0.24)
    cases = [
        ("场景 1", "“我帮你把这两条自动合并了”", "越界：合并须人工复核确认"),
        ("场景 2", "“我把这份清单发给合作方吧”", "越界：不得对外发送"),
        ("场景 3", "“我算一下总金额”", "越界：AI 不计算金额，系统按分存整数"),
    ]
    for i, (tag, quote, why) in enumerate(cases):
        cx = M + i * 4.12
        card(s, cx, 3.05, 3.85, 1.75, tag,
             quote + "\n→ " + why, tsize=15, bsize=12.5, accent=(i == 2))
    band(s, M, 5.05, CW, 0.90,
         "请一一识别越界，并说出该用什么正确做法替代。", 16, NAVY, LIGHT, bold=True)
    text(s, M, 6.12, CW, 0.32,
         "凡涉及写库、对外发送、金额，必须有人或系统规则兜底。", 13, SLATE)


@page("P27", "全课知识地图与易错点", "总结", "把整门课收成一张主线图")
def p27(s):
    row1 = ["字段字典\n定规则", "导入中心\n搬数据", "查重合并\n处理重复", "质量报告\n看健康"]
    row2 = ["导出审批与水印\n管出口", "操作审计\n留痕迹", "AI 工具\n守边界"]
    for i, t in enumerate(row1):
        cx = 0.72 + i * 3.11
        rect(s, cx, 1.95, 2.75, 0.70, fill=LIGHT, line=SLATE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
        text(s, cx, 1.95, 2.75, 0.70, t, 13, NAVY, True,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line=1.1)
        if i < 3:
            arrow(s, cx + 2.80, 2.18, 0.24, 0.24)
    for i, t in enumerate(row2):
        cx = 0.72 + i * 3.11
        rect(s, cx, 2.82, 2.75, 0.70, fill=LIGHT, line=SLATE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
        text(s, cx, 2.82, 2.75, 0.70, t, 13, NAVY, True,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, line=1.1)
        if i < 2:
            arrow(s, cx + 2.80, 3.05, 0.24, 0.24)
    traps = [
        ("易错点一 · SoR", "六类对象的唯一可信源搞不清 → 冲突时以 SoR 为准。"),
        ("易错点二 · 查重", "把疑似当确定直接合并 → 规范化同名只是疑似，必须人工复核。"),
        ("易错点三 · AI", "把 AI 建议当指令执行 → 进复核队列、不直写、不外发、不算金额。"),
    ]
    for i, (t, d) in enumerate(traps):
        cx = M + i * 4.12
        card(s, cx, 3.80, 3.85, 1.68, t, d, tsize=14.5, bsize=12.5, accent=(i == 2))
    band(s, M, 5.72, CW, 0.68,
         "守住这三处，你就具备了独立上岗的基本判断力。", 16, NAVY, LIGHT, bold=True)


@page("P28", "综合认证题与判分说明", "认证", "三问 · 每问 1 分，共 3 分")
def p28(s):
    band(s, M, 1.90, CW, 1.10,
         "情景（虚构）：会员数据导入后重复率异常升高；两条记录企业名规范化后相同，但统一社会信用代码不同，其中一条缺必填字段；"
         "你想用 AI 直接完成合并和金额汇总。",
         14, NAVY, LIGHT)
    qa = [
        ("（1）重复类别与处理", "疑似重复，须人工复核；因存在未决字段，挂起而不合并。"),
        ("（2）缺必填字段的导入处置", "导入时即被校验报告拦截，修正后重传。"),
        ("（3）AI 工具的越界点", "建议进复核队列，不得直接写入、不得对外发送、不得计算金额。"),
    ]
    for i, (t, d) in enumerate(qa):
        cx = M + i * 4.12
        card(s, cx, 3.22, 3.85, 2.00, t, d, tsize=14.5, bsize=13)
    band(s, M, 5.44, CW, 0.86,
         "判分：每问 1 分共 3 分；漏“挂起”扣 0.5；AI 三红线少一条扣 0.5；类别答“确定重复”不得分。",
         14, TEXT, LIGHT)
    text(s, M, 6.38, CW, 0.32, "守住边界，就是数据管理员的价值。", 14, NAVY, True)


# ---- 讲解词逐字提取 ------------------------------------------------------
def extract_notes():
    notes = {}
    cur = None
    with open(LESSON, "r", encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            m = re.match(r"^###\s+(P\d{2})[\s\u3000]", line)
            if m:
                cur = m.group(1)
                continue
            if cur and line.startswith("**讲解词：**"):
                notes[cur] = line[len("**讲解词：**"):].strip()
                cur = None
    expected = ["P%02d" % i for i in range(1, 29)]
    missing = [p for p in expected if p not in notes]
    if missing:
        raise RuntimeError("讲解词提取缺失: " + ", ".join(missing))
    return notes


def attach_notes(slide, note_text):
    tf = slide.notes_slide.notes_text_frame
    tf.text = note_text
    for p in tf.paragraphs:
        _ppr(p, PP_ALIGN.LEFT, 1.2)
        for r in p.runs:
            _font(r, 11, TEXT)


def seal_line_break(prs):
    """收尾扫描：所有 a:p（含线框空段落、表格单元格）补齐避头尾属性。"""
    for slide in prs.slides:
        for p_el in slide._element.iter(qn("a:p")):
            pPr = p_el.find(qn("a:pPr"))
            if pPr is None:
                pPr = p_el.makeelement(qn("a:pPr"), {})
                p_el.insert(0, pPr)
            pPr.set("eaLnBrk", "1")
            pPr.set("hangingPunct", "1")


def main():
    notes = extract_notes()
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(notes, f, ensure_ascii=False, indent=2)

    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    blank = prs.slide_layouts[6]

    for pid, title, chapter, subtitle, fn in PAGES:
        slide = prs.slides.add_slide(blank)
        frame(slide, pid, title, chapter, subtitle)
        fn(slide)
        attach_notes(slide, notes[pid])

    if len(prs.slides._sldIdLst) != 28:
        raise RuntimeError("页数不是 28")
    seal_line_break(prs)
    prs.save(OUT_PPTX)
    print("生成完成：%s (%d 页)" % (os.path.basename(OUT_PPTX), len(prs.slides._sldIdLst)))
    print("生成完成：%s (%d 条)" % (os.path.basename(OUT_JSON), len(notes)))


if __name__ == "__main__":
    main()
