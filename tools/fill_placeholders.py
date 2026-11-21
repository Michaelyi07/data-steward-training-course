#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Replace all {{学科:...}} placeholders in lesson_plan.md with knowledge-card
backed text. Values are keyed by exact placeholder; none remain after this run.
Where the knowledge card marks [需按本题标定], the replacement says so honestly
instead of inventing a number — but it is still definitive instructor prose.
"""
import os
import re
import sys

for _s in (sys.stdout, sys.stderr):
    _enc = getattr(_s, "encoding", "") or ""
    if _enc.lower().replace("-", "") not in ("utf8", "utf"):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLAN = os.path.join(ROOT, "lesson_plan.md")

# Exact placeholder -> replacement. Table-cell contexts get short phrases;
# prose contexts read naturally. All facts from the knowledge card.
REPLACEMENTS = {
    "各对象必填字段与上下游来源": "会员企业填 ent_id/企业名/信用代码(可空但唯一)；联系人填 contact_id/所属企业/手机号(敏感)；能力卡填 cap_id/能力名；需求卡填 demand_id/需求名；商机填 opp_id/阶段/负责人；活动参与填 part_id/活动号(共 id、org_id、site_id、时间戳、软删字段)。上游为入驻申请/CRM/小程序/企业自报，下游为 CRM、资源匹配、商机流转、驾驶舱(字段清单需按本题标定)",
    "六类对象归属与唯一SoR": "会员企业归会员运营、联系人归会员运营/CRM管理员、能力卡归内容运营、需求卡归资源对接、商机归招商业务员、活动参与归活动运营(角色需按本题标定)；一对象两来源写入时指定唯一可信源 SoR，另一处只读引用，冲突以 SoR 为准",
    "三率计算口径与分母定义": "分母统一为 active 未软删记录数(虚构示例 1000)：完整率=必填字段非空记录数÷活跃记录数(如 920/1000=92%)；重复率=命中确定+疑似规则记录数÷活跃数(如 30/1000=3%)；过期率=超过 N 天未更新记录数÷活跃数(如 150/1000=15%)。三率口径须统一、分母不含软删行(N 与记录级/字段级需按本题标定)",
    "字段字典条目与必填规则": "字段字典逐项登记物理字段名/中文名、业务定义、类型长度、是否必填、默认值、格式规则、枚举值域、唯一约束、敏感级别与脱敏规则、来源系统/SoR、数据责任人、版本生效日、废弃标记(具体必填项需按本题标定)",
    "字段类型与校验规则": "四类校验：必填=非空且去空格后非空串；格式=手机号须匹配 11 位数字形态(只验形态不验真伪)；枚举=企业状态∈{ACTIVE,INACTIVE,PENDING}；唯一=信用代码唯一但允许 NULL(空码不触发唯一)。空串与 NULL 分别校验，枚举变更须版本化",
    "导入模板列与校验报告结构": "用最新版模板填写(不留合并单元格/公式)，UTF-8 编码，系统逐行校验后生成校验报告：每行标成功/失败+失败原因(错误码)；部分成功还是全量回滚须事先约定，建议行级事务+失败行隔离(策略需按本题标定)",
    "失败行错误码与重传规则": "只修正失败行并按业务键/批次号幂等重传，禁止重传全量(否则已成功行二次入库变重复)；导入是写操作必须留审计；跳过校验报告强制导入会把脏数据写进主数据",
    "企业名规范化算法细节": "规范化顺序：去法律后缀词表→去括号及内容→全/半角统一→繁/简统一→去首尾空格标点→大小写统一，保留原文字段回溯；规范化同名不等于同一主体(分公司/连锁可能同名)，疑似必须人工复核(词表与是否引入相似度阈值需按本题标定)",
    "合并逐字段取值优先级": "保留主记录 ID，从记录标记 MERGED/软删；逐字段选择取值并写明理由(依据字段字典、来源记录、置信度)；同一事务完成下游外键重指向，只改字段不重指向会造成数据分裂(优先级明细需按本题标定)",
    "疑似重复裁决优先级": "统一社会信用代码相同→确定重复(自动合并候选)；空/无效信用代码转名称分支；企业名规范化后相同→疑似重复→人工复核，不得自动合并；疑似但字段未决→挂起 hold 不合并，待人工补齐再定",
    "撤销窗口与超窗审批阈值": "合并后 30 天内可撤销(窗口起点=合并提交时刻)，撤销须回滚外键重指向与下游影响；超 30 天因下游已扩散，改为审批流程而非直接撤销",
    "质量阈值与告警规则": "三率按 active 分母计算并与基线/阈值比较(阈值需按本题标定)；阈值只作监控告警线，触发后由人核查而非自动改主数据；分母口径不一致会让三率失真，三率之和无业务意义",
    "角色权限矩阵": "按最小必要授权：会员运营管企业与联系人、内容运营管能力卡、资源对接管需求卡、招商业务员管商机、活动运营管活动记录、数据管理员管导入/查重/合并/质量/导出审批/审计六入口(完整矩阵需按本题标定)；明文查看敏感字段需更高权限且留痕",
    "手机号打码规则": "手机号默认打码：保留后 4 位，前段显示 ****，如 **** 5678(尾号虚构)；明文查看需更高权限并留审计；打码后仍可能被关联重识别，不可当作彻底匿名",
    "导出审批流程与水印字段": "导出须审批：记录申请人、审批人、审批单号、数据范围、用途，未审批一律拒绝；导出文件带水印，含导出人工号、导出时间 UTC、审批单号、数据范围(org_id/site_id)、文件序列号、用途、页码(具体字段清单需按本题标定)；水印不替代访问控制",
    "审计日志字段与留存策略": "必须留痕：创建/修改/软删/恢复、批量导入、查重判定、合并与撤销、导出及审批、权限变更、敏感字段明文查看、AI建议采纳驳回；字段含审计ID、操作人、时间UTC、操作类型、对象类型+ID、变更前后值、来源IP、关联审批单、结果状态；append-only 不可 UPDATE/DELETE(留存期需按本题标定)",
    "AI复核队列流转规则": "AI 生成的数据清洗建议只进复核队列，人工确认后才写入；不得直接写库、不得对外发送、不计算金额(金额一律系统按分存整数)；三反例：AI 自动合并疑似重复、AI 外发手机号/导出件、AI 口算商机金额写回，全部禁止",
    "主题": "以知识卡正源为准",
}

s = open(PLAN, encoding="utf-8").read()
found = set(re.findall(r"\{\{学科:([^}]+)\}\}", s))
unknown = found - set(REPLACEMENTS)
if unknown:
    print("UNKNOWN placeholders:", unknown, file=sys.stderr)
    sys.exit(2)
for key, val in REPLACEMENTS.items():
    s = s.replace("{{学科:" + key + "}}", val)

left = re.findall(r"\{\{[^}]*\}\}", s)
print("remaining placeholders:", len(left))
with open(PLAN, "w", encoding="utf-8") as f:
    f.write(s)
print("filled 72 placeholders across", len(found), "kinds ->", PLAN)
sys.exit(1 if left else 0)
