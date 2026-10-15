# -*- coding: utf-8 -*-
"""实验一报告：四、实验交付物 之后的小节序号改为交付物序号 1-6。"""
import copy
import docx
from docx.oxml.ns import qn

SRC = 'E:/软件项目管理实验/实验1/软件项目管理-实验一(最终提交物).docx'

# 交付物 1-6 的新序号（旧文本 -> 新文本）
RENUMBER = {
    '二、需求分解清单': '1. 需求分解清单',
    '2.1  Epic 描述': '1.1  Epic 描述',
    '2.2  Feature 分解与初步优先级（MoSCoW）': '1.2  Feature 分解与初步优先级（MoSCoW）',
    '2.3  User Story 列表（共 8 个）': '1.3  User Story 列表（共 8 个）',
    '三、验收标准文档': '2. 验收标准文档',
    '3.1  核心故事一：US-04 在线提交宿舍报修': '2.1  核心故事一：US-04 在线提交宿舍报修',
    '3.2  核心故事二：US-05 宿管员处理报修工单': '2.2  核心故事二：US-05 宿管员处理报修工单',
    '3.3  核心故事三：US-03 宿舍长登记晚归 / 未归': '2.3  核心故事三：US-03 宿舍长登记晚归 / 未归',
    '四、故事地图': '3. 故事地图',
    '4.1  用户故事叙事': '3.1  用户故事叙事',
    # 4.2 拆成独立一节「4. 骨架行说明」，正文标题另建 4.1 / 4.2
    '4.2  骨架行（Backbone）及提取理由': '4. 骨架行说明',
    '骨架行提取理由：': '4.2  提取理由',
    '4.3  故事地图全图': '3.2  故事地图全图',
    '五、MoSCoW 分类表与版本规划（交付物 5）': '5. MoSCoW 分类表与版本规划',
    '六、实验心得体会': '6. 实验心得体会',
}

d = docx.Document(SRC)
body = d.element.body


def ptext(el):
    return ''.join(t.text or '' for t in el.iter(qn('w:t')))


def paras():
    return [el for el in body.iterchildren() if el.tag == qn('w:p')]


def find_para(old):
    hits = [el for el in paras() if ptext(el).strip() == old]
    assert len(hits) == 1, f'expected 1 match for {old!r}, got {len(hits)}'
    return hits[0]


def set_text(p, new):
    runs = p.findall(qn('w:r'))
    assert runs, 'paragraph has no runs'
    first = runs[0]
    ts = first.findall(qn('w:t'))
    assert ts, 'first run has no w:t'
    ts[0].text = new
    ts[0].set(qn('xml:space'), 'preserve')
    for t in ts[1:]:
        first.remove(t)
    for r in runs[1:]:
        p.remove(r)


def restyle(p, template):
    """让 p 的段落/字符格式与 template 标题完全一致。"""
    pPr = p.find(qn('w:pPr'))
    if pPr is not None:
        p.remove(pPr)
    tPr = template.find(qn('w:pPr'))
    if tPr is not None:
        p.insert(0, copy.deepcopy(tPr))
    t_runs = template.findall(qn('w:r'))
    runs = p.findall(qn('w:r'))
    if t_runs and runs:
        t_rPr = t_runs[0].find(qn('w:rPr'))
        rPr = runs[0].find(qn('w:rPr'))
        if rPr is not None:
            runs[0].remove(rPr)
        if t_rPr is not None:
            runs[0].insert(0, copy.deepcopy(t_rPr))


# --- 0. 先抓住格式模板（改文本之前） -------------------------------
top_tpl = find_para('四、故事地图')          # 一级交付物标题格式
sub_tpl = find_para('4.1  用户故事叙事')     # 二级小节标题格式

# --- 1. 重编号 -----------------------------------------------------
for old, new in RENUMBER.items():
    set_text(find_para(old), new)

# --- 2. 「4. 骨架行说明」升为一级标题格式 ---------------------------
sec4 = find_para('4. 骨架行说明')
restyle(sec4, top_tpl)

# --- 3. 正文小节标题「4.2  提取理由」套用二级标题格式 ---------------
sub42 = find_para('4.2  提取理由')
restyle(sub42, sub_tpl)

# --- 4. 在骨架行表格前插入「4.1  骨架行（Backbone）」 ---------------
backbone_tbl = None
for nxt in paras():
    if ptext(nxt).strip() == '4. 骨架行说明':
        sib = nxt.getnext()
        while sib is not None and sib.tag != qn('w:tbl'):
            sib = sib.getnext()
        backbone_tbl = sib
        break
assert backbone_tbl is not None, 'backbone table not found'

new41 = copy.deepcopy(sub_tpl)
set_text(new41, '4.1  骨架行（Backbone）')
backbone_tbl.addprevious(new41)

# --- 5. 把「3.2 故事地图全图」+ 全图表格 上移到「用户故事叙事」之后 ---
p_narr = None
for el in paras():
    if ptext(el).strip().startswith('"我是大二学生小林'):
        p_narr = el
        break
assert p_narr is not None, 'narration paragraph not found'

p_map = find_para('3.2  故事地图全图')
map_tbl = p_map.getnext()
assert map_tbl is not None and map_tbl.tag == qn('w:tbl'), 'full-map table not found'

body.remove(p_map)
body.remove(map_tbl)
p_narr.addnext(p_map)
p_map.addnext(map_tbl)

d.save(SRC)
print('saved OK')
