# -*- coding: utf-8 -*-
"""tnu.py — 天津师范大学 PPT 模板的「原页编辑」引擎。

设计原则（这是本技能存在的理由）：
    模板页已经排好版。正确做法是 **改它已有形状里的文字**，不是新建文本框。
    上一版之所以难看，就是因为用 add_textbox 在模板上又画了一套版式。

四个子命令
    inventory  把模板/成品解析成 JSON（每页每形状的 id、几何、占位符、字体）
    catalog    生成人可读的页面目录（找「哪一页适合放什么」用）
    fill       按 spec.json 把文字/图表/表格写进**既有**形状，并裁剪排序页面
    verify     对比模板与成品：新增形状数必须为 0；越界/溢出/重叠体检
    render     调 PowerPoint COM 导出逐页 PNG（本机只有 powershell，没有 pwsh）

用法
    python -X utf8 tnu.py inventory <pptx> [--out inv.json] [--pages 1,6,14]
    python -X utf8 tnu.py catalog   <pptx> [--out catalog.md]
    python -X utf8 tnu.py fill      <spec.json> [--out 成品.pptx]
    python -X utf8 tnu.py verify    <模板.pptx> <成品.pptx> [--json]
    python -X utf8 tnu.py render    <pptx> <outdir> [--w 1600] [--h 900]
"""
import argparse
import copy
import json
import os
import subprocess
import sys

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

EMU = 914400.0
GROUP = 6  # MSO_SHAPE_TYPE.GROUP


# --------------------------------------------------------------- 形状查找
def iter_shapes(shapes):
    """递归遍历（含 GROUP 内嵌套），模板里有大量组合。"""
    for s in shapes:
        yield s
        if s.shape_type == GROUP:
            for t in iter_shapes(s.shapes):
                yield t


def shp(slide, sid):
    """按 shape_id 找形状；也接受 "ph:11" 这种按**占位符 idx** 的寻址。

    占位符用 idx 寻址比 shape_id 稳：模板里几十页共用同一套 shape_id，
    而封面 / 目录 / 过渡页的 idx 是版式语义（11 主标题、12 副题、13 英文…）。
    """
    key = str(sid)
    if key.startswith("ph:"):
        idx = int(key[3:])
        for s in slide.shapes:
            try:
                if s.is_placeholder and s.placeholder_format.idx == idx:
                    return s
            except Exception:
                pass
        raise KeyError("模板页 P%s 上没有 placeholder idx=%d"
                       % (getattr(slide, "_tnu_no", "?"), idx))
    for s in iter_shapes(slide.shapes):
        if s.shape_id == int(key):
            return s
    raise KeyError("模板页 P%s 上没有 shape_id=%s"
                   % (getattr(slide, "_tnu_no", "?"), key))


def ids(slide):
    return {s.shape_id for s in iter_shapes(slide.shapes)}


# --------------------------------------------------------------- 文字写入
def _first_nonempty_p(txBody):
    ps = txBody.findall(qn("a:p"))
    for p in ps:
        if "".join(t.text or "" for t in p.iter(qn("a:t"))).strip():
            return p
    return ps[0]


def set_paras(shape, paras, size=None, bold=None, wrap=None, spacing=None,
              align=None, color=None, font=None):
    """替换文字但**保留原段落/字符格式**。

    三个必须踩对的坑：
      1. 以 txBody 里第一个非空段落作模板 deepcopy——它的 pPr/rPr 带着
         模板给定的字体、颜色、行距，扔掉就等于自己重排版。
      2. a:endParaRPr 必须是 <a:p> 的**最后一个**子元素，否则 PowerPoint
         会忽略这个 run，表现为「文字填了但不显示」。
      3. 一个段落只留一个 run，其余删除；多 run 会导致字数统计与换行错乱。
    """
    tf = shape.text_frame
    if wrap is not None:
        tf.word_wrap = wrap
    txBody = tf._txBody
    if not txBody.findall(qn("a:p")):
        tf.text = ""
    tpl = copy.deepcopy(_first_nonempty_p(txBody))
    for p in txBody.findall(qn("a:p")):
        txBody.remove(p)

    n = len(paras)
    sizes = size if isinstance(size, (list, tuple)) else [size] * n
    bolds = bold if isinstance(bold, (list, tuple)) else [bold] * n

    for txt, sz, bd in zip(paras, sizes, bolds):
        p = copy.deepcopy(tpl)
        runs = p.findall(qn("a:r"))
        if not runs:
            r = p.makeelement(qn("a:r"), {})
            r.append(r.makeelement(qn("a:t"), {}))
            p.append(r)
            runs = p.findall(qn("a:r"))
        for r in runs[1:]:
            p.remove(r)
        r = runs[0]
        for t in r.findall(qn("a:t")):
            r.remove(t)
        t = r.makeelement(qn("a:t"), {})
        t.text = txt
        r.append(t)
        endPr = p.find(qn("a:endParaRPr"))
        if endPr is not None:
            p.remove(endPr)
            p.append(endPr)
        rPr = r.find(qn("a:rPr"))
        if rPr is None:
            rPr = r.makeelement(qn("a:rPr"), {})
            r.insert(0, rPr)
        if sz is not None:
            rPr.set("sz", str(int(round(float(sz) * 100))))
        if bd is not None:
            rPr.set("b", "1" if bd else "0")
        if color is not None:
            sf = rPr.makeelement(qn("a:solidFill"), {})
            c = sf.makeelement(qn("a:srgbClr"), {"val": color.lstrip("#").upper()})
            sf.append(c)
            for old in rPr.findall(qn("a:solidFill")):
                rPr.remove(old)
            rPr.insert(0, sf)
        if font is not None:
            for tag in ("a:latin", "a:ea", "a:cs"):
                for old in rPr.findall(qn(tag)):
                    rPr.remove(old)
                e = rPr.makeelement(qn(tag), {"typeface": font})
                rPr.append(e)
        pPr = p.find(qn("a:pPr"))
        if align is not None or spacing is not None:
            if pPr is None:
                pPr = p.makeelement(qn("a:pPr"), {})
                p.insert(0, pPr)
        if align is not None:
            pPr.set("algn", align)
        if spacing is not None:
            ln = pPr.find(qn("a:lnSpc"))
            if ln is None:
                ln = pPr.makeelement(qn("a:lnSpc"), {})
                pPr.insert(0, ln)
            for c in list(ln):
                ln.remove(c)
            ln.append(ln.makeelement(qn("a:spcPct"),
                                     {"val": str(int(round(spacing * 1000)))}))
        txBody.append(p)
    return shape


def fill(slide, mapping, *, quiet=False):
    """mapping: {shape_id: "文字" | {"paras": [...], "size":.., "bold":..}}"""
    for sid, spec in mapping.items():
        s = shp(slide, sid)
        if isinstance(spec, str):
            spec = {"paras": [spec]}
        elif isinstance(spec, (list, tuple)):
            spec = {"paras": list(spec)}
        set_paras(s, spec["paras"], size=spec.get("size"), bold=spec.get("bold"),
                  wrap=spec.get("wrap"), spacing=spec.get("spacing"),
                  align=spec.get("align"), color=spec.get("color"),
                  font=spec.get("font"))
        if not quiet:
            print("  fill sid=%-5s %s" % (sid, spec["paras"][0][:28]))


def drop(slide, sid):
    s = shp(slide, sid)
    s._element.getparent().remove(s._element)


def place(slide, sid, left=None, top=None, width=None, height=None):
    """只改 off/ext；这是「不新建形状」前提下唯一的排版手段。"""
    s = shp(slide, sid)
    for attr, val in (("left", left), ("top", top),
                      ("width", width), ("height", height)):
        if val is not None:
            setattr(s, attr, Emu(int(float(val) * EMU)))
    return s


def anchor(slide, sid, how="t"):
    shp(slide, sid).text_frame.vertical_anchor = {"t": 1, "m": 3, "b": 4}[how]


def header(slide, zh, en, no):
    """页眉三占位符按 placeholder idx 定位：11 中文短名 / 12 序号 / 13 英文。"""
    got = {}
    for s in slide.shapes:
        try:
            if s.is_placeholder:
                got[s.placeholder_format.idx] = s
        except Exception:
            pass
    for idx, txt, sz in ((11, zh, 14), (13, en, 11), (12, no, 14)):
        if idx in got and txt:
            set_paras(got[idx], [txt], size=sz,
                      bold=(idx != 13), align="ctr" if idx == 12 else None)
    return sorted(got)


def widen_title(slide, sid, left=1.665, width=10.0, size=None, text=None):
    """把模板那个 3.33 英寸宽的居中标题框，以版面中心 6.665 为轴左右扩开。

    模板原文是「在此输入您的标题」，宽度只够 5 个字，陈述句标题必然折行。
    这是「不新建形状」前提下唯一需要改几何的地方。
    """
    place(slide, sid, left=left, width=width)
    s = shp(slide, sid)
    txt = text if text is not None else s.text_frame.text
    set_paras(s, [txt], size=size, bold=True, align="ctr")


def set_chart(slide, sid, categories=None, series=None, title=None, title_size=16):
    ch = shp(slide, sid).chart
    if categories is not None:
        cd = CategoryChartData()
        cd.categories = list(categories)
        for name, vals in (series or {}).items():
            cd.add_series(name, tuple(vals))
        ch.replace_data(cd)
    if title:
        ch.has_title = True
        tf = ch.chart_title.text_frame
        tf.text = title
        for p in tf.paragraphs:
            for r in p.runs:
                r.font.size = Pt(title_size)
                r.font.bold = True
    else:
        ch.has_title = False
    return ch


def set_table(slide, sid, rows, size=13, allow_empty_pad=True):
    """覆写模板表格。

    ★ 必须清空没用到的格子。模板 5×5 里装着「类别5 / 85 / 21 …」这类示例
      数字；如果只写前 4 列，第 5 列会**静默保留假数据**，成品看起来像真的。
      所以：行数不足也照样把剩余行清空；列数不足按 allow_empty_pad 补空，
      超出行/列才报错（模板表格不能增行增列）。
    """
    tbl = shp(slide, sid).table
    nrow, ncol = len(tbl.rows), len(tbl.columns)
    need_r = len(rows)
    need_c = max((len(r) for r in rows), default=0)
    if need_r > nrow or need_c > ncol:
        raise ValueError("表格是 %dx%d，给的是 %dx%d——模板表格不能增行增列，"
                         "要么删要点，要么换页" %
                         (nrow, ncol, need_r, need_c))
    if need_c < ncol and not allow_empty_pad:
        raise ValueError("表格有 %d 列，只给了 %d 列；模板多余列会残留示例数据"
                         % (ncol, need_c))
    for i in range(nrow):
        row = rows[i] if i < need_r else []
        for j in range(ncol):
            cell = tbl.cell(i, j)
            cell.text_frame.text = row[j] if j < len(row) else ""
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(size)
    return tbl


# --------------------------------------------------------------- inventory
def _geom(s):
    try:
        return [round(s.left / EMU, 2), round(s.top / EMU, 2),
                round(s.width / EMU, 2), round(s.height / EMU, 2)]
    except TypeError:
        return None


def _sizes(shape):
    out = []
    try:
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size:
                    out.append(round(r.font.size.pt, 1))
    except Exception:
        pass
    return sorted(set(out))


def shape_info(s):
    d = {"shape_id": s.shape_id, "name": s.name, "type": str(s.shape_type),
         "geom": _geom(s)}
    if s.shape_type == GROUP:
        d["children"] = [shape_info(t) for t in s.shapes]
        return d
    try:
        if s.is_placeholder:
            d["ph"] = {"idx": s.placeholder_format.idx,
                       "type": str(s.placeholder_format.type)}
    except Exception:
        pass
    if s.has_text_frame and s.text_frame.text.strip():
        d["text"] = s.text_frame.text.strip()
        d["sizes"] = _sizes(s)
        d["paras"] = len(s.text_frame.paragraphs)
    if getattr(s, "has_chart", False):
        d["chart"] = {"type": str(s.chart.chart_type),
                      "series": len(s.chart.plots[0].series)}
    if getattr(s, "has_table", False):
        d["table"] = {"rows": len(s.table.rows), "cols": len(s.table.columns)}
    if getattr(s, "has_text_frame", False):
        d["wrap"] = s.text_frame.word_wrap
    return d


def inventory(path, pages=None):
    prs = Presentation(path)
    out = {"file": path, "slides": len(prs.slides),
           "canvas": [round(prs.slide_width / EMU, 3),
                      round(prs.slide_height / EMU, 3)],
           "layouts": [l.name for l in prs.slide_masters[0].slide_layouts],
           "pages": []}
    for i, sl in enumerate(prs.slides, 1):
        if pages and i not in pages:
            continue
        out["pages"].append({
            "no": i,
            "layout": sl.slide_layout.name,
            "shapes": [shape_info(s) for s in sl.shapes],
        })
    return out


def cmds_inventory(a):
    pages = {int(x) for x in a.pages.split(",")} if a.pages else None
    inv = inventory(a.pptx, pages=pages)
    txt = json.dumps(inv, ensure_ascii=False, indent=1)
    if a.out:
        with open(a.out, "w", encoding="utf-8") as f:
            f.write(txt)
        print("written", a.out, len(txt), "chars")
    else:
        print(txt)


def _walk(shapes, depth=0, lid=""):
    lines = []
    for s in shapes:
        g = _geom(s)
        head = "%s%s%s" % ("  " * depth,
                           "<GROUP>" if s.shape_type == GROUP else "",
                           "%s %s" % (s.shape_id, s.name))
        info = ""
        try:
            if s.is_placeholder:
                info += " [PH idx=%d %s]" % (s.placeholder_format.idx,
                                             s.placeholder_format.type)
        except Exception:
            pass
        if g:
            info += "  %.2f,%.2f  %.2fx%.2f" % tuple(g)
        if getattr(s, "has_chart", False):
            info += "  [CHART %s]" % s.chart.chart_type
        if getattr(s, "has_table", False):
            info += "  [TABLE %dx%d]" % (len(s.table.rows), len(s.table.columns))
        try:
            if s.has_text_frame and s.text_frame.text.strip():
                t = s.text_frame.text.strip().replace("\n", " / ")
                sz = _sizes(s)
                info += '  "%s"%s' % (t[:38], (" @%spt" % "/".join(map(str, sz))) if sz else "")
        except Exception:
            pass
        lines.append(head + info)
        if s.shape_type == GROUP:
            lines += _walk(s.shapes, depth + 1, lid)
    return lines


def cmds_catalog(a):
    prs = Presentation(a.pptx)
    L = []
    L.append("# 模板页面目录（自动生成，勿手改）\n")
    L.append("源文件：`%s`　页数：%d　画布：%.3f×%.3f 英寸\n"
             % (a.pptx, len(prs.slides), prs.slide_width / EMU,
                prs.slide_height / EMU))
    L.append("**用法**：先在本表里按「文字框容量 / 是否带图表表格」挑页，"
             "再按 `shape_id` 填 spec.json。\n")
    L.append("## 版式一览（%d 个）\n" % len(prs.slide_masters[0].slide_layouts))
    for i, l in enumerate(prs.slide_masters[0].slide_layouts):
        n = sum(1 for s in prs.slides if s.slide_layout.name == l.name)
        L.append("- `%d` %s（模板中被用 %d 次）" % (i, l.name, n))
    L.append("")
    for i, sl in enumerate(prs.slides, 1):
        txts = []
        for s in iter_shapes(sl.shapes):
            try:
                if s.has_text_frame and s.text_frame.text.strip():
                    txts.append(s.text_frame.text.strip().replace("\n", " / ")[:30])
            except Exception:
                pass
        L.append("\n## P%02d　%s" % (i, sl.slide_layout.name))
        if txts:
            L.append("> " + " ｜ ".join(txts[:4]))
        L.append("\n```")
        L += _walk(sl.shapes)
        L.append("```")
    txt = "\n".join(L)
    if a.out:
        with open(a.out, "w", encoding="utf-8") as f:
            f.write(txt)
        print("written", a.out, len(txt), "chars")
    else:
        print(txt)


# --------------------------------------------------------------- fill
def _load_json(path):
    """读 spec；必须容忍 BOM。

    Windows PowerShell 5.1 的 `Set-Content -Encoding utf8` 与
    `ConvertTo-Json | Out-File` 都会写 BOM，而 json.load(encoding="utf-8")
    遇到 BOM 直接抛 "Unexpected UTF-8 BOM"。所以这里一律用 utf-8-sig。
    """
    with open(path, encoding="utf-8-sig") as f:
        return json.load(f)


def cmds_fill(a):
    spec = _load_json(a.spec)
    prs = Presentation(spec.get("template") or a.template)
    S = list(prs.slides)
    if spec.get("keep"):
        order = [int(x) for x in spec["keep"]]
    else:
        order = list(range(1, len(S) + 1))

    drops = spec.get("drop") or {}
    places = spec.get("place") or {}
    anchors = spec.get("anchor") or {}
    fills = spec.get("fill") or {}
    charts = spec.get("chart") or {}
    tables = spec.get("table") or {}
    headers = spec.get("header") or {}
    titles = spec.get("widen_title") or {}

    for key in order:
        sl = S[key - 1]
        sl._tnu_no = key
        if str(key) in drops:
            for sid in drops[str(key)]:
                try:
                    drop(sl, sid)
                except KeyError:
                    pass
        for sid, g in (places.get(str(key)) or {}).items():
            place(sl, sid, **g)
        for sid, how in (anchors.get(str(key)) or {}).items():
            anchor(sl, sid, how)
        if str(key) in headers:
            h = headers[str(key)]
            header(sl, h[0], h[1], str(h[2]) if len(h) > 2 else "")
        if str(key) in fills:
            print("P%02d (%s)" % (key, sl.slide_layout.name))
            fill(sl, fills[str(key)])
        for sid, c in (charts.get(str(key)) or {}).items():
            set_chart(sl, sid, c.get("categories"), c.get("series"),
                      c.get("title"), c.get("title_size", 16))
        for sid, t in (tables.get(str(key)) or {}).items():
            set_table(sl, sid, t["rows"], t.get("size", 13))
        if str(key) in titles:
            t = titles[str(key)]
            widen_title(sl, t["sid"], t.get("left", 1.665),
                        t.get("width", 10.0), t.get("size"), t.get("text"))

    # 裁剪 + 排序（未保留页必须 drop_rel，否则包内残留 slide 关系）
    lst = prs.slides._sldIdLst
    ids_all = list(lst)
    keep = [ids_all[i - 1] for i in order]
    for el in ids_all:
        if el not in keep:
            prs.part.drop_rel(el.get(qn("r:id")))
            lst.remove(el)
    for el in keep:
        lst.remove(el)
        lst.append(el)

    # 备注（可选）
    notes = spec.get("notes") or {}
    if notes:
        for i, sl in enumerate(prs.slides, 1):
            t = notes.get(str(i))
            if t:
                sl.notes_slide.notes_text_frame.text = t

    out = a.out or spec["output"]
    prs.save(out)
    print("saved:", out)
    print("slides:", len(prs.slides._sldIdLst))
    return out


# --------------------------------------------------------------- verify
def _rec_ids(shapes):
    s = set()
    for x in iter_shapes(shapes):
        s.add(x.shape_id)
    return s


CJK = lambda c: 0x2E80 <= ord(c) <= 0x9FFF or 0xFF00 <= ord(c) <= 0xFFEF

# 模板自带的占位/示例文案。成品里如果还留着这些，说明那一格忘了填，
# 而它看起来完全正常——是这套工作流里最贵的一类错误。
RESIDUE = (
    "请在此输入标题", "在此输入您的标题", "在此输入标题", "请在此处输入标题",
    "请输入您的标题", "请您输入标题", "输入标题",
    "请输入文字内容", "请您输入文字", "请输入文字内容请输入文字内容",
    "Click here to add your t", "ENTER ENGLISH", "Please enter a title",
    "请输入标题", "请输入内容", "输入文字", "可输入描述内容",
    "类别1", "类别2", "类别3", "类别4", "类别5",
    "请在此处输入标题", "学校现有27个党政职能部门", "天津师范大学是天津市属重点院校",
    "请输入您的标题", "请在此输入", "NO.1 ", "NO.2 ", "NO.3 ",
)


def _text_est(shape):
    """估算文字块的高度（英寸）。CJK em≈1.0，拉丁≈0.52，行距按 1.25。

    ★ 必须相对模板同一形状做校准。模板自带的东西本来就会「越界/贴边」：
      内容页右上装饰图 `图片 13` 的几何是 (11.34, 5.97) 2.15×1.23，右缘到
      13.49 英寸，超出 13.33 的画布；时间线页标题框 height=0（靠自动调整）；
      直观数据页 48pt 数字框只有 1.22×0.91。这些都不是错误。
    """
    g = _geom(shape)
    if not g or not getattr(shape, "has_text_frame", False):
        return 0.0
    x, y, w, h = g
    txt = shape.text_frame.text
    if not txt.strip():
        return 0.0
    szs = _sizes(shape)
    pt = max(szs) if szs else 18.0
    em_w = pt / 72.0
    per_line = max(1.0, w / em_w)
    lines = 0.0
    for para in txt.split("\n"):
        units = sum(1.0 if CJK(c) else 0.52 for c in para)
        lines += max(1.0, -(-units // per_line))
    return lines * em_w * 1.25


def _overflow(shape, canvas_w, canvas_h, base=None):
    """base ＝ 模板中同 shape_id 的形状。给了 base 就只报「比模板更糟」的项。"""
    out = []
    g = _geom(shape)
    if not g:
        return out
    x, y, w, h = g
    bg = _geom(base) if base is not None else None
    moved = bg is not None and tuple(bg) != tuple(g)
    outside = (x < -0.05 or y < -0.05
               or x + w > canvas_w + 0.05 or y + h > canvas_h + 0.05)
    if outside and (bg is None or moved):
        out.append("OUT-OF-CANVAS sid=%s %s" % (shape.shape_id, g))

    est = _text_est(shape)
    if est <= 0:
        return out
    base_est = _text_est(base) if base is not None else 0.0
    limit = max(h * 1.08 + 0.02, base_est * 1.12)
    if est > limit:
        out.append("OVERFLOW est %.2f（框 %.2f，模板基准 %.2f）sid=%s %r"
                   % (est, h, base_est, shape.shape_id,
                      shape.text_frame.text.strip()[:26]))
    return out


def _residue(shape, base=None):
    """模板示例文案残留检测。

    只在「这段文字与模板同 shape_id 的原件**完全一样**」时才报，
    因为改过字的框是有意的，没改的框才是漏填。单纯关键词匹配会在
    作者特意保留某句模板文案时误报。
    """
    hits = []
    try:
        if not shape.has_text_frame:
            return hits
        txt = shape.text_frame.text.strip()
    except Exception:
        return hits
    if not txt:
        return hits
    if base is not None:
        try:
            if base.text_frame.text.strip() != txt:
                return hits
        except Exception:
            pass
    for pat in RESIDUE:
        if pat in txt:
            hits.append(txt[:32])
            break
    return hits


def _residue_table(slide):
    """表格每个格子单独查——残留常常只在一个 cell 里。"""
    hits = []
    for s in iter_shapes(slide.shapes):
        try:
            if not s.has_table:
                continue
        except Exception:
            continue
        for row in s.table.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                if txt and any(p in txt for p in RESIDUE):
                    hits.append("sid=%s %r" % (s.shape_id, txt[:20]))
    return hits


def cmds_verify(a):
    tpl = Presentation(a.template)
    built = Presentation(a.built)
    T = list(tpl.slides)
    B = list(built.slides)
    tp = {i + 1: s for i, s in enumerate(T)}

    report = {"template_slides": len(T), "built_slides": len(B),
              "pages": [], "problems": []}
    # 顺序对齐：spec 的 keep 顺序在成品里看不出来，只能靠「形状 id 集合是模板
    # 某页的子集」反查来源页。同 id 集合法会同时命中多页时取候选里最靠前的，
    # 并把这个歧义记下来（TNU 模板大量页共用同一套 shape_id）。
    used = set()
    for k, bs in enumerate(B, 1):
        bid = _rec_ids(bs.shapes)
        cands = [t for t in tp if t not in used
                 and bid and bid <= _rec_ids(tp[t].shapes)]
        if not cands:
            report["problems"].append(
                "P%d 找不到对应的模板页——说明新增了模板里没有的形状"
                "（这正是本技能要禁止的事）" % k)
            report["pages"].append({"built": k, "template": None,
                                    "added": [], "removed": []})
            continue
        # 候选可能不止一页：模板大量页共用同一套 shape_id（如 20 号在内容是
        # 组合、在表格页是表格）。挑「形状集合最接近」的那页，否则后续的
        # 基准对齐会错位，掩盖真正的越界。
        cands.sort(key=lambda t: len(_rec_ids(tp[t].shapes)))
        src = cands[0]
        used.add(src)
        base = {s.shape_id: s for s in iter_shapes(tp[src].shapes)}
        added = bid - set(base)
        removed = set(base) - bid
        report["pages"].append({"built": k, "template": src,
                                "ambiguous_sources": [t for t in cands[1:]],
                                "added": sorted(added),
                                "removed": sorted(removed)})
        if added:
            report["problems"].append(
                "P%d（模板P%d）新增形状 %d 个：%s"
                % (k, src, len(added), sorted(added)[:8]))
        for s in iter_shapes(bs.shapes):
            report["problems"] += ["P%d %s" % (k, m) for m in
                                   _overflow(s, built.slide_width / EMU,
                                             built.slide_height / EMU,
                                             base.get(s.shape_id))]
            for r in _residue(s, base.get(s.shape_id)):
                report["problems"].append("P%d 模板原文残留：%r（sid=%s）"
                                          % (k, r, s.shape_id))
        for r in _residue_table(bs):
            report["problems"].append("P%d 表格残留：%r" % (k, r))

    report["added_total"] = sum(len(p["added"]) for p in report["pages"])
    report["removed_total"] = sum(len(p["removed"]) for p in report["pages"])
    report["ok"] = not report["problems"]

    if a.json:
        print(json.dumps(report, ensure_ascii=False, indent=1))
    else:
        print("模板 %d 页 → 成品 %d 页" % (report["template_slides"], len(B)))
        print("新增形状合计：%d　删除形状合计：%d"
              % (report["added_total"], report["removed_total"]))
        for p in report["pages"]:
            flag = "✗" if p["added"] else "✓"
            print("  %s P%02d ← 模板P%02d　新增 %d　删 %d %s"
                  % (flag, p["built"], p["template"], len(p["added"]),
                     len(p["removed"]),
                     ("删:" + str(p["removed"][:10])) if p["removed"] else ""))
        print()
        if report["problems"]:
            print("可疑项 %d 条（启发式，先渲染核对再改）：" % len(report["problems"]))
            for m in report["problems"]:
                print("  " + m)
        else:
            print("未发现越界/溢出/新增形状。")
    return 0 if report["ok"] else 1


# --------------------------------------------------------------- render
RENDER_PS1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "render.ps1")


def cmds_render(a):
    os.makedirs(a.outdir, exist_ok=True)
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
           "-File", RENDER_PS1, "-Src", a.pptx, "-Out", a.outdir,
           "-W", str(a.w), "-H", str(a.h)]
    print(" ".join(cmd))
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       errors="replace")
    print(r.stdout.strip())
    if r.returncode != 0:
        print(r.stderr.strip()[:2000])
    return r.returncode


# --------------------------------------------------------------- main
def main(argv=None):
    ap = argparse.ArgumentParser(prog="tnu.py", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("inventory"); p.add_argument("pptx")
    p.add_argument("--out"); p.add_argument("--pages")
    p.set_defaults(func=cmds_inventory)

    p = sub.add_parser("catalog"); p.add_argument("pptx")
    p.add_argument("--out"); p.set_defaults(func=cmds_catalog)

    p = sub.add_parser("fill"); p.add_argument("spec")
    p.add_argument("--out"); p.add_argument("--template")
    p.set_defaults(func=cmds_fill)

    p = sub.add_parser("verify")
    p.add_argument("template"); p.add_argument("built")
    p.add_argument("--json", action="store_true"); p.set_defaults(func=cmds_verify)

    p = sub.add_parser("render")
    p.add_argument("pptx"); p.add_argument("outdir")
    p.add_argument("--w", type=int, default=1600)
    p.add_argument("--h", type=int, default=900)
    p.set_defaults(func=cmds_render)

    a = ap.parse_args(argv)
    # 注意：sys.exit(字符串) 会把它打到 stderr 并以退出码 1 结束。
    # 而 cmds_fill 要返回产物路径给调用方，所以这里必须强制收敛成 int。
    rc = a.func(a)
    return rc if isinstance(rc, int) else 0


if __name__ == "__main__":
    sys.exit(main())
