"""Build PowerPoint templates (.potx) matching the unibucbeamer Beamer design.

Usage (from the repo root):
    pip install python-pptx
    python PowerPoint/build_potx.py

Writes PowerPoint/UB-template-en.potx and PowerPoint/UB-template-ro.potx.
The slide master and layouts carry the design (bars, colours, fonts, logo,
title-page artwork); the few example slides only show how to use it.
"""

import copy
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.slide import SlideLayout
from pptx.util import Emu, Inches, Pt
from pptx.dml.color import RGBColor

HERE = Path(__file__).resolve().parent
RES = HERE.parent / "inst" / "resources"
LOGO, BACKGROUND = RES / "logo.png", RES / "background.png"

# Colours (as in presentation_setup.tex) -----------------------------------
NAVY, TEAL, BLOCK = "022169", "00AECA", "3333B2"
FONT = "Source Sans 3"  # free (SIL OFL): https://github.com/adobe-fonts/source-sans

# Geometry (16:9, proportions measured on the Beamer PDF) -------------------
W, H = 12192000, 6858000
TOP_H = int(H * 0.042)          # top bar
FOOT_H = int(H * 0.036)         # footline
FOOT_Y = H - FOOT_H
SPLIT_TOP = W // 2              # navy | teal in the top bar
SPLIT_FOOT = W * 2 // 3         # navy | teal in the footline
MX = Inches(0.35)               # left/right margin

NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')

TEXT = {
    "en": dict(
        title="Presentation title", subtitle="Subtitle", author="Author name",
        date="dd.mm.yyyy", footer="Author name  ·  Presentation title",
        p_author="Author name", p_date="Date", p_subtitle="Subtitle",
        section="Using the template",
        section_sub="The UB layouts are under Home › New Slide.",
        s1_title="Slides and blocks",
        s1_body=["Layouts are chosen under Home › New Slide.",
                 "The footer is set under Insert › Header & Footer › Apply to All.",
                 "The UB colours appear in every colour menu."],
        s1_sub="Sub-headings",
        s1_sub_text="Navy, bold, no box.",
        block_title="Blocks", block_body="This box can be copied for highlighted content.",
        s2_title="Lists", s2_left=["First item", "Second item", ("Nested item", 1), "Third item"],
        s2_right=["First step", "Second step", "Third step"],
        s3_title="Tables", table=[["Name", "Age"], ["Mihai", "25"], ["Adriana", "30"]],
        table_number="Table 1", table_title="Age of the Example Participants",
        table_note=("Note.", " This is an APA-style table."),
    ),
    "ro": dict(
        title="Titlul prezentării", subtitle="Subtitlu", author="Numele autorului",
        date="zz.ll.aaaa", footer="Numele autorului  ·  Titlul prezentării",
        p_author="Numele autorului", p_date="Data", p_subtitle="Subtitlu",
        section="Utilizarea șablonului",
        section_sub="Layout-urile UB se găsesc în Home › New Slide.",
        s1_title="Slide-uri și blocuri",
        s1_body=["Layout-ul se alege din Home › New Slide.",
                 "Subsolul se completează din Insert › Header & Footer › Apply to All.",
                 "Culorile UB apar în toate meniurile de culori."],
        s1_sub="Subtitluri",
        s1_sub_text="Bleumarin, îngroșat, fără chenar.",
        block_title="Blocuri", block_body="Caseta se poate copia pentru conținut evidențiat.",
        s2_title="Liste", s2_left=["Primul element", "Al doilea element", ("Element imbricat", 1),
                                   "Al treilea element"],
        s2_right=["Primul pas", "Al doilea pas", "Al treilea pas"],
        s3_title="Tabele", table=[["Nume", "Vârstă"], ["Mihai", "25"], ["Adriana", "30"]],
        table_number="Tabelul 1", table_title="Vârsta participanților din exemplu",
        table_note=("Notă.", " Acesta este un tabel în stil APA."),
    ),
}


# XML helpers -------------------------------------------------------------
class Ids:
    def __init__(self):
        self.n = 1

    def __call__(self):
        self.n += 1
        return self.n


def xfrm(x, y, cx, cy):
    return (f'<a:xfrm><a:off x="{int(x)}" y="{int(y)}"/>'
            f'<a:ext cx="{int(cx)}" cy="{int(cy)}"/></a:xfrm>')


def rect(ids, name, x, y, cx, cy, clr):
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{ids()}" name="{name}"/><p:cNvSpPr/>
<p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr>{xfrm(x, y, cx, cy)}
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill>{clr}</a:solidFill>
<a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>
<a:p><a:endParaRPr lang="en-GB"/></a:p></p:txBody></p:sp>'''


def pic(ids, name, rid, x, y, cx, cy):
    return f'''<p:pic><p:nvPicPr><p:cNvPr id="{ids()}" name="{name}" descr="{name}"/>
<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr userDrawn="1"/></p:nvPicPr>
<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
<p:spPr>{xfrm(x, y, cx, cy)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'''


def ph(ids, name, ph_attrs, pos=None, body_pr="", lst="", text="", geom=False, para=None):
    """Placeholder. ph_attrs e.g. 'type="title"'; text = prompt text."""
    sppr = (xfrm(*pos) if pos else "") + (
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' if geom else "")
    para = para or (f'<a:p><a:r><a:rPr lang="en-GB"/><a:t>{text}</a:t></a:r></a:p>' if text
                    else '<a:p><a:endParaRPr lang="en-GB"/></a:p>')
    return f'''<p:sp><p:nvSpPr><p:cNvPr id="{ids()}" name="{name}"/>
<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph {ph_attrs}/></p:nvPr></p:nvSpPr>
<p:spPr>{sppr}</p:spPr><p:txBody><a:bodyPr {body_pr}/><a:lstStyle>{lst}</a:lstStyle>
{para}</p:txBody></p:sp>'''


def lvl1(algn="l", sz=None, clr=None, bold=False, nobullet=False):
    rpr = f' sz="{sz}"' if sz else ""
    rpr += ' b="1"' if bold else ""
    fill = f"<a:solidFill>{clr}</a:solidFill>" if clr else ""
    bu = '<a:buNone/>' if nobullet else ""
    mar = ' marL="0" indent="0"' if nobullet else ""
    return f'<a:lvl1pPr algn="{algn}"{mar}>{bu}<a:defRPr{rpr}>{fill}</a:defRPr></a:lvl1pPr>'


S_NAVY = '<a:schemeClr val="tx2"/>'     # dk2 = navy
S_TEAL = '<a:schemeClr val="accent2"/>'
S_WHITE = '<a:schemeClr val="bg1"/>'
S_BLACK = '<a:schemeClr val="tx1"/>'
S_GREY = '<a:srgbClr val="595959"/>'

INSET0 = 'lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"'


BLEED = Pt(2)  # bars run past the slide edge so no white hairline shows


def bars(ids):
    B = BLEED
    return (rect(ids, "Top bar navy", -B, -B, SPLIT_TOP + B, TOP_H + B, S_NAVY)
            + rect(ids, "Top bar teal", SPLIT_TOP, -B, W - SPLIT_TOP + B, TOP_H + B, S_TEAL)
            + rect(ids, "Footline navy", -B, FOOT_Y, SPLIT_FOOT + B, FOOT_H + B, S_NAVY)
            + rect(ids, "Footline teal", SPLIT_FOOT, FOOT_Y, W - SPLIT_FOOT + B, FOOT_H + B, S_TEAL))


SLIDENUM = ('<a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum">'
            '<a:rPr lang="en-GB"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-GB"/></a:p>')


def footers(ids, t, with_pos=False, lst=True):
    """Date / footer / slide-number placeholders. The default texts are stored
    here too: PowerPoint copies them from the layout onto every new slide."""
    third = (W - SPLIT_FOOT) // 2
    white = lambda: lvl1("ctr", 1100, S_WHITE) if lst else ""
    p = (lambda *a: a) if with_pos else (lambda *a: None)
    return (ph(ids, "Date", 'type="dt" sz="half" idx="10"', p(SPLIT_FOOT, FOOT_Y, third, FOOT_H),
               INSET0, white(), t["date"], geom=with_pos)
            + ph(ids, "Footer", 'type="ftr" sz="quarter" idx="11"', p(0, FOOT_Y, SPLIT_FOOT, FOOT_H),
                 INSET0, white(), t["footer"], geom=with_pos)
            + ph(ids, "Slide number", 'type="sldNum" sz="quarter" idx="12"',
                 p(SPLIT_FOOT + third, FOOT_Y, third, FOOT_H), INSET0, white(), geom=with_pos, para=SLIDENUM))


# Content area
TITLE_POS = (MX, TOP_H + Inches(0.12), W - 2 * MX, Inches(0.85))
BODY_Y = TOP_H + Inches(1.15)
BODY_POS = (MX, BODY_Y, W - 2 * MX, FOOT_Y - Inches(0.2) - BODY_Y)


# Theme ---------------------------------------------------------------------
THEME = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Universitatea din București">
<a:themeElements>
<a:clrScheme name="UniBuc">
<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="{NAVY}"/></a:dk2><a:lt2><a:srgbClr val="E8ECF3"/></a:lt2>
<a:accent1><a:srgbClr val="{NAVY}"/></a:accent1><a:accent2><a:srgbClr val="{TEAL}"/></a:accent2>
<a:accent3><a:srgbClr val="{BLOCK}"/></a:accent3><a:accent4><a:srgbClr val="7A8DB8"/></a:accent4>
<a:accent5><a:srgbClr val="80D6E4"/></a:accent5><a:accent6><a:srgbClr val="7F7F7F"/></a:accent6>
<a:hlink><a:srgbClr val="{TEAL}"/></a:hlink><a:folHlink><a:srgbClr val="{BLOCK}"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="UniBuc">
<a:majorFont><a:latin typeface="{FONT}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="{FONT}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
<a:ln w="28575"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
<a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'''


def bullet_lvl(n, mar, ind, sz, char="●", pct=70000):
    return (f'<a:lvl{n}pPr marL="{mar}" indent="{-ind}" algn="l" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
            f'<a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef>'
            f'<a:buClr>{S_NAVY}</a:buClr><a:buSzPct val="{pct}"/><a:buFont typeface="Arial"/>'
            f'<a:buChar char="{char}"/><a:defRPr sz="{sz}" kern="1200">'
            f'<a:solidFill>{S_BLACK}</a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl{n}pPr>')


def master_xml(layout_id_lst, t):
    ids = Ids()
    title_lst = ""
    tree = (bars(ids)
            + ph(ids, "Title", 'type="title"', TITLE_POS, 'lIns="0" anchor="ctr"', title_lst,
                 "Click to edit title", geom=True)
            + ph(ids, "Text", 'type="body" idx="1"', BODY_POS, 'lIns="0"', "",
                 "Click to edit text", geom=True)
            + footers(ids, t, with_pos=True))
    body_levels = "".join(bullet_lvl(n, m, i, s, c) for n, m, i, s, c in [
        (1, 285750, 285750, 2200, "●"), (2, 685800, 285750, 2000, "●"),
        (3, 1085850, 285750, 1800, "●"), (4, 1485900, 285750, 1600, "–"),
        (5, 1885950, 285750, 1600, "–")])
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster {NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
{tree}</p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3"
 accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
{layout_id_lst}
<p:hf hdr="0"/>
<p:txStyles>
<p:titleStyle><a:lvl1pPr algn="l" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>
<a:defRPr sz="2800" kern="1200"><a:solidFill>{S_NAVY}</a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr>
</a:lvl1pPr></p:titleStyle>
<p:bodyStyle>{body_levels}</p:bodyStyle>
<p:otherStyle><a:defPPr><a:defRPr lang="en-GB"/></a:defPPr>
<a:lvl1pPr marL="0" algn="l" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
<a:defRPr sz="1800" kern="1200"><a:solidFill>{S_BLACK}</a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr>
</p:otherStyle></p:txStyles></p:sldMaster>'''.encode("utf-8")


def layout_xml(name, tree, type_attr="", show_master=True):
    smsp = "" if show_master else ' showMasterSp="0"'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout {NS}{(" " + type_attr) if type_attr else ""} preserve="1"{smsp}>
<p:cSld name="{name}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
{tree}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'''.encode("utf-8")


def build_layouts(parts, t):
    """parts: dict name -> SlideLayoutPart. Returns nothing; rewrites XML."""
    # Title slide: no bars, artwork pinned bottom-right at full height,
    # logo centred at the top (as in the Beamer title page).
    p = parts["title"]
    _, rid_bg = p.get_or_add_image_part(str(BACKGROUND))
    _, rid_logo = p.get_or_add_image_part(str(LOGO))
    bg_w = H * 3508 // 2481
    logo_w = Inches(2.5)
    logo_h = logo_w * 478 // 578
    ids = Ids()
    cw, cx = Inches(8), (W - Inches(8)) // 2
    nb = 'lIns="0" rIns="0" anchor="ctr"'
    tree = (pic(ids, "Title artwork", rid_bg, W - bg_w, 0, bg_w, H)
            + pic(ids, "UB logo", rid_logo, (W - logo_w) // 2, Inches(0.75), logo_w, logo_h)
            + ph(ids, "Title", 'type="ctrTitle"', (cx, Inches(3.03), cw, Inches(0.8)), nb,
                 lvl1("ctr", 3600, S_NAVY), "Click to edit title")
            + ph(ids, "Subtitle", 'type="subTitle" idx="1"', (cx, Inches(3.82), cw, Inches(0.55)), nb,
                 lvl1("ctr", 2400, S_NAVY, nobullet=True), t["p_subtitle"])
            + ph(ids, "Author", 'type="body" sz="quarter" idx="13" hasCustomPrompt="1"',
                 (cx, Inches(4.4), cw, Inches(0.55)), nb, lvl1("ctr", 2400, S_BLACK, nobullet=True),
                 t["p_author"])
            + ph(ids, "Date", 'type="body" sz="quarter" idx="14" hasCustomPrompt="1"',
                 (cx, Inches(4.98), cw, Inches(0.55)), nb, lvl1("ctr", 2400, S_BLACK, nobullet=True),
                 t["p_date"]))
    p._element = parse_xml(layout_xml("Title Slide", tree, 'type="title"', show_master=False))

    def std(name, type_attr, body):
        ids = Ids()
        tree = ph(ids, "Title", 'type="title"', text="Click to edit title") + body(ids) + footers(ids, t, lst=False)
        parts[name]._element = parse_xml(layout_xml(LAYOUT_NAMES[name], tree, type_attr))

    std("content", 'type="obj"', lambda ids: ph(ids, "Content", 'idx="1"', text="Click to edit text"))

    half = (W - 2 * MX - Inches(0.4)) // 2
    std("two", 'type="twoObj"', lambda ids:
        ph(ids, "Left", 'sz="half" idx="1"', (MX, BODY_POS[1], half, BODY_POS[3]), text="Click to edit text")
        + ph(ids, "Right", 'sz="half" idx="2"', (W - MX - half, BODY_POS[1], half, BODY_POS[3]),
             text="Click to edit text"))

    std("titleonly", 'type="titleOnly"', lambda ids: "")

    # Section header: big navy title, black subtitle underneath.
    ids = Ids()
    sx = Inches(0.9)
    tree = (ph(ids, "Title", 'type="title"', (sx, Inches(2.3), W - 2 * sx, Inches(1.4)),
               'lIns="0" anchor="b"', lvl1("l", 4000, S_NAVY), "Click to edit title")
            + ph(ids, "Text", 'type="body" idx="1"', (sx, Inches(3.8), W - 2 * sx, Inches(0.9)),
                 'lIns="0" anchor="t"', lvl1("l", 2000, S_BLACK, nobullet=True), "Click to edit text")
            + footers(ids, t, lst=False))
    parts["section"]._element = parse_xml(layout_xml("Section Header", tree, 'type="secHead"'))

    ids = Ids()
    parts["blank"]._element = parse_xml(layout_xml("Blank", footers(ids, t, lst=False), 'type="blank"'))


LAYOUT_NAMES = {"title": "Title Slide", "content": "Title and Content", "section": "Section Header",
                "two": "Two Content", "titleonly": "Title Only", "blank": "Blank"}
KEEP = {"Title Slide": "title", "Title and Content": "content", "Section Header": "section",
        "Two Content": "two", "Title Only": "titleonly", "Blank": "blank"}


# Example slides ---------------------------------------------------------------
def add_footer_phs(slide, layout_part, t):
    """python-pptx does not copy date/footer/number placeholders; add them."""
    texts = {"dt": t["date"], "ftr": t["footer"]}
    for sp in layout_part._element.cSld.spTree.iterchildren():
        phs = sp.xpath("./p:nvSpPr/p:nvPr/p:ph")
        if not phs or phs[0].get("type") not in ("dt", "ftr", "sldNum"):
            continue
        kind = phs[0].get("type")
        new = copy.deepcopy(sp)
        new.find(".//{*}spPr").clear()
        lst = new.find(".//{*}lstStyle")
        for c in list(lst):
            lst.remove(c)
        txb = new.find(".//{*}txBody")
        for para in txb.findall("{*}p"):
            txb.remove(para)
        if kind == "sldNum":
            para = parse_xml(f'<a:p {NS}><a:fld id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum">'
                             '<a:rPr lang="en-GB"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-GB"/></a:p>')
        else:
            para = parse_xml(f'<a:p {NS}><a:r><a:rPr lang="ro-RO"/><a:t>{texts[kind]}</a:t></a:r></a:p>')
        txb.append(para)
        slide.shapes._spTree.append(new)


def set_paras(tf, items, numbered=False):
    tf.clear()
    for k, item in enumerate(items):
        text, level = item if isinstance(item, tuple) else (item, 0)
        para = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        para.text, para.level = text, level
        if numbered:
            pPr = para._p.get_or_add_pPr()
            for tag in ("buChar", "buNone", "buAutoNum"):
                for el in pPr.findall("{*}" + tag):
                    pPr.remove(el)
            pPr.append(parse_xml(f'<a:buFont {NS} typeface="+mj-lt"/>'))
            pPr.append(parse_xml(f'<a:buAutoNum {NS} type="arabicPeriod"/>'))


def add_examples(prs, layouts, parts, t):
    def new(key):
        s = prs.slides.add_slide(layouts[key])
        if key != "title":
            add_footer_phs(s, parts[key], t)
        return s

    s = new("title")
    vals = {0: t["title"], 1: t["subtitle"], 13: t["author"], 14: t["date"]}
    for shape in s.placeholders:
        shape.text_frame.text = vals[shape.placeholder_format.idx]

    s = new("section")
    s.shapes.title.text = t["section"]
    s.placeholders[1].text_frame.text = t["section_sub"]

    s = new("content")
    s.shapes.title.text = t["s1_title"]
    body = s.placeholders[1]
    body.left, body.top, body.width = BODY_POS[0], BODY_POS[1], BODY_POS[2]
    body.height = Inches(2.9)
    tf = body.text_frame
    set_paras(tf, t["s1_body"])
    head = tf.add_paragraph()
    head.text = t["s1_sub"]
    pPr = head._p.get_or_add_pPr()
    pPr.set("marL", "0"); pPr.set("indent", "0")
    pPr.insert(0, parse_xml(f'<a:spcBef {NS}><a:spcPts val="1800"/></a:spcBef>'))
    pPr.append(parse_xml(f'<a:buNone {NS}/>'))
    head.runs[0].font.bold = True
    head.runs[0].font.color.rgb = RGBColor.from_string(NAVY)
    txt = tf.add_paragraph()
    txt.text = t["s1_sub_text"]
    p2 = txt._p.get_or_add_pPr(); p2.set("marL", "0"); p2.set("indent", "0")
    p2.append(parse_xml(f'<a:buNone {NS}/>'))
    add_block(s, MX, Inches(4.55), W - 2 * MX, Inches(1.25), t["block_title"], t["block_body"])

    s = new("two")
    s.shapes.title.text = t["s2_title"]
    set_paras(s.placeholders[1].text_frame, t["s2_left"])
    set_paras(s.placeholders[2].text_frame, t["s2_right"], numbered=True)

    s = new("titleonly")
    s.shapes.title.text = t["s3_title"]
    add_apa_table(s, t)


def textbox(slide, x, y, cx, cy, runs, size=18):
    """runs: list of (text, bold, italic) in one paragraph."""
    tb = slide.shapes.add_textbox(x, y, cx, cy)
    tf = tb.text_frame
    tf.word_wrap = True
    bp = tf._txBody.bodyPr
    for side in ("lIns", "rIns"):
        bp.set(side, "0")
    para = tf.paragraphs[0]
    for text, bold, italic in runs:
        r = para.add_run()
        r.text = text
        r.font.size, r.font.bold, r.font.italic = Pt(size), bold, italic
        r.font.color.rgb = RGBColor.from_string("000000")
    return tb


def add_apa_table(slide, t):
    """APA 7 table: bold number, italic title above; three horizontal rules
    (from the APA table style); italic 'Note.' below."""
    rows = t["table"]
    tw, row_h = Inches(6), Inches(0.42)
    x, y = (W - tw) // 2, Inches(1.55)
    textbox(slide, x, y, tw, Inches(0.4), [(t["table_number"], True, False)])
    textbox(slide, x, y + Inches(0.4), tw, Inches(0.4), [(t["table_title"], False, True)])
    ty = y + Inches(0.95)
    gt = slide.shapes.add_table(len(rows), len(rows[0]), x, ty, tw, row_h * len(rows))
    tbl = gt.table
    tbl.first_row, tbl.horz_banding = True, False
    tbl._tbl.tblPr.find("{*}tableStyleId").text = APA_STYLE_ID
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = val
            para = cell.text_frame.paragraphs[0]
            para.alignment = 1 if c == 0 else 2   # stub left, other columns centred
            para.runs[0].font.size = Pt(18)
    note, rest = t["table_note"]
    textbox(slide, x, ty + row_h * len(rows) + Inches(0.12), tw, Inches(0.7),
            [(note, False, True), (rest, False, False)], size=14)


def add_block(slide, x, y, cx, cy, title, body):
    """Beamer-style block: white box with a soft shadow, coloured title."""
    shp = slide.shapes.add_shape(1, x, y, cx, cy)  # rectangle
    shp.name = "Block"
    sp = shp._element
    spPr = sp.spPr
    for tag in ("solidFill", "ln"):
        for el in spPr.findall("{*}" + tag):
            spPr.remove(el)
    spPr.append(parse_xml(f'<a:solidFill {NS}><a:schemeClr val="bg1"/></a:solidFill>'))
    spPr.append(parse_xml(f'<a:ln {NS}><a:noFill/></a:ln>'))
    spPr.append(parse_xml(
        f'<a:effectLst {NS}><a:outerShdw blurRad="76200" dist="38100" dir="2700000" algn="tl" rotWithShape="0">'
        '<a:srgbClr val="000000"><a:alpha val="45000"/></a:srgbClr></a:outerShdw></a:effectLst>'))
    style = sp.find("{*}style")
    if style is not None:
        sp.remove(style)
    tf = shp.text_frame
    tf.word_wrap = True
    bp = tf._txBody.bodyPr
    bp.set("anchor", "t"); bp.set("lIns", str(Inches(0.15))); bp.set("tIns", str(Inches(0.08)))
    tf.text = title
    p0 = tf.paragraphs[0]
    p0.runs[0].font.size = Pt(22)
    p0.runs[0].font.color.rgb = RGBColor.from_string(BLOCK)
    p1 = tf.add_paragraph()
    p1.text = body
    p1.runs[0].font.size = Pt(20)
    p1.runs[0].font.color.rgb = RGBColor.from_string("000000")


# APA 7 table style ---------------------------------------------------------
# Custom table style shipped in the template and set as the default, so every
# new table gets it: rules above and below the table and under the column
# headings; no vertical lines, no shading, no bold headings.
APA_STYLE_ID = "{6E4A1C2B-3D5F-4A7B-8C9D-0E1F2A3B4C5D}"
RULE = '<a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>'
NONE = '<a:ln><a:noFill/></a:ln>'
TABLE_STYLES = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{APA_STYLE_ID}">
<a:tblStyle styleId="{APA_STYLE_ID}" styleName="APA">
<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="tx1"/></a:tcTxStyle>
<a:tcStyle><a:tcBdr><a:left>{NONE}</a:left><a:right>{NONE}</a:right><a:top>{RULE}</a:top>
<a:bottom>{RULE}</a:bottom><a:insideH>{NONE}</a:insideH><a:insideV>{NONE}</a:insideV></a:tcBdr>
<a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl>
<a:firstRow><a:tcTxStyle b="off"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="tx1"/></a:tcTxStyle>
<a:tcStyle><a:tcBdr><a:top>{RULE}</a:top><a:bottom>{RULE}</a:bottom></a:tcBdr><a:fill><a:noFill/></a:fill></a:tcStyle></a:firstRow>
<a:lastRow><a:tcStyle><a:tcBdr><a:bottom>{RULE}</a:bottom></a:tcBdr><a:fill><a:noFill/></a:fill></a:tcStyle></a:lastRow>
</a:tblStyle></a:tblStyleLst>'''.encode("utf-8")


# Build -------------------------------------------------------------------------
TEMPLATE_CT = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"


def build(lang):
    t = TEXT[lang]
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(W), Emu(H)

    master = prs.slide_master
    for layout in list(master.slide_layouts):
        if layout.name not in KEEP:
            master.slide_layouts.remove(layout)
    parts = {KEEP[l.name]: l.part for l in master.slide_layouts}

    # Theme
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    theme_part = master.part.part_related_by(RT.THEME)
    theme_part._blob = THEME.encode("utf-8")
    prs.part.part_related_by(RT.TABLE_STYLES)._blob = TABLE_STYLES

    # Master (keep the layout id list python-pptx maintains)
    id_lst = etree.tostring(master._element.find("{*}sldLayoutIdLst")).decode()
    master.part._element = parse_xml(master_xml(id_lst, t))
    master.part.__dict__.pop("slide_master", None)

    build_layouts(parts, t)
    layouts = {k: SlideLayout(p._element, p) for k, p in parts.items()}

    add_examples(prs, layouts, parts, t)

    cp = prs.core_properties
    cp.title = "Universitatea din București – " + ("template" if lang == "en" else "șablon")
    cp.author = "George Gunnesch-Luca"
    cp.language = "en-GB" if lang == "en" else "ro-RO"

    out = HERE / f"UB-template-{lang}.potx"
    prs.save(out)
    mark_as_template(out)
    return out


def mark_as_template(path):
    """python-pptx always writes a .pptx; switch the main part to template."""
    import shutil, tempfile, zipfile
    tmp = Path(tempfile.mkstemp(suffix=".potx")[1])
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(b"presentationml.presentation.main+xml",
                                    TEMPLATE_CT.split("officedocument.")[1].encode())
            zout.writestr(item, data)
    shutil.move(tmp, path)
    path.chmod(0o644)


if __name__ == "__main__":
    for lang in ("en", "ro"):
        print("wrote", build(lang))
