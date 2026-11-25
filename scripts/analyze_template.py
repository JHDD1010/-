"""템플릿 .pptx 의 슬라이드 크기, 마스터/레이아웃/placeholder, 테마 색·폰트, 텍스트 스타일,
실제 슬라이드의 폰트·크기·색 사용 빈도를 JSON 으로 출력한다.

    python scripts/analyze_template.py templates/본사.pptx > analysis.json
"""
import sys, json, collections, zipfile, re
from pptx import Presentation
from pptx.util import Emu
from lxml import etree

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

def cm(v): return round(Emu(v).cm, 2) if v is not None else None

def theme_info(master):
    for rel in master.part.rels.values():
        if rel.reltype.endswith("/theme"):
            x = etree.fromstring(rel.target_part.blob)
            name = x.get("name")
            cs = x.find(".//a:clrScheme", NS)
            colors = {}
            for c in cs:
                tag = etree.QName(c).localname
                v = c[0].get("val") or c[0].get("lastClr")
                colors[tag] = v
            fs = x.find(".//a:fontScheme", NS)
            fonts = {}
            for kind in ("majorFont", "minorFont"):
                f = fs.find(f"a:{kind}", NS)
                fonts[kind] = {"latin": f.find("a:latin", NS).get("typeface"),
                               "ea": f.find("a:ea", NS).get("typeface"),
                               "Hang": next((s.get("typeface") for s in f.findall("a:font", NS) if s.get("script") == "Hang"), None)}
            return {"theme": name, "colorScheme": cs.get("name"), "colors": colors, "fontScheme": fs.get("name"), "fonts": fonts}

def lvl_styles(el):
    out = []
    if el is None: return out
    for i in range(1, 10):
        l = el.find(f"a:lvl{i}pPr", NS)
        if l is None: continue
        d = {"lvl": i}
        if l.get("marL"): d["marL_cm"] = cm(int(l.get("marL")))
        if l.get("indent"): d["indent_cm"] = cm(int(l.get("indent")))
        r = l.find("a:defRPr", NS)
        if r is not None:
            if r.get("sz"): d["sz"] = int(r.get("sz")) / 100
            if r.get("b"): d["b"] = r.get("b")
            sc = r.find(".//a:solidFill/*", NS)
            if sc is not None: d["color"] = sc.get("val")
            lat = r.find("a:latin", NS)
            if lat is not None: d["latin"] = lat.get("typeface")
        bu = l.find("a:buChar", NS)
        if bu is not None: d["bullet"] = bu.get("char")
        if l.find("a:buNone", NS) is not None: d["bullet"] = None
        bc = l.find("a:buClr/*", NS)
        if bc is not None: d["buColor"] = bc.get("val")
        sb = l.find("a:spcBef/a:spcPts", NS)
        if sb is not None: d["spcBef"] = int(sb.get("val")) / 100
        out.append(d)
        if i >= 5: break
    return out

def ph_desc(ph):
    pf = ph.placeholder_format
    try: t = pf.type.name if pf.type is not None else "None"
    except Exception: t = str(pf.type)
    txt = ph.text_frame.text[:60].replace("\n", " / ") if ph.has_text_frame else ""
    return {"idx": pf.idx, "type": t, "name": ph.name,
            "box_cm": [cm(ph.left), cm(ph.top), cm(ph.width), cm(ph.height)], "prompt": txt}

def nonph_shapes(shapes):
    out = []
    for s in shapes:
        if s.is_placeholder: continue
        d = {"kind": str(s.shape_type).split(".")[-1].split(" ")[0], "name": s.name,
             "box_cm": [cm(s.left), cm(s.top), cm(s.width), cm(s.height)]}
        if s.has_text_frame and s.text_frame.text.strip():
            d["text"] = s.text_frame.text[:60]
        out.append(d)
    return out

def run_stats(prs):
    fonts, sizes, colors, kinds = (collections.Counter() for _ in range(4))
    for s in prs.slides:
        for el in s._element.iter():
            q = etree.QName(el).localname
            if q in ("rPr", "defRPr", "endParaRPr"):
                if el.get("sz"): sizes[int(el.get("sz")) / 100] += 1
                for f in ("latin", "ea"):
                    x = el.find(f"a:{f}", NS)
                    if x is not None: fonts[f"{f}:{x.get('typeface')}"] += 1
                c = el.find("a:solidFill/*", NS)
                if c is not None: colors[c.get("val")] += 1
        for sh in s.shapes:
            kinds[str(sh.shape_type).split(".")[-1].split(" ")[0]] += 1
    return fonts, sizes, colors, kinds

def main(path):
    prs = Presentation(path)
    R = {"file": path, "slide_size_cm": [cm(prs.slide_width), cm(prs.slide_height)], "slides": len(prs.slides)}
    z = zipfile.ZipFile(path)
    progids = collections.Counter(re.findall(rb'progId="([^"]+)"', b"".join(z.read(n) for n in z.namelist() if n.startswith("ppt/slide") and n.endswith(".xml"))))
    R["ole_progids"] = {k.decode(): v for k, v in progids.items()}
    R["embedded_fonts"] = [n for n in z.namelist() if "fonts/" in n]
    tags = [n for n in z.namelist() if n.startswith("ppt/tags/")]
    R["tags_parts"] = len(tags)
    R["custom_xml"] = [n for n in z.namelist() if n.startswith("customXml/") and "item" in n and n.endswith(".xml")][:5]
    R["masters"] = []
    used = collections.Counter(s.slide_layout.name for s in prs.slides)
    for m in prs.slide_masters:
        tx = m._element.find("p:txStyles", NS)
        md = {"name": m.name if hasattr(m, "name") else None, "theme": theme_info(m),
              "placeholders": [ph_desc(p) for p in m.placeholders],
              "decorations": nonph_shapes(m.shapes),
              "titleStyle": lvl_styles(tx.find("p:titleStyle", NS))[:1],
              "bodyStyle": lvl_styles(tx.find("p:bodyStyle", NS)),
              "layouts": []}
        md["name"] = m._element.find("p:cSld", NS).get("name")
        for l in m.slide_layouts:
            md["layouts"].append({"name": l.name, "used_by_slides": used.get(l.name, 0),
                                  "placeholders": [ph_desc(p) for p in l.placeholders],
                                  "decorations": nonph_shapes(l.shapes)})
        R["masters"].append(md)
    fonts, sizes, colors, kinds = run_stats(prs)
    R["slide_font_usage"] = fonts.most_common(15)
    R["slide_size_usage"] = sorted(sizes.most_common(15))
    R["slide_color_usage"] = colors.most_common(15)
    R["slide_shape_kinds"] = kinds.most_common()
    R["slide_list"] = []
    for i, s in enumerate(prs.slides, 1):
        title = s.shapes.title.text_frame.text[:80].replace("\n", " ") if s.shapes.title is not None and s.shapes.title.has_text_frame else ""
        R["slide_list"].append([i, s.slide_layout.name, title])
    return R

if __name__ == "__main__":
    print(json.dumps(main(sys.argv[1]), ensure_ascii=False, indent=1))
