import copy, io
from pptx import Presentation
from pptx.oxml.ns import qn
from lxml import etree

import sys

# usage: place_on_template.py <content.pptx> <rosatom_template.pptx> <out.pptx>
SRC, DST, OUT = sys.argv[1:4]
src = Presentation(SRC); dst = Presentation(DST)
k = dst.slide_width / src.slide_width
print('scale', k)
ss, ds = src.slides[0], dst.slides[0]
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
SCHEME = {'accent1': '4F81BD', 'lt1': 'FFFFFF'}
SHADOW = '<a:effectLst xmlns:a="%s"><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst>' % A
LEN_ATTR = {'lIns', 'tIns', 'rIns', 'bIns', 'marL', 'indent', 'blurRad', 'dist', 'rad'}
SP_ORDER = ['xfrm', 'custGeom', 'prstGeom', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'ln']

def sc(v): return str(int(round(int(v) * k)))

missing_latin = 0
for sh in ss.shapes:
    el = copy.deepcopy(sh._element)
    # images
    for blip in el.iter(qn('a:blip')):
        rid = blip.get(qn('r:embed'))
        part = ss.part.related_part(rid)
        new_rid = ds.part.get_or_add_image_part(io.BytesIO(part.blob))[1]
        blip.set(qn('r:embed'), new_rid)
    for e in el.iter():
        tag = etree.QName(e).localname if isinstance(e.tag, str) else ''
        if tag in ('off', 'chOff'):
            e.set('x', sc(e.get('x')))
            e.set('y', sc(e.get('y')))
        elif tag in ('ext', 'chExt') and e.get('cx') is not None:
            e.set('cx', sc(e.get('cx'))); e.set('cy', sc(e.get('cy')))
        elif tag == 'ln' and e.get('w'):
            e.set('w', sc(e.get('w')))
        elif tag in ('rPr', 'defRPr', 'endParaRPr') and e.get('sz'):
            e.set('sz', sc(e.get('sz')))
        elif tag in ('spcPts', 'buSzPts') and e.get('val'):
            e.set('val', sc(e.get('val')))
        elif tag == 'schemeClr' and e.get('val') in SCHEME:
            e.tag = qn('a:srgbClr'); e.set('val', SCHEME[e.get('val')])
        for a in LEN_ATTR:
            if e.get(a) is not None and tag != 'gd':
                e.set(a, sc(e.get(a)))
        if tag in ('rPr', 'endParaRPr') and e.get('sz') is None:
            e.set('sz', sc(1800))
        if tag == 'bodyPr':
            for a, d in (('lIns', 91440), ('rIns', 91440), ('tIns', 45720), ('bIns', 45720)):
                if e.get(a) is None:
                    e.set(a, sc(d))
        if tag in ('rPr', 'endParaRPr') and e.find(qn('a:latin')) is None:
            missing_latin += 1
    # explicit shadow where style supplied one
    spPr = el.find(qn('p:spPr'))
    if el.find(qn('p:style')) is not None and spPr is not None and spPr.find(qn('a:effectLst')) is None:
        eff = etree.fromstring(SHADOW)
        for c in eff.iter():
            for a in ('blurRad', 'dist'):
                if c.get(a): c.set(a, sc(c.get(a)))
        last = None
        for c in spPr:
            if etree.QName(c).localname in SP_ORDER: last = c
        (last.addnext(eff) if last is not None else spPr.insert(0, eff))
    ds.shapes._spTree.append(el)
print('rPr without latin:', missing_latin)
dst.save(OUT)
