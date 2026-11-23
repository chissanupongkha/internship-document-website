# -*- coding: utf-8 -*-
"""Edit a Word template inside the system and save it as a new version.

The editor shows the paragraphs of the .docx (and the text inside its tables). Staff change the
wording, bold / italic / underline, alignment and tags, then save. The system rebuilds the file from
the previous version, so pictures, headers, page setup, tables and fonts stay exactly as they were.
"""
import io
import json
import os
import re
from copy import deepcopy

from django.contrib import messages
from django.core.files.base import ContentFile
from django.http import JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.views.decorators.http import require_POST
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph

from .decorators import staff_required
from .docgen import AVAILABLE_SOURCE_FIELDS, TemplateFileMissing, extract_fields_from_docx, template_bytes
from .models import LetterTemplate, Logo, TemplateFieldMapping
from .views import log_activity

ALIGN_OUT = {'left': 'left', 'center': 'center', 'right': 'right', 'both': 'both', 'justify': 'both', 'distribute': 'both'}
TAG_RE = re.compile(r'\{\{\s*([A-Za-z_][\w.]*)\s*\}\}')


# ----------------------------------------------------------------------
# .docx  ->  editable items
# ----------------------------------------------------------------------
def _run_text(r):
    out = []
    for ch in r:
        t = ch.tag
        if t == qn('w:t'):
            out.append(ch.text or '')
        elif t == qn('w:tab'):
            out.append('\t')
        elif t in (qn('w:br'), qn('w:cr')):
            out.append('\n')
    return ''.join(out)


def _is_on(rpr, *tags):
    if rpr is None:
        return False
    for tag in tags:
        el = rpr.find(qn(tag))
        if el is not None and el.get(qn('w:val')) not in ('0', 'false', 'none'):
            return True
    return False


def _val(rpr, tag, attr='w:val'):
    el = rpr.find(qn(tag)) if rpr is not None else None
    return el.get(qn(attr)) if el is not None else None


def _run_fmt(r):
    """Direct formatting of one run: bold, italic, underline, size (pt), colour (hex), font name."""
    rpr = r.find(qn('w:rPr'))
    sz = _val(rpr, 'w:sz') or _val(rpr, 'w:szCs')
    col = _val(rpr, 'w:color')
    rf = rpr.find(qn('w:rFonts')) if rpr is not None else None
    font = None
    if rf is not None:
        font = rf.get(qn('w:ascii')) or rf.get(qn('w:cs')) or rf.get(qn('w:hAnsi')) or rf.get(qn('w:eastAsia'))
    return (_is_on(rpr, 'w:b', 'w:bCs'), _is_on(rpr, 'w:i', 'w:iCs'), _is_on(rpr, 'w:u'),
            int(sz) / 2 if sz and sz.isdigit() else None,
            col if col and col not in ('auto',) and re.fullmatch(r'[0-9A-Fa-f]{6}', col) else None,
            font)


def _para_runs(p_el):
    runs = []
    for r in p_el.iter(qn('w:r')):
        text = _run_text(r)
        if not text:
            continue
        fmt = _run_fmt(r)
        if runs and runs[-1][1:] == list(fmt):
            runs[-1][0] += text
        else:
            runs.append([text, *fmt])
    return [{'t': t, 'b': b, 'i': i, 'u': u, 'sz': sz, 'col': col, 'font': font}
            for t, b, i, u, sz, col, font in runs]


def _has_picture(p_el):
    return p_el.find('.//' + qn('w:drawing')) is not None or p_el.find('.//' + qn('w:pict')) is not None


def _has_page_break(p_el):
    return any(br.get(qn('w:type')) == 'page' for br in p_el.iter(qn('w:br')))


def _has_section(p_el):
    ppr = p_el.find(qn('w:pPr'))
    return ppr is not None and ppr.find(qn('w:sectPr')) is not None


def _is_pagebreak_only(p_el):
    return _has_page_break(p_el) and not _has_picture(p_el) and not _has_section(p_el) \
        and not ''.join(_run_text(r) for r in p_el.iter(qn('w:r'))).replace('\n', '').strip()


def _para_locked(p_el):
    """Pictures, section settings and page breaks that carry text are kept exactly as they are."""
    return _has_picture(p_el) or _has_section(p_el) or (_has_page_break(p_el) and not _is_pagebreak_only(p_el))


def _para_align(p_el, doc=None):
    """Alignment of a paragraph: its own setting, else its style chain (Word centres many logos
    through a style, not on the paragraph), else a floating picture's position, else left."""
    ppr = p_el.find(qn('w:pPr'))
    jc = ppr.find(qn('w:jc')) if ppr is not None else None
    if jc is not None:
        return ALIGN_OUT.get(jc.get(qn('w:val')), 'left')
    if doc is not None:
        try:
            st = Paragraph(p_el, doc._body).style
        except Exception:
            st = None
        while st is not None:
            sppr = st.element.find(qn('w:pPr'))
            sjc = sppr.find(qn('w:jc')) if sppr is not None else None
            if sjc is not None:
                return ALIGN_OUT.get(sjc.get(qn('w:val')), 'left')
            st = st.base_style
    wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    for anc in p_el.iter('{%s}anchor' % wp):          # floating picture placed "centred" on the page / margin
        ph = anc.find('{%s}positionH' % wp)
        al = ph.find('{%s}align' % wp) if ph is not None else None
        if al is not None and (al.text or '').strip() in ('left', 'center', 'right'):
            return al.text.strip()
    return 'left'


def _tw(v):                                    # twips -> pt
    return round(int(v) / 20, 1) if v not in (None, '') else None


def _para_props(p_el, doc):
    """Effective paragraph layout: direct values, else the paragraph style, else the document defaults."""
    para = Paragraph(p_el, doc._body)
    ppr = p_el.find(qn('w:pPr'))
    ind = ppr.find(qn('w:ind')) if ppr is not None else None
    first = left = None
    if ind is not None:
        fl, hg = ind.get(qn('w:firstLine')), ind.get(qn('w:hanging'))
        lf = ind.get(qn('w:left')) or ind.get(qn('w:start'))
        if fl:
            first = _tw(fl)
        elif hg:
            first = -_tw(hg)
        left = _tw(lf)
    before = after = line = None
    chain, st = [], None
    try:
        st = para.style
    except Exception:
        st = None
    sp = ppr.find(qn('w:spacing')) if ppr is not None else None
    def read_sp(sp_el):
        b, a, l, rule = (sp_el.get(qn('w:before')), sp_el.get(qn('w:after')),
                         sp_el.get(qn('w:line')), sp_el.get(qn('w:lineRule')))
        return b, a, l, rule
    sources = []
    if sp is not None:
        sources.append(read_sp(sp))
    while st is not None:
        sppr = st.element.find(qn('w:pPr'))
        if sppr is not None:
            ssp = sppr.find(qn('w:spacing'))
            if ssp is not None:
                sources.append(read_sp(ssp))
            sind = sppr.find(qn('w:ind'))
            if sind is not None:
                if first is None:
                    fl, hg = sind.get(qn('w:firstLine')), sind.get(qn('w:hanging'))
                    first = _tw(fl) if fl else (-_tw(hg) if hg else None)
                if left is None:
                    left = _tw(sind.get(qn('w:left')) or sind.get(qn('w:start')))
        st = st.base_style
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is not None:
        d = dd.find('.//' + qn('w:pPrDefault') + '/' + qn('w:pPr') + '/' + qn('w:spacing'))
        if d is not None:
            sources.append(read_sp(d))
    rule = None
    for b, a, l, r in sources:
        if before is None and b is not None:
            before = _tw(b)
        if after is None and a is not None:
            after = _tw(a)
        if line is None and l is not None:
            rule = r or 'auto'
            line = round(int(l) / 240, 2) if rule == 'auto' else None
    return {'align': _para_align(p_el, doc), 'first': first or 0, 'left': left or 0,
            'before': before or 0, 'after': after or 0, 'line': line if rule in (None, 'auto') else None}


def _defaults(doc):
    size = name = None
    try:
        st = doc.styles['Normal']
        size = st.font.size.pt if st.font.size else None
        name = st.font.name
    except Exception:
        pass
    dd = doc.styles.element.find(qn('w:docDefaults'))
    if dd is not None:
        rpr = dd.find('.//' + qn('w:rPrDefault') + '/' + qn('w:rPr'))
        if rpr is not None:
            if size is None:
                v = _val(rpr, 'w:sz') or _val(rpr, 'w:szCs')
                size = int(v) / 2 if v and v.isdigit() else None
            if name is None:
                rf = rpr.find(qn('w:rFonts'))
                if rf is not None:
                    name = rf.get(qn('w:ascii')) or rf.get(qn('w:cs')) or rf.get(qn('w:hAnsi'))
    return {'size': size or 14, 'font': name or 'TH Sarabun New'}


def _picture(p_el, doc):
    """The first picture of a paragraph as a data: URL, with its size in cm."""
    blip = p_el.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
    if blip is None:
        return None
    rid = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
    try:
        part = doc.part.related_parts[rid]
    except Exception:
        return None
    blob = part.blob
    if len(blob) > 3_000_000:
        return None
    import base64
    ext = p_el.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
    w = h = None
    if ext is not None:
        w, h = round(int(ext.get('cx')) / 360000, 2), round(int(ext.get('cy')) / 360000, 2)
    return {'src': f"data:{part.content_type};base64,{base64.b64encode(blob).decode()}", 'w': w, 'h': h}


def _table_info(tbl):
    grid = tbl.find(qn('w:tblGrid'))
    widths = [round(int(g.get(qn('w:w'), 0)) / 567, 2) for g in grid.findall(qn('w:gridCol'))] if grid is not None else []
    tblpr = tbl.find(qn('w:tblPr'))
    borders = False
    if tblpr is not None:
        tb = tblpr.find(qn('w:tblBorders'))
        if tb is not None:
            borders = any(ch.get(qn('w:val')) not in (None, 'nil', 'none') for ch in tb)
        st = _val(tblpr, 'w:tblStyle')
        if st and 'grid' in st.lower():
            borders = True
    return widths, borders


def _cell_item(p_el, doc):
    return {'props': _para_props(p_el, doc), 'runs': _para_runs(p_el), 'locked': _para_locked(p_el)}


def docx_to_items(data, pics=True):
    """-> (items, page) where page holds the sheet size, margins and default font for the editor."""
    doc = Document(io.BytesIO(data))
    sec = doc.sections[0]
    page = {'w': round(sec.page_width.cm, 2), 'h': round(sec.page_height.cm, 2),
            'ml': round(sec.left_margin.cm, 2), 'mr': round(sec.right_margin.cm, 2),
            'mt': round(sec.top_margin.cm, 2), 'mb': round(sec.bottom_margin.cm, 2), **_defaults(doc)}
    items, pi, ti = [], 0, 0
    for child in doc.element.body:
        if child.tag == qn('w:p'):
            if _is_pagebreak_only(child):
                items.append({'kind': 'pb', 'src': pi})
            elif _para_locked(child):
                pic = _picture(child, doc) if (pics and _has_picture(child)) else None
                items.append({'kind': 'p', 'src': pi, 'locked': True, 'props': _para_props(child, doc), 'runs': [],
                              'pic': pic, 'note': 'Picture (logo) kept as is' if _has_picture(child)
                              else ('Section settings kept as is' if _has_section(child) else 'Page break with text kept as is')})
            else:
                items.append({'kind': 'p', 'src': pi, 'locked': False, 'props': _para_props(child, doc),
                              'runs': _para_runs(child)})
            pi += 1
        elif child.tag == qn('w:tbl'):
            widths, borders = _table_info(child)
            rows = []
            for ri, tr in enumerate(child.findall(qn('w:tr'))):
                cells = []
                for tc in tr.findall(qn('w:tc')):
                    span = _val(tc.find(qn('w:tcPr')), 'w:gridSpan')
                    cells.append({'span': int(span) if span and span.isdigit() else 1,
                                  'paras': [_cell_item(p, doc) for p in tc.findall(qn('w:p'))]})
                rows.append({'src': ri, 'cells': cells})
            items.append({'kind': 'table', 'src': ti, 'rows': rows, 'widths': widths, 'borders': borders})
            ti += 1
    return items, page


# ----------------------------------------------------------------------
# edited items  ->  .docx (built from the previous file)
# ----------------------------------------------------------------------
STRIP_FROM_BASE = ('w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:u', 'w:sz', 'w:szCs', 'w:color', 'w:rFonts')


def _apply_props(p_el, doc, props):
    """Write only the layout values that differ from what the paragraph already has."""
    if not props:
        return
    para = Paragraph(p_el, doc._body)
    cur = _para_props(p_el, doc)
    pf = para.paragraph_format
    align = {'left': 0, 'center': 1, 'right': 2, 'both': 3}.get(ALIGN_OUT.get(props.get('align', cur['align']), 'left'))
    if align is not None and props.get('align') != cur['align']:
        para.alignment = align
    def num(k):
        try:
            return float(props.get(k))
        except (TypeError, ValueError):
            return None
    first, left, before, after, line = num('first'), num('left'), num('before'), num('after'), num('line')
    if first is not None and abs(first - cur['first']) > 0.6:
        ind = p_el.get_or_add_pPr().find(qn('w:ind'))
        if ind is not None:
            for a in ('w:firstLineChars', 'w:hangingChars', 'w:firstLine', 'w:hanging'):
                ind.attrib.pop(qn(a), None)
        pf.first_line_indent = Pt(first)
    if left is not None and abs(left - cur['left']) > 0.6:
        ind = p_el.get_or_add_pPr().find(qn('w:ind'))
        if ind is not None:
            for a in ('w:leftChars', 'w:startChars'):
                ind.attrib.pop(qn(a), None)
        pf.left_indent = Pt(left)
    if before is not None and abs(before - cur['before']) > 0.6:
        pf.space_before = Pt(before)
    if after is not None and abs(after - cur['after']) > 0.6:
        pf.space_after = Pt(after)
    if line and (cur['line'] is None or abs(line - cur['line']) > 0.01):
        pf.line_spacing = line


def _set_flag(rpr, tag, on):
    el = rpr.find(qn(tag))
    if el is not None:
        rpr.remove(el)
    if on:
        rpr.append(OxmlElement(tag))


def _rewrite_paragraph(p_el, parent, item, doc):
    base = None
    for r in p_el.iter(qn('w:r')):
        rpr = r.find(qn('w:rPr'))
        if rpr is not None and _run_text(r):
            base = deepcopy(rpr)
            break
    for child in list(p_el):
        if child.tag != qn('w:pPr'):
            p_el.remove(child)
    para = Paragraph(p_el, parent)
    props = item.get('props') or {'align': item.get('align', 'left')}
    _apply_props(p_el, doc, props)
    for seg in item.get('runs', []):
        text = str(seg.get('t', ''))
        if not text:
            continue
        run = para.add_run()
        old = run._r.find(qn('w:rPr'))
        if old is not None:
            run._r.remove(old)
        rpr = deepcopy(base) if base is not None else OxmlElement('w:rPr')
        for tag in STRIP_FROM_BASE:
            el = rpr.find(qn(tag))
            if el is not None:
                rpr.remove(el)
        run._r.insert(0, rpr)
        font = str(seg.get('font') or '').strip()
        if font:
            rf = OxmlElement('w:rFonts')
            for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
                rf.set(qn(a), font)
            rpr.insert(0, rf)
        if seg.get('b'):
            _set_flag(rpr, 'w:b', True); _set_flag(rpr, 'w:bCs', True)
        if seg.get('i'):
            _set_flag(rpr, 'w:i', True); _set_flag(rpr, 'w:iCs', True)
        try:
            sz = float(seg.get('sz')) if seg.get('sz') else None
        except (TypeError, ValueError):
            sz = None
        if sz and 4 <= sz <= 96:
            for tag in ('w:sz', 'w:szCs'):
                el = OxmlElement(tag); el.set(qn('w:val'), str(int(round(sz * 2)))); rpr.append(el)
        col = str(seg.get('col') or '').lstrip('#')
        if re.fullmatch(r'[0-9A-Fa-f]{6}', col):
            el = OxmlElement('w:color'); el.set(qn('w:val'), col.upper()); rpr.append(el)
        if seg.get('u'):
            el = OxmlElement('w:u'); el.set(qn('w:val'), 'single'); rpr.append(el)
        run.text = text
        # python-docx's text setter keeps rPr first; make sure the order is still valid for Word
        _sort_rpr(run._r.find(qn('w:rPr')))


RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow',
             'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern',
             'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs',
             'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']


def _sort_rpr(rpr):
    if rpr is None:
        return
    kids = list(rpr)
    def key(e):
        name = e.tag.split('}')[1]
        return RPR_ORDER.index(name) if name in RPR_ORDER else len(RPR_ORDER)
    for k in kids:
        rpr.remove(k)
    for k in sorted(kids, key=key):
        rpr.append(k)


def _new_pagebreak(model):
    p = OxmlElement('w:p')
    r = OxmlElement('w:r'); br = OxmlElement('w:br'); br.set(qn('w:type'), 'page')
    r.append(br); p.append(r)
    return p


def items_to_docx(data, items):
    doc = Document(io.BytesIO(data))
    body = doc.element.body
    paras = [c for c in body if c.tag == qn('w:p')]
    tables = [c for c in body if c.tag == qn('w:tbl')]
    sect = body.find(qn('w:sectPr'))
    order, used_p, last = [], set(), None
    for it in items:
        kind = it.get('kind')
        if kind == 'table':
            tbl = tables[int(it['src'])]
            trs = tbl.findall(qn('w:tr'))
            new_trs = []
            for row in it.get('rows', []):
                if row.get('src') is not None and int(row['src']) < len(trs):
                    tr = trs[int(row['src'])]
                elif row.get('copy') is not None and int(row['copy']) < len(trs):
                    tr = deepcopy(trs[int(row['copy'])])
                else:
                    continue
                for tc, cell in zip(tr.findall(qn('w:tc')), row.get('cells', [])):
                    for p_el, pit in zip(tc.findall(qn('w:p')), cell.get('paras', [])):
                        if not _para_locked(p_el):
                            _rewrite_paragraph(p_el, doc._body, pit, doc)
                new_trs.append(tr)
            if not new_trs:
                raise ValueError("A table needs at least one row.")
            for tr in trs:
                tbl.remove(tr)
            for tr in new_trs:
                tbl.append(tr)
            order.append(tbl)
            continue
        if kind == 'pb':
            src = it.get('src')
            if src is not None and int(src) < len(paras) and int(src) not in used_p and _is_pagebreak_only(paras[int(src)]):
                el = paras[int(src)]
                used_p.add(int(src))
            else:
                el = _new_pagebreak(last)
            order.append(el)
            continue
        src = it.get('src')
        if src is not None and int(src) < len(paras) and int(src) not in used_p:
            el = paras[int(src)]
            used_p.add(int(src))
            if not _para_locked(el):
                _rewrite_paragraph(el, doc._body, it, doc)
            elif _has_picture(el) and it.get('props'):      # the logo itself stays; its alignment can change
                _apply_props(el, doc, {'align': (it['props'] or {}).get('align')})
        else:                                       # a paragraph the user added: copy the look of the one before it
            model = last if last is not None else (paras[0] if paras else None)
            if model is None:
                continue
            el = deepcopy(model)
            for br in el.findall('.//' + qn('w:br')):
                br.getparent().remove(br)
            if _para_locked(el) or _is_pagebreak_only(el):
                el = OxmlElement('w:p')
            _rewrite_paragraph(el, doc._body, it, doc)
        order.append(el)
        last = el
    for i, p_el in enumerate(paras):                # pictures and section settings the user cannot remove stay in place
        if i not in used_p and (_has_picture(p_el) or _has_section(p_el)):
            order.append(p_el)
    for c in list(body):
        if c is not sect:
            body.remove(c)
    for el in order:
        if sect is not None:
            sect.addprevious(el)
        else:
            body.append(el)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def _check_tags_render(data):
    """Raise ValueError when a {{ tag }} is typed wrongly, so a broken template is never saved."""
    from docxtpl import DocxTemplate
    try:
        tpl = DocxTemplate(io.BytesIO(data))
        tpl.render({})
    except Exception as e:
        raise ValueError(str(e))


def _read(t):
    return template_bytes(t)


def _new_version(src, data, user_note=''):
    last = (LetterTemplate.objects.filter(letter_type=src.letter_type)
            .order_by('-version').first())
    base = re.sub(r'(_v\d+)?\.docx$', '', os.path.basename(src.name or 'template'), flags=re.I)
    ver = (last.version if last else src.version) + 1
    new = LetterTemplate(letter_type=src.letter_type, version=ver,
                         name=f"{base}_v{ver}.docx", logo=src.logo, is_active=False)
    new.file.save(f"{base}_v{ver}.docx", ContentFile(data), save=True)
    new.file.open('rb')
    try:
        tags = extract_fields_from_docx(new.file).get('template_variables') or []
    finally:
        new.file.close()
    old = {m.placeholder: m for m in src.field_mappings.all()}
    for tag in tags:
        m = old.get(tag)
        TemplateFieldMapping.objects.create(
            template=new, placeholder=tag,
            source_type=m.source_type if m else TemplateFieldMapping.SOURCE_UNMAPPED,
            source_value=m.source_value if m else '')
    return new, tags


# ----------------------------------------------------------------------
# Views
# ----------------------------------------------------------------------
@staff_required
def template_editor(request, template_id):
    t = get_object_or_404(LetterTemplate, id=template_id)
    try:
        data = _read(t)
    except TemplateFileMissing as e:
        messages.error(request, str(e))
        return redirect('letters:template_list')

    if request.method == 'POST':
        try:
            payload = json.loads(request.body.decode('utf-8'))
            items = payload['items']
            new_data = items_to_docx(data, items)
            _check_tags_render(new_data)
        except ValueError as e:
            return JsonResponse({'ok': False, 'error': f"A tag is typed wrongly, so nothing was saved. Check that every tag looks like {{{{ name }}}}. ({e})"}, status=400)
        except Exception as e:
            return JsonResponse({'ok': False, 'error': f"Could not build the Word file: {e}"}, status=400)
        new, tags = _new_version(t, new_data)
        unmapped = new.field_mappings.filter(source_type=TemplateFieldMapping.SOURCE_UNMAPPED).count()
        log_activity(request, 'EDIT_TEMPLATE', f"Edited '{t.name}' in the editor and saved version {new.version} of '{t.letter_type}'")
        messages.success(request, f"Saved as version {new.version}." + (f" {unmapped} tag(s) still need mapping before it can be activated." if unmapped else " Every tag is mapped; you can preview and activate it."))
        return JsonResponse({'ok': True, 'url': request.build_absolute_uri(
            f"/templates/{new.id}/editor/")})

    items, page = docx_to_items(data)
    versions = LetterTemplate.objects.filter(letter_type=t.letter_type).order_by('-version')
    next_version = (versions.first().version if versions else t.version) + 1
    mapped = set(t.field_mappings.exclude(source_type=TemplateFieldMapping.SOURCE_UNMAPPED).values_list('placeholder', flat=True))
    return render(request, 'letters/template_editor.html', {
        'template': t, 'items': items, 'page': page, 'versions': versions, 'next_version': next_version,
        'tags': [{'tag': k, 'label': v} for k, v in AVAILABLE_SOURCE_FIELDS],
        'mapped': sorted(mapped), 'logos': Logo.objects.all(),
    })


@staff_required
@require_POST
def template_restore(request, template_id):
    t = get_object_or_404(LetterTemplate, id=template_id)
    new, tags = _new_version(t, _read(t))
    log_activity(request, 'EDIT_TEMPLATE', f"Restored version {t.version} of '{t.letter_type}' as version {new.version}")
    messages.success(request, f"Version {t.version} was copied as the new version {new.version}. Preview and activate it when ready.")
    return redirect('letters:template_editor', template_id=new.id)


# ----------------------------------------------------------------------
# small read-only picture of a template (cards on the Document templates page)
# ----------------------------------------------------------------------
def mini_preview_html(data, max_items=22):
    """HTML of the first page of a template, drawn at full A4 width (the page scales it down with CSS)."""
    from django.utils.html import escape
    from django.utils.safestring import mark_safe
    items, page = docx_to_items(data, pics=False)

    def para(pr, runs, empty_h=True):
        pr = pr or {}
        st = (f"text-align:{'justify' if pr.get('align') == 'both' else pr.get('align', 'left')};"
              f"text-indent:{pr.get('first', 0)}pt;padding-left:{max(pr.get('left', 0), 0)}pt;"
              f"margin:{pr.get('before', 0)}pt 0 {pr.get('after', 0)}pt;line-height:{(pr.get('line') or 1) * 1.1:.2f};")
        out = []
        for r in runs or []:
            css = ''.join([
                'font-weight:700;' if r.get('b') else '', 'font-style:italic;' if r.get('i') else '',
                'text-decoration:underline;' if r.get('u') else '',
                f"font-size:{r['sz']}pt;" if r.get('sz') else '', f"color:#{r['col']};" if r.get('col') else ''])
            txt = escape(r['t']).replace('\n', '<br>')
            txt = re.sub(r'(\{\{.*?\}\})', r'<span class="mtag">\1</span>', txt)
            out.append(f'<span style="{css}">{txt}</span>' if css else txt)
        return f'<div style="{st}white-space:pre-wrap;min-height:1em">{"".join(out)}</div>'

    html, n = [], 0
    for it in items:
        if n >= max_items:
            break
        n += 1
        if it['kind'] == 'pb':
            break
        if it['kind'] == 'table':
            rows = []
            for row in it['rows']:
                cells = ''.join(f'<td style="padding:1px 4px;border:1px solid #999;vertical-align:top">'
                                + ''.join(para(p['props'], p['runs']) for p in c['paras']) + '</td>' for c in row['cells'])
                rows.append(f'<tr>{cells}</tr>')
            html.append(f'<table style="border-collapse:collapse;width:100%;table-layout:auto">{"".join(rows)}</table>')
        elif it.get('locked'):
            if 'Picture' in (it.get('note') or ''):
                al = (it.get('props') or {}).get('align', 'left')
                html.append(f'<div style="text-align:{al};margin:2pt 0"><span class="mlogo">LOGO</span></div>')
        else:
            html.append(para(it.get('props'), it.get('runs')))
    wrapper = (f'<div class="minisheet" style="padding:{page["mt"]}cm {page["mr"]}cm {page["mb"]}cm {page["ml"]}cm;'
               f'font-family:\'{page["font"]}\',\'TH Sarabun New\',serif;font-size:{page["size"]}pt">{"".join(html)}</div>')
    return mark_safe(wrapper)
