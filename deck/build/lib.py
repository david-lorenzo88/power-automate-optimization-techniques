"""Small helper layer over python-pptx for the Experts Live template (20in x 11.25in)."""
import copy
from lxml import etree
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

W, H = 20.0, 11.25
FONT = 'Codec Pro'
CODE = 'Courier New'
ICON_DIR = None  # set by build

NAVY = '040E3D'; LAV = 'F1F0FF'; BLUE = '1E6FFF'; CYAN = '3FD0E0'; VIOLET = '7B61FF'
INK = '040E3D'; MUTED_L = '4A5578'; MUTED_D = 'B4BCE6'
CARD_D = '0D1B5E'; CARD_D_LINE = '2B3FA0'; CARD_L = 'FFFFFF'; CARD_L_LINE = 'D5D8F2'
SLOW = 'E5484D'; FAST = '0E9F73'; FAST_D = '35E3A8'; CODE_BG = '0A1442'; CODE_FG = 'E8ECFF'


def rgb(h):
    return RGBColor.from_string(h)


def _apply_run(run, size=22, color=INK, font=FONT, bold=False, italic=False):
    f = run.font
    f.size = Pt(size); f.name = font; f.bold = bold; f.italic = italic
    f.color.rgb = rgb(color)
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set('typeface', font)


def text(slide, x, y, w, h, paras, size=22, color=INK, font=FONT, align='l', anchor='t',
         bold=False, line=None, space_after=0, name=None, italic=False):
    """paras: str | list of (str | list[(text, opts)] | dict(runs=..., bullet=, align=, space_after=))"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {'t': MSO_ANCHOR.TOP, 'm': MSO_ANCHOR.MIDDLE, 'b': MSO_ANCHOR.BOTTOM}[anchor]
    if isinstance(paras, str):
        paras = [paras]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        popts = {}
        if isinstance(p, dict):
            popts = p
            runs = p['runs']
        else:
            runs = p
        if isinstance(runs, str):
            runs = [(runs, {})]
        para.alignment = {'l': PP_ALIGN.LEFT, 'c': PP_ALIGN.CENTER, 'r': PP_ALIGN.RIGHT}[popts.get('align', align)]
        if line:
            para.line_spacing = line
        sa = popts.get('space_after', space_after)
        if sa:
            para.space_after = Pt(sa)
        if popts.get('space_before'):
            para.space_before = Pt(popts['space_before'])
        for t, o in runs:
            r = para.add_run()
            r.text = t
            _apply_run(r, size=o.get('size', popts.get('size', size)), color=o.get('color', popts.get('color', color)),
                       font=o.get('font', font), bold=o.get('bold', bold), italic=o.get('italic', italic))
        if popts.get('bullet'):
            pPr = para._p.get_or_add_pPr()
            ind = popts.get('indent', 0.4)
            pPr.set('marL', str(int(Inches(ind)))); pPr.set('indent', str(int(-Inches(ind))))
            buClr = etree.SubElement(pPr, qn('a:buClr'))
            etree.SubElement(buClr, qn('a:srgbClr')).set('val', popts.get('bullet_color', BLUE))
            etree.SubElement(pPr, qn('a:buFont')).set('typeface', 'Arial')
            etree.SubElement(pPr, qn('a:buChar')).set('char', '•')
    return tb


def shape(slide, kind, x, y, w, h, fill=None, line=None, line_w=1.0, radius=None, name=None):
    shp = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        shp.name = name
    if fill:
        shp.fill.solid(); shp.fill.fore_color.rgb = rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = rgb(line); shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    if radius is not None and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = radius
    # no text, no shadow from theme
    spPr = shp._element.spPr
    eff = etree.SubElement(spPr, qn('a:effectLst'))
    shp.text_frame.text = ''
    return shp


def card(slide, x, y, w, h, dark=False, fill=None, line=None, radius=0.08, name=None):
    return shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h,
                 fill=fill or (CARD_D if dark else CARD_L), line=line or (CARD_D_LINE if dark else CARD_L_LINE),
                 line_w=1.0, radius=radius, name=name)


def icon(slide, key, x, y, d, circle=BLUE, glyph='w', pad=0.24):
    if circle:
        shape(slide, MSO_SHAPE.OVAL, x, y, d, d, fill=circle, name=f'Icon circle {key}')
    p = d * pad
    pic = slide.shapes.add_picture(f'{ICON_DIR}/{key}_{glyph}.png', Inches(x + p), Inches(y + p), Inches(d - 2 * p), Inches(d - 2 * p))
    pic.name = f'Icon {key}'
    return pic


def pill(slide, label, x=None, y=1.85, w=None, h=0.62, fill=CYAN, fg=NAVY, icon_key='play'):
    w = w or (1.15 + len(label) * 0.152)
    x = x if x is not None else W - 1.12 - w
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=fill, radius=0.5, name='Demo pill')
    ic = slide.shapes.add_picture(f'{ICON_DIR}/{icon_key}_n.png', Inches(x + 0.22), Inches(y + 0.17), Inches(h - 0.34), Inches(h - 0.34))
    ic.name = 'Demo pill icon'
    text(slide, x + 0.22 + h - 0.2, y, w - (h + 0.25), h, label, size=17, color=fg, anchor='m', name='Demo pill text')


def arrow(slide, x1, y1, x2, y2, color=BLUE, width=2.25):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = rgb(color); c.line.width = Pt(width)
    ln = c.line._get_or_add_ln()
    tail = etree.SubElement(ln, qn('a:tailEnd')); tail.set('type', 'triangle'); tail.set('w', 'med'); tail.set('len', 'med')
    return c


def code(slide, x, y, w, h, lines, size=17, bg=CODE_BG, line=CARD_D_LINE, pad=0.35, name='Code'):
    """lines: list of str or list of (text, color) runs."""
    card(slide, x, y, w, h, fill=bg, line=line, radius=0.04, name=name + ' background')
    paras = []
    for ln in lines:
        if isinstance(ln, str):
            col = '8E9BD6' if ln.strip().startswith('//') else CODE_FG
            paras.append([(ln if ln else ' ', {'color': col})])
        else:
            paras.append([(t, {'color': c}) for t, c in ln])
    return text(slide, x + pad, y + pad * 0.8, w - 2 * pad, h - 1.6 * pad, paras, size=size, font=CODE, color=CODE_FG, line=1.05, name=name)


def table(slide, x, y, w, col_w, rows, header_fill=BLUE, header_fg='FFFFFF', body_fill='FFFFFF', alt_fill='F6F6FF',
          fg=INK, size=18, row_h=0.58, header_size=None, code_cols=(), line_color='D5D8F2', name='Table', dark=False):
    nrows, ncols = len(rows), len(rows[0])
    gs = slide.shapes.add_table(nrows, ncols, Inches(x), Inches(y), Inches(w), Inches(row_h * nrows))
    gs.name = name
    tbl = gs.table
    tblPr = tbl._tbl.tblPr
    tblPr.set('firstRow', '0'); tblPr.set('bandRow', '0')
    style = tblPr.find(qn('a:tableStyleId'))
    if style is not None:
        style.text = '{5940675A-B579-460E-94D1-54222C63F5DA}'  # No Style, Table Grid
    for i, cw in enumerate(col_w):
        tbl.columns[i].width = Inches(cw)
    for r in range(nrows):
        tbl.rows[r].height = Inches(row_h)
        for c in range(ncols):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(header_fill if r == 0 else (alt_fill if r % 2 == 0 else body_fill))
            cell.margin_left = cell.margin_right = Inches(0.15)
            cell.margin_top = cell.margin_bottom = Inches(0.06)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            val = rows[r][c]
            p = tf.paragraphs[0]
            run = p.add_run(); run.text = val
            is_code = r > 0 and c in code_cols
            _apply_run(run, size=(header_size or size) if r == 0 else (size - 2 if is_code else size),
                       color=header_fg if r == 0 else fg, font=CODE if is_code else FONT, bold=(r == 0))
            # borders
            tcPr = cell._tc.get_or_add_tcPr()
            for edge in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
                ln = etree.SubElement(tcPr, qn(edge)); ln.set('w', '9525')
                sf = etree.SubElement(ln, qn('a:solidFill')); etree.SubElement(sf, qn('a:srgbClr')).set('val', line_color)
            # schema order: borders must precede fill in tcPr
            fill_el = tcPr.find(qn('a:solidFill'))
            if fill_el is not None:
                tcPr.remove(fill_el); tcPr.append(fill_el)
    return gs


def kicker_title(slide, kicker, title, dark, centered=False, title_size=46, y=1.95, x=1.12, w=13.4):
    col = LAV if dark else INK
    kcol = CYAN if dark else BLUE
    if centered:
        w = 12.0; x = (W - w) / 2
    al = 'c' if centered else 'l'
    text(slide, x, y, w, 0.4, kicker.upper(), size=17, color=kcol, align=al, name='Kicker')
    text(slide, x, y + 0.42, w, 1.05, title, size=title_size, color=col, align=al, name='Title')


def notes(slide, s):
    slide.notes_slide.notes_text_frame.text = s


def remove_text_shapes(slide):
    for shp in list(slide.shapes):
        if shp.has_text_frame and shp.text_frame.text.strip():
            shp._element.getparent().remove(shp._element)


def set_runs(shp, paras_text, size=None, color=None):
    """Replace text of an existing template text box, keeping first run formatting."""
    tf = shp.text_frame
    first_p = tf.paragraphs[0]._p
    proto_r = copy.deepcopy(first_p.find(qn('a:r')))
    proto_pPr = first_p.find(qn('a:pPr'))
    txBody = first_p.getparent()
    for p in txBody.findall(qn('a:p')):
        txBody.remove(p)
    for t in paras_text:
        p = etree.SubElement(txBody, qn('a:p'))
        if proto_pPr is not None:
            p.append(copy.deepcopy(proto_pPr))
        r = copy.deepcopy(proto_r)
        r.find(qn('a:t')).text = t
        rPr = r.find(qn('a:rPr'))
        if size:
            rPr.set('sz', str(int(size * 100)))
        if color:
            sf = rPr.find(qn('a:solidFill'))
            sf.find(qn('a:srgbClr')).set('val', color)
        p.append(r)
