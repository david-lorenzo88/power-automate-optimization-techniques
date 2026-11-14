"""Builds the session deck on the Experts Live Emirates template.

    python deck/build_deck.py <skeleton.pptx> <out.pptx>

<skeleton.pptx> is the template with its frames duplicated into the 24-slide order
(see deck/README.md). Every number on the slides is a measured result from the lab
(lab/RESULTS.md, poc/RESULTS.md) — update RESULTS below, rebuild, don't hand-edit.
"""
import sys

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

# ---------------------------------------------------------------- palette (sampled from the template)
NAVY = RGBColor(0x03, 0x0E, 0x3C)
NAVY2 = RGBColor(0x17, 0x26, 0x69)
BLUE = RGBColor(0x2F, 0x6B, 0xFF)
SKY = RGBColor(0x8F, 0xB8, 0xFF)
LAV = RGBColor(0xF1, 0xF0, 0xFE)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED = RGBColor(0x55, 0x5E, 0x80)
MUTED_D = RGBColor(0xB8, 0xC0, 0xE0)
SLOW = RGBColor(0xD9, 0x3F, 0x55)
SLOW_T = RGBColor(0xFC, 0xE4, 0xE8)
FAST = RGBColor(0x0B, 0x8F, 0x7E)
FAST_T = RGBColor(0xDC, 0xF4, 0xEF)
FAST_D = RGBColor(0x3D, 0xDC, 0xC4)
CODE_BG = RGBColor(0x0B, 0x16, 0x4F)

HEAD = 'Codec Pro'  # embedded in the template
BODY = 'Calibri'
MONO = 'Consolas'

# ---------------------------------------------------------------- measured results (PREMIUM-US, Nov 2026)
R = {
    'UC1': dict(vol=500, slow_s=113.9, fast_s=0.9, slow_act=664, fast_act=17, iters=500, x=130, checksum='74|46009.79'),
    'UC2': dict(vol=200, slow_s=181.2, fast_s=1.5, slow_calls=201, fast_calls=2, x=124, checksum='200|99799.00|66'),
    'UC3': dict(vol=100, slow_s=254.7, fast_s=1.4, slow_act=947, fast_act=20, iters=505, x=177),
    'UC4': dict(vol=1000, slow_s=665.3, mid_s=83.0, fast_s=44.9, slow_calls=1001, mid_calls=1001, fast_calls=5),
    'UC5': dict(single_ms=182.1, cm_ms=10.1, um_ms=10.7, up_ms=34.6, slow_s_100=28.8, mid_s_100=1.9),
    'UC6': dict(vol=2000, slow_s=59.6, fast_s=1.1, slow_kb=5943781 / 1024, fast_kb=11623 / 1024, slow_rows=2000, fast_rows=100, slow_calls=101, x=54),
    'UC7': dict(changes=100, slow_runs=100, fast_runs=10),
    'UC8': dict(vol=1000, slow_rows=100, slow_sum='12,163.15', fast_sum='499,995.00'),
    'UC10': dict(vol=200, slow_s=76.0, fast_s=6.2, x=12),
}

prs = None


def inch(v):
    return Inches(v)


def text(slide, x, y, w, h, paras, *, anchor=MSO_ANCHOR.TOP, name=None, margin=0):
    """paras: list of dicts {t, size, color, bold, font, align, space_after, runs:[{t,color,bold,font,size}]}"""
    tb = slide.shapes.add_textbox(inch(x), inch(y), inch(w), inch(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ('margin_left', 'margin_right', 'margin_top', 'margin_bottom'):
        setattr(tf, side, inch(margin))
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get('align', PP_ALIGN.LEFT)
        if 'space_after' in p:
            para.space_after = Pt(p['space_after'])
        if 'line' in p:
            para.line_spacing = p['line']
        for r in p.get('runs', [p]):
            run = para.add_run()
            run.text = r['t']
            f = run.font
            f.size = Pt(r.get('size', p.get('size', 24)))
            f.bold = r.get('bold', p.get('bold', False))
            f.name = r.get('font', p.get('font', BODY))
            f.color.rgb = r.get('color', p.get('color', NAVY))
    return tb


def box(slide, x, y, w, h, fill, *, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06, line=None, name=None, shadow=False):
    s = slide.shapes.add_shape(shape, inch(x), inch(y), inch(w), inch(h))
    if name:
        s.name = name
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line
        s.line.width = Pt(1.25)
    else:
        s.line.fill.background()
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    if not shadow:
        s.shadow.inherit = False
    s.text_frame.text = ''
    return s


def title(slide, t, *, dark=False, y=2.15, x=1.12, w=17.6, size=50, align=PP_ALIGN.LEFT):
    return text(slide, x, y, w, 1.25, [{'t': t, 'size': size, 'font': HEAD, 'color': WHITE if dark else NAVY, 'align': align}], name='Title')


def kicker(slide, t, *, dark=False, y=1.82, x=1.12, color=None):
    return text(slide, x, y, 12, 0.45, [{'t': t.upper(), 'size': 18, 'bold': True, 'color': color or (SKY if dark else BLUE)}], name='Kicker')


def footnote(slide, t, *, dark=False, y=10.35):
    return text(slide, 1.12, y, 17.6, 0.45, [{'t': t, 'size': 18, 'color': MUTED_D if dark else MUTED}], name='Footnote')


def code(slide, x, y, w, h, lines, *, size=17, name='Code'):
    box(slide, x, y, w, h, CODE_BG, radius=0.08, name=f'{name} background')
    paras = []
    for ln in lines:
        if isinstance(ln, tuple):
            paras.append({'t': ln[0], 'size': size, 'font': MONO, 'color': ln[1], 'space_after': 4})
        else:
            paras.append({'t': ln, 'size': size, 'font': MONO, 'color': WHITE, 'space_after': 4})
    return text(slide, x + 0.35, y + 0.25, w - 0.7, h - 0.5, paras, anchor=MSO_ANCHOR.MIDDLE, name=name)


def arrow(slide, x, y, w, h, color):
    s = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, inch(x), inch(y), inch(w), inch(h))
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    s.shadow.inherit = False
    s.name = 'Arrow'
    return s


NB = '\u00a0'  # keep numbers and units together


def secs(v):
    return '–' if v is None else (f'{v:.1f}{NB}s' if v < 120 else f'{v / 60:.1f}{NB}min')


def num(v):
    return f'{v:,.0f}'


def factor_badge(slide, cx, cy, factor, label='faster', d=2.4):
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, inch(cx - d / 2), inch(cy - d / 2), inch(d), inch(d))
    c.name = 'Speed-up badge'
    c.fill.solid()
    c.fill.fore_color.rgb = NAVY
    c.line.fill.background()
    c.shadow.inherit = False
    tf = c.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = factor
    r.font.size = Pt(48 if len(factor) <= 3 else 42)
    r.font.bold = True
    r.font.name = HEAD
    r.font.color.rgb = WHITE
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = label
    r2.font.size = Pt(18)
    r2.font.name = BODY
    r2.font.color.rgb = SKY


def variant_card(slide, x, y, w, h, variant, technique, metrics, *, tint, color):
    box(slide, x, y, w, h, tint, radius=0.05, name=f'{variant} card')
    text(slide, x + 0.45, y + 0.35, w - 0.9, 0.45, [{'t': variant, 'size': 20, 'bold': True, 'color': color}], name=f'{variant} label')
    text(slide, x + 0.45, y + 0.85, w - 0.9, 1.2, [{'t': technique, 'size': 28, 'bold': True, 'color': NAVY, 'line': 0.95}], name=f'{variant} technique')
    my = y + 2.35
    # columns sized by the length of their value, so "4.2 min" or "5.7 MB" never wraps
    weights = [len(v) + 3 for v, _ in metrics]
    widths = [(w - 0.9) * k / sum(weights) for k in weights]
    for i, (value, label) in enumerate(metrics):
        cx = x + 0.45 + sum(widths[:i])
        col_w = widths[i]
        text(slide, cx, my, col_w - 0.1, h - (my - y) - 0.3, [
            {'t': value, 'size': 40, 'bold': True, 'color': color, 'font': HEAD, 'space_after': 2},
            {'t': label, 'size': 19, 'color': MUTED},
        ], name=f'{variant} metric {i + 1}')


def pair(slide, slow, fast, factor, *, y=3.35, h=4.6):
    variant_card(slide, 1.12, y, 7.55, h, 'SLOW', slow[0], slow[1], tint=SLOW_T, color=SLOW)
    variant_card(slide, 11.33, y, 7.55, h, 'FAST', fast[0], fast[1], tint=FAST_T, color=FAST)
    factor_badge(slide, 10.0, y + h / 2, factor)


def stat_card(slide, x, y, w, h, value, label, sub, *, dark=False, color=None, vsize=84):
    box(slide, x, y, w, h, NAVY2 if dark else WHITE, radius=0.06, name=f'Stat {value}')
    text(slide, x + 0.4, y + 0.35, w - 0.8, h - 0.7, [
        {'t': value, 'size': vsize, 'bold': True, 'font': HEAD, 'color': color or (WHITE if dark else BLUE), 'space_after': 4},
        {'t': label, 'size': 27, 'bold': True, 'color': WHITE if dark else NAVY, 'space_after': 8},
        {'t': sub, 'size': 21, 'color': MUTED_D if dark else MUTED},
    ], name=f'Stat {value} text')


def notes(slide, t):
    slide.notes_slide.notes_text_frame.text = t


def set_runs(shape, texts):
    """Replace the text of an existing template text box, keeping its formatting (one run per paragraph)."""
    paras = shape.text_frame.paragraphs
    for p, t in zip(paras, texts):
        runs = p.runs
        if runs:
            runs[0].text = t
            for r in runs[1:]:
                r.text = ''
    for p in paras[len(texts):]:
        for r in p.runs:
            r.text = ''


def section(slide, number, name, sub):
    # The template's text box on this frame draws navy text on navy: remove it, add our own.
    for sh in [sh for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text.strip()]:
        sh._element.getparent().remove(sh._element)
    text(slide, 2.32, 3.55, 6, 0.9, [{'t': number, 'size': 36, 'bold': True, 'color': SKY, 'font': HEAD}], name='Section number')
    text(slide, 2.32, 4.4, 13.0, 4.2, [{'t': name, 'size': 80, 'color': WHITE, 'font': HEAD, 'line': 0.9, 'space_after': 18},
                                        {'t': sub, 'size': 30, 'color': MUTED_D}], name='Section title')


# ================================================================ build
def build(src, out):
    global prs
    prs = Presentation(src)
    S = list(prs.slides)
    assert len(S) == 24, len(S)

    # 1 ---------------------------------------------------------- title
    s = S[0]
    boxes = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
    t_title = next(b for b in boxes if 'Project' in b.text_frame.text)
    t_sub = next(b for b in boxes if 'Innovating' in b.text_frame.text)
    t_by = next(b for b in boxes if 'Presented' in b.text_frame.text)
    set_runs(t_title, ['From Hours', 'to Seconds'])
    t_sub.width = inch(9.6)
    set_runs(t_sub, ['Power Automate optimization techniques that actually work'])
    t_by.width = inch(9.6)
    set_runs(t_by, ['David Lorenzo López  ·  24 Nov 2026'])
    notes(s, 'Welcome. 50 minutes, almost all of it demos. Every number you will see today was measured on a real tenant this month, '
             'with a slow and a fast version of the same flow producing the same result.')

    # 2 ---------------------------------------------------------- hook (light)
    s = S[1]
    kicker(s, 'Measured, not promised')
    title(s, 'Same input. Same result. Up to 177× faster')
    for i, (v, lbl, sub) in enumerate([
        ('177×', 'Aggregation per region', f"{secs(R['UC3']['slow_s'])} → {secs(R['UC3']['fast_s'])}  ·  {R['UC3']['vol']} orders"),
        ('130×', 'Data shaping', f"{secs(R['UC1']['slow_s'])} → {secs(R['UC1']['fast_s'])}  ·  {R['UC1']['vol']} orders"),
        ('124×', 'Join / enrichment', f"{secs(R['UC2']['slow_s'])} → {secs(R['UC2']['fast_s'])}  ·  {R['UC2']['vol']} orders"),
    ]):
        stat_card(s, 1.12 + i * 6.05, 3.7, 5.65, 3.9, v, lbl, sub, vsize=110)
    footnote(s, 'PREMIUM-US demo environment, November 2026. Each slow/fast pair returns an identical checksum.', y=8.05)
    notes(s, 'The hook: these are three flows from today. Same input, same output (we checksum the result), and the fast version is two orders of magnitude faster. '
             'Nothing exotic: Select, Filter array, a dictionary object and xpath. Standard connectors only.')

    # 3 ---------------------------------------------------------- cost (dark)
    s = S[2]
    kicker(s, 'Why it matters', dark=True)
    title(s, 'Slow flows cost more than time', dark=True)
    cards = [
        ('40,000', 'requests per user per day', 'Every action execution counts: each loop iteration too (Premium user)'),
        ('1', 'default loop concurrency', 'Out of the box, every Apply to each runs one item at a time'),
        ('600', 'SharePoint calls per minute', 'Per connection: a loop of Create item hits throttling fast'),
        ('120 s', 'synchronous limit', 'Power Apps waits 2 minutes for a flow response, then gives up'),
    ]
    for i, (v, lbl, sub) in enumerate(cards):
        stat_card(s, 1.12 + i * 4.5, 3.4, 4.15, 4.85, v, lbl, sub, dark=True, color=SKY, vsize=66)
    text(s, 1.12, 8.65, 17.6, 1.4, [{'runs': [
        {'t': 'A loop over 5,000 rows with 3 actions inside ≈ 15,000 requests: ', 'size': 27, 'color': WHITE, 'bold': True},
        {'t': 'more than a third of a user\'s daily allowance, for one run.', 'size': 27, 'color': MUTED_D},
    ]}], name='Key message')
    notes(s, 'Power Platform request limits are per action execution, whether the action does real work or not. '
             'Loops are serial by default. SharePoint throttles at 600 calls per minute per connection. And anything called from Power Apps has 120 seconds.')

    # 4 ---------------------------------------------------------- agenda (spotlight)
    s = S[3]
    title(s, 'Five ways to go from hours to seconds', y=2.2, align=PP_ALIGN.CENTER, x=1.0, w=18.0)
    items = [('01', 'Kill the loops', 'Select, Filter array, dictionary joins'), ('02', 'Aggregate without iterating', 'xml() + xpath, min/max'),
             ('03', 'Go bulk', 'SharePoint $batch, Dataverse CreateMultiple'), ('04', 'Filter at the source', '$filter, $select, $expand, trigger conditions'),
             ('05', 'Tune the platform', 'Pagination, concurrency, fan-out, respond early')]
    for i, (n, t, sub) in enumerate(items):
        x = 1.12 + i * 3.6
        box(s, x, 3.9, 3.3, 4.7, WHITE, radius=0.06, name=f'Agenda {n}')
        text(s, x + 0.3, 4.2, 2.7, 4.2, [
            {'t': n, 'size': 56, 'bold': True, 'font': HEAD, 'color': BLUE, 'space_after': 10},
            {'t': t, 'size': 28, 'bold': True, 'color': NAVY, 'space_after': 12},
            {'t': sub, 'size': 21, 'color': MUTED},
        ], name=f'Agenda {n} text')
    footnote(s, 'Then: the scoreboard of every run, and what to take home.', y=8.95)
    notes(s, 'Agenda: five technique blocks, each with a live slow-vs-fast demo launched from a Code App, then the scoreboard.')

    # 5 ---------------------------------------------------------- find the bottleneck
    s = S[4]
    kicker(s, 'Diagnose before you optimize')
    title(s, 'Find the slow action, not the slow flow')
    steps = [('Run history', 'Open a slow run and sort by action duration. The loop or the connector call that dominates is your target.'),
             ('flowrun table', 'Dataverse keeps every run: status, start, end and duration (in milliseconds). Query it to find your slowest flows.'),
             ('Automation Center', 'Tenant-wide view of failing and slow flows, plus the Analytics tab per flow.')]
    for i, (t, d) in enumerate(steps):
        x = 1.12 + i * 6.05
        box(s, x, 3.45, 5.25, 4.75, WHITE, radius=0.06, name=f'Step {i + 1}')
        text(s, x + 0.4, 3.75, 4.45, 4.3, [
            {'t': f'{i + 1}', 'size': 48, 'bold': True, 'font': HEAD, 'color': BLUE, 'space_after': 8},
            {'t': t, 'size': 30, 'bold': True, 'color': NAVY, 'space_after': 12},
            {'t': d, 'size': 23, 'color': MUTED},
        ], name=f'Step {i + 1} text')
        if i < 2:
            arrow(s, x + 5.33, 5.6, 0.62, 0.5, SKY)
    footnote(s, 'Measured in our tenant: flowrun.duration is in milliseconds, and new runs appear there after a few minutes.', y=8.6)
    notes(s, 'Before touching anything, find where the time goes. Run history shows action durations; the flowrun table in Dataverse gives you every run with duration in ms '
             '(we verified the unit: it matches the flow\'s own measurement to within half a second). Automation Center for the tenant view.')

    # 6 ---------------------------------------------------------- how we measure
    s = S[5]
    kicker(s, 'The lab')
    title(s, 'Same input, same result, different cost')
    flow = [('Code App', 'Run SLOW / FAST\nwith a correlation ID'), ('Flow', 'Logs “Running”,\nresponds early'), ('Work', 'The technique\nunder test'),
            ('Telemetry', 'Duration, actions,\ncalls, checksum'), ('Scoreboard', 'Speed-up and\n“same result?” check')]
    for i, (t, d) in enumerate(flow):
        x = 1.12 + i * 3.6
        box(s, x, 3.5, 3.0, 3.1, NAVY if i in (0, 4) else WHITE, radius=0.08, name=f'Lab {t}')
        text(s, x + 0.25, 3.65, 2.5, 2.8, [
            {'t': t, 'size': 26, 'bold': True, 'color': WHITE if i in (0, 4) else NAVY, 'space_after': 8},
            {'t': d, 'size': 20, 'color': MUTED_D if i in (0, 4) else MUTED},
        ], anchor=MSO_ANCHOR.MIDDLE, name=f'Lab {t} text')
        if i < 4:
            arrow(s, x + 3.06, 4.85, 0.48, 0.4, BLUE)
    text(s, 1.12, 7.05, 17.6, 3.0, [
        {'runs': [{'t': 'Every pair returns the same checksum. ', 'bold': True, 'size': 22}, {'t': 'If the fast flow is wrong, the scoreboard shows ❌.', 'size': 22, 'color': MUTED}], 'space_after': 8},
        {'runs': [{'t': 'Every flow responds early. ', 'bold': True, 'size': 22}, {'t': 'In our test, Run() from Power Apps failed at 120 s even though the flow finished.', 'size': 22, 'color': MUTED}], 'space_after': 8},
        {'runs': [{'t': '30 flows, generated as code. ', 'bold': True, 'size': 22}, {'t': 'Same telemetry contract everywhere; volume is a parameter.', 'size': 22, 'color': MUTED}]},
    ], name='Lab principles')
    notes(s, 'How the demos work. A Power Apps Code App launches each variant with a correlation ID. The flow logs a Running row, responds immediately, does the work, '
             'and writes duration, estimated actions, connector calls, payload and a checksum. The app polls Dataverse and compares.')

    # 7 ---------------------------------------------------------- section 01
    section(S[6], '01', 'Kill the loops', 'Data work doesn\'t need Apply to each')
    notes(S[6], 'Block 1: most loops in customer tenants only reshape, filter or join data. None of that needs a loop.')

    # 8 ---------------------------------------------------------- UC1
    s = S[7]
    kicker(s, 'UC1 · Data shaping')
    title(s, 'Reshape 500 orders in 2 actions, not 664')
    u = R['UC1']
    pair(s, ('Apply to each → Condition → Compose → Append to array variable', [(secs(u['slow_s']), 'duration'), (num(u['slow_act']), 'actions (est.)'), (num(u['iters']), 'loop iterations')]),
         ('Filter array + Select', [(secs(u['fast_s']), 'duration'), (num(u['fast_act']), 'actions (est.)'), ('0', 'loop iterations')]), f"{u['x']}×")
    footnote(s, f"{u['vol']} SharePoint orders, keep the Open ones and add a VAT column. Both return checksum {u['checksum']}.", y=8.4)
    notes(s, 'Demo UC1. Slow: loop, condition, compose, append — about 3 actions per open order and 1 per closed one, all serial. Fast: Filter array then Select. Two actions regardless of volume.')

    # 9 ---------------------------------------------------------- UC2 dictionary
    s = S[8]
    kicker(s, 'UC2 · Join / enrichment')
    title(s, 'Join without nested loops: the dictionary pattern')
    u = R['UC2']
    code(s, 1.12, 3.4, 10.2, 5.5, [
        ('1  Select customers → "C0001": {…}', SKY),
        "   concat('\"', item()?['Title'], '\":', string(item()))",
        ('2  Compose: one JSON object keyed by code', SKY),
        "   json(concat('{', join(body('Dict_parts'), ','), '}'))",
        ('3  Select orders + look up the customer', SKY),
        "   outputs('Dict')?[item()?['CustCode']]?['CustomerName']",
    ], size=19)
    for i, (lbl, sv, fv) in enumerate([('Duration', secs(u['slow_s']), secs(u['fast_s'])), ('Connector calls', num(u['slow_calls']), num(u['fast_calls']))]):
        y = 3.4 + i * 2.85
        box(s, 11.75, y, 7.13, 2.6, WHITE, radius=0.06, name=f'{lbl} card')
        text(s, 12.15, y + 0.25, 6.4, 0.45, [{'t': lbl.upper(), 'size': 18, 'bold': True, 'color': MUTED}], name=f'{lbl} label')
        text(s, 12.15, y + 0.8, 3.1, 1.6, [{'t': sv, 'size': 44, 'bold': True, 'font': HEAD, 'color': SLOW}, {'t': 'SLOW: query per order', 'size': 18, 'color': MUTED}], name=f'{lbl} slow')
        text(s, 15.45, y + 0.8, 3.1, 1.6, [{'t': fv, 'size': 44, 'bold': True, 'font': HEAD, 'color': FAST}, {'t': 'FAST: dictionary', 'size': 18, 'color': MUTED}], name=f'{lbl} fast')
    footnote(s, f"{u['vol']} orders enriched with 2,000 customers: {u['x']}× faster, same checksum {u['checksum']}.", y=9.35)
    notes(s, 'Demo UC2, usually the biggest single gain in enterprise flows. Instead of a lookup per order, read all customers once, turn them into one JSON object keyed by customer code, '
             'and look up inside a Select. N+1 calls become 2.')

    # 10 --------------------------------------------------------- decision framework (spotlight)
    s = S[9]
    title(s, 'Do you need a loop at all?', y=2.2, align=PP_ALIGN.CENTER, x=1.0, w=18.0)
    qs = [('Side effect per item?', 'Send a mail, call an API that takes one record: loop, but make it concurrent and thin.', SLOW),
          ('Only reshaping, filtering, joining, aggregating?', 'No loop: Select, Filter array, a dictionary object, xpath.', FAST),
          ('Target accepts many records per call?', 'Loop over chunks, not items: SharePoint $batch, Dataverse CreateMultiple.', BLUE),
          ('Can the source do the work?', 'Push it down: $filter, $select, $expand, FetchXML aggregates, trigger conditions.', NAVY2)]
    for i, (q, a, c) in enumerate(qs):
        x = 1.6 + (i % 2) * 8.6
        y = 3.75 + (i // 2) * 3.05
        box(s, x, y, 8.2, 2.75, WHITE, radius=0.06, name=f'Question {i + 1}')
        dot = s.shapes.add_shape(MSO_SHAPE.OVAL, inch(x + 0.4), inch(y + 0.42), inch(0.55), inch(0.55))
        dot.fill.solid(); dot.fill.fore_color.rgb = c; dot.line.fill.background(); dot.shadow.inherit = False; dot.name = f'Question {i + 1} marker'
        text(s, x + 1.2, y + 0.3, 6.7, 2.3, [{'t': q, 'size': 27, 'bold': True, 'color': NAVY, 'space_after': 8}, {'t': a, 'size': 22, 'color': MUTED}], name=f'Question {i + 1} text')
    notes(s, 'Takeaway slide for block 1: four questions in order. Only a real side effect per item justifies a loop, and even then it should be concurrent and thin.')

    # 11 --------------------------------------------------------- section 02
    section(S[10], '02', 'Aggregate without iterating', 'Count, sum, min, max: one expression each')
    notes(S[10], 'Block 2: aggregations. The xml/xpath trick.')

    # 12 --------------------------------------------------------- UC3
    s = S[11]
    kicker(s, 'UC3 · Aggregation per region')
    title(s, 'Totals per region with xpath, not nested loops')
    u = R['UC3']
    pair(s, ('Loop per region × loop over orders, Increment / Set variable', [(secs(u['slow_s']), 'duration'), (num(u['slow_act']), 'actions (est.)'), (num(u['iters']), 'iterations')]),
         ('xml() + xpath count / sum, min() / max(), union() group-by', [(secs(u['fast_s']), 'duration'), (num(u['fast_act']), 'actions (est.)'), ('0', 'iterations')]), f"{u['x']}×", h=4.55)
    code(s, 1.12, 8.25, 17.76, 1.35, ["xpath(xml(json(concat('{\"r\":{\"o\":', string(body('Orders')), '}}'))), 'sum(/r/o[Region=\"EMEA\"]/Amount)')"], size=19)
    notes(s, 'Demo UC3. Wrap the array in a root object, convert to XML, and xpath gives you count and sum per group in one expression. '
             'min and max come from max(json(concat(\'[\', join(xpath(... /text()), \',\'), \']\'))). Distinct regions with union(). Same totals to the cent.')

    # 13 --------------------------------------------------------- section 03
    section(S[12], '03', 'Go bulk', 'Many records per call, and the limits nobody tells you about')
    notes(S[12], 'Block 3: bulk APIs. Big gains, and some surprises we found on the way.')

    # 14 --------------------------------------------------------- UC4 SharePoint
    s = S[13]
    kicker(s, 'UC4 · SharePoint bulk create')
    title(s, 'SharePoint $batch: fast, if you respect its limits')
    u = R['UC4']
    rows = [('SLOW', 'Create item, concurrency 1', secs(u['slow_s']), num(u['slow_calls'])),
            ('MID', 'Create item, concurrency 50', secs(u['mid_s']), num(u['mid_calls'])),
            ('FAST', '$batch, 250 per batch, 4 in parallel', secs(u['fast_s']), num(u['fast_calls']))]
    tbl = s.shapes.add_table(4, 4, inch(1.12), inch(3.4), inch(10.4), inch(4.2)).table
    for j, (hdr, wdt) in enumerate([('', 1.4), ('Technique', 5.0), ('1,000 items', 2.0), ('HTTP calls', 2.0)]):
        tbl.columns[j].width = inch(wdt)
        c = tbl.cell(0, j)
        c.text = hdr
    for i, r in enumerate(rows, start=1):
        for j, v in enumerate(r):
            tbl.cell(i, j).text = v
    for i in range(4):
        for j in range(4):
            c = tbl.cell(i, j)
            c.fill.solid()
            c.fill.fore_color.rgb = NAVY if i == 0 else WHITE
            c.margin_left = c.margin_right = inch(0.15)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in c.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(23 if i else 19)
                    r.font.name = BODY
                    r.font.bold = (i == 0) or j in (0, 2)
                    r.font.color.rgb = WHITE if i == 0 else ({1: SLOW, 2: NAVY2, 3: FAST}[i] if j in (0, 2) else NAVY)
                p.alignment = PP_ALIGN.RIGHT if j >= 2 else PP_ALIGN.LEFT
    lessons = [('~90 s per batch', 'SharePoint works through a batch one item at a time and stops after about 90 s: 1,000 items in one batch lost 51 to timeouts.'),
               ('Not atomic', 'One bad item in a changeset: the other 99 were still created. Check every part\'s status.'),
               ('Concurrency helps too', '50 parallel Create item calls were fine at 1,000 items, but 600 calls per minute is the ceiling.')]
    for i, (h, d) in enumerate(lessons):
        y = 3.4 + i * 2.1
        box(s, 11.95, y, 6.93, 1.95, WHITE, radius=0.06, name=f'Lesson {i + 1}')
        text(s, 12.3, y + 0.15, 6.3, 1.7, [{'t': h, 'size': 23, 'bold': True, 'color': NAVY, 'space_after': 4}, {'t': d, 'size': 18, 'color': MUTED}], name=f'Lesson {i + 1} text')
    footnote(s, 'Same checksum for all three: 1000|499995.00. Batch bodies are built with chunk() + Select + join(), no loop per item.', y=9.85)
    notes(s, 'Demo UC4. What we learned building this: a SharePoint $batch is processed serially on the server, roughly 100–170 ms per item, and stops after about 90 seconds — '
             'items after that come back as 500 timeouts while the batch itself says 200. Changesets are not transactional. So: batches of 250, a few in parallel, and count every part.')

    # 15 --------------------------------------------------------- UC5 Dataverse
    s = S[14]
    kicker(s, 'UC5 · Dataverse bulk create')
    title(s, 'CreateMultiple: 18× per record, and all-or-nothing')
    u = R['UC5']
    cd = CategoryChartData()
    cd.categories = ['Single POST', 'CreateMultiple', 'UpdateMultiple', 'UpsertMultiple']
    cd.add_series('ms per record', (u['single_ms'], u['cm_ms'], u['um_ms'], u['up_ms']))
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, inch(1.12), inch(3.3), inch(10.0), inch(5.6), cd)
    gf.name = 'Dataverse ms per record chart'
    ch = gf.chart
    ch.has_legend = False
    ch.has_title = True
    ch.chart_title.text_frame.text = 'Milliseconds per record (1,000 records, Web API)'
    tr = ch.chart_title.text_frame.paragraphs[0].runs[0].font
    tr.size = Pt(21); tr.bold = True; tr.color.rgb = NAVY; tr.name = BODY
    plot = ch.plots[0]
    plot.gap_width = 60
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = '0.0" ms"'
    dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(20); dl.font.bold = True; dl.font.color.rgb = NAVY; dl.font.name = BODY
    ser = plot.series[0]
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb = FAST
    pt = ser.points[0]
    pt.format.fill.solid(); pt.format.fill.fore_color.rgb = SLOW
    ca = ch.category_axis
    ca.reverse_order = True
    ca.tick_labels.font.size = Pt(20); ca.tick_labels.font.color.rgb = NAVY; ca.tick_labels.font.name = BODY
    ca.format.line.fill.background()
    va = ch.value_axis
    va.visible = False
    va.has_major_gridlines = False
    box(s, 11.6, 3.3, 7.28, 5.6, WHITE, radius=0.06, name='Rollback card')
    text(s, 12.0, 3.6, 6.5, 5.1, [
        {'t': 'ONE BAD ROW', 'size': 18, 'bold': True, 'color': MUTED, 'space_after': 8},
        {'runs': [{'t': '19 of 20', 'size': 44, 'bold': True, 'font': HEAD, 'color': SLOW}]},
        {'t': 'rows created by the loop (Add a new row)', 'size': 20, 'color': MUTED, 'space_after': 12},
        {'runs': [{'t': '0 of 100', 'size': 44, 'bold': True, 'font': HEAD, 'color': FAST}]},
        {'t': 'rows created by a $batch changeset: rolled back', 'size': 20, 'color': MUTED, 'space_after': 12},
        {'t': 'CreateMultiple isn\'t in “Perform a bound action”: call it with HTTP with Microsoft Entra ID.', 'size': 20, 'bold': True, 'color': NAVY},
    ], name='Rollback text')
    footnote(s, f"Flows at 100 rows: loop {secs(u['slow_s_100'])}, loop with concurrency 50 {secs(u['mid_s_100'])}. Chart: Web API timings from our PoC script.", y=9.35)
    notes(s, 'Demo UC5. CreateMultiple sends up to 1,000 records per request and is one transaction. In our PoC, 10 ms per record versus 182 ms for single creates. '
             'It is collection-bound, so the Dataverse connector\'s bound-action picker can\'t call it: use HTTP with Microsoft Entra ID. Changesets roll back; loops leave you with partial data.')

    # 16 --------------------------------------------------------- section 04
    section(S[15], '04', 'Filter at the source', 'Don\'t move data you are going to throw away')
    notes(S[15], 'Block 4: let SharePoint and Dataverse do the filtering, the column selection, the lookups and even the decision whether to run at all.')

    # 17 --------------------------------------------------------- UC6
    s = S[16]
    kicker(s, 'UC6 · Filter at source')
    title(s, 'Ask for 100 rows and 4 columns, not everything')
    u = R['UC6']
    pair(s, ('Get items (all columns) → Filter array → Get item per row', [(secs(u['slow_s']), 'duration'), (f"{u['slow_kb'] / 1024:.1f}{NB}MB", 'payload read'), (num(u['slow_calls']), 'connector calls')]),
         ('One REST call: $filter + $select + $expand', [(secs(u['fast_s']), 'duration'), (f"{u['fast_kb']:.0f}{NB}KB", 'payload read'), ('1', 'connector call')]), f"{u['x']}×", h=4.55)
    code(s, 1.12, 8.25, 17.76, 1.35, ["items?$filter=Region eq 'EMEA' and Status eq 'Open'&$select=Title,Amount,Customer/CustomerName&$expand=Customer"], size=18)
    notes(s, 'Demo UC6. The slow flow reads 2,000 wide rows (26 columns), filters in the flow, then calls Get item per match for the customer. '
             'The fast one filters, selects and expands the lookup in one call. Put the indexed column first in $filter on big lists.')

    # 18 --------------------------------------------------------- UC7
    s = S[17]
    kicker(s, 'UC7 · Trigger conditions')
    title(s, 'The cheapest run is the one that never starts')
    u = R['UC7']
    for i, (v, lbl, sub, c) in enumerate([(str(u['slow_runs']), 'runs', 'Condition + Terminate inside the flow', SLOW), (str(u['fast_runs']), 'runs', 'Trigger condition on Status = Approved', FAST)]):
        x = 1.12 + i * 9.0
        box(s, x, 3.4, 8.76, 3.6, WHITE, radius=0.06, name=f'UC7 {lbl} {i}')
        text(s, x + 0.5, 3.6, 7.8, 3.2, [
            {'runs': [{'t': v, 'size': 88, 'bold': True, 'font': HEAD, 'color': c}, {'t': f'  {lbl}', 'size': 30, 'bold': True, 'color': NAVY}], 'space_after': 4},
            {'t': sub, 'size': 25, 'color': MUTED},
        ], anchor=MSO_ANCHOR.MIDDLE, name=f'UC7 stat {i}')
    code(s, 1.12, 7.4, 17.76, 1.35, ["@equals(triggerOutputs()?['body/Status'], 'Approved')"], size=24)
    footnote(s, f"{u['changes']} list item changes, 1 in 10 relevant. Every run that starts is billed, even if its first action is Terminate.", y=9.15)
    notes(s, 'Demo UC7. A driver flow changes 100 items; 10 become Approved. Both automated flows watch the list. Without a trigger condition: 100 runs, 90 of them doing nothing but billed. With it: 10.')

    # 19 --------------------------------------------------------- section 05
    section(S[18], '05', 'Tune the platform', 'Defaults that quietly cost you')
    notes(S[18], 'Block 5: pagination, concurrency and its trap, fan-out, and the 120-second limit.')

    # 20 --------------------------------------------------------- UC8 traps
    s = S[19]
    kicker(s, 'UC8 · Two silent traps')
    title(s, 'Fast is useless if the result is wrong')
    u = R['UC8']
    traps = [('No pagination', f"Get items returned {u['slow_rows']} of {num(u['vol'])} rows. No error, no warning.", 'Turn on pagination with a realistic threshold, or page with $top + $skiptoken.'),
             ('Concurrent loop + variable', 'Read a variable, add, set it back: parallel iterations overwrite each other and lose updates.', 'Don\'t accumulate in variables: Select + xpath, or Compose inside the loop.')]
    for i, (h, d, fix) in enumerate(traps):
        x = 1.12 + i * 9.0
        box(s, x, 3.35, 8.76, 3.8, WHITE, radius=0.06, name=f'Trap {i + 1}')
        text(s, x + 0.45, 3.6, 7.9, 3.4, [
            {'t': h, 'size': 32, 'bold': True, 'color': SLOW, 'space_after': 10},
            {'t': d, 'size': 24, 'color': NAVY, 'space_after': 14},
            {'runs': [{'t': 'Fix: ', 'size': 23, 'bold': True, 'color': FAST}, {'t': fix, 'size': 23, 'color': MUTED}]},
        ], name=f'Trap {i + 1} text')
    for i, (lbl, v, c) in enumerate([('SLOW checksum', f"100 | {u['slow_sum']}", SLOW), ('FAST checksum', f"1,000 | {u['fast_sum']}", FAST)]):
        x = 1.12 + i * 9.0
        box(s, x, 7.55, 8.76, 1.5, CODE_BG, radius=0.08, name=f'{lbl} box')
        text(s, x + 0.45, 7.6, 7.9, 1.4, [{'runs': [{'t': f'{lbl}  ', 'size': 20, 'bold': True, 'color': MUTED_D}, {'t': v, 'size': 30, 'bold': True, 'font': MONO, 'color': FAST_D if c == FAST else RGBColor(0xFF, 0x8A, 0x9A)}]}],
             anchor=MSO_ANCHOR.MIDDLE, name=f'{lbl} text')
    notes(s, 'Demo UC8 — the scoreboard shows a red cross here on purpose. Get items without pagination quietly returns 100 rows. And a concurrent loop that reads and sets a variable loses updates. '
             'Concurrency is great, but not with shared state.')

    # 21 --------------------------------------------------------- fan-out + respond early
    s = S[20]
    kicker(s, 'UC10 + the 120-second rule')
    title(s, 'Parallelize the work, answer the caller first')
    u = R['UC10']
    box(s, 1.12, 3.5, 8.76, 5.0, WHITE, radius=0.06, name='Fan-out card')
    text(s, 1.57, 3.8, 7.9, 4.5, [
        {'t': 'FAN-OUT TO CHILD FLOWS', 'size': 18, 'bold': True, 'color': MUTED, 'space_after': 8},
        {'runs': [{'t': secs(u['slow_s']), 'size': 54, 'bold': True, 'font': HEAD, 'color': SLOW}, {'t': '  →  ', 'size': 40, 'color': MUTED}, {'t': secs(u['fast_s']), 'size': 54, 'bold': True, 'font': HEAD, 'color': FAST}], 'space_after': 14},
        {'t': f"{u['vol']} records: one sequential loop vs chunk() into 10 chunks, Apply to each with concurrency 10, one child flow per chunk.", 'size': 23, 'color': NAVY, 'space_after': 12},
        {'t': 'Collect the child results with result(\'loop\') — no shared variable.', 'size': 21, 'color': MUTED},
    ], name='Fan-out text')
    box(s, 10.12, 3.5, 8.76, 5.0, NAVY, radius=0.06, name='Respond early card')
    text(s, 10.57, 3.8, 7.9, 4.5, [
        {'t': 'RESPOND EARLY', 'size': 18, 'bold': True, 'color': SKY, 'space_after': 8},
        {'runs': [{'t': f'120{NB}s', 'size': 54, 'bold': True, 'font': HEAD, 'color': RGBColor(0xFF, 0x8A, 0x9A)}, {'t': '  →  ', 'size': 40, 'color': MUTED_D}, {'t': f'~1{NB}s', 'size': 54, 'bold': True, 'font': HEAD, 'color': FAST_D}], 'space_after': 14},
        {'t': 'Respond at the end of a 150 s flow: Run() fails at 120 s, and the run shows as Failed even though the work finished.', 'size': 23, 'color': WHITE, 'space_after': 12},
        {'t': 'Respond first, then work, and report progress through Dataverse.', 'size': 21, 'color': MUTED_D},
    ], name='Respond early text')
    notes(s, 'Fan-out: split into chunks, run child flows in parallel — 12× here. And from the PoC: any flow called from Power Apps that can take longer than 2 minutes must respond early. '
             'If it responds late, the caller gets an error at 120 s and the run is marked failed because the Response action fails.')

    # 22 --------------------------------------------------------- scoreboard (dark, chart)
    s = S[21]
    kicker(s, 'Scoreboard', dark=True)
    title(s, 'Every pair, same result, measured', dark=True)
    board = [('UC3 Aggregation (100)', R['UC3']['x']), ('UC1 Data shaping (500)', R['UC1']['x']), ('UC2 Join (200)', R['UC2']['x']),
             ('UC6 Filter at source (2,000)', R['UC6']['x']), ('UC10 Fan-out (200)', R['UC10']['x'])]
    if R['UC4']['slow_s']:
        board.append(('UC4 SharePoint bulk (1,000)', round(R['UC4']['slow_s'] / R['UC4']['fast_s'])))
    cd = CategoryChartData()
    cd.categories = [b[0] for b in board]
    cd.add_series('Speed-up (×)', [b[1] for b in board])
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, inch(1.12), inch(3.25), inch(13.4), inch(7.0), cd)
    gf.name = 'Speed-up chart'
    ch = gf.chart
    ch.has_legend = False
    ch.has_title = False
    plot = ch.plots[0]
    plot.gap_width = 55
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = '0"×"'
    dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(24); dl.font.bold = True; dl.font.color.rgb = WHITE; dl.font.name = BODY
    ser = plot.series[0]
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb = FAST_D
    ca = ch.category_axis
    ca.reverse_order = True
    ca.tick_labels.font.size = Pt(21); ca.tick_labels.font.color.rgb = WHITE; ca.tick_labels.font.name = BODY
    ca.format.line.fill.background()
    va = ch.value_axis
    va.visible = False
    va.has_major_gridlines = False
    notes(s, 'The scoreboard from the app, live during the session. Speed-up is slow duration divided by fast duration at the same volume; every pair matched its checksum.')

    # 23 --------------------------------------------------------- takeaways (spotlight)
    s = S[22]
    title(s, 'Take home', y=2.2, align=PP_ALIGN.CENTER, x=1.0, w=18.0)
    tk = [('Measure first', 'Action durations in run history; flowrun for the whole tenant.'),
          ('No loops for data work', 'Select, Filter array, dictionary objects, xpath.'),
          ('Go bulk, know the limits', '$batch ≤ 250–500 and parallel; CreateMultiple for Dataverse.'),
          ('Push work to the source', '$filter, $select, $expand, aggregates, trigger conditions.'),
          ('Tune, then verify', 'Pagination on, no shared variables in parallel loops, respond early.')]
    for i, (h, d) in enumerate(tk):
        y = 3.65 + i * 1.32
        dot = s.shapes.add_shape(MSO_SHAPE.OVAL, inch(3.2), inch(y + 0.12), inch(0.75), inch(0.75))
        dot.fill.solid(); dot.fill.fore_color.rgb = BLUE; dot.line.fill.background(); dot.shadow.inherit = False; dot.name = f'Takeaway {i + 1} marker'
        tf = dot.text_frame; tf.margin_left = tf.margin_right = 0; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER; r = p.add_run(); r.text = str(i + 1); r.font.size = Pt(22); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = HEAD
        text(s, 4.3, y, 12.5, 1.1, [{'runs': [{'t': f'{h}  ', 'size': 26, 'bold': True, 'color': NAVY}, {'t': d, 'size': 22, 'color': MUTED}]}], anchor=MSO_ANCHOR.MIDDLE, name=f'Takeaway {i + 1}')
    notes(s, 'Five things to take home. Measure, remove loops from data work, use bulk APIs within their limits, push work to the source, and tune the platform without breaking correctness.')

    # 24 --------------------------------------------------------- thank you
    notes(S[23], 'Thank you. Questions. The flows, the Code App and the measurement scripts are in the session repository.')

    prs.save(out)


if __name__ == '__main__':
    build(sys.argv[1], sys.argv[2])
