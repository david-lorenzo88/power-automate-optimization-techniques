import sys
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.oxml.ns import qn
import lib
from lib import *

base, out, lib.ICON_DIR = sys.argv[1], sys.argv[2], sys.argv[3]
prs = Presentation(base)
ALL = list(prs.slides)
# 32 slides: two slides added after the first share (measured results, UC10 + 120 s rule);
# S keeps the original 30-slide numbering so the content below stays readable.
S = [ALL[0], ALL[1]] + ALL[3:27] + ALL[28:32]
NEW_MEASURED, NEW_UC10 = ALL[2], ALL[27]
DEMO_X, DEMO_Y = W - 1.12 - 3.9, 2.45


def demo(slide, label):
    pill(slide, label)


def divider(slide, num, title, sub, note):
    remove_text_shapes(slide)
    text(slide, 2.3, 2.55, 6, 2.2, num, size=150, color=CYAN, name='Section number')
    text(slide, 2.3, 4.95, 13.6, 1.3, title, size=66, color=LAV, name='Title')
    text(slide, 2.3, 6.45, 11.0, 1.0, sub, size=26, color=MUTED_D, name='Subtitle')
    notes(slide, note)


# ------------------------------------------------------------------ 1 · Title
s = S[0]
for shp in s.shapes:
    if not shp.has_text_frame:
        continue
    t = shp.text_frame.text
    if t.startswith('Project'):
        set_runs(shp, ['From Hours', 'to Seconds'])
    elif t.startswith('Innovating'):
        shp.width = Inches(9.6)
        set_runs(shp, ['Power Automate optimization techniques that actually work'], size=30)
    elif t.startswith('Presented'):
        shp.width = Inches(7.0)
        set_runs(shp, ['Presented by: David Lorenzo López & Samir', 'Experts Live Emirates  ·  24 Nov 2026'], size=21)
notes(s, "Welcome. 50 minutes, almost all demos. Every technique is shown as a SLOW flow and a FAST flow with the same "
         "input and the same result checksum, launched from the Flow Optimization Lab app so you see real timings.")

# ------------------------------------------------------------------ 2 · Hook (dark)
s = S[1]
kicker_title(s, 'Why we are here', 'The flow that took five hours', dark=True)
demo(s, 'LIVE · SLOW vs FAST')
# left: the anti-pattern
text(s, 1.12, 4.1, 7.6, 0.5, 'What the slow version does', size=24, color=CYAN)
items = ['Loops over every record, one at a time',
         'Calls the API once per record',
         'Retrieves every row and every column',
         'Runs even when nothing relevant changed']
text(s, 1.12, 4.85, 7.8, 3.6, [dict(runs=i, bullet=True, space_after=14, bullet_color=SLOW) for i in items], size=24, color=LAV)
text(s, 1.12, 7.85, 8.2, 1.0, 'Rebuilt with the techniques in this session: same output, about thirty seconds.',
     size=22, color=MUTED_D, italic=True)
# right: stat comparison
card(s, 10.0, 4.1, 3.2, 3.7, dark=True)
text(s, 10.0, 4.45, 3.2, 1.6, '5 h', size=88, color=SLOW, align='c')
text(s, 10.0, 6.35, 3.2, 1.0, 'before', size=22, color=MUTED_D, align='c')
arrow(s, 13.35, 5.95, 14.05, 5.95, color=CYAN, width=3)
card(s, 14.2, 4.1, 3.2, 3.7, dark=True)
text(s, 14.2, 4.45, 3.2, 1.6, '30 s', size=88, color=FAST_D, align='c')
text(s, 14.2, 6.35, 3.2, 1.0, 'after', size=22, color=MUTED_D, align='c')
text(s, 10.0, 8.2, 7.4, 0.8, '600× faster, far fewer billed actions', size=24, color=LAV, align='c')
notes(s, "Hook. Start the SLOW and FAST versions of UC9 (Excel rows) or UC1 (data shaping) from the Lab app with the "
         "Race button and let them run while you talk. The slow one will still be running when we come back to it. "
         "Tell your own 5 h -> 30 s story here; the demo flows reproduce the same pattern at a smaller volume.")

# ------------------------------------------------------------------ 3 · Why it matters (light)
s = S[2]
kicker_title(s, 'Why it matters', 'Slow flows are expensive flows', dark=False)
cards = [
    ('coins', 'They cost money', 'Every action run is a billed Power Platform request: 40,000 per day for a Premium user.'),
    ('speed', 'They get throttled', 'SharePoint allows 600 calls per minute per connection. Dataverse: 6,000 requests per 5 minutes per user.'),
    ('clock', 'They delay the business', 'A nightly job that finishes at 9 a.m. is not a nightly job anymore.'),
    ('robot', 'They slow your agents', 'Agents and Copilot call flows in real time. A slow flow is a slow agent.'),
]
cw, gap, x0, y0, ch = 4.1, 0.43, 1.12, 3.95, 5.0
for i, (ic, head, body) in enumerate(cards):
    x = x0 + i * (cw + gap)
    card(s, x, y0, cw, ch)
    icon(s, ic, x + 0.4, y0 + 0.45, 1.1)
    text(s, x + 0.4, y0 + 1.8, cw - 0.8, 1.0, head, size=25, color=INK, anchor='b')
    text(s, x + 0.4, y0 + 3.0, cw - 0.8, 1.9, body, size=19, color=MUTED_L, line=1.1)
notes(s, "Frame the cost: actions are billed and throttled the same whether they do useful work or not. "
         "A loop over 5,000 rows with three actions inside is 15,000 requests - over a third of a user's daily allowance.")

# ------------------------------------------------------------------ 4 · Six numbers (dark)
s = S[3]
kicker_title(s, 'Platform facts', 'Six numbers to know by heart', dark=True)
stats = [('1', 'Default concurrency of Apply to each: every loop runs one item at a time'),
         ('100', 'Items SharePoint Get items returns when pagination is off'),
         ('600', 'SharePoint connector calls per minute, per connection'),
         ('1,000', 'Requests you can pack into one Dataverse $batch'),
         ('40,000', 'Power Platform requests per day for a Premium user'),
         ('120 s', 'Timeout for a synchronous request or response')]
tw, th, gx, gy = 5.0, 2.65, 0.45, 0.4
for i, (num, lab) in enumerate(stats):
    x = 1.12 + (i % 3) * (tw + gx); y = 3.75 + (i // 3) * (th + gy)
    card(s, x, y, tw, th, dark=True)
    text(s, x + 0.4, y + 0.3, tw - 0.8, 1.2, num, size=60, color=CYAN)
    text(s, x + 0.4, y + 1.5, tw - 0.8, 1.05, lab, size=18, color=LAV, line=1.05)
text(s, 1.12, 10.25, 15.0, 0.4, 'Source: Microsoft Learn, Power Automate limits and configuration; Dataverse service protection API limits (Nov 2026)',
     size=13, color=MUTED_D)
notes(s, "These six numbers explain most of the techniques today. Concurrency defaults to 1. Pagination is off, so "
         "Get items silently stops at 100. SharePoint throttles at 600 calls per minute per connection. A Dataverse $batch "
         "takes up to 1,000 requests. Premium users get 40,000 requests a day. Synchronous calls time out at 120 seconds.")

# ------------------------------------------------------------------ 5 · Agenda (light spotlight)
s = S[4]
kicker_title(s, 'Agenda', 'Six moves to a faster flow', dark=False, centered=True)
steps = [('search', 'Diagnose', 'the bottleneck'), ('loop', 'Kill loops', 'Select · Filter'), ('sum', 'Aggregate', 'with XPath'),
         ('layers', 'Go bulk', '$batch · bulk APIs'), ('filter', 'Filter', 'at the source'), ('sliders', 'Tune', 'the platform')]
sw, sg = 2.15, 0.27
x0 = (W - (6 * sw + 5 * sg)) / 2
for i, (ic, a, b) in enumerate(steps):
    x = x0 + i * (sw + sg)
    card(s, x, 4.3, sw, 3.75, radius=0.12)
    icon(s, ic, x + (sw - 1.2) / 2, 4.6, 1.2, circle=BLUE if i % 2 == 0 else VIOLET)
    text(s, x, 6.0, sw, 0.5, f'0{i + 1}', size=20, color=BLUE, align='c')
    text(s, x + 0.1, 6.45, sw - 0.2, 0.6, a, size=24, color=INK, align='c')
    text(s, x + 0.1, 7.05, sw - 0.2, 0.9, b, size=17, color=MUTED_L, align='c')
text(s, 4.0, 8.6, 12.0, 0.6, 'Every move: a SLOW flow, a FAST flow, same checksum, measured live', size=20, color=MUTED_L, align='c')
notes(s, "The order follows how you would actually work on a slow flow: find the bottleneck first, then remove loops, "
         "then batch writes, then ask the source for less, and finally tune settings.")

# ------------------------------------------------------------------ 6 · The lab (dark)
s = S[5]
kicker_title(s, 'How we measure', 'The Flow Optimization Lab', dark=True)
demo(s, 'LIVE · Lab tour')
boxes = [('mobile', 'Code App', 'Pick a scenario, a variant and a volume'),
         ('bolt', 'SLOW or FAST flow', 'Same payload, same expected result'),
         ('database', 'Telemetry row', 'Timings and counters written to Dataverse'),
         ('chart', 'Scoreboard', 'Speed-up, plus links to the run and the editor')]
bw, bg_ = 3.55, 0.6
for i, (ic, head, body) in enumerate(boxes):
    x = 1.12 + i * (bw + bg_)
    card(s, x, 3.85, bw, 3.0, dark=True)
    icon(s, ic, x + 0.35, 4.15, 0.95, circle=BLUE)
    text(s, x + 0.35, 5.25, bw - 0.6, 0.5, head, size=23, color=LAV)
    text(s, x + 0.35, 5.8, bw - 0.6, 1.0, body, size=17, color=MUTED_D, line=1.05)
    if i < 3:
        arrow(s, x + bw + 0.08, 5.35, x + bw + bg_ - 0.08, 5.35, color=CYAN, width=2.5)
text(s, 1.12, 7.35, 10, 0.5, 'What every run reports', size=22, color=CYAN)
kpis = ['Duration', 'Actions billed', 'Connector calls', 'Payload size', 'Checksum match', 'Speed-up']
kx = 1.12
for k in kpis:
    kw = 0.32 + len(k) * 0.155
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, kx, 7.95, kw + 0.4, 0.6, fill=CARD_D, line=CARD_D_LINE, radius=0.5)
    text(s, kx, 7.95, kw + 0.4, 0.6, k, size=18, color=LAV, align='c', anchor='m')
    kx += kw + 0.4 + 0.25
text(s, 1.12, 9.15, 15.5, 0.6, 'Same input. Same checksum. Different cost.', size=30, color=LAV)
notes(s, "Tour the app: scenario cards, volume selector, Run SLOW / Run FAST / Race. Each flow writes a telemetry row "
         "(correlation id, duration, actions, connector calls, payload, checksum, run link). The app polls that row, so "
         "slow flows that run longer than the 120 s limit are no problem. Click 'Open run' to jump to the run history.")

# ------------------------------------------------------------------ 7 · Section 01
divider(S[6], '01', 'Find the real bottleneck', 'Run history, action timings and the flowrun table',
        "Section 1. Never optimize blind: measure first.")

# ------------------------------------------------------------------ 8 · Diagnose (light)
s = S[7]
kicker_title(s, 'Diagnose', 'Optimize the action that owns the run', dark=False)
steps = [('Open the slowest run', 'Scan the duration of every action'),
         ('Expand the loop', 'Time per iteration × number of iterations'),
         ('Look for retries', 'HTTP 429 and Retry-After in the raw outputs'),
         ('Count the runs', 'Too many runs? Add trigger conditions'),
         ('Count the actions', 'Actions × runs = requests consumed'),
         ('Watch the trend', 'flowrun table and Automation Center')]
for i, (a, b) in enumerate(steps):
    y = 3.75 + i * 0.98
    shape(s, MSO_SHAPE.OVAL, 1.12, y, 0.72, 0.72, fill=BLUE)
    text(s, 1.12, y, 0.72, 0.72, str(i + 1), size=22, color='FFFFFF', align='c', anchor='m')
    text(s, 2.15, y - 0.04, 8.6, 0.45, a, size=23, color=INK)
    text(s, 2.15, y + 0.42, 8.6, 0.4, b, size=17, color=MUTED_L)
# illustrative run timeline
card(s, 11.2, 3.9, 7.4, 5.25)
text(s, 11.6, 4.2, 6.6, 0.5, 'Illustrative run: where the time goes', size=19, color=MUTED_L)
bars = [('Trigger', 0.03, BLUE), ('Get items', 0.06, BLUE), ('Apply to each', 0.86, SLOW), ('Update status', 0.03, BLUE), ('Send email', 0.02, BLUE)]
for i, (lab, frac, col) in enumerate(bars):
    y = 4.95 + i * 0.75
    text(s, 11.6, y, 2.3, 0.5, lab, size=17, color=INK, anchor='m')
    shape(s, MSO_SHAPE.RECTANGLE, 14.0, y + 0.1, max(0.08, 4.2 * frac), 0.32, fill=col)
text(s, 11.6, 8.55, 6.6, 0.5, 'One loop, 86% of the run time', size=19, color=SLOW)
text(s, 1.12, 9.75, 10.2, 0.6, 'Measured in our tenant: flowrun.duration is in milliseconds, and new runs show up there after a few minutes.', size=15, color=MUTED_L)
notes(s, "Open the SLOW run of the hook demo via the 'Open run' link in the app. Show the action durations, expand the loop, "
         "show the per-iteration time. Mention the flowrun table in Dataverse (28-day default retention) and Automation Center "
         "for trends across flows. The bar chart on this slide is illustrative.")

# ------------------------------------------------------------------ 9 · Section 02
divider(S[8], '02', 'Kill unnecessary loops', 'Select, Filter array and expression-based data shaping',
        "Section 2. The single biggest win in most enterprise flows.")

# ------------------------------------------------------------------ 10 · Before / after (light spotlight)
s = S[9]
kicker_title(s, 'Data shaping', '130× faster, same result', dark=False, centered=True)
demo(s, 'LIVE · UC1 shaping')
for col, (x, head, hcol, acts, foot) in enumerate([
    (3.4, 'SLOW', SLOW, [('Apply to each', '500 orders', 0), ('Condition', "Status = 'Open'", 1), ('Compose', 'build the new object', 1), ('Append to array variable', 'collect the result', 1)],
     'Measured: 113.9 s and 664 actions for 500 orders'),
    (10.6, 'FAST', FAST, [('Filter array', "Status eq 'Open'", 0), ('Select', 'map to the new shape', 0)],
     'Measured: 0.9 s and 17 actions, same checksum')]):
    card(s, x, 4.1, 6.0, 5.6)
    text(s, x + 0.4, 4.35, 5.0, 0.5, head, size=22, color=hcol)
    for j, (a, b, ind) in enumerate(acts):
        y = 5.0 + j * 0.95
        ax = x + 0.4 + ind * 0.55
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, ax, y, 5.2 - ind * 0.55, 0.78, fill='F4F5FF', line='C9CDF0', radius=0.15)
        text(s, ax + 0.25, y, 2.9, 0.78, a, size=18, color=INK, anchor='m')
        text(s, ax + 2.85 - ind * 0.3, y, 2.2, 0.78, b, size=15, color=MUTED_L, anchor='m')
    text(s, x + 0.4, 8.85, 5.3, 0.7, foot, size=18, color=hcol)
    if head == 'FAST':
        text(s, x + 0.4, 7.05, 5.2, 1.5, [[('Advanced mode keeps every column:', {})], [("addProperty(item(), 'Total', …)", {'font': CODE, 'size': 15, 'color': INK})]],
             size=16, color=MUTED_L, line=1.15)
notes(s, "Demo UC1. Open the SLOW flow (Edit link in the app): Apply to each + Condition + Compose + Append to array variable. "
         "Then the FAST flow: Filter array + Select (advanced mode, addProperty). Run both at volume 5,000. "
         "Point at the checksum: identical. Point at actions billed.")

# ------------------------------------------------------------------ 11 · Toolbox (light)
s = S[10]
kicker_title(s, 'Toolbox', 'Expressions that replace loops', dark=False)
rows = [['Instead of looping to…', 'Use', 'Example'],
        ['reshape every object', 'Select + addProperty', "addProperty(item(), 'Total', mul(…))"],
        ['keep the matching rows', 'Filter array', "@greater(item()?['Amount'], 1000)"],
        ['extract one column', 'Select (text mode) + join', "join(body('Select_Emails'), ';')"],
        ['remove duplicates', 'union', "union(body('S'), body('S'))"],
        ['compare two lists', 'Filter array + contains', "not(contains(body('Ids'), item()?['Id']))"],
        ['sort', 'sort + reverse', "reverse(sort(body('S'), 'Amount'))"],
        ['take the first match', 'first / last / take', "first(body('Filter_array'))"],
        ['split into batches', 'chunk', "chunk(body('Select'), 1000)"]]
table(s, 1.12, 3.75, 17.76, [4.6, 4.6, 8.56], rows, size=18, row_h=0.6, code_cols=(2,))
notes(s, "Reference slide: share it, don't read it. Highlight union for distinct values, contains for diffs, and chunk, "
         "which feeds the bulk patterns in section 4.")

# ------------------------------------------------------------------ 12 · Dictionary join (dark)
s = S[11]
kicker_title(s, 'Joins', 'Joins without nested loops', dark=True)
demo(s, 'LIVE · UC2 join')
tiers = [('Nested loop, 5,000 × 2,000', '10,000,000', 'iterations: never finishes', SLOW),
         ('Query the customer per order', '3.0 min', 'measured: 201 connector calls for 200 orders', 'F5A524'),
         ('Dictionary lookup', '1.5 s', 'measured: 2 connector calls, 124× faster', FAST_D)]
for i, (lab, num, sub, col) in enumerate(tiers):
    y = 3.85 + i * 2.05
    card(s, 1.12, y, 6.6, 1.8, dark=True)
    text(s, 1.5, y + 0.2, 6.0, 0.45, lab, size=19, color=MUTED_D)
    text(s, 1.5, y + 0.62, 6.0, 0.8, num, size=42, color=col)
    text(s, 1.5, y + 1.32, 6.0, 0.4, sub, size=16, color=LAV)
code(s, 8.3, 3.85, 9.0, 5.9, [
    '// 1 · Select customers (text mode)',
    "concat('\"', item()?['Id'], '\":',",
    "       string(item()))",
    '',
    '// 2 · Compose Dict',
    "json(concat('{',",
    "  join(body('Select_KV'), ','), '}'))",
    '',
    '// 3 · Select orders, look up by key',
    "outputs('Dict')?[item()?['CustomerId']]",
    "               ?['Name']",
], size=19)
notes(s, "Demo UC2. Build a JSON object keyed by CustomerId (Select + join + json), then a single Select over the orders "
         "reads the customer with outputs('Dict')?[key]. O(n x m) becomes O(n) and the API calls disappear.")

# ------------------------------------------------------------------ 13 · Decision framework (light spotlight)
s = S[12]
kicker_title(s, 'Decision framework', 'To loop or not to loop', dark=False, centered=True)
qs = [('1', 'Can the source do the work?', 'Push it down', 'OData · FetchXML · trigger conditions', BLUE),
      ('2', 'Is it only reshaping, filtering, joining or summing?', 'No loop', 'Select · Filter array · xpath', VIOLET),
      ('3', 'Does the target accept many records per call?', 'Batch it', 'chunk() + $batch / CreateMultiple', BLUE),
      ('4', 'Still need a side effect per item?', 'Loop, thin and concurrent', 'Concurrency 20–50, no variables', SLOW)]
for i, (n, q, a, b, col) in enumerate(qs):
    y = 4.05 + i * 1.42
    card(s, 3.2, y, 8.3, 1.18)
    shape(s, MSO_SHAPE.OVAL, 3.45, y + 0.24, 0.7, 0.7, fill=col)
    text(s, 3.45, y + 0.24, 0.7, 0.7, n, size=20, color='FFFFFF', align='c', anchor='m')
    text(s, 4.4, y, 6.9, 1.18, q, size=20, color=INK, anchor='m')
    arrow(s, 11.65, y + 0.59, 12.25, y + 0.59, color=col)
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 12.4, y, 4.4, 1.18, fill=col, radius=0.15)
    text(s, 12.65, y + 0.14, 4.0, 0.5, a, size=20, color='FFFFFF')
    text(s, 12.65, y + 0.62, 4.0, 0.45, b, size=15, color='FFFFFF')
notes(s, "Takeaway 1: ask these four questions in order. Most loops die at question 1 or 2. Only real per-item side effects "
         "(an email per person, an API without batch support) deserve a loop, and then make it concurrent and thin.")

# ------------------------------------------------------------------ 14 · Section 03
divider(S[13], '03', 'Aggregate without loops', 'The XPath trick: sum, count, min, max and average',
        "Section 3. Aggregations without a single Apply to each.")

# ------------------------------------------------------------------ 15 · XPath (dark)
s = S[14]
kicker_title(s, 'The XPath trick', 'Sum, count, average: zero loops', dark=True)
demo(s, 'LIVE · UC3 xpath')
code(s, 1.12, 3.85, 11.4, 6.3, [
    '// 1 · Convert once (needs a single root)',
    "xml(json(concat('{\"root\":{\"r\":',",
    "   string(body('Select_Amounts')), '}}')))",
    '',
    '// 2 · Aggregate in one expression each',
    [('Sum      ', CYAN), ("xpath(Xml, 'sum(/root/r/Amount)')", CODE_FG)],
    [('Count    ', CYAN), ("xpath(Xml, 'count(/root/r)')", CODE_FG)],
    [('Count if ', CYAN), ("xpath(Xml, 'count(/root/r[Status=\"Open\"])')", CODE_FG)],
    [('Sum if   ', CYAN), ("xpath(Xml, 'sum(/root/r[Region=\"EMEA\"]/Amount)')", CODE_FG)],
    [('Average  ', CYAN), ('div(Sum, Count)', CODE_FG)],
    [('Max      ', CYAN), ("max(body('Select_Amounts_as_numbers'))", CODE_FG)],
], size=20)
facts = [('177× faster', 'measured: 4.2 min → 1.4 s, same totals'), ('947 → 20', 'actions per run (estimated by the flows)'), ('No variables', 'nothing to lock, nothing to race')]
for i, (a, b) in enumerate(facts):
    y = 3.85 + i * 2.15
    card(s, 13.0, y, 4.3, 1.9, dark=True)
    text(s, 13.35, y + 0.3, 3.7, 0.7, a, size=30, color=CYAN)
    text(s, 13.35, y + 1.05, 3.7, 0.7, b, size=17, color=LAV)
notes(s, "Demo UC3. Select keeps only the columns you need and renames them to valid XML names, xml(json(...)) converts once, "
         "then each aggregate is one expression. Show the SLOW flow with Increment variable / Set variable in a loop for comparison. "
         "Totals match to the cent.")

# ------------------------------------------------------------------ 16 · Group by + gotchas (light)
s = S[15]
kicker_title(s, 'Group by and gotchas', 'Totals per region, still no loop', dark=False)
code(s, 1.12, 3.8, 9.6, 3.1, [
    '// Distinct keys',
    "union(body('Select_Regions'), body('Select_Regions'))",
    '',
    '// Select over the keys: one total per key',
    "xpath(outputs('Xml'), concat(",
    "  'sum(/root/r[Region=\"', item(), '\"]/Amount)'))",
], size=17)
text(s, 1.12, 7.25, 9.6, 1.5, 'Three regions or three hundred: the flow has the same number of actions.', size=21, color=MUTED_L)
gotchas = [('Element names must be valid XML', 'No spaces, no leading digit: rename with Select first'),
           ('sum() returns a float', 'Round it with formatNumber()'),
           ('Empty arrays', 'Guard the conversion with empty()'),
           ('Max and Min', 'max() / min() on a number array beat the XPath idiom'),
           ('Expression size', 'Stay under 8,192 characters')]
for i, (a, b) in enumerate(gotchas):
    y = 3.8 + i * 1.08
    icon(s, 'warning', 11.4, y + 0.05, 0.7, circle='F5A524', glyph='w', pad=0.22)
    text(s, 12.35, y - 0.02, 6.4, 0.45, a, size=20, color=INK)
    text(s, 12.35, y + 0.42, 6.4, 0.45, b, size=16, color=MUTED_L)
notes(s, "Group by: union gives the distinct keys, then a Select over the keys computes each group's total with xpath. "
         "Then the gotchas: XML-safe names, float results, empty arrays, prefer max()/min() for extremes on big arrays.")

# ------------------------------------------------------------------ 17 · Section 04
divider(S[16], '04', 'Go bulk or go home', '$batch, bulk operation messages and changesets',
        "Section 4. Stop writing one record per HTTP request.")

# ------------------------------------------------------------------ 18 · Bulk landscape (light spotlight)
s = S[17]
kicker_title(s, 'Bulk options', 'One request, many records', dark=False, centered=True)
rows = [['Option', 'Per request', 'All-or-nothing', 'Best for'],
        ['SharePoint $batch', '≤ 500 (~90 s server limit)', 'No: each part commits', 'List items, standard licence'],
        ['Dataverse $batch', 'Up to 1,000', 'Yes, per changeset', 'Mixed operations, parent + children'],
        ['CreateMultiple · UpdateMultiple · UpsertMultiple', '100–1,000', 'Yes (standard tables)', 'Fastest; call via HTTP with Entra ID'],
        ['DeleteMultiple', '100 (elastic only)', 'Partial success', 'Elastic tables'],
        ['BulkDelete', 'Query-based job', 'Async system job', 'Large deletes on standard tables'],
        ['Office Scripts (Excel)', 'Whole table, one call', 'n/a', 'Replacing Add a row loops'],
        ['Microsoft Graph $batch', '20', 'No', 'Users, Teams, Planner']]
table(s, 2.35, 3.85, 15.3, [5.4, 3.0, 3.0, 3.9], rows, size=17, row_h=0.72)
notes(s, "Map of the bulk surfaces. Dataverse: bulk messages are the fastest path for create/update/upsert; $batch when you "
         "need mixed operations or Content-ID references. SharePoint $batch works with the standard 'Send an HTTP request to "
         "SharePoint' action. Excel: Office Scripts setValues() replaces thousands of 'Add a row' calls.")

# ------------------------------------------------------------------ 19 · $batch anatomy (dark)
s = S[18]
kicker_title(s, 'Anatomy', 'A $batch request, built without a loop', dark=True)
demo(s, 'LIVE · UC4 + UC5 bulk')
card(s, 1.12, 3.85, 8.4, 5.9, dark=True)
text(s, 1.5, 4.1, 7.8, 0.5, 'POST _api/$batch   multipart/mixed', size=18, color=CYAN, font=CODE)
shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 1.5, 4.8, 7.65, 4.6, fill='13246F', line=CARD_D_LINE, radius=0.05)
text(s, 1.85, 5.0, 7.0, 0.5, '--changeset   (SharePoint: not atomic)', size=17, color=VIOLET, font=CODE)
for j, t in enumerate(['POST …/items  {"Title":"#1"}', 'POST …/items  {"Title":"#2"}', 'POST …/items  {"Title":"#3"}', '… one part per record']):
    y = 5.65 + j * 0.92
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 1.85, y, 6.95, 0.72, fill='1C2F86' if j < 3 else '13246F', line=CARD_D_LINE, radius=0.15)
    text(s, 2.1, y, 6.5, 0.72, t, size=16, color=LAV, font=CODE, anchor='m')
steps = [('chunk(items, 250)', 'loop over chunks, 4 in parallel'),
         ('Select', 'one MIME part per record'),
         ('join()', 'parts + boundaries = request body'),
         ('Send an HTTP request', 'one call per chunk'),
         ("split(…, 'HTTP/1.1 201')", 'count successes, no loop')]
for i, (a, b) in enumerate(steps):
    y = 3.85 + i * 1.25
    shape(s, MSO_SHAPE.OVAL, 10.1, y + 0.1, 0.65, 0.65, fill=BLUE)
    text(s, 10.1, y + 0.1, 0.65, 0.65, str(i + 1), size=19, color='FFFFFF', align='c', anchor='m')
    text(s, 11.0, y, 6.3, 0.5, a, size=20, color=CYAN, font=CODE)
    text(s, 11.0, y + 0.5, 6.3, 0.45, b, size=17, color=LAV)
text(s, 1.12, 9.95, 16.2, 0.5, 'Measured, 1,000 SharePoint items: SLOW 11.1 min · MID (concurrency 50) 83 s · FAST 44.9 s → 15× faster. One batch stops after ~90 s of server time.',
     size=16, color=MUTED_D)
notes(s, "Demo UC4 (SharePoint, 1,000 items): SLOW is Create item in a loop; MID is the same loop at concurrency 50 - watch the "
         "429s; FAST builds the multipart body with Select + join and sends it in a couple of calls. Then UC5 on Dataverse "
         "with CreateMultiple / UpsertMultiple at 1,000 records per call.")

# ------------------------------------------------------------------ 20 · Changesets (light)
s = S[19]
kicker_title(s, 'Transactions', 'All or nothing with changesets', dark=False)
rowsdef = [('Dataverse CreateMultiple, 1 invalid record', 'Measured: 0 of 100 written, the whole request rolled back', ['bad'] + ['rolled'] * 9, SLOW),
           ('Loop of Add a new row, 1 invalid record', 'Measured: 19 of 20 written, the flow ends Partial', ['ok'] * 4 + ['bad'] + ['ok'] * 5, FAST)]
for i, (head, foot, cells, col) in enumerate(rowsdef):
    y = 3.85 + i * 2.85
    card(s, 1.12, y, 10.2, 2.55)
    text(s, 1.5, y + 0.25, 9.4, 0.5, head, size=21, color=INK)
    for k, st in enumerate(cells):
        fill = {'ok': FAST, 'bad': SLOW, 'rolled': 'D7DAEE'}[st]
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 1.5 + k * 0.92, y + 0.95, 0.75, 0.75, fill=fill, radius=0.18)
    text(s, 1.5, y + 1.85, 9.4, 0.5, foot, size=17, color=col)
text(s, 1.12, 9.6, 10.2, 0.4, 'Each square ≈ 10% of the batch', size=14, color=MUTED_L)
pts = [('Atomic in Dataverse', 'One failure rolls back the whole changeset'),
       ('Content-ID references', 'Create a parent and its children in one transaction'),
       ('Throttling vs entitlement', 'A bulk call is one request for throttling; each record still counts as a Power Platform request'),
       ('SharePoint is different', 'Not atomic: with 1 bad item, the other 99 were still committed (measured)')]
for i, (a, b) in enumerate(pts):
    y = 3.85 + i * 1.4
    icon(s, ['lock', 'key', 'speed', 'list'][i], 12.0, y, 0.75, circle=BLUE)
    text(s, 13.0, y - 0.03, 5.9, 0.45, a, size=20, color=INK)
    text(s, 13.0, y + 0.43, 5.9, 0.9, b, size=16, color=MUTED_L, line=1.05)
notes(s, "Demo the rollback: run UC5 with 'inject 1 invalid record' inside a changeset - zero rows written. Same batch without "
         "a changeset and with Prefer: odata.continue-on-error - 99 rows and one error. Be honest about entitlement: bulk saves "
         "time and throttling, but each record still counts as a Power Platform request.")

# ------------------------------------------------------------------ 21 · Section 05
divider(S[20], '05', 'Filter at the source', 'OData filters, $select, $expand and trigger conditions',
        "Section 5. The cheapest data is the data you never download.")

# ------------------------------------------------------------------ 22 · Funnel (light spotlight)
s = S[21]
kicker_title(s, 'Get less', 'Ask the source for less', dark=False, centered=True)
demo(s, 'LIVE · UC6 filtering')
layers = [('Get items: 20,000 rows × 25 columns', 13.0, 'C9CDF0', INK),
          ("$filter   Status eq 'Open' and Amount gt 1000", 11.4, 'A9B3F5', INK),
          ('$select   Title, Amount, Customer', 9.8, '7F8CF0', 'FFFFFF'),
          ('$expand   Customer/Title: no Get item per row', 8.2, BLUE, 'FFFFFF'),
          ('Only the rows and columns you use', 6.6, VIOLET, 'FFFFFF')]
for i, (lab, w, fill, fg) in enumerate(layers):
    y = 3.95 + i * 1.18
    x = (W - w) / 2
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, 1.0, fill=fill, radius=0.2)
    text(s, x, y, w, 1.0, lab, size=19 if i else 20, color=fg, align='c', anchor='m', font=CODE if 0 < i < 4 else FONT)
text(s, 3.5, 9.95, 13.0, 0.5, 'Measured: 59.6 s → 1.1 s and 5.7 MB → 11 KB (53×). Dataverse FetchXML aggregate: 10.1 MB → 1 KB (36×)', size=17, color=MUTED_L, align='c')
notes(s, "Demo UC6. SLOW: Get items with pagination at 100,000 and all columns, then Filter array and a Get item per row for the "
         "customer. FAST: Filter Query + $select + $expand on the lookup + $top. Show the payload size KPI in the app. "
         "Mention indexed columns for lists over 5,000 items and the Dataverse Row count option.")

# ------------------------------------------------------------------ 23 · Trigger conditions (dark)
s = S[22]
kicker_title(s, 'Trigger conditions', 'The fastest run never starts', dark=True)
demo(s, 'LIVE · UC7 triggers')
for i, (lab, num, sub, col, n_on) in enumerate([
        ('Condition inside the flow', '100 runs', '90 of them only check and exit', SLOW, 100),
        ('Trigger condition', '10 runs', 'the other 90 never start', FAST_D, 10)]):
    y = 3.85 + i * 3.1
    card(s, 1.12, y, 8.6, 2.8, dark=True)
    text(s, 1.5, y + 0.25, 7.9, 0.45, lab, size=19, color=MUTED_D)
    text(s, 1.5, y + 0.7, 4.0, 0.9, num, size=46, color=col)
    text(s, 1.5, y + 1.7, 7.9, 0.45, sub, size=17, color=LAV)
    for k in range(20):
        on = k < n_on // 5
        shape(s, MSO_SHAPE.OVAL, 5.6 + (k % 10) * 0.37, y + 0.85 + (k // 10) * 0.37, 0.26, 0.26, fill=col if on else '24347E')
code(s, 10.4, 3.85, 6.9, 2.35, [
    '// Settings › Trigger conditions',
    "@equals(triggerOutputs()?",
    "  ['body/Status/Value'],",
    "  'Approved')",
], size=18)
text(s, 10.4, 6.6, 6.9, 2.8, [dict(runs='Dataverse triggers: use Filter rows and Select columns as well', bullet=True, bullet_color=CYAN, space_after=10),
                                dict(runs='No run means no billed actions and no history noise', bullet=True, bullet_color=CYAN)],
     size=19, color=LAV)
text(s, 1.12, 10.1, 9, 0.4, 'Measured: 100 item edits, 1 in 10 relevant. Each dot = 5 runs', size=14, color=MUTED_D)
notes(s, "Demo UC7. The app launches a driver flow that edits 100 SharePoint items; only 10 become 'Approved'. "
         "The flow without a trigger condition runs 100 times; the one with the condition runs 10 times. Show the run count in the app.")

# ------------------------------------------------------------------ 24 · Section 06
divider(S[23], '06', 'Tune the platform', 'Concurrency, pagination, Compose vs variables, and the traps',
        "Section 6. Settings that help, and the ones that bite.")

# ------------------------------------------------------------------ 25 · Knobs (light)
s = S[24]
kicker_title(s, 'Settings', 'Six settings worth a second look', dark=False)
knobs = [('random', 'Concurrency', 'Default 1. Raise to 20–50 for independent per-item work.'),
         ('pages', 'Pagination', 'Off by default. Set a threshold that matches your data.'),
         ('cube', 'Compose vs variables', 'Compose for values set once. Variables lock inside concurrent loops.'),
         ('stream', 'Parallel branches', 'Independent calls side by side: total time = the slowest branch.'),
         ('redo', 'Retry policy', 'Honour 429 and Retry-After. Keep batch calls idempotent.'),
         ('bolt', 'Respond early', 'A late response fails at 120 s and marks the run Failed. Answer first.')]
kw, kh, kgx, kgy = 5.6, 2.55, 0.48, 0.38
for i, (ic, a, b) in enumerate(knobs):
    x = 1.12 + (i % 3) * (kw + kgx); y = 3.75 + (i // 3) * (kh + kgy)
    card(s, x, y, kw, kh)
    icon(s, ic, x + 0.35, y + 0.35, 0.95, circle=BLUE if i % 2 == 0 else VIOLET)
    text(s, x + 1.55, y + 0.5, kw - 1.85, 0.6, a, size=23, color=INK)
    text(s, x + 0.35, y + 1.45, kw - 0.7, 1.0, b, size=17, color=MUTED_L, line=1.05)
notes(s, "Quick tour of the knobs. Concurrency and pagination are the two most common surprises. Respond early is the pattern "
         "the Lab app itself uses: the flow writes a telemetry row and the app polls it, so long runs never hit the 120 s timeout.")

# ------------------------------------------------------------------ 26 · Concurrency bites (dark)
s = S[25]
kicker_title(s, 'The trap', 'When concurrency bites', dark=True)
demo(s, 'LIVE · UC8 race')
card(s, 1.12, 3.85, 8.8, 6.0, dark=True)
text(s, 1.5, 4.1, 8.0, 0.5, 'Apply to each, concurrency 20', size=20, color=CYAN)
for j, (a, b) in enumerate([('Set variable', "Current = item()"), ('Compose', "uses variables('Current')")]):
    y = 4.8 + j * 0.95
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 1.5, y, 8.0, 0.75, fill='1C2F86', line=CARD_D_LINE, radius=0.15)
    text(s, 1.8, y, 3.0, 0.75, a, size=18, color=LAV, anchor='m')
    text(s, 4.6, y, 4.8, 0.75, b, size=16, color=MUTED_D, anchor='m', font=CODE)
for j, (lab, vals, col) in enumerate([('Expected', ['A', 'B', 'C', 'D'], FAST_D), ('Got', ['C', 'C', 'D', 'D'], SLOW)]):
    y = 7.0 + j * 1.25
    text(s, 1.5, y, 2.0, 0.8, lab, size=19, color=MUTED_D, anchor='m')
    for k, v in enumerate(vals):
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 3.5 + k * 1.05, y, 0.85, 0.8, fill=col, radius=0.18)
        text(s, 3.5 + k * 1.05, y, 0.85, 0.8, v, size=24, color=NAVY, align='c', anchor='m')
text(s, 1.12, 10.05, 9.0, 0.45, 'Measured checksum: SLOW 100 | 12,163.15 vs FAST 1,000 | 499,995.00', size=15, color=MUTED_D)
fixes = [('Keep state per iteration', 'Use item() or a Compose inside the loop, never a shared variable'),
         ('No side effect? No loop', 'A Select does the same job in one action'),
         ('Respect the target', 'SharePoint: 600 calls/min. Dataverse: 52 concurrent requests'),
         ('Order matters?', 'Keep concurrency at 1, or sort afterwards')]
for i, (a, b) in enumerate(fixes):
    y = 3.85 + i * 1.5
    icon(s, 'check', 10.6, y + 0.02, 0.7, circle=FAST, glyph='w')
    text(s, 11.55, y - 0.03, 5.8, 0.45, a, size=21, color=LAV)
    text(s, 11.55, y + 0.45, 5.8, 0.9, b, size=17, color=MUTED_D, line=1.05)
notes(s, "Demo UC8: the concurrent loop with Set variable returns wrong values; same loop with item()/Compose is correct. "
         "Then the pagination trap: Get items without pagination silently returns 100 items - a correctness bug, not just speed.")

# ------------------------------------------------------------------ 27 · Scoreboard (light spotlight, measured)
s = S[26]
kicker_title(s, 'Scoreboard', 'Every pair, measured', dark=False, centered=True)
score = [('UC3 Aggregation (100)', 177), ('UC1 Data shaping (500)', 130), ('UC2 Join (200)', 124), ('UC6 Filter at source (2,000)', 53),
         ('UC9 Excel (500)', 43), ('UC6 Dataverse aggregate (5,000)', 36), ('UC5 Dataverse bulk (1,000)', 32),
         ('UC4 SharePoint bulk (1,000)', 15), ('UC10 Fan-out (200)', 12)]
cd = CategoryChartData()
cd.categories = [c for c, _ in score]
cd.add_series('Speed-up', [v for _, v in score])
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(3.2), Inches(3.7), Inches(13.6), Inches(5.9), cd)
gf.name = 'Speed-up chart'
ch = gf.chart
ch.has_legend = False
ch.has_title = False
ch.font.name = FONT; ch.font.size = Pt(16); ch.font.color.rgb = rgb(INK)
plot = ch.plots[0]; plot.gap_width = 45
plot.has_data_labels = True
dl = plot.data_labels; dl.number_format = '0"×"'; dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.OUTSIDE_END; dl.font.size = Pt(16); dl.font.color.rgb = rgb(INK)
ser = plot.series[0]; ser.format.fill.solid(); ser.format.fill.fore_color.rgb = rgb(BLUE)
ca = ch.category_axis; ca.reverse_order = True; ca.tick_labels.font.size = Pt(16); ca.format.line.color.rgb = rgb('C9CDF0')
ca.has_major_gridlines = False
va = ch.value_axis; va.visible = False; va.has_major_gridlines = False; va.minimum_scale = 0; va.maximum_scale = 200
text(s, 3.2, 9.75, 13.6, 0.6, 'SLOW ÷ FAST duration, same checksum for every pair. PREMIUM-US demo environment, November 2026. Volume in brackets.',
     size=15, color=MUTED_L, align='c')
notes(s, "Every pair measured in the lab (lab/RESULTS.md), and every pair returned the same checksum. "
         "Switch to the scoreboard in the Lab app to show today's live runs next to these.")

# ------------------------------------------------------------------ 28 · Takeaways (light)
s = S[27]
kicker_title(s, 'Takeaways', 'Five things to do tomorrow morning', dark=False)
tks = [('Find the action that owns the run', 'Run history, action timings, flowrun table'),
       ('Replace data-shaping loops', 'Select, Filter array, dictionary joins, xpath'),
       ('Batch your writes', '$batch ≤ 500 per call, CreateMultiple for Dataverse'),
       ('Ask the source for less', '$filter, $select, $expand, trigger conditions'),
       ('Tune deliberately', 'Concurrency, pagination, Compose over variables')]
for i, (a, b) in enumerate(tks):
    y = 3.75 + i * 1.18
    shape(s, MSO_SHAPE.OVAL, 1.12, y + 0.05, 0.85, 0.85, fill=BLUE if i % 2 == 0 else VIOLET)
    text(s, 1.12, y + 0.05, 0.85, 0.85, str(i + 1), size=26, color='FFFFFF', align='c', anchor='m')
    text(s, 2.3, y, 9.5, 0.5, a, size=25, color=INK)
    text(s, 2.3, y + 0.52, 9.5, 0.45, b, size=18, color=MUTED_L)
card(s, 12.6, 3.85, 6.0, 5.1)
icon(s, 'check', 13.0, 4.25, 1.1, circle=FAST)
text(s, 13.0, 5.6, 5.2, 1.4, 'Same output, a fraction of the actions', size=28, color=INK, line=1.05)
text(s, 13.0, 7.2, 5.2, 1.5, 'Faster runs, lower request consumption, fewer throttling errors, flows your agents can rely on.',
     size=17, color=MUTED_L, line=1.1)
notes(s, "Recap the five moves. Encourage people to pick their slowest flow tomorrow and apply just step 1 and step 2.")

# ------------------------------------------------------------------ 29 · Resources (dark)
s = S[28]
kicker_title(s, 'Take it home', 'Everything from today, ready to reuse', dark=True)
res = [('github', 'The Lab', ['All SLOW and FAST flows', 'The Code App and setup scripts', 'github.com/david-lorenzo88/', 'power-automate-optimization-techniques']),
       ('book', 'Microsoft Learn', ['Power Automate limits and configuration', 'Dataverse bulk operation messages', 'Execute batch operations (Web API)', 'Cloud flow run history in Dataverse']),
       ('flag', 'Start tomorrow', ['Pick your slowest flow', 'Measure: run history and flowrun', 'Rebuild one loop', 'Compare the before and after'])]
rw, rgap = 5.05, 0.45
for i, (ic, head, lines) in enumerate(res):
    x = 1.12 + i * (rw + rgap)
    card(s, x, 3.85, rw, 5.6, dark=True)
    icon(s, ic, x + 0.4, 4.2, 1.0, circle=BLUE if i != 2 else VIOLET)
    text(s, x + 0.4, 5.45, rw - 0.8, 0.6, head, size=26, color=LAV)
    paras = [dict(runs=l, bullet=(i != 0 or j < 2), bullet_color=CYAN, space_after=8,
                  color=(CYAN if (i == 0 and j >= 2) else LAV), size=(15 if (i == 0 and j >= 2) else 18)) for j, l in enumerate(lines)]
    text(s, x + 0.4, 6.25, rw - 0.7, 3.0, paras, size=18, color=LAV)
notes(s, "Point to the repo for the flows, the Code App and the scripts. Make sure the repository is public before the session.")

# ------------------------------------------------------------------ 30 · Thank you
s = S[29]
text(s, 1.55, 7.45, 9.0, 0.8, 'Questions?', size=40, color=CYAN)
notes(s, "Thank the audience. Leave the Lab scoreboard open on screen during Q&A.")

# ------------------------------------------------------------------ 3 (new) · Measured, not promised (light spotlight)
s = NEW_MEASURED
kicker_title(s, 'Measured, not promised', 'Same result, up to 177× faster', dark=False, centered=True)
top = [('177×', 'Aggregation per region', '4.2 min → 1.4 s · 100 orders'),
       ('130×', 'Data shaping', '113.9 s → 0.9 s · 500 orders'),
       ('124×', 'Join / enrichment', '3.0 min → 1.5 s · 200 orders')]
mw, mg = 4.3, 0.5
mx0 = (W - (3 * mw + 2 * mg)) / 2
for i, (num, lab, sub) in enumerate(top):
    x = mx0 + i * (mw + mg)
    card(s, x, 4.3, mw, 4.3, radius=0.08)
    text(s, x + 0.45, 4.65, mw - 0.9, 1.4, num, size=80, color=BLUE)
    text(s, x + 0.45, 6.25, mw - 0.9, 0.9, lab, size=22, color=INK)
    text(s, x + 0.45, 7.55, mw - 0.9, 0.7, sub, size=17, color=MUTED_L)
text(s, 3.5, 9.05, 13.0, 0.6, 'PREMIUM-US demo environment, November 2026. Every SLOW/FAST pair returns an identical checksum.',
     size=17, color=MUTED_L, align='c')
notes(s, "Not a promise, a measurement: three of the pairs we will run today, measured in our demo tenant last week. "
         "Same input, same checksum, only the technique changes. The full scoreboard comes at the end.")

# ------------------------------------------------------------------ 28 (new) · UC10 + 120-second rule (light)
s = NEW_UC10
kicker_title(s, 'Fan-out and the 120-second rule', 'Parallelize the work, answer the caller first', dark=False)
demo(s, 'LIVE · UC10 fan-out')
card(s, 1.12, 3.8, 8.6, 5.3)
text(s, 1.55, 4.15, 7.8, 0.5, 'Fan-out to child flows', size=20, color=BLUE)
text(s, 1.55, 4.75, 7.8, 1.2, [[('76.0 s', {'color': SLOW}), (' → ', {'color': MUTED_L}), ('6.2 s', {'color': FAST})]], size=60)
text(s, 1.55, 6.25, 7.8, 1.5, '200 records: one sequential loop vs chunk() into 10 chunks, Apply to each with concurrency 10, one child flow per chunk.',
     size=18, color=INK, line=1.1)
text(s, 1.55, 7.9, 7.8, 1.0, 'Fan-out saves wall-clock time, not actions: the work moves into the child flows.', size=16, color=MUTED_L, line=1.1)
card(s, 10.3, 3.8, 8.6, 5.3, fill=NAVY, line=NAVY)
text(s, 10.75, 4.15, 7.8, 0.5, 'Respond early', size=20, color=CYAN)
text(s, 10.75, 4.75, 7.8, 1.2, [[('120 s', {'color': 'FF8A8E'}), (' → ', {'color': MUTED_D}), ('~1 s', {'color': FAST_D})]], size=60)
text(s, 10.75, 6.25, 7.8, 1.5, 'Respond at the end of a 150 s flow: Run() fails at 120 s, and the run shows Failed even though the work finished.',
     size=18, color=LAV, line=1.1)
text(s, 10.75, 7.9, 7.8, 1.0, 'Respond first, then work, and report progress through Dataverse.', size=16, color=MUTED_D, line=1.1)
notes(s, "Demo UC10: the same 200 records, sequential vs fan-out to child flows. Then the 120-second rule we hit while building the Lab app: "
         "a flow called from Power Apps must answer within 120 s, so the Lab flows respond in about a second and keep working.")

prs.save(out)
print('saved', out)
