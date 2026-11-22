import sys
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
import lib
from lib import *

base, out, lib.ICON_DIR = sys.argv[1], sys.argv[2], sys.argv[3]
prs = Presentation(base)
s = prs.slides[0]
kicker_title(s, 'Take it home', 'These practices, as an agent skill', dark=False)
pill(s, 'GitHub Copilot + Claude Code', icon_key='github')

steps = [('robot', 'You describe the flow', 'In GitHub Copilot or Claude Code: "Build a flow that…"'),
         ('bolt', 'power-automate plugin (Microsoft)', 'Builds and deploys with FlowAgent: exact connectors and connections'),
         ('sliders', 'flow-best-practices skill', 'Decision framework, 9 rules, patterns measured in our lab'),
         ('check', 'Lint, then validate', 'lint-flow.mjs runs 16 checks before validate_flow')]
for i, (ic, a, b) in enumerate(steps):
    y = 3.8 + i * 1.42
    icon(s, ic, 1.12, y, 0.95, circle=BLUE if i != 2 else VIOLET)
    text(s, 2.4, y - 0.02, 7.6, 0.5, a, size=23, color=INK)
    text(s, 2.4, y + 0.5, 7.6, 0.5, b, size=17, color=MUTED_L)
    if i < 3:
        arrow(s, 1.595, y + 1.0, 1.595, y + 1.38, color=BLUE, width=2)

card(s, 10.6, 3.8, 8.3, 2.75)
text(s, 11.0, 4.05, 7.5, 0.5, 'Two skills', size=17, color=BLUE)
text(s, 11.0, 4.55, 7.5, 0.5, 'flow-best-practices', size=22, color=INK, font=CODE)
text(s, 11.0, 5.05, 7.5, 0.45, 'Loads automatically whenever a flow is built or edited', size=16, color=MUTED_L)
text(s, 11.0, 5.6, 7.5, 0.5, '/optimize-flow <flow>', size=22, color=INK, font=CODE)
text(s, 11.0, 6.1, 7.5, 0.45, 'Audits a flow, proposes fixes, measures before and after', size=16, color=MUTED_L)
code(s, 10.6, 6.85, 8.3, 2.15, [
    '// Install',
    '/plugin marketplace add',
    '  david-lorenzo88/power-automate-optimization-techniques',
    '/plugin install',
    '  power-automate-best-practices@power-automate-optimization',
], size=15, pad=0.3)
text(s, 1.12, 9.65, 10.5, 0.5, 'Every SLOW lab flow is flagged, every FAST flow passes. MIT licensed.', size=17, color=MUTED_L)
notes(s, "Everything from this session is packaged as an agent skill for GitHub Copilot and Claude Code. It works next to "
         "Microsoft's power-automate plugin from microsoft/power-platform-skills: that plugin makes a flow valid, this one makes it "
         "efficient. flow-best-practices loads automatically whenever the assistant builds or edits a flow: the decision framework, "
         "the rules and the patterns we measured today. Before validate_flow, it runs lint-flow.mjs, 16 checks, and we tested it "
         "against our lab: every SLOW flow is flagged with its anti-pattern, every FAST flow passes. /optimize-flow audits an existing "
         "flow and proposes the rewrites. Install it with the two commands on the slide; the repo is MIT licensed.")
prs.save(out)
print('saved')
