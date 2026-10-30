---
max_turns: 30
timeout_seconds: 900
allowed_tools: [Skill, Bash, Read, Write, Edit, Glob, Grep]
tags: [edge]
---

First run this to create my current deck: `python3 -c "from pptx import Presentation; from pptx.util import Inches; p=Presentation(); s=p.slides.add_slide(p.slide_layouts[5]); s.shapes.title.text='Three priorities'; [s.shapes.add_textbox(Inches(0.5+3.1*i), Inches(2), Inches(2.8), Inches(2)).text_frame.__setattr__('text', t) for i, t in enumerate(['Hire two engineers', 'Ship the mobile app', 'Open the Lyon office'])]; p.save('plain.pptx')"`

Then give plain.pptx a more organic style.
