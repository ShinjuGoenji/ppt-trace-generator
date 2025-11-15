from pptx import Presentation
import pptx
import pptx.presentation
import pptx.slide
from pptx.util import Inches, Pt

from pptx.dml.color import RGBColor
from componets import Component, SRAM


# --- parameters --- #
OUTPUT_FILE = "generated_trace.pptx"
TITLE = "CRT Trace"

# Componenets
components = []
sram = SRAM(
    row=4,
    col=32,
    left=Inches(0.5),
    top=Inches(1),
    width=Inches(9),
    height=Inches(1),
    interleave=True,
)
components.append(sram)


# --- static elements --- #


# --- MAIN LOGIC --- #
def main_logic(prs: pptx.presentation.Presentation):

    MAX_CYCLE = 10
    for cycle in range(MAX_CYCLE):
        sram.write(
            int(cycle / 32),
            cycle % 32,
            f"x{cycle}",
            color=RGBColor(
                255 * (cycle % 3 == 0),
                255 * (cycle % 3 == 1),
                255 * (cycle % 3 == 2),
            ),
        )

        render(prs)


# --- PPT settings --- #
def init_ppt():
    TEMPLATE_FILE = "template.pptx"
    prs = Presentation(TEMPLATE_FILE)

    # Add title slide
    title_layout = prs.slide_layouts[0]
    title_slide = prs.slides.add_slide(title_layout)
    title = title_slide.shapes.title
    subtitle = title_slide.placeholders[1]

    title.text = TITLE
    subtitle.text = "Student: Chao-En Kuo\nAdvisor: Hsie-Chia Chang"

    return prs


def render(prs: pptx.presentation.Presentation):
    trace_layout = prs.slide_layouts[2]
    slide = prs.slides.add_slide(trace_layout)

    for component in components:
        component.render(slide)


if __name__ == "__main__":
    prs = init_ppt()
    main_logic(prs)
    prs.save(OUTPUT_FILE)
