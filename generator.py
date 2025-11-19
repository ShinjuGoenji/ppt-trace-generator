from pptx import Presentation
import pptx
import pptx.presentation
from pptx.util import Cm

from pptx.dml.color import RGBColor
from componets import SRAM, PE, FLAG
from util import COLORS


# --- parameters --- #
OUTPUT_FILE = "generated_trace.pptx"
TITLE = "CRT Trace"

# Componenets
sram = SRAM(
    row=7,
    col=32,
    left=Cm(6),
    top=Cm(3),
    width=Cm(21),
    height=Cm(7),
    interleave=True,
)

add_mul_small = PE(
    name="ADD_MUL_SMALL",
    left=Cm(12),
    top=Cm(12),
    width=Cm(4),
    height=Cm(4),
    PE_num=4,
)

mod_small_unsigned = PE(
    name="MOD_SMALL_UNSIGNED",
    left=Cm(18),
    top=Cm(12),
    width=Cm(4),
    height=Cm(4),
    PE_num=4,
)

r_flag_0 = FLAG(
    name="Read", color=RGBColor(178, 178, 190), flag=False, left=Cm(2), top=Cm(3)
)

w_flag_0 = FLAG(
    name="Write", color=RGBColor(178, 178, 190), flag=False, left=Cm(2), top=Cm(4.5)
)

r_flag_1 = FLAG(
    name="Read", color=RGBColor(227, 229, 237), flag=False, left=Cm(2), top=Cm(6)
)

w_flag_1 = FLAG(
    name="Write", color=RGBColor(227, 229, 237), flag=False, left=Cm(2), top=Cm(7.5)
)

write_buffer = SRAM(
    row=3,
    col=2,
    left=Cm(6.5),
    top=Cm(12),
    width=Cm(3),
    height=Cm(1),
    interleave=True,
    bottom_start=False,
)

components = [
    sram,
    add_mul_small,
    mod_small_unsigned,
    r_flag_0,
    r_flag_1,
    w_flag_0,
    w_flag_1,
    write_buffer,
]


# --- static elements --- #


# --- MAIN LOGIC --- #
def main_logic(prs: pptx.presentation.Presentation):
    inputs = {}
    for i in range(32 * 4):
        inputs[i] = (f"x{int(i/32)}_{i%32}", COLORS["BLUE"])
    input_cnt = 0

    mod_small_unsigned_ptr = 0
    add_mul_small_ptr = 32
    write_buffer_ptr = [0, 0]

    MAX_CYCLE = 46
    for cycle in range(MAX_CYCLE):

        input_data, input_color = inputs[input_cnt]
        u = int(input_cnt / 32)
        v = input_cnt % 32

        w_flag_0.flag = False
        w_flag_1.flag = False
        r_flag_0.flag = False
        r_flag_1.flag = False
        add_mul_small.count()
        mod_small_unsigned.count()

        sram_write = []
        mod_small_unsigned_write = []
        write_buffer_write = []

        # inputs
        if (
            mod_small_unsigned.ready()
            and mod_small_unsigned_ptr < input_cnt
            and mod_small_unsigned_ptr < add_mul_small_ptr
        ):
            read_data, read_color = sram.read(
                int(mod_small_unsigned_ptr / 32), mod_small_unsigned_ptr % 32
            )
            mod_small_unsigned_write.append(
                (
                    read_data,
                    (1 + 3 * (int(mod_small_unsigned_ptr / 32) + 1) + 1),
                    read_color,
                )
            )
            if mod_small_unsigned_ptr % 2 == 0:
                r_flag_0.flag = True
            else:
                r_flag_1.flag = True

        if input_cnt < len(inputs):
            if u == 0:
                if mod_small_unsigned.ready() and mod_small_unsigned_ptr == input_cnt:
                    mod_small_unsigned_write.append(
                        (
                            input_data,
                            5,
                            input_color,
                        )
                    )
                else:
                    sram_write.append((u, v, input_data, input_color))
                    if v % 2 == 0:
                        w_flag_0.flag = True
                    else:
                        w_flag_1.flag = True
            else:
                if add_mul_small.ready() and add_mul_small_ptr == input_cnt:
                    add_mul_small.write(
                        text=input_data, color=input_color, cnt=(3 + u + 1)
                    )
                else:
                    sram_write.append((u, v, input_data, input_color))
                    if v % 2 == 0:
                        w_flag_0.flag = True
                    else:
                        w_flag_1.flag = True

        if w_flag_0.flag == False and write_buffer_ptr[0] > 0:
            for k in range(write_buffer_ptr[0]):
                text, color = write_buffer.read(k, 0)
                if k == 0:
                    i = int(text.replace("x", "").split("_")[0])
                    j = int(text.split("_")[1])
                    sram_write.append((i, j, text, color))
                    w_flag_0.flag = True
                else:
                    write_buffer_write.append((k - 1, 0, text, color))
            write_buffer_ptr[0] -= 1
        if w_flag_1.flag == False and write_buffer_ptr[1] > 0:
            for k in range(write_buffer_ptr[1]):
                text, color = write_buffer.read(k, 1)
                if k == 0:
                    i = int(text.replace("x", "").split("_")[0])
                    j = int(text.split("_")[1])
                    sram_write.append((i, j, text, color))
                    w_flag_1.flag = True
                else:
                    write_buffer_write.append((k - 1, 1, text, color))
            write_buffer_ptr[1] -= 1

        for k, data in enumerate(mod_small_unsigned.data):
            text = data[0]

            if text == "" or mod_small_unsigned.cnt[k] > 0:
                continue

            i = int(text.replace("x", "").split("_")[0])
            j = int(text.split("_")[1])
            if j % 2 == 0:
                if w_flag_0.flag:
                    write_buffer_write.append(
                        (write_buffer_ptr[j % 2], j % 2, text, COLORS["YELLOW"])
                    )
                    write_buffer_ptr[j % 2] += 1
                else:
                    sram_write.append((i, j, text, COLORS["YELLOW"]))
                    w_flag_0.flag = True
            else:
                if w_flag_1.flag:
                    write_buffer_write.append(
                        (write_buffer_ptr[j % 2], j % 2, text, COLORS["YELLOW"])
                    )
                    write_buffer_ptr[j % 2] += 1
                else:
                    sram_write.append((i, j, text, COLORS["YELLOW"]))
                    w_flag_1.flag = True
            mod_small_unsigned.data[k] = ("", RGBColor(0, 0, 0))

        for k in sram_write:
            sram.write(*k)
        for k in mod_small_unsigned_write:
            mod_small_unsigned.write(*k)
            mod_small_unsigned_ptr += 1
        for k in write_buffer_write:
            write_buffer.write(*k)

        # counter
        input_cnt += 1

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
    trace_layout = prs.slide_layouts[3]
    slide = prs.slides.add_slide(trace_layout)

    for component in components:
        component.render(slide)


if __name__ == "__main__":
    prs = init_ppt()
    main_logic(prs)
    prs.save(OUTPUT_FILE)
