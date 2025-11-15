import pptx
import pptx.presentation
import pptx.slide
from pptx.util import Inches, Pt
from pptx.table import Table
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


class Component:
    def __init__(self):
        pass

    def render(self, slide: pptx.slide.Slide):
        raise NotImplementedError


class SRAM(Component):
    def __init__(
        self,
        row: int,
        col: int,
        left: Inches,
        top: Inches,
        width: Inches,
        height: Inches,
        interleave: bool = False,
        bottom_start: bool = True,
    ):
        self.row = row
        self.col = col
        self.left = left
        self.top = top
        self.width = width
        self.height = height
        self.interleave = interleave  # 是否啟用帶狀欄
        self.bottom_start = bottom_start  # 座標是否從底部開始計算

        # 將 data store 改為儲存 (text, color) 的元組 (tuple)
        self.data = [[(None, None) for _ in range(col)] for _ in range(row)]

    def get_position(self, i: int, j: int):
        """
        獲取 (i, j) 儲存格的左上角 (x, y) 座標。
        i, j 皆為 0-based 索引 (從 top-left 開始)。
        - i: row 索引 (0 to self.row - 1)
        - j: col 索引 (0 to self.col - 1)

        會根據 self.bottom_start 決定 y 座標的對應方式。
        """
        x = self.left + j * self.width / self.col

        cell_height = self.height / self.row
        if self.bottom_start:
            # bottom-up 映射: data[0] (頂部資料) 對應到
            # 視覺上的最後一列 (visual_row_index = self.row - 1)
            visual_row_index = (self.row - 1) - i
            y = self.top + (visual_row_index * cell_height)
        else:
            # top-down 映射: data[0] 對應到 視覺上的第一列 (i=0)
            y = self.top + i * cell_height

        return x, y

    def render(self, slide: pptx.slide.Slide) -> Table:
        """
        將这个 SRAM 表格新增到指定的投影片上
        """
        self.slide = slide  # 儲存 slide 物件
        shapes = slide.shapes
        table = shapes.add_table(
            self.row,
            self.col,
            self.left,
            self.top,
            self.width,
            self.height,
        ).table

        # 根據 self.interleave 啟用帶狀欄
        if self.interleave:
            table.vert_banding = True  # 啟用帶狀欄 (Banded Columns)
            table.horz_banding = False
        else:
            table.vert_banding = False
            table.horz_banding = False

        col_width = self.width // self.col
        for i in range(self.col):
            table.columns[i].width = col_width

        row_height = self.height // self.row
        for i in range(self.row):
            table.rows[i].height = row_height

        self.render_data(slide)

    def write(self, i: int, j: int, text: str, color: RGBColor = None):
        """
        將資料 (text, color) 寫入內部的 self.data 暫存
        i, j 為 0-based 索引 (從 top-left 開始)
        """
        if 0 <= i < self.row and 0 <= j < self.col:
            self.data[i][j] = (text, color)
        else:
            print(f"SRAM write 錯誤: 索引 ({i}, {j}) 超出範圍")

    def render_data(self, slide: pptx.slide.Slide):
        """
        遍歷 self.data，將所有資料繪製為文字方塊
        """
        if slide is None:
            print("SRAM 錯誤: 必須先呼叫 .add(slide) 才能 render_data")
            return

        cell_width = self.width / self.col
        cell_height = self.height / self.row

        for i in range(self.row):
            for j in range(self.col):
                text, color = self.data[i][j]

                # 如果 data 不是 None 或空字串
                if text:
                    x, y = self.get_position(i, j)

                    tb = slide.shapes.add_textbox(x, y, cell_width, cell_height)
                    tf = tb.text_frame

                    # 清除邊界並置中
                    tf.margin_bottom = Inches(0)
                    tf.margin_top = Inches(0)
                    tf.margin_left = Inches(0)
                    tf.margin_right = Inches(0)
                    tf.word_wrap = False  # 避免自動換行

                    p = tf.paragraphs[0]
                    p.alignment = PP_ALIGN.CENTER  # 水平置中
                    # (垂直置中是 tf.vertical_anchor)

                    p.text = text
                    p.font.name = "Tahoma"  # 設定字體
                    p.font.size = Pt(12)  # 設定字體大小

                    if color:
                        p.font.color.rgb = color
