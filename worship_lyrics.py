from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.slide import SlideLayout
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor


def read_data(filetext: str) -> tuple:
    filetext = filetext.replace("\r\n", "\n")
    data = tuple(tuple(s.strip().split("\n")) for s in filetext.split("\n\n"))
    return data


def validate_data(data: tuple):
    if len(data[0]) != 4:
        raise TypeError(
            "First data (name, copyright, image_path, text_color) must be 4")
    return True


def initialize_pptx() -> Presentation:
    pre = Presentation()
    pre.slide_height = Inches(9)
    pre.slide_width = Inches(16)

    return pre


BLANK_SLIDE_LAYOUT = 6  # Blank
FONT = "微軟正黑體"


def add_slide(prs, lyrics, header_data):
    slide = prs.slides.add_slide(prs.slide_layouts[BLANK_SLIDE_LAYOUT])
    shapes = slide.shapes
    color = RGBColor.from_string(header_data[3])
    # Add background image
    shapes.add_picture(
        header_data[2], 0, 0,
        height=prs.slide_height, width=prs.slide_width
    )
    # Add lyrics text
    if lyrics[0] != "#EMPTY!":  # Use "#EMPTY!" to skip a slide
        text_shape = shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(
                1.5), Inches(15), Inches(6)
        )
        text_shape.fill.background()
        text_shape.line.fill.background()
        text_frame = text_shape.text_frame
        text_frame.clear()
        text_frame.vertical_anchor = MSO_ANCHOR.TOP
        text_paragraph = text_frame.paragraphs[0]
        text_paragraph.text = "\n".join(lyrics)
        text_paragraph.alignment = PP_ALIGN.CENTER
        text_paragraph.font.size = Pt(60)
        text_paragraph.font.name = FONT
        text_paragraph.font.bold = True
        text_paragraph.font.color.rgb = color
    # Add copyright text
    cpr_shape = shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(6.5), Inches(10), Inches(2)
    )
    cpr_shape.fill.background()
    cpr_shape.line.fill.background()
    cpr_frame = cpr_shape.text_frame
    cpr_frame.clear()
    cpr_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
    cpr_paragraph = cpr_frame.paragraphs[0]
    cpr_paragraph.text = "\n".join(header_data[:2])
    cpr_paragraph.alignment = PP_ALIGN.LEFT
    cpr_paragraph.font.size = Pt(24)
    cpr_paragraph.font.name = FONT
    cpr_paragraph.font.bold = True
    cpr_paragraph.font.color.rgb = color


def workflow(filename, savefilename):
    with open(filename, mode='rb') as f:
        b = f.read()
        filetext = str(b, encoding='utf-8')
    data = read_data(filetext)
    validate_data(data)
    prs = initialize_pptx()
    for lyrics in data[1:]:
        add_slide(prs, lyrics, data[0])
    prs.save(savefilename)


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='Create worship lyrics PPTX')
    parser.add_argument('filename')
    parser.add_argument('savefilename')
    args = parser.parse_args()

    workflow(args.filename, args.savefilename)

    from sys import exit
    exit(0)
