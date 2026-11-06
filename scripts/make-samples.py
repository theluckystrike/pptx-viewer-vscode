"""Generate the sample decks used by the tests and the README screenshots.

Run: python3 scripts/make-samples.py   (needs python-pptx and Pillow)
Writes test/fixtures/*.pptx. Each file is well under 1 MB.
"""
import io
import pathlib
import zipfile

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

OUT = pathlib.Path(__file__).resolve().parents[1] / "test" / "fixtures"
OUT.mkdir(parents=True, exist_ok=True)


def photo_png():
    """A small landscape picture: sky, sun, hills."""
    img = Image.new("RGB", (480, 300), (110, 170, 230))
    d = ImageDraw.Draw(img)
    d.ellipse((360, 30, 430, 100), fill=(255, 215, 80))
    d.polygon([(0, 220), (140, 120), (300, 230), (480, 140), (480, 300), (0, 300)], fill=(60, 140, 80))
    d.polygon([(0, 260), (200, 200), (480, 270), (480, 300), (0, 300)], fill=(40, 110, 60))
    buf = io.BytesIO()
    img.save(buf, "PNG")
    buf.seek(0)
    return buf


def quarterly_review():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = "Q3 Platform Review"
    s.placeholders[1].text = "Latency, uptime and the plan for Q4"
    s.notes_slide.notes_text_frame.text = "Welcome everyone. Keep the intro under two minutes."

    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = "Summary"
    tf = s.placeholders[1].text_frame
    tf.text = "p95 latency fell from 420 ms to 180 ms"
    for line, lvl in [("Cache hit rate is now 93%", 1), ("Uptime was 99.97% across all regions", 0),
                      ("Two incidents, both under 15 minutes", 0), ("Cost per request fell 22%", 0)]:
        p = tf.add_paragraph()
        p.text = line
        p.level = lvl
    s.notes_slide.notes_text_frame.text = "Stress the latency win first. The cache work is what made it possible."

    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Latency by region"
    rows = [("Region", "p50 (ms)", "p95 (ms)", "Change"), ("us-east", "62", "171", "-58%"),
            ("eu-west", "71", "188", "-55%"), ("ap-south", "95", "240", "-49%")]
    tbl = s.shapes.add_table(len(rows), 4, Inches(1), Inches(1.8), Inches(11.3), Inches(2.6)).table
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            tbl.cell(r, c).text = v
    s.notes_slide.notes_text_frame.text = "ap-south still lags. The new edge node ships in October."

    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Where the time goes"
    s.shapes.add_picture(photo_png(), Inches(0.8), Inches(1.7), Inches(5.6))
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7), Inches(1.9), Inches(5.4), Inches(1.2))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0x25, 0x63, 0xEB)
    box.text_frame.text = "Edge cache: 93% hit rate"
    box.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(24)
    box.text_frame.paragraphs[0].runs[0].font.bold = True
    ell = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7), Inches(3.5), Inches(2.4), Inches(2.4))
    ell.fill.solid()
    ell.fill.fore_color.rgb = RGBColor(0x16, 0xA3, 0x4A)
    ell.text_frame.text = "-22% cost"
    tb = s.shapes.add_textbox(Inches(9.8), Inches(3.8), Inches(3), Inches(2))
    tb.text_frame.word_wrap = True
    tb.text_frame.text = "Database time is down 40%. Network time is flat."
    s.notes_slide.notes_text_frame.text = "The photo is a placeholder for the architecture diagram."

    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Requests per day"
    cd = CategoryChartData()
    cd.categories = ["Jul", "Aug", "Sep"]
    cd.add_series("Requests (M)", (41, 47, 55))
    s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.7), Inches(8), Inches(5), cd)
    s.notes_slide.notes_text_frame.text = "Growth is steady. Mention the September launch."

    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = "Next steps"
    tf = s.placeholders[1].text_frame
    tf.text = "Ship the ap-south edge node"
    for line in ["Move session store to the new cluster", "Review latency budget in December"]:
        tf.add_paragraph().text = line
    s.notes_slide.notes_text_frame.text = "Ask for owners for each item before we close."
    prs.save(OUT / "quarterly-review.pptx")


def hostile():
    """Text that looks like HTML or script. The viewer must show it as plain text."""
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = 'Title <img src=x onerror="document.title=\'PWNED\'">'
    s.placeholders[1].text = 'Body <script>document.title="PWNED2"</script>\nA & B < C > D "quoted"'
    s.notes_slide.notes_text_frame.text = "<b>not bold</b> javascript:alert(1)"
    prs.save(OUT / "hostile.pptx")


def many_slides(n=120):
    prs = Presentation()
    for i in range(1, n + 1):
        s = prs.slides.add_slide(prs.slide_layouts[1])
        s.shapes.title.text = f"Slide {i}"
        s.placeholders[1].text = f"Item {i} alpha\nItem {i} beta"
        if i % 10 == 0:
            s.notes_slide.notes_text_frame.text = f"Checkpoint {i}"
    prs.save(OUT / "many-slides.pptx")


def ppsx_from(src, dst, content_type_from, content_type_to):
    """Copy a .pptx to .ppsx/.potx by changing the main part content type (that is the only difference)."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(content_type_from.encode(), content_type_to.encode())
            zout.writestr(item, data)


def smartart_stub():
    """A deck whose slide 1 has a SmartArt graphic frame (diagram parts) and a video, so the viewer shows the 'cannot draw' notice."""
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Org chart"
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    frame = (
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Diagram 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
        '<p:xfrm><a:off x="914400" y="1600200"/><a:ext cx="7315200" cy="4114800"/></p:xfrm>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
        '<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId9" r:lo="rId10" r:qs="rId11" r:cs="rId12"/>'
        '</a:graphicData></a:graphic></p:graphicFrame>'
        '<p:pic><p:nvPicPr><p:cNvPr id="10" name="Video 1"/><p:cNvPicPr/><p:nvPr><a:videoFile r:link="rId13"/></p:nvPr></p:nvPicPr>'
        '<p:blipFill/><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr></p:pic>'
    )
    out = OUT / "smartart.pptx"
    with zipfile.ZipFile(buf) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "ppt/slides/slide1.xml":
                data = data.replace(b"</p:spTree>", frame.encode() + b"</p:spTree>")
            zout.writestr(item, data)


if __name__ == "__main__":
    quarterly_review()
    hostile()
    many_slides()
    smartart_stub()
    main = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
    ppsx_from(OUT / "quarterly-review.pptx", OUT / "quarterly-review.ppsx", main,
              "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml")
    ppsx_from(OUT / "quarterly-review.pptx", OUT / "template.potx", main,
              "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml")
    for f in sorted(OUT.iterdir()):
        print(f"{f.name}\t{f.stat().st_size}")
