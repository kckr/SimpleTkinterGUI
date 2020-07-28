import reportlab
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate, Paragraph, Table, NextPageTemplate, \
    PageBreak, KeepInFrame, FrameBreak, Spacer, Image
from functools import partial

styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='Content',
                          fontName='Helvetica',
                          fontSize=14,
                          spaceBefore=6,
                          leading=20,
                          spaceAfter=12,
                          textColor=colors.HexColor("#3577D5")))
styles.add(ParagraphStyle(name='Content1',
                          fontName='Helvetica',
                          spaceBefore=6,
                          spaceAfter=12,
                          borderWidth=6,
                          borderPadding=3,
                          backColor=colors.HexColor("#DEEEEB"),
                          fontSize=12,
                          textColor=colors.HexColor("#222927")))
styles.add(ParagraphStyle(name='Content2',
                          fontName='Helvetica',
                          spaceBefore=6,
                          spaceAfter=12,
                          fontSize=14,
                          textColor=colors.HexColor("#FC7701")))
styles.add(ParagraphStyle(name='Content3',
                          fontName='Helvetica',
                          fontSize=14,
                          spaceBefore=6,
                          spaceAfter=12,
                          textColor=colors.HexColor("#4a5357")))


# pdf header across allpages
def header(canvas, doc, content):
    canvas.saveState()
    doc.height = 2 * inch
    doc.width = 2 * inch
    canvas.drawImage(content, 70, 650, 2 * inch, 2 * inch)
    canvas.restoreState()


def buildtabledata(table1record):
    data = [['{}'.format(x) for x in ['Number', 'Description', 'Qty']]]
    for record in table1record:
        recordtoappend = [record[1], record[2], record[3]]
        data.append(recordtoappend)

    return Table(data, repeatRows=1)


def text(my_canvas, assitemno, assitemdesc, ready, shipment, table1record, noteorder, comment):
    doc = BaseDocTemplate(my_canvas, pagesize=letter, allowSplitting=1)
    header_content = 'logo.jpg'
    flowables = []
    flowables.append(Paragraph(
        '<b>Assembly Item Number</b>' + '&nbsp' * 15 + ' <b>Assembly Item Description</b>',
        styles["Content"]))
    flowables.append(Paragraph(assitemno + ' &nbsp ' * (32 - len(assitemno)) + assitemdesc, styles["Content1"]))
    text = """ <b> Ready </b> """ + '&nbsp' * 43 + """<b> Shipment </b>"""
    para = Paragraph(text, styles["Content2"])
    flowables.append(para)
    text1 = ready + 29 * """ &nbsp """ + shipment
    para1 = Paragraph(text1, styles["Content1"])
    flowables.append(para1)
    flowables.append(Paragraph('<b> Notes </b>', styles["Content3"]))
    flowables.append(Paragraph(noteorder, styles["Content1"]))
    flowables.append(Paragraph('<b> Comments </b>', styles["Content3"]))
    flowables.append(Paragraph(comment, styles["Content1"]))
    datatable = buildtabledata(table1record)
    flowables.append(datatable)

    tblstyle = reportlab.platypus.TableStyle(
        [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.HexColor("#17061f")), ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
         ('BOX', (0, 0), (-1, -1), 0.25, colors.HexColor("#17061f")),
         ('GRID', (0, 0), (-1, -1), 0.01 * inch, (0, 0, 0,)), ])
    tblstyle1 = reportlab.platypus.TableStyle(
        [('FONT', (0, 0), (-1, -1), 'Helvetica', 10), ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor("#4a5357")),
         ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#137d31")),
         ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 14)])
    datatable.setStyle(tblstyle)
    datatable.setStyle(tblstyle1)

    up_frame = Frame(70, 0, width=6 * inch, height=9 * inch, showBoundary=0)

    first_template = PageTemplate(id='first', frames=[up_frame], pagesize=letter,
                                  onPage=partial(header, content=header_content))
    doc.addPageTemplates(first_template)
    flowables.append(NextPageTemplate('first'))
    flowables.append(Spacer(inch, 2 * inch))
    flowables.append(FrameBreak())

    doc.build(flowables)
