from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Flowable, Paragraph, SimpleDocTemplate,\
     Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
import PyPDF4
import re
import os


class LineDrw(Flowable):

    def __init__(self, width, height=0):
        Flowable.__init__(self)
        self.width = width
        self.height = height

    def __repr__(self):
        return "Line(w=%s)" % self.width

    def draw(self):
        self.canv.line(0, self.height, self.width, self.height)


class Report:

    def __init__(self, datatable, rptdata, Colnames, Colwdths, ttl, filename):
        # (table name, table data, table column names, table column
        # widths, name of PDF file)
        self.datatable = datatable
        self.rptdata = rptdata
        self.columnames = Colnames
        self.colms = Colwdths
        self.filename = filename
        self.ttl = ttl

        self.width, self.height = letter
        self.textAdjust = 6.5

    def create_pdf(self):
        body = []

        styles = getSampleStyleSheet()
        spacer1 = Spacer(0, 0.25*inch)
        spacer2 = Spacer(0, 0.5*inch)

        tblstyle = TableStyle([
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')])

        for s in range(len(self.datatable)):
            if self.rptdata[s] is not None:
                if self.datatable[s].find("_") != -1:
                    title = (self.datatable[s].replace("_", " "))
                else:
                    title = (' '.join(re.findall('([A-Z][a-z]*)',
                             self.datatable[s])))

                # provide the table name for the page header
                ptext = '<font size=14>%s</font>' % title + self.ttl
                body.append(Paragraph(ptext, styles["Heading2"]))
                line = LineDrw(500)
                body.append(line)
                body.append(spacer2)

                for pg in range(0, len(self.rptdata[s]), 3):
                    n = 0
                    for i in self.columnames[s][0]:
                        # add the requirement process type and notes
                        ptxt = ("<b>{0}:</b>      {1}<br/>".
                                format(i, self.rptdata[s][pg][n]))
                        body.append(Paragraph(ptxt, styles["Normal"]))
                        body.append(spacer1)
                        n += 1
                    if self.datatable[s] == 'Piping':
                        txt = 'Pipe Material'
                    else:
                        txt = 'Tube Material'
                    ptext = '<font size=14>%s</font>' % txt
                    body.append(Paragraph(ptext, styles["Heading3"]))
                    line = LineDrw(100)
                    body.append(line)
                    body.append(spacer1)

                    # this populates the upper table
                    tbldata = []
                    tbldata.append(self.columnames[s][1])
                    rowdata = self.rptdata[s][pg+1]

                    if rowdata != []:
                        for seg in rowdata:
                            m = 0
                            seg = list(seg)
                            for item in seg:
                                # wrap any text which is longer than
                                # 10 characters
                                if type(item) == str:
                                    if len(item) >= 10:
                                        item = Paragraph(item,
                                                         styles['Normal'])
                                        seg[m] = item
                                m += 1
                            tbldata.append(tuple(seg))

                            colwdth1 = []
                            for i in self.colms[s][1]:
                                colwdth1.append(i * self.textAdjust)

                        tbl1 = Table(tbldata, colWidths=colwdth1)

                        tbl1.setStyle(tblstyle)
                        body.append(tbl1)
                        body.append(spacer2)

                    if self.datatable[s] == 'Piping':
                        txt = 'Pipe Nipples'
                    else:
                        txt = 'Tube Valves'
                    ptext = '<font size=14>%s</font>' % txt
                    body.append(Paragraph(ptext, styles["Heading3"]))
                    line = LineDrw(100)
                    body.append(line)
                    body.append(spacer1)

                    # this populates the lower table
                    tbldata = []
                    tbldata.append(self.columnames[s][2])
                    rowdata = self.rptdata[s][pg+2]

                    if rowdata != []:
                        for seg in rowdata:
                            if seg is None:
                                continue
                            m = 0
                            seg = list(seg)
                            for item in seg:
                                # wrap any text which is longer than
                                # 10 characters
                                if type(item) == str:
                                    if len(item) >= 10:
                                        item = Paragraph(item,
                                                         styles['Normal'])
                                        seg[m] = item
                                m += 1
                            tbldata.append(tuple(seg))

                            colwdth1 = []
                            for i in self.colms[s][2]:
                                colwdth1.append(i * self.textAdjust)

                        tbl1 = Table(tbldata, colWidths=colwdth1)

                        tbl1.setStyle(tblstyle)
                        body.append(tbl1)

                    body.append(PageBreak())
                w, h = tbl1.wrap(15, 15)
            else:
                if self.datatable[s].find("_") != -1:
                    title = (self.datatable[s].replace("_", " "))
                else:
                    title = (' '.join(re.findall('([A-Z][a-z]*)',
                             self.datatable[s])))

                # provide the table name for the page header
                ptext = '<font size=14>%s</font>' % title + self.ttl
                body.append(Paragraph(ptext, styles["Heading2"]))
                line = LineDrw(500)
                body.append(line)

                txt = 'Not Data Set Up for this Item'
                ptext = '<font size=12>%s</font>' % txt
                body.append(Paragraph(ptext, styles["Heading2"]))

        doc = SimpleDocTemplate(
            'tmp_rot_file.pdf', pagesize=landscape(letter),
            rightMargin=.5*inch, leftMargin=.5*inch, topMargin=.75*inch,
            bottomMargin=.5*inch)

        doc.build(body)

        pdf_old = open('tmp_rot_file.pdf', 'rb')
        pdf_reader = PyPDF4.PdfFileReader(pdf_old)
        pdf_writer = PyPDF4.PdfFileWriter()

        for pagenum in range(pdf_reader.numPages):
            page = pdf_reader.getPage(pagenum)
            page.rotateCounterClockwise(90)
            pdf_writer.addPage(page)

        pdf_out = open(self.filename, 'wb')
        pdf_writer.write(pdf_out)
        pdf_out.close()
        pdf_old.close()
        os.remove('tmp_rot_file.pdf')
