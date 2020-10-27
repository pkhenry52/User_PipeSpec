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
    def __init__(self, datatable, rptdata, Colnames, Colwdths,
                 filename, exclud, ttl=None):
        # (table name, table data, table column names, table column
        # widths, name of PDF file)
        self.datatable = datatable
        self.rptdata = rptdata
        self.columnames = Colnames
        self.colms = Colwdths
        self.filename = filename
        self.ttl = ttl
        self.exclud = exclud

        self.width, self.height = letter
        self.textAdjust = 6.5

    def create_pdf(self):
        body = []
        tblttl = []

        styles = getSampleStyleSheet()
        spacer1 = Spacer(0, 0.25*inch)
        spacer2 = Spacer(0, 0.5*inch)
        # for each table returned convert table name to report title
        for item in self.datatable:
            tbltitl = (' '.join(re.findall('([A-Z][a-z]*)', item)))
            tblttl.append(tbltitl)

        tblstyle = TableStyle([
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')])
        pgnum = 1

        for n in range(len(self.datatable)):
            if self.rptdata[n] != []:
                # use the table name to create the report title
                if self.datatable[n].find("_") != -1:
                    self.tblname = (self.datatable[n].replace("_", " "))
                else:
                    self.tblname = (' '.join(re.findall('([A-Z][a-z]*)',
                                    self.datatable[n])))

                txt = self.tblname + ' ' + self.ttl
                ptext = '<font size=14>%s</font>' % txt
                body.append(Paragraph(ptext, styles["Heading2"]))

                line = LineDrw(500)

                body.append(line)
                body.append(spacer1)

                # provide the list of items included in fittings
                if self.datatable[n] == 'Fittings':
                    txt = '''NOTE: Fittings applies to:<br/>
                                90deg Elbows,<br/>
                                45deg Elbows,<br/>
                                Tees,<br/>
                                Couplings<br/>
                                and Laterals'''
                    ptxt = '<font size=12>%s</font>' % txt
                    body.append(Paragraph(ptxt, styles["Normal"]))
                    body.append(spacer2)

                colwdth1 = []
                colwdth2 = []
                for i in self.colms[n][0]:
                    colwdth1.append(i * self.textAdjust)
                for i in self.colms[n][1]:
                    colwdth2.append(i * self.textAdjust)

                NumTbls = len(self.rptdata[n])
                for pg in range(0, NumTbls, 2):
                    rptdata1 = []
                    rptdata2 = []
                    rptdata1.append(list(self.columnames[n][0]))
                    rptdata1.append(self.rptdata[n][pg])
                    rptdata2.append(list(self.columnames[n][1]))

                    m = 0
                    # remove the individual rows of data from data string
                    for seg in range(
                            len(self.rptdata[n][pg+1])//len(colwdth2)):
                        temp_data = list(self.rptdata[n][pg+1]
                                            [m * len(colwdth2):
                                            (m * len(colwdth2) +
                                            len(colwdth2))])
                        p = 0
                        for item in temp_data:
                            if type(item) == str:
                                if len(item) >= 15:
                                    item = Paragraph(item,
                                                        styles['Normal'])
                                    temp_data[p] = item
                            p += 1

                        rptdata2.append(temp_data)

                        m += 1

                    tbl1 = Table(rptdata1, colWidths=colwdth1)
                    tbl2 = Table(rptdata2, colWidths=colwdth2)

                    tbl1.setStyle(tblstyle)
                    body.append(tbl1)
                    body.append(spacer2)

                    tbl2.setStyle(tblstyle)
                    body.append(tbl2)
                    body.append(spacer2)

                    if pgnum == 2 or self.datatable[n] == 'Fittings':
                        body.append(PageBreak())
                        pgnum = 0

                    pgnum += 1
                w, h = tbl2.wrap(15, 15)

        for nodata in self.exclud:
            txt = nodata + ' ' + self.ttl
            ptext = '<font size=14>%s</font>' % txt
            body.append(Paragraph(ptext, styles["Heading2"]))

            line = LineDrw(500)

            body.append(line)

            txt = nodata + ' has not been set up'
            ptext = '<font size=14>%s</font>' % txt
            body.append(Paragraph(ptext, styles["Heading2"]))
            body.append(spacer1)

        doc = SimpleDocTemplate(
            'tmp_rot_file.pdf', pagesize=landscape(letter),
            rightMargin=.5*inch, leftMargin=.5*inch, topMargin=.75*inch,
            bottomMargin=.5*inch)

        doc.build(body)

        pdf_old = open('tmp_rot_file.pdf', 'rb')
        pdf_reader = PyPDF4.PdfFileReader(pdf_old)
        pdf_writer = PyPDF4.PdfFileWriter()
        # this will remove the blank page if there is no data present
        m = 0
        if len(self.exclud) == 6:
            m = 1

        for pagenum in range(pdf_reader.numPages-m):
            page = pdf_reader.getPage(pagenum)
            page.rotateCounterClockwise(90)
            pdf_writer.addPage(page)

        pdf_out = open(self.filename, 'wb')
        pdf_writer.write(pdf_out)
        pdf_out.close()
        pdf_old.close()
        os.remove('tmp_rot_file.pdf')
