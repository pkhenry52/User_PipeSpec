
import wx
from wx.core import DefaultSize
import wx.dataview as dv
import textwrap
import sqlite3
import re
import datetime
import os
import webbrowser
# import pdfkit
import wx.lib.scrolledpanel as sc
from PyPDF4 import PdfFileMerger
from bs4 import BeautifulSoup
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.figure import Figure


# this initializes the database and the cursor
def connect_db(filename):
    global cursr, db
    db = sqlite3.connect(filename)
    with db:
        cursr = db.cursor()
        cursr.execute('PRAGMA foreign_keys=ON')


class StrUpFrm(wx.Frame):
    # the main stsart up form which begins with the selection of the database
    # followed with the login and then selection of the user screens
    def __init__(self):

        super(StrUpFrm, self).__init__(None, wx.ID_ANY,
                                       "Pipe Specification Start-Up",
                                       size=(250, 160),
                                       style=wx.DEFAULT_FRAME_STYLE &
                                       ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX |
                                         wx.MINIMIZE_BOX))

        self.Bind(wx.EVT_CLOSE, self.OnClosePrt)

        self.go_value = False
        self.currentDirectory = os.getcwd()
        # create the buttons and bindings
        self.b2 = wx.Button(self, label="  Open\nDatabase", size=(80, 50))
        self.b2.Bind(wx.EVT_BUTTON, self.onOpenFile)

        self.b1 = wx.Button(self, label="Cancel", size=(60, 30))
        self.Bind(wx.EVT_BUTTON, self.OnClosePrt, self.b1)
        self.b1.SetForegroundColour((255, 0, 0))

        sizer = wx.BoxSizer(wx.VERTICAL)

        sizer.Add(self.b2, 0, wx.ALL | wx.CENTER, 8)
        sizer.Add((10,10))
        sizer.Add(self.b1, 0, wx.ALL | wx.CENTER, 5)
        self.SetSizer(sizer)

    def onOpenFile(self, evt):
        self.path = ''
        dlg = wx.FileDialog(
            self, message="Choose a file",
            defaultDir=self.currentDirectory,
            defaultFile="",
            wildcard="SQLite file (*.db)|*.db",
            style=wx.FD_OPEN | wx.FD_MULTIPLE | wx.FD_CHANGE_DIR
            )
        if dlg.ShowModal() == wx.ID_OK:
            self.path = dlg.GetPaths()[0]

        dlg.Destroy()

        if self.path != '':
            connect_db(self.path)
            self.go_value = True
            SpecFrm(self)

    def OnClosePrt(self, evt):
        if self.go_value:
            cursr.close()
            db.close()
        self.Destroy()


class SpecFrm(wx.Frame):
    def __init__(self, parent):
        self.parent = parent
        ttl = 'Commodity Properties and Related Piping Specification'
        super(SpecFrm, self).__init__(parent,
                                      title=ttl)

        self.Maximize(True)
        self.SetSizeHints(minW=1125, minH=750)
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.InitUI()

    def InitUI(self):
        self.pnl = BldFrm(self)

        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_1.Add(self.pnl, 1, wx.EXPAND | wx.ALL, 0)
        self.SetSizer(sizer_1)
        self.Layout()

        menuBar = wx.MenuBar()

        # File Menu
        fyl = wx.Menu()
        html = fyl.Append(wx.ID_ANY, "Open HTML file")
        self.Bind(wx.EVT_MENU, self.pnl.OnHTML, html)
        pdf = fyl.Append(wx.ID_ANY, "Open PDF file")
        self.Bind(wx.EVT_MENU, self.pnl.OnPDF, pdf)
        mrg = fyl.Append(wx.ID_ANY, "Merge PDF files")
        self.Bind(wx.EVT_MENU, self.pnl.OnMerge, mrg)
        html2pdf = fyl.Append(wx.ID_ANY, "Convert HTML to PDF")
        self.Bind(wx.EVT_MENU, self.pnl.OnConvert, html2pdf)
        dlet = fyl.Append(wx.ID_ANY, "Delete Document")
        self.Bind(wx.EVT_MENU, self.pnl.OnDelete, dlet)
        fyl.AppendSeparator()
        exxt = fyl.Append(wx.ID_EXIT, "E&xit")
        self.Bind(wx.EVT_MENU, self.OnClose, exxt)
        menuBar.Append(fyl, "&File")

        forms = wx.Menu()
        sow = forms.Append(200, 'Scope of Work')
        self.Bind(wx.EVT_MENU, self.pnl.OnForm, sow)
        htr = forms.Append(201, 'Hydro Test Report')
        self.Bind(wx.EVT_MENU, self.pnl.OnForm, htr)
        htw = forms.Append(202, 'Hydro Test Waiver')
        self.Bind(wx.EVT_MENU, self.pnl.OnForm, htw)
        ncr = forms.Append(203, 'Nonconformance Report')
        self.Bind(wx.EVT_MENU, self.pnl.OnForm, ncr)
        msr = forms.Append(204, 'Material Substitution Request')
        self.Bind(wx.EVT_MENU, self.pnl.OnForm, msr)
        rns = forms.Append(205, 'Request for New Specification')
        self.Bind(wx.EVT_MENU, self.pnl.OnForm, rns)
        its = forms.Append(206, 'Inspection Travel Sheet')
        self.Bind(wx.EVT_MENU, self.pnl.OnITS, its)
        menuBar.Append(forms, 'F&orms')

        frmhelp = wx.Menu()
        calc = frmhelp.Append(300, 'Misc Calculations')
        self.Bind(wx.EVT_MENU, self.pnl.OnCalcs, calc)
        hlp = frmhelp.Append(301, 'Help for this Form')
        self.Bind(wx.EVT_MENU, self.pnl.OnHelp, hlp)
        abut = frmhelp.Append(302, 'About')
        self.Bind(wx.EVT_MENU, self.pnl.OnAbout, abut)
        menuBar.Append(frmhelp, '&Help')

        self.SetMenuBar(menuBar)

        # add these following lines since this is a call up form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnClose(self, evt):
        self.GetParent().Enable(True)   # add this line for child
        self.__eventLoop.Exit()     # add this line for child
        self.Destroy()


class BldFrm(wx.Panel):

    def __init__(self, parent, model=None):
        '''Routine to build form and populate grid'''
        super(BldFrm, self).__init__(parent)

        self.btns = []
        self.lctrls = []
        self.tblname = 'CommodityProperties'
        self.model = model

        self.data = self.LoadComdData()

        self.Srch = False
        self.NewSpec = False
        self.ComdPrtyID = None

        self.InitUI()

    def InitUI(self):
        # set up the table column names, width and if
        # column can be edited ie primary autoincrement
        tblinfo = []

        tblinfo = Dbase(self.tblname).Fld_Size_Type()
        ID_col = tblinfo[1]
        autoincrement = tblinfo[2]

        self.columnames = ['ID', 'Commodity\nCode', 'Commodity Description',
                           'Pipe Material\nSpecification', 'Fluid Category',
                           'Design\nPressure', 'Min. Design\nTemperature',
                           'Max. Design\nTemperature', 'End Connections',
                           'Specification\nCode', 'Pending', 'Note']

        self.colwdth = [6, 10, 35, 15, 17, 10, 11, 11, 18, 11, 6, 30]

        # Create a dataview control
        self.dvc = dv.DataViewCtrl(self, wx.ID_ANY, wx.DefaultPosition,
                                   wx.Size(500, 300),
                                   style=wx.BORDER_THEME
                                   | dv.DV_ROW_LINES
                                   | dv.DV_VERT_RULES
                                   | dv.DV_HORIZ_RULES
                                   | dv.DV_SINGLE
                                   )
        self.dvc.SetMinSize = (wx.Size(100, 200))
        self.dvc.SetMaxSize = (wx.Size(500, 400))

        # if autoincrement is false then the data can be sorted based on ID_col
        if autoincrement == 0:
            self.data.sort(key=lambda tup: tup[ID_col])

        # use the sorted data to load the dataviewlistcontrol
        if self.model is None:
            self.model = DataMods(self.tblname, self.data)
        self.dvc.AssociateModel(self.model)

        n = 0
        for colname in self.columnames:
            self.dvc.AppendTextColumn(colname, n,
                                      width=wx.LIST_AUTOSIZE_USEHEADER,
                                      mode=dv.DATAVIEW_CELL_INERT)
            n += 1

        # make columns not sortable and but reorderable.
        for c in self.dvc.Columns:
            c.Sortable = False
            c.Reorderable = True
            c.Resizeable = True

        self.dvc.Columns[(1)].Sortable = True
        self.dvc.Columns[(3)].Sortable = True

        # change to not let the ID col be moved.
        self.dvc.Columns[(ID_col)].Reorderable = False
        self.dvc.Columns[(ID_col)].Resizeable = False

        # Bind some events so we can see what the DVC sends us
        self.Bind(dv.EVT_DATAVIEW_ITEM_ACTIVATED, self.OnGridSelect, self.dvc)

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        # set up first row of combo boxesand label
        self.lblsizer1 = wx.BoxSizer(wx.HORIZONTAL)
        self.cmbsizer1 = wx.BoxSizer(wx.HORIZONTAL)
        self.lblsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        self.cmbsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        self.dvsizer = wx.BoxSizer(wx.HORIZONTAL)

        self.dvsizer.Add(10, -1, 0)
        self.dvsizer.Add(self.dvc, 1, wx.ALL | wx.EXPAND, 5)
        self.dvsizer.Add(10, -1, 0)

        self.Codelbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.Codelbl.SetLabel('Commodity\nCode')
        self.Codelbl.SetForegroundColour((255, 0, 0))

        self.Speclbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.Speclbl.SetLabel("Piping\nMat'r Spec")
        self.Speclbl.SetForegroundColour((255, 0, 0))

        self.Fluidlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.Fluidlbl.SetLabel('Fluid\nCategory')
        self.Fluidlbl.SetForegroundColour((255, 0, 0))

        self.DPlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.DPlbl.SetLabel('Design\nPressure')
        self.DPlbl.SetForegroundColour((255, 0, 0))

        self.DTlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.DTlbl.SetLabel(' Design Temperature\nMin.           Max.')
        self.DTlbl.SetForegroundColour((255, 0, 0))

        self.Endlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.Endlbl.SetLabel('Commodity End\nConnections')
        self.Endlbl.SetForegroundColour((255, 0, 0))

        self.lblsizer1.Add(self.Codelbl, 0, wx.ALL, 5)
        self.lblsizer1.Add(50, -1, 0)
        self.lblsizer1.Add(self.Speclbl, 0, wx.ALL, 5)
        self.lblsizer1.Add(90, -1, 0)
        self.lblsizer1.Add(self.Fluidlbl, 0, wx.ALL, 5)
        self.lblsizer1.Add(80, -1, 0)
        self.lblsizer1.Add(self.DPlbl, 0, wx.ALL, 5)
        self.lblsizer1.Add(15, -1, 0)
        self.lblsizer1.Add(self.DTlbl, 0, wx.ALL, 5)
        self.lblsizer1.Add(100, -1, 0)
        self.lblsizer1.Add(self.Endlbl, 0, wx.ALL, 5)
        self.lblsizer1.Add(130, -1, 0)

        self.textDP = wx.TextCtrl(self, size=(80, -1),  style=wx.CB_READONLY)
        self.textDP.SetHint('psig')

        self.textMinT = wx.TextCtrl(self, size=(80, -1),  style=wx.CB_READONLY)
        self.textMinT.SetHint('Deg F')

        self.textMaxT = wx.TextCtrl(self, size=(80, -1),  style=wx.CB_READONLY)
        self.textMaxT.SetHint('Deg F')

        # Start the generation of the required combo boxes
        # using a dictionary of column names and table names

        self.cmbCode = wx.TextCtrl(self, size=(80, -1), style=wx.CB_READONLY)
        self.cmbPipe = wx.TextCtrl(self, size=(90, -1), style=wx.CB_READONLY)
        self.cmbFluid = wx.TextCtrl(self, size=(150, -1), style=wx.CB_READONLY)
        self.cmbEnd = wx.TextCtrl(self, size=(300, -1), style=wx.CB_READONLY)
        self.txtDscrpt = wx.TextCtrl(self, size=(300, -1),
                                     style=wx.TE_CENTER)
        self.textSpec = wx.TextCtrl(self, size=(120, -1),
                                    style=wx.CB_READONLY)

        self.chkPend = wx.CheckBox(self)
        self.chkPend.SetValue(False)
        self.chkPend.SetForegroundColour(wx.Colour(255, 0, 0))
        self.chkPend.Disable()

        self.notes = wx.TextCtrl(self, size=(700, 40), value='',
                                 style=wx.TE_MULTILINE | wx.TE_LEFT |
                                 wx.CB_READONLY)

        self.cmbsizer1.Add(self.cmbCode, 0)
        self.cmbsizer1.Add((55, 5))
        self.cmbsizer1.Add(self.cmbPipe, 0)
        self.cmbsizer1.Add((45, 5))
        self.cmbsizer1.Add(self.cmbFluid, 0)
        self.cmbsizer1.Add((45, 5))
        self.cmbsizer1.Add(self.textDP, 0)
        self.cmbsizer1.Add(self.textMinT, 0)
        self.cmbsizer1.Add(self.textMaxT, 0)
        self.cmbsizer1.Add((45, 5))
        self.cmbsizer1.Add(self.cmbEnd, 0)

        self.Dscrptlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.Dscrptlbl.SetLabel('Commodity\nDescription')
        self.Dscrptlbl.SetForegroundColour((255, 0, 0))

        self.Speclbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.Speclbl.SetLabel('Piping\nSpecificaion')
        self.Speclbl.SetForegroundColour((255, 0, 0))

        self.Pendlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.Pendlbl.SetLabel('Specificaion\nPending Approval')
        self.Pendlbl.SetForegroundColour((255, 0, 0))

        self.lblsizer2.Add(self.Dscrptlbl, 0, wx.TOP, 15)
        self.lblsizer2.Add(330, -1, 0)
        self.lblsizer2.Add(self.Speclbl, 0, wx.TOP, 15)
        self.lblsizer2.Add(25, -1, 0)
        self.lblsizer2.Add(self.Pendlbl, 0, wx.TOP, 15)
        self.lblsizer2.Add(10, -1, 0)

        # add a button to call main form to search combo list data
        self.b6 = wx.Button(self, label="<-- Search\nDescription", size=(80,40))
        self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b6)

        self.b5 = wx.Button(self, label="Reset", size=(40,30))
        self.Bind(wx.EVT_BUTTON, self.OnRestoreBoxs, self.b5)

        self.cmbsizer2.Add(15, -1, 0)
        self.cmbsizer2.Add(self.txtDscrpt, 0, wx.ALIGN_CENTRE)
        self.cmbsizer2.Add(10, -1, 0)
        self.cmbsizer2.Add(self.b6, 0, wx.BOTTOM | wx.ALIGN_CENTRE, 5)
        self.cmbsizer2.Add(45, -1, 0)
        self.cmbsizer2.Add(self.textSpec, 0, wx.BOTTOM | wx.CENTER, 5)
        self.cmbsizer2.Add(40, -1, 0)
        self.cmbsizer2.Add(self.chkPend, 0, wx.BOTTOM, 40)
        self.cmbsizer2.Add(40, -1, 0)
        self.cmbsizer2.Add(self.b5, 0, wx.BOTTOM | wx.ALIGN_CENTER, 5)

        self.Notelbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER)
        self.Notelbl.SetLabel('Fluid Specific Notes')
        self.Notelbl.SetForegroundColour((255, 0, 0))

        self.notebox = wx.BoxSizer(wx.HORIZONTAL)
        self.notebox.Add(self.Notelbl, 1, wx.ALL | wx.ALIGN_CENTER, 25)
        self.notebox.Add(self.notes, 0, wx.ALIGN_CENTER, 5)

        self.tbllist = ['Piping', 'Fittings', 'Flanges', 'OrificeFlanges',
                        'GasketPacks', 'Fasteners', 'Unions', 'OLets',
                        'GrooveClamps', 'WeldRequirements', 'Specials',
                        'Tubing', 'BranchChart', 'Notes', 'InspectionPacks',
                        'PaintSpec', 'Insulation', 'GateValve', 'GlobeValve',
                        'PlugValve', 'BallValve', 'ButterflyValve',
                        'PistonCheckValve', 'SwingCheckValve']

        rptlist = [0, 3, 3, 3, 1, 4, 3, 3, 3, 5, 8, 0, 7, 8, 1, 1, 6, 2, 2, 2,
                   2, 2, 2, 2]

        chklbl = []
        for tblName in self.tbllist:
            if tblName.find("_") != -1:
                chkName = (tblName.replace("_", " "))
            else:
                chkName = (' '.join(re.findall('([A-Z][a-z]*)', tblName)))
            chklbl.append(chkName)
        # create a dictionary of checkbox labels and table names
        self.tbl_rpt = dict(zip(self.tbllist, rptlist))

        self.chksizers = []
        self.chkboxs = []
        self.Sizer1 = wx.BoxSizer(wx.HORIZONTAL)
        n = 0
        for a in range(4):
            chksizer = wx.BoxSizer(wx.VERTICAL)
            self.chksizers.append(chksizer)
            for b in range(6):
                lbl = (' '.join(re.findall('([A-Z][a-z]*)', self.tbllist[n])))
                chkbox = wx.CheckBox(self, id=n, label=lbl,
                                     name=self.tbllist[n],
                                     style=wx.ALIGN_RIGHT)
                self.chkboxs.append(chkbox)
                chksizer.Add(chkbox, 0, wx.ALIGN_RIGHT)
                n += 1
            self.Sizer1.Add(chksizer, 0, wx.ALL, 10)

        self.prtsizer = wx.BoxSizer(wx.VERTICAL)
        self.b1 = wx.Button(self, label="Build Report for\nSelected Items", size=(120,40))
        self.Bind(wx.EVT_BUTTON, self.OnPrintItems, self.b1)
        self.b1.Enable(False)
        self.b2 = wx.Button(self, label='Build Scope\nof Work', size=(120,40))
        self.Bind(wx.EVT_BUTTON, self.OnPrintScope, self.b2)
        self.b2.Enable(False)
        self.prtsizer.Add(self.b1, 0, wx.ALIGN_CENTER)
        self.prtsizer.Add((10,10))
        self.prtsizer.Add(self.b2, 0, wx.ALIGN_CENTER)
        self.Sizer1.Add(self.prtsizer, 0, wx.ALIGN_CENTER)

        self.Sizer.Add((20, 15))
        self.Sizer.Add(self.lblsizer1, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.cmbsizer1, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.lblsizer2, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.cmbsizer2, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.notebox, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.dvsizer, 1, wx.EXPAND)
        self.Sizer.Add((20, 15))
        self.Sizer.Add(self.Sizer1, 0, wx.ALIGN_CENTER)
        self.Sizer.Add((20, 30))

    def LoadComdData(self):
        # provides the actual data from the table depending on the valve style
        # note the order of the fields is important and reflects the order set
        # up in the SQLite database there should be a more secure way of doing
        # this that does not depend on the database structure
        self.Dsql = ('''SELECT v.CommodityPropertyID, a.Commodity_Code,
         a.Commodity_Description, b.Pipe_Material_Spec,
         d.Designation, v.Design_Pressure, v.Minimum_Design_Temperature,
         v.Maximum_Design_Temperature, c.Commodity_End, v.Pipe_Code,
         v.Pending, v.Note''')

        self.Dsql = self.Dsql + ' FROM ' + self.tblname + ' v' + '\n'

        # the index fields are returned from the PRAGMA statement tbldata
        # in reverse order so the alpha characters need to count up not down
        # join above list of grid columns + INNER JOIN Frg_Tbl 'alpha' ON
        # v.LinkField = 'alpha'.Frg_Fld where alpha is incremented
        n = 0
        tbldata = Dbase().Dtbldata(self.tblname)
        for element in tbldata:
            alpha = chr(96-n+len(tbldata))
            self.Dsql = (self.Dsql + ' INNER JOIN ' + element[2] + ' ' + alpha
                         + ' ON v.' + element[3] + ' = ' + alpha + '.' +
                         element[4] + '\n')
            n += 1
        data = Dbase().Dsqldata(self.Dsql)
        return data

    def FylDilog(self, wildcard, msg, stl):
        self.currentDirectory = os.getcwd()
        path = 'No File'
        dlg = wx.FileDialog(
            self, message=msg,
            defaultDir=self.currentDirectory,
            defaultFile="",
            wildcard=wildcard,
            style=stl
            )
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPaths()[0]
        dlg.Destroy()
        return path

    def OnForm(self, evt):
        slectID = evt.GetId()
        fileid = str(slectID)[-1]
        html_path = ''
        filename = ''
        filenames = ['Original_SOW.html', 'Original_HTR.html',
                     'Original_HTW.html', 'Original_NCR.html',
                     'Original_MSR.html', 'Original_RPS.html']
        # this should be the default location for the files

        filename = (os.getcwd() + os.sep + 'Forms'
                    + os.sep + filenames[int(fileid)])

        if 'Data' in filename:
            os.chdir('..')
            filename = (os.getcwd() + os.sep + 'Forms'
                        + os.sep + filenames[int(fileid)])

        html_path = 'file:' + os.sep*2 + filename

        # if default location does not work then open directory finder
        if os.path.isfile(filename) is False:
            filename = ''
            dlg = wx.DirDialog(self,
                               'Select the template forms Directory:',
                               style=wx.DD_DEFAULT_STYLE)

            if dlg.ShowModal() == wx.ID_OK:
                html_path = dlg.GetPath()
            dlg.Destroy()

            if html_path != '':
                filename = ''

        if filename == '':
            wx.MessageBox('Problem Locating HTML File', 'Error', wx.OK)
        else:
            brwsrs = ['firefox', 'safari', 'chrome', 'opera',
                      'netscape', 'google-chrome', 'lynx',
                      'mozilla', 'galeon', 'chromium',
                      'chromium-browser', 'windows-default',
                      'w3m', 'no browser']
            for brwsr in brwsrs:
                if brwsr == 'no browser':
                    wx.MessageBox('Problem Locating Web Browser',
                                  'Error', wx.OK)
                else:
                    try:
                        webbrowser.get(using=brwsr).open(html_path, new=2)
                        break
                    except Exception:
                        pass

    def OnHTML(self, evt):
        # show the html file in browser
        wildcard = "HTML file (*.html)|*.html"
        msg = 'Select HTML File to View'
        filename = self.FylDilog(wildcard, msg, wx.FD_OPEN |
                                 wx.FD_CHANGE_DIR)

        if filename != 'No File':
            brwsrs = ['firefox', 'safari', 'chrome', 'opera',
                      'netscape', 'google-chrome', 'lynx',
                      'mozilla', 'galeon', 'chromium',
                      'chromium-browser', 'windows-default', 'w3m']
            for brwsr in brwsrs:
                try:
                    webbrowser.get(using=brwsr).open(filename, new=2)
                    break
                except Exception:
                    pass

    def OnPDF(self, evt):
        PDFFrm(self)

    def OnMerge(self, evt):
        MergeFrm(self)

    def OnITS(self, evt):
        BldTrvlSht(self)

    def OnConvert(self, evt):
        # select the html file to convert
        wildcard = "HTML file (*.html)|*.html"
        msg = 'Select HTML to Convert'
        filename = self.FylDilog(wildcard, msg, wx.FD_OPEN |
                                 wx.FD_FILE_MUST_EXIST)

        if filename != 'No File':

            with open(filename) as f:
                soup = BeautifulSoup(f, "html.parser")

            id_txt = []
            vl_txt = []
            id_txtar = []
            vl_txtar = []
            chk_chkd = []

            # get the name of the form for proper selection of pdf formate
            titl = [ttl.get_text() for ttl in soup.select('h1')][0]

            # get all the text input values
            for item in soup.find_all("input", {"type": "text"}):
                id_txt.append(item.get('id'))
                if item.get('value') is None:
                    vl_txt.append('TBD')
                else:
                    vl_txt.append(item.get('value'))
            txt_inpt = dict(zip(id_txt, vl_txt))
            txt_boxes = txt_inpt

            # get all the text area input
            for item in soup.find_all('textarea'):
                id_txtar.append(item.get('id'))
                if item.contents != []:
                    vl_txtar.append(item.contents[0])
                else:
                    vl_txtar.append('')
            txtar_inpt = dict(zip(id_txtar, vl_txtar))
            txt_area = txtar_inpt

            # collect the checkboxes which are checked only
            for item in soup.find_all('input', checked=True):
                chk_chkd.append(item.get('id'))
            chkd_boxes = chk_chkd

            # open dialog to specify new pdf file name
            wildcrd = "PDF file (*.pdf)|*.pdf"
            mssg = 'Save new PDF'
            pdf_file = self.FylDilog(wildcrd, mssg, wx.FD_SAVE)
            # if pdf extention is not specified add it to the file name
            if pdf_file.find(".pdf") == -1:
                pdf_file = pdf_file + '.pdf'

            # use the form title to select the proper convertion program
            if titl.find('Nonconformance') != -1:
                from NCR_pdf import NCRForm
                NCRForm(pdf_file, txt_boxes, txt_area, chkd_boxes).create()
            elif titl.find('Specification') != -1:
                from RPS_pdf import RPSForm
                RPSForm(pdf_file, txt_boxes, txt_area, chkd_boxes).create()
            elif titl.find('Scope') != -1:
                from SOW_pdf import SOWForm
                SOWForm(pdf_file, txt_boxes, txt_area, chkd_boxes).create()
            elif titl.find('Material') != -1:
                from MSR_pdf import MSRForm
                MSRForm(pdf_file, txt_boxes, txt_area, chkd_boxes).create()
            elif titl.find('Waiver') != -1:
                from HTW_pdf import HTWForm
                HTWForm(pdf_file, txt_boxes, txt_area, chkd_boxes).create()
            elif titl.find('Hydrostatic') != -1:
                from HTR_pdf import HTRForm
                HTRForm(pdf_file, txt_boxes, txt_area, chkd_boxes).create()
            else:
                msg = 'Program can only convert the following HTML Documents:\n\
                Hydrostatic Test Report,\n\
                Hydrostatic Test Waiver,\n\
                Nonconformance Report,\n\
                Scope Of Work,\n\
                Material Substitution or\n\
                Request For New Specification'
                dlg = wx.MessageDialog(self, message=msg,
                                       caption="Invalid Document",
                                       style=wx.OK)
                dlg.ShowModal()
                dlg.Destroy()

    def OnDelete(self, evt):
        wildcard = "HTML file (*.html)|*.html|"\
                   "PDF files (*.pdf)|*.pdf"
        msg = 'Select File to Delete'
        filename = self.FylDilog(wildcard, msg, wx.FD_OPEN |
                                 wx.FD_FILE_MUST_EXIST)
        if filename != 'No File':
            msg = 'Confirm deletion of file?\n' + filename
            dlg = wx.MessageDialog(
                self, message=msg, caption="Question",
                style=wx.OK | wx.CANCEL | wx.ICON_QUESTION)

            if dlg.ShowModal() == wx.ID_OK:
                os.remove(filename)

    def OnHelp(self, evt):
        msg = ('''
    \tThis form is designed to generate PDF files regarding the
    piping specification for a selected commodity property.  It
    will use this information to develop the Scope of work if
    requested.\n
    \tSelect the required commodity property by double clicking
    the commodity in the table. The items to be printed can then
    be selected from the check boxes at the bottom of the screen.\n
    \tTo help search for a specific commodity property, the sort
    order of the Commodity Code and Pipe Material Spec columns
    can be changed by double clicking the table header.  You can
    also search based on words in the commodity description.
    Simply type word or phrase in the Commodity Description box then
    press Search Description.
    ''')

        dlg = wx.MessageDialog(self, message=msg,
                               caption='Use of Form',
                               style=wx.ICON_INFORMATION | wx.STAY_ON_TOP
                               | wx.CENTRE)
        dlg.ShowModal()
        dlg.Destroy()

    def OnAbout(self, evt):
        msg = ('''
    This program was written as open source, the program code can
    be used downloaded and distributed freely.\n
    Please note; all data was entered for demonstration use only.\n
    If the user wishes data can be input at a cost either as data
    supplied by the user or developed on a fee base.\n
    kphprojects@gmail.com''')
        wx.MessageBox(msg, 'Title', wx.OK)

    def OnCalcs(self, evt):
        CalcFrm(self)

    def OnPrintItems(self, evt):
        # get a specific user file name
        msg = 'Save Report as PDF.'
        wildcard = 'PDF (*.pdf)|*.pdf'
        filename = self.FylDilog(wildcard, msg, wx.FD_SAVE |
                                 wx.FD_OVERWRITE_PROMPT)
        if filename != 'No File':
            # confirm an extension was specifed as pdf if not
            # add pdf extension
            if filename[-4:].lower() != '.pdf':
                filename += '.pdf'
            rpts = []
            # build the needed empty array to be populated later
            for n in range(9):
                rpt = []
                rpts.append(rpt)
            # populate the rpts array with the table
            # names of all the check boxes
            # tbl_rpt is a dictionary of (table name:report type number)
            for n in range(len(self.chkboxs)):
                if self.chkboxs[n].GetValue() is True:
                    tblrpt = self.chkboxs[n].GetName()
                    rpts[self.tbl_rpt[tblrpt]].append(tblrpt)

            self.rptSelection(rpts, filename)
            PDFFrm(self, filename)

    def OnPrintScope(self, evt):
        msg = 'Save Report as PDF.'
        wildcard = 'PDF (*.pdf)|*.pdf'
        filename = self.FylDilog(wildcard, msg, wx.FD_SAVE |
                                 wx.FD_OVERWRITE_PROMPT)
        if filename != 'No File':
            # confirm an extension was specifed as pdf if not
            # add pdf extension
            if filename[-4:].lower() != '.pdf':
                filename += '.pdf'

            rpts = []
            # build the needed empty array to be populated later
            for n in range(9):
                rpt = []
                rpts.append(rpt)
            # this seperates the check box table names
            # into the correct sublist by report type number
            for tblrpt in self.tbllist:
                rpts[self.tbl_rpt[tblrpt]].append(tblrpt)

            self.rptSelection(rpts, filename)
            PDFFrm(self, filename)

    def rptSelection(self, rpts, filename):
        # this loop will be used to generate the first report
        # to which all following reports will be merged
        # use os.path.basename('filepath') to get the file name
        # to get full directory path excluding the ending "/"
        # use os.path.dirname('filename')
        n = 0
        for n in range(len(rpts)):
            # using just the list of fitting related tables call the
            # fittings report builder

            # cycle through each list of table names in the rpts list
            # remember rpts list is a list of a list of table names or
            # an empty list used to hold the location
            # for each report type pass the list of table names, comd_prty
            # pipe code, file name and item ID text.
            if rpts[n] != []:
                if n == 0:  # for piping and tubing
                    Conduit(rpts[0], self.ComdPrtyID,
                            self.cmbCode.GetValue(),
                            self.cmbPipe.GetValue(),
                            (os.path.dirname(filename) + os.sep
                            + 'PDFtmp1.pdf'),
                            self.textSpec.GetValue()).CondBld()
                elif n == 1:  # for paint, gaskets and inspection
                    Fabricate(rpts[1], self.ComdPrtyID,
                              self.cmbCode.GetValue(),
                              self.cmbPipe.GetValue(),
                              (os.path.dirname(filename) + os.sep
                              + 'PDFtmp4.pdf'),
                              self.textSpec.GetValue()).FabBld()
                elif n == 2:  # for all the valves
                    ValvesRpt(rpts[2], self.ComdPrtyID,
                              self.cmbCode.GetValue(),
                              self.cmbPipe.GetValue(),
                              (os.path.dirname(filename) + os.sep
                              + 'PDFtmp3.pdf'),
                              self.textSpec.GetValue()).VlvBld()
                elif n == 3:  # for fittings, unions, flanges, o-lets, clamps
                    FittingsRpt(rpts[3], self.ComdPrtyID,
                                self.cmbCode.GetValue(),
                                self.cmbPipe.GetValue(),
                                (os.path.dirname(filename) + os.sep
                                + 'PDFtmp2.pdf'),
                                self.textSpec.GetValue()).FitngBld()
                elif n == 4:  # for fasteners
                    FastenersRpt(rpts[4], self.ComdPrtyID,
                                 self.cmbCode.GetValue(),
                                 self.cmbPipe.GetValue(),
                                 (os.path.dirname(filename) + os.sep
                                 + 'PDFtmp5.pdf'),
                                 self.textSpec.GetValue()).FstBld()
                elif n == 5:  # for weld requirements
                    WeldingRpt(rpts[5], self.ComdPrtyID,
                               self.cmbCode.GetValue(),
                               self.cmbPipe.GetValue(),
                               (os.path.dirname(filename) + os.sep
                               + 'PDFtmp6.pdf'),
                               self.textSpec.GetValue()).WldBld()
                elif n == 6:  # for insulation requirements
                    InsulRpt(rpts[6], self.ComdPrtyID,
                             self.cmbCode.GetValue(),
                             self.cmbPipe.GetValue(),
                             (os.path.dirname(filename) + os.sep
                             + 'PDFtmp8.pdf'),
                             self.textSpec.GetValue()).InsulBld()
                elif n == 7:  # for branch table
                    BranchRpt(rpts[7], self.ComdPrtyID,
                              self.cmbCode.GetValue(),
                              self.cmbPipe.GetValue(),
                              self.cmbEnd.GetValue(),
                              (os.path.dirname(filename) + os.sep
                              + 'PDFtmp9.pdf'),
                              self.textSpec.GetValue()).BrchBld()
                elif n == 8:  # for specials and notes
                    SpcNotesRpt(rpts[8], self.ComdPrtyID,
                                self.cmbCode.GetValue(),
                                self.cmbPipe.GetValue(),
                                (os.path.dirname(filename) + os.sep
                                + 'PDFtmp7.pdf'),
                                self.textSpec.GetValue()).SpcNtsBld()
            n += 1

        self.merge_pdf(filename)

    def merge_pdf(self, targetfile):
        import glob
        PDF_merge = PdfFileMerger()
        sorc = glob.glob(os.path.dirname(targetfile) + os.sep + 'PDFtmp*.pdf')
        sorc.sort()
        for pdffile in sorc:
            PDF_merge.append(open(pdffile, 'rb'))
            os.remove(pdffile)

        with open(targetfile, 'wb') as fileobj:
            PDF_merge.write(fileobj)

    def Search_Restore(self):
        if self.Srch is True:
            self.b6.SetLabel("<-- Search\nDescription")
            self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b6)
            self.RestoreBoxs()
            self.Srch = False
        else:
            self.b6.SetLabel("Restore\nTable")
            self.Bind(wx.EVT_BUTTON, self.OnRestoreTbl, self.b6)
            self.Srch = True

    def OnRestoreBoxs(self, evt):
        self.RestoreBoxs()
        for chkbox in self.chkboxs:
            chkbox.SetValue(0)

    def OnRestoreTbl(self, evt):
        self.RestoreTbl()

    def RestoreBoxs(self):
        self.cmbCode.ChangeValue('')
        self.cmbPipe.ChangeValue('')
        self.cmbFluid.ChangeValue('')
        self.textDP.ChangeValue('')
        self.textMinT.ChangeValue('')
        self.textMaxT.ChangeValue('')
        self.cmbEnd.ChangeValue('')
        self.txtDscrpt.ChangeValue('')
        self.textSpec.ChangeValue('')
        self.chkPend.SetValue(False)
        self.notes.ChangeValue('')

        self.textMaxT.SetHint('Deg F')
        self.textMinT.SetHint('Deg F')
        self.textDP.SetHint('psig')

        self.b1.Enable(False)
        self.b2.Enable(False)

    def RestoreTbl(self):
        self.data = Dbase().Dsqldata(self.Dsql)
        self.model = DataMods(self.tblname, self.data)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh
        self.Search_Restore()

    def OnSearch(self, evt):
        srchstrg = ''
        qry = ('''SELECT * FROM CommodityCodes WHERE Commodity_Description
                LIKE '%''' + self.txtDscrpt.GetValue() + "%' COLLATE NOCASE")
        srchdata = Dbase().Search(qry)
        if srchdata:
            n = len(srchdata)
            m = 1
            for item in srchdata:
                srchstrg = srchstrg + "'" + item[0] + "'"
                if m < n:
                    srchstrg = srchstrg + " OR a.Commodity_Code = "
                m += 1
            ShQuery = self.Dsql + '\n WHERE a.Commodity_Code = ' + srchstrg
            srchdata = Dbase().Search(ShQuery)
        else:
            srchdata = []

        self.data = srchdata
        self.model = DataMods(self.tblname, self.data)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh
        self.Search_Restore()

    def OnGridSelect(self, evt):
        item = self.dvc.GetSelection()
        rowGS = self.model.GetRow(item)

        self.ComdPrtyID = self.data[rowGS][0]
        self.cmbCode.ChangeValue(self.data[rowGS][1])
        self.txtDscrpt.ChangeValue(self.data[rowGS][2])
        self.cmbPipe.ChangeValue(self.data[rowGS][3])
        self.cmbFluid.ChangeValue(self.data[rowGS][4])
        self.textDP.ChangeValue(str(self.data[rowGS][5]))
        self.textMinT.ChangeValue(str(self.data[rowGS][6]))
        self.textMaxT.ChangeValue(str(self.data[rowGS][7]))
        self.cmbEnd.ChangeValue(self.data[rowGS][8])
        self.textSpec.ChangeValue(str(self.data[rowGS][9]))
        self.chkPend.SetValue(self.data[rowGS][10])
        self.notes.ChangeValue(str(self.data[rowGS][11]))

        self.b1.Enable()
        self.b2.Enable()
        self.b6.SetLabel("<-- Search\nDescription")
        self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b6)
        self.Srch = False
        self.data = Dbase().Dsqldata(self.Dsql)
        self.model = DataMods(self.tblname, self.data)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh

    def OnCloseFrm(self, evt):
        # Dbase().close_database()
        self.GetParent().Destroy()


class MergeFrm(wx.Frame):
    def __init__(self, parent):

        super(MergeFrm, self).__init__(parent,
                                       title='Select pdf files to merge',
                                       size=(725, 550))
        self.Bind(wx.EVT_CLOSE, self.OnExit)
        self.parent = parent
        self.InitUI()

    def InitUI(self):

        from wx.lib.itemspicker import ItemsPicker, \
             EVT_IP_SELECTION_CHANGED, IP_REMOVE_FROM_CHOICES

        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        pick_sizer = wx.BoxSizer(wx.VERTICAL)
        self.ip = ItemsPicker(self, -1,
                              [],
                              'All pdf files in\nselected directory:',
                              'Files to be merged\nin order shown:',
                              size=(700, 400), ipStyle=IP_REMOVE_FROM_CHOICES)
        # event occures when bAdd is pressed
        self.ip.Bind(EVT_IP_SELECTION_CHANGED, self.OnSelectionChange)
        self.ip._source.SetMinSize((-1, 150))

        self.ip.bAdd.SetLabel('Add =>')
        self.ip.bRemove.SetLabel('<= Remove')
        pick_sizer.Add(self.ip, 0, wx.ALL, 10)

        btnsizer = wx.BoxSizer(wx.HORIZONTAL)
        b1 = wx.Button(self, label="Select File\nDirectory")
        self.Bind(wx.EVT_BUTTON, self.OnAdd, b1)

        b2 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnExit, b2)

        b3 = wx.Button(self, label='Merge\nPDF files')
        self.Bind(wx.EVT_BUTTON, self.OnMerge, b3)

        btnsizer.Add(b1, 0, wx.ALL, 5)
        btnsizer.Add((340, 10))
        btnsizer.Add(b3, 0, wx.ALL, 5)
        btnsizer.Add(b2, 0, wx.ALL, 5)
        self.Sizer.Add(pick_sizer, 0, wx.ALL, 10)
        self.Sizer.Add(btnsizer, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnSelectionChange(self, evt):
        # items in left side box
        self.pdf_list = self.ip.GetSelections()

    def OnMerge(self, evt):
        self.MergePDF()

    def MergePDF(self):
        # get a specific user file name
        msg = 'Save Report as PDF.'
        wildcard = 'PDF (*.pdf)|*.pdf'
        filename = self.FylDilog(wildcard, msg, wx.FD_SAVE |
                                 wx.FD_OVERWRITE_PROMPT)
        if filename != 'No File':
            # confirm an extension was specifed as pdf if not
            # add pdf extension
            if filename[-4:].lower() != '.pdf':
                filename += '.pdf'

            merger = PdfFileMerger()
            for pdf in self.pdf_list:
                pdf_file = self.path + os.sep + pdf
                merger.append(open(pdf_file, 'rb'))

            with open(filename, "wb") as fout:
                merger.write(fout)

    def OnAdd(self, evt):
        dlg = wx.DirDialog(self, "Choose a directory containing PDF files:",
                           style=wx.DD_DEFAULT_STYLE)

        if dlg.ShowModal() == wx.ID_OK:

            self.path = dlg.GetPath()
            files = [a for a in os.listdir(self.path) if a.endswith(".pdf")]
        dlg.Destroy()
        self.ip.SetItems(files)

    def FylDilog(self, wildcard, msg, stl):
        self.currentDirectory = os.getcwd()
        path = 'No File'
        dlg = wx.FileDialog(
            self, message=msg,
            defaultDir=self.currentDirectory,
            defaultFile="",
            wildcard=wildcard,
            style=stl
            )
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPaths()[0]
        dlg.Destroy()

        return path

    def OnExit(self, evt):
        self.GetParent().Enable(True)   # add for child form
        self.__eventLoop.Exit()        # add for child form
        self.Destroy()


class PDFFrm(wx.Frame):
    def __init__(self, parent, filename=None):
        super(PDFFrm, self).__init__(parent)

        self.Maximize(True)
        self.parent = parent
        self.filename = filename

        self.Bind(wx.EVT_CLOSE, self.OnCloseFrm)
        self.InitUI()

    def InitUI(self):
        from wx.lib.pdfviewer import pdfViewer, pdfButtonPanel

        hsizer = wx.BoxSizer(wx.HORIZONTAL)
        vsizer = wx.BoxSizer(wx.VERTICAL)
        self.buttonpanel = pdfButtonPanel(self, wx.ID_ANY,
                                          wx.DefaultPosition,
                                          wx.DefaultSize, 0)
        vsizer.Add(self.buttonpanel, 0,
                   wx.GROW)
        self.viewer = pdfViewer(self, wx.ID_ANY, wx.DefaultPosition,
                                wx.DefaultSize, wx.HSCROLL |
                                wx.VSCROLL | wx.SUNKEN_BORDER)

        vsizer.Add(self.viewer, 1, wx.GROW)
        loadbutton = wx.Button(self, wx.ID_ANY, "Load PDF file",
                               wx.DefaultPosition, wx.DefaultSize, 0)
        loadbutton.SetForegroundColour((255, 0, 0))
        vsizer.Add(loadbutton, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        hsizer.Add(vsizer, 1, wx.GROW)
        self.SetSizer(hsizer)
        self.SetAutoLayout(True)

        # introduce buttonpanel and viewer to each other
        self.buttonpanel.viewer = self.viewer
        self.viewer.buttonpanel = self.buttonpanel

        self.Bind(wx.EVT_BUTTON, self.OnLoadButton, loadbutton)

        if self.filename is not None:
            self.viewer.LoadFile(self.filename)

        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnLoadButton(self, event):
        dlg = wx.FileDialog(self, wildcard="*.pdf")
        if dlg.ShowModal() == wx.ID_OK:
            wx.BeginBusyCursor()
            self.viewer.LoadFile(dlg.GetPath())
            wx.EndBusyCursor()
        dlg.Destroy()

    def OnCloseFrm(self, evt):
        self.GetParent().Enable(True)   # add for child form
        self.__eventLoop.Exit()        # add for child form
        # Dbase().close_database()
        self.Destroy()


class CalcFrm(wx.Frame):
    def __init__(self, parent):
        self.lctrls = []
        self.parent = parent
        ttl = 'Wall Thickness and Hydro-Test calculation'
        super(CalcFrm, self).__init__(parent, title=ttl,
                                      size=(580, 720),
                                      style=wx.DEFAULT_FRAME_STYLE &
                                      ~ (wx.RESIZE_BORDER |
                                         wx.MAXIMIZE_BOX |
                                         wx.MINIMIZE_BOX))

        self.FrmSizer = wx.BoxSizer(wx.VERTICAL)
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.pnl = CalcPnl(self)
        self.FrmSizer.Add(self.pnl, 1, wx.EXPAND)
        self.FrmSizer.Add((35, 10))
        self.FrmSizer.Add((10, 20))
        self.pnl.b4.Bind(wx.EVT_BUTTON, self.OnClose)
        self.SetSizer(self.FrmSizer)

        # add these 5 following lines to child parent form
        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnClose(self, evt):
        self.GetParent().Enable(True)   # add for child form
        self.__eventLoop.Exit()        # add for child form
        self.Destroy()


class CalcPnl(sc.ScrolledPanel):

    def __init__(self, parent):
        super(CalcPnl, self).__init__(parent, size=(560, 630))

        self.lctrls = []
        self.InitUI()

    def InitUI(self):
        font = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.NORMAL)
        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.BOLD)

        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        DTSizer = wx.BoxSizer(wx.HORIZONTAL)
        DTbox_border = wx.StaticBox(self, -1)
        DTbox = wx.StaticBoxSizer(DTbox_border, wx.VERTICAL)
        text = wx.StaticText(self, label="Allowed Stress")
        text.SetForegroundColour('red')
        DTbox.Add(text, 0, wx.ALIGN_LEFT)
        Stressnote = wx.StaticText(self, label='Sd = ', style=wx.ALIGN_LEFT)
        self.Stresstxt = wx.TextCtrl(self, size=(80, -1), value='',
                                     style=wx.TE_CENTER)
        self.Stresstxt.SetForegroundColour('red')
        self.Stresstxt.SetFont(font)
        DTempnote = wx.StaticText(self, label='psi @ Design Temperature',
                                  style=wx.ALIGN_LEFT)
        DTSizer.Add(DTbox, 0, wx.BOTTOM | wx.LEFT, border=10)
        DTSizer.Add(Stressnote, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL,
                    border=20)
        DTSizer.Add(self.Stresstxt, 0, wx.LEFT | wx.ALIGN_CENTER, border=10)
        DTSizer.Add(DTempnote, 0, wx.LEFT | wx.ALIGN_CENTER, border=25)

        HTSizer = wx.BoxSizer(wx.HORIZONTAL)
        Hydronote = wx.StaticText(self, label='Sh = ', style=wx.ALIGN_LEFT)
        self.Hydrotxt = wx.TextCtrl(self, size=(80, -1), value='',
                                    style=wx.TE_CENTER)
        self.Hydrotxt.SetForegroundColour('red')
        self.Hydrotxt.SetFont(font)
        HTnote = wx.StaticText(self, label='psi @ Hydro Test Temperature',
                               style=wx.ALIGN_LEFT)
        HTSizer.Add(Hydronote, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL,
                    border=135)
        HTSizer.Add(self.Hydrotxt, 0, wx.LEFT | wx.ALIGN_CENTER, border=10)
        HTSizer.Add(HTnote, 0, wx.LEFT | wx.ALIGN_CENTER, border=25)

        # draw a line between upper and lower section
        ln1 = wx.StaticLine(self, 0, size=(560, 2), style=wx.LI_VERTICAL)
        ln1.SetBackgroundColour('Black')

        DataSizer = wx.BoxSizer(wx.HORIZONTAL)
        Databox_border = wx.StaticBox(self, -1)
        Databox = wx.StaticBoxSizer(Databox_border, wx.VERTICAL)
        text = wx.StaticText(self, label="Design Data")
        text.SetForegroundColour('red')
        Databox.Add(text, 0, wx.ALIGN_LEFT)
        DPnote = wx.StaticText(self, label='Design Pressure Pd = ',
                               style=wx.ALIGN_LEFT)
        self.DPtxt = wx.TextCtrl(self, size=(80, -1), value='',
                                 style=wx.TE_CENTER)
        self.DPtxt.SetForegroundColour('red')
        self.DPtxt.SetFont(font)
        Unitsnote = wx.StaticText(self, label='psig', style=wx.ALIGN_LEFT)
        b3 = wx.Button(self, label="View Flange\nRatings")
        self.Bind(wx.EVT_BUTTON, self.OnFlange, b3)
        DataSizer.Add(Databox, 0, wx.BOTTOM | wx.LEFT, border=10)
        DataSizer.Add(DPnote, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, border=20)
        DataSizer.Add(self.DPtxt, 0, wx.LEFT | wx.ALIGN_CENTER, border=10)
        DataSizer.Add(Unitsnote, 0, wx.LEFT | wx.ALIGN_CENTER, border=25)
        DataSizer.Add(b3, 0, wx.LEFT | wx.ALIGN_CENTER, 15)

        # draw a line between upper and lower section
        ln2 = wx.StaticLine(self, 0, size=(560, 2), style=wx.LI_VERTICAL)
        ln2.SetBackgroundColour('Black')

        self.cmbsizer1 = wx.BoxSizer(wx.HORIZONTAL)
        noteCor = wx.StaticText(
            self, label='Corrosion Allowance c (inches) = ',
            style=wx.ALIGN_LEFT)
        noteCor.SetFont(font)
        self.cmbCor = wx.ComboCtrl(self, pos=(10, 10), size=(100, -1),
                                   style=wx.CB_READONLY)
        self.cmbCor.SetPopupControl(ListCtrlComboPopup('CorrosionAllowance',
                                                       showcol=1,
                                                       lctrls=self.lctrls))
        self.cmbsizer1.Add(noteCor, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 15)
        self.cmbsizer1.Add(self.cmbCor, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 5)

        self.cmbsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        noteOD = wx.StaticText(
            self, label='Pipe OD (inches) = ',
            style=wx.ALIGN_LEFT)
        noteOD.SetFont(font)
        noteWall = wx.StaticText(self, label='Wall Thk.',
                                 style=wx.ALIGN_LEFT)
        noteWall.SetFont(font)
        self.cmbOD = wx.ComboCtrl(self, pos=(10, 10),
                                  size=(100, -1), style=wx.CB_READONLY)
        self.cmbOD.SetPopupControl(ListCtrlComboPopup('Pipe_OD',
                                                      showcol=1,
                                                      lctrls=self.lctrls))
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.CutDpth, self.cmbOD)
        FillQuery = (
            "SELECT Wall,Sch FROM PipeDimensions WHERE Size = '" +
            self.cmbOD.GetValue() + "'")
        self.cmbSch = wx.ComboCtrl(self, pos=(10, 10),
                                   size=(100, -1), style=wx.CB_READONLY)
        self.cmbSch.SetPopupControl(
            ListCtrlComboPopup('PipeDimensions', FillQuery,
                               showcol=0,
                               lctrls=self.lctrls))
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.WallThk, self.cmbSch)
        self.cmbSch.Disable()
        self.noteSch = wx.TextCtrl(self, size=(100, -1), value='Sch',
                                   style=wx.TE_READONLY | wx.TE_LEFT)
        self.cmbsizer2.Add(noteOD, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 15)
        self.cmbsizer2.Add(self.cmbOD, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 5)
        self.cmbsizer2.Add(noteWall, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 15)
        self.cmbsizer2.Add(self.cmbSch, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 5)
        self.cmbsizer2.Add(self.noteSch, 0, wx.LEFT |
                           wx.ALIGN_CENTER_VERTICAL, 15)

        radioboxsizer1 = wx.BoxSizer(wx.HORIZONTAL)
        noteDpth = wx.StaticText(self, label='Specify pipe end: ',
                                 style=wx.ALIGN_LEFT)
        noteDpth.SetFont(font)
        rdBoxSzr1 = wx.BoxSizer(wx.VERTICAL)
        lblDpth = wx.StaticText(self, label='Cut Depth Cd',
                                style=wx.ALIGN_LEFT)
        lblDpth.SetFont(font)
        self.txtThrd = wx.TextCtrl(self, size=(80, -1), value='',
                                   style=wx.TE_CENTER | wx.TE_READONLY)
        self.txtThrd.SetFont(font)
        self.txtGrv = wx.TextCtrl(self, size=(80, -1), value='',
                                  style=wx.TE_CENTER | wx.TE_READONLY)
        self.txtGrv.SetFont(font)
        rdBoxSzr1.Add(lblDpth, 0, wx.TOP | wx.LEFT, 10)
        rdBoxSzr1.Add((15, 10))
        rdBoxSzr1.Add(self.txtThrd, 0, wx.TOP | wx.LEFT | wx.BOTTOM, 5)
        rdBoxSzr1.Add(self.txtGrv, 0, wx.ALL, 5)

        self.rdBox1 = wx.RadioBox(self, size=(110, 120),
                                  choices=['Plain', '\nThreaded\n', 'Grooved'],
                                  majorDimension=1, style=wx.RA_SPECIFY_COLS)
        self.rdBox1.SetSelection(0)
        self.Bind(wx.EVT_RADIOBOX, self.CutDpth, self.rdBox1)
        radioboxsizer1.Add(noteDpth, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        radioboxsizer1.Add((15, 10))
        radioboxsizer1.Add(self.rdBox1, 0, wx.ALL | wx.ALIGN_TOP, 5)
        radioboxsizer1.Add((5, 10))
        radioboxsizer1.Add(rdBoxSzr1, 0, wx.TOP | wx.ALIGN_TOP, 8)

        radioboxsizer = wx.BoxSizer(wx.HORIZONTAL)
        noteEff = wx.StaticText(
            self, label='E = Pipe Quality Factor, default values: ',
            style=wx.ALIGN_LEFT)
        noteEff.SetFont(font)
        rdBoxSzr = wx.BoxSizer(wx.VERTICAL)
        self.ERWtxt = wx.TextCtrl(self, size=(80, -1), value='.085',
                                  style=wx.TE_CENTER)
        self.ERWtxt.SetFont(font)
        self.Smlstxt = wx.TextCtrl(self, size=(80, -1), value='1.00',
                                   style=wx.TE_CENTER)
        self.Smlstxt.SetFont(font)
        rdBoxSzr.Add(self.ERWtxt, 0, wx.ALL, 5)
        rdBoxSzr.Add(self.Smlstxt, 0, wx.ALL, 5)

        self.rdBox = wx.RadioBox(self, size=(110, 80),
                                 choices=['ERW', '\nSeamless\n'],
                                 majorDimension=1, style=wx.RA_SPECIFY_COLS)
        self.rdBox.SetSelection(0)
        radioboxsizer.Add(noteEff, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 10)
        radioboxsizer.Add(rdBoxSzr, 0, wx.TOP | wx.ALIGN_TOP, 8)
        radioboxsizer.Add(self.rdBox, 0, wx.ALL | wx.ALIGN_TOP, 5)

        self.StrFactsizer = wx.BoxSizer(wx.HORIZONTAL)
        noteStrsFact = wx.StaticText(
            self, label=('''Y = stress - temperature compensation factor,
                          default value: '''), style=wx.ALIGN_LEFT)
        noteStrsFact.SetFont(font)
        self.StrsFact = wx.TextCtrl(self, size=(80, -1), value='0.40',
                                    style=wx.TE_CENTER)
        self.StrsFact.SetFont(font)
        self.StrFactsizer.Add(noteStrsFact, 0, wx.LEFT |
                              wx.ALIGN_CENTER_VERTICAL, 15)
        self.StrFactsizer.Add(self.StrsFact, 0, wx.ALL, 10)

        # draw a line between upper and lower section
        ln3 = wx.StaticLine(self, 0, size=(560, 2), style=wx.LI_VERTICAL)
        ln3.SetBackgroundColour('Black')

        notesizer = wx.BoxSizer(wx.VERTICAL)
        note1 = wx.StaticText(
            self, label=('''Minimum Pipe Wall Thickness tm = (Pd * Pipe OD) / 2
                           (Sd * E + Pd * Y) + c + Cd'''), style=wx.ALIGN_LEFT)
        note1.SetFont(font)
        note1.SetForegroundColour(('red'))
        notesizer.Add(note1, 0, wx.TOP | wx.LEFT | wx.ALIGN_LEFT, border=10)

        CalcWallszr = wx.BoxSizer(wx.HORIZONTAL)
        self.b1 = wx.Button(self, label="Calculate Minimum Wall Thickness")
        self.b1.SetForegroundColour('green')
        self.Bind(wx.EVT_BUTTON, self.CalcWall, self.b1)
        self.noteCalcWall = wx.StaticText(self, label=' ', style=wx.ALIGN_LEFT)
        self.noteCalcWall.SetFont(font1)
        CalcWallszr.Add(self.b1, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        CalcWallszr.Add(self.noteCalcWall, 0, wx.ALL |
                        wx.ALIGN_CENTER_VERTICAL, 5)

        # draw a line between upper and lower section
        ln4 = wx.StaticLine(self, 0, size=(560, 2), style=wx.LI_VERTICAL)
        ln4.SetBackgroundColour('Black')

        notesizer1 = wx.BoxSizer(wx.VERTICAL)
        note3 = wx.StaticText(
            self, label='Required Hydro Test Pressure = 1.5 * Pd * Sh / Sd',
            style=wx.ALIGN_LEFT)
        note3.SetFont(font)
        note3.SetForegroundColour(('red'))
        notesizer1.Add(note3, 0, wx.TOP | wx.LEFT | wx.ALIGN_LEFT, border=10)

        CalcHydroszr = wx.BoxSizer(wx.HORIZONTAL)
        self.b2 = wx.Button(self, label="Calculate Hydro Test Pressure")
        self.b2.SetForegroundColour('green')
        self.Bind(wx.EVT_BUTTON, self.CalcHydro, self.b2)
        self.noteCalcHydro = wx.StaticText(self, label=' ',
                                           style=wx.ALIGN_LEFT)
        self.noteCalcHydro.SetFont(font1)
        CalcHydroszr.Add(self.b2, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        CalcHydroszr.Add(self.noteCalcHydro, 0, wx.ALL |
                         wx.ALIGN_CENTER_VERTICAL, 5)

        btnbox = wx.BoxSizer(wx.VERTICAL)
        self.b4 = wx.Button(self, label="Exit")
        self.b4.SetForegroundColour('red')
        btnbox.Add(self.b4, 0, wx.RIGHT | wx.ALIGN_CENTER, 15)

        self.Sizer.Add(DTSizer, 0, wx.ALL, 5)
        self.Sizer.Add(HTSizer, 0, wx.ALL, 5)
        self.Sizer.Add(ln1, 0, wx.ALIGN_LEFT)
        self.Sizer.Add(DataSizer, 0, wx.ALL, 5)
        self.Sizer.Add(ln2, 0, wx.ALIGN_LEFT)
        self.Sizer.Add(self.cmbsizer1, 0, wx.TOP | wx.ALIGN_LEFT, 15)
        self.Sizer.Add(self.cmbsizer2, 0, wx.TOP | wx.ALIGN_LEFT, 15)
        self.Sizer.Add(radioboxsizer1, 0, wx.ALL, 5)
        self.Sizer.Add(radioboxsizer, 0, wx.ALL, 5)
        self.Sizer.Add(self.StrFactsizer, 0, wx.ALL | wx.ALIGN_LEFT)
        self.Sizer.Add(ln3, 0, wx.ALIGN_LEFT)
        self.Sizer.Add(notesizer, 0, wx.ALIGN_LEFT)
        self.Sizer.Add(CalcWallszr, 0, wx.ALL, 5)
        self.Sizer.Add(ln4, 0, wx.ALIGN_LEFT)
        self.Sizer.Add(notesizer1, 0, wx.ALIGN_LEFT)
        self.Sizer.Add(CalcHydroszr, 0, wx.ALL, 5)
        self.Sizer.Add(btnbox, 0, wx.ALIGN_RIGHT)
        self.SetupScrolling()

    def CutDpth(self, evt):
        if self.cmbOD.GetValue() != '':
            if self.rdBox1.GetSelection() == 1:
                qry = ("SELECT ThreadDpth FROM PipeDimensions WHERE Size = '"
                       + self.cmbOD.GetValue() + "'")
                self.txtThrd.ChangeValue(Dbase().Dsqldata(qry)[0][0])
                self.txtGrv.ChangeValue('')
                self.Cd = eval(self.txtThrd.GetValue())
            elif self.rdBox1.GetSelection() == 2:
                qry = ("SELECT GrooveDpth FROM PipeDimensions WHERE Size = '"
                       + self.cmbOD.GetValue() + "'")
                Dbase().Dsqldata(qry)
                self.txtGrv.ChangeValue(Dbase().Dsqldata(qry)[0][0])
                self.txtThrd.ChangeValue('')
                self.Cd = eval(self.txtGrv.GetValue())
            elif self.rdBox1.GetSelection() == 0:
                self.txtThrd.ChangeValue('')
                self.txtGrv.ChangeValue('')
                self.Cd = 0

            # fill the schedule drop down based on pipe OD
            ReFillQuery = (
                "SELECT Wall,Sch FROM PipeDimensions WHERE Size = '" +
                self.cmbOD.GetValue() + "'")
            self.cmbSch.Enable()

            lctr = self.lctrls[2]
            lctr.DeleteAllItems()
            index = 0

            for values in Dbase().Dsqldata(ReFillQuery):
                col = 0
                for value in values:
                    if col == 0:
                        lctr.InsertItem(index, str(value))
                    else:
                        lctr.SetItem(index, col, str(value))
                    col += 1
                index += 1

    def WallThk(self, evt):
        qry = ('SELECT Sch FROM PipeDimensions WHERE Wall = "'
               + str(self.cmbSch.GetValue())) + '"'
        self.noteSch.ChangeValue(Dbase().Dsqldata(qry)[0][0])

    def CalcWall(self, evt):
        chck = self.ValWallData()
        if chck != []:
            self.DataWarn(chck)
        else:
            if self.rdBox.GetStringSelection() == 'ERW':
                E = eval(self.ERWtxt.GetValue())
            else:
                E = eval(self.Smlstxt.GetValue())

            Pd = eval(self.DPtxt.GetValue())
            pipeOD = eval(self.cmbOD.GetValue()[:-1].replace('-', '+'))
            Sd = eval(self.Stresstxt.GetValue())
            Y = eval(self.StrsFact.GetValue())
            c = eval(self.cmbCor.GetValue())

            WT1 = Pd * pipeOD
            WT2 = 2 * (Sd * E + pipeOD * Y)
            WT = WT1 / WT2 + c + self.Cd
            self.noteCalcWall.SetLabel(str('{0:5.3f}'.format(WT)) + ' inches')

    def CalcHydro(self, evt):
        chckH = self.ValHydroData()
        if chckH != []:
            self.DataWarn(chckH)
        else:
            Pd = eval(self.DPtxt.GetValue())
            Sh = eval(self.Hydrotxt.GetValue())
            Sd = eval(self.Stresstxt.GetValue())
            HP1 = Sh / Sd
            if Sh > Sd:
                HP1 = 1
            HP = 1.5 * Pd * HP1
            self.noteCalcHydro.SetLabel(str('{0:5.3f}'.format(HP)) + ' psig')

    def ValWallData(self):
        chck = []
        if self.Stresstxt.GetValue() == '':
            chck.append(1)
        if self.DPtxt.GetValue() == '':
            chck.append(3)
        if self.cmbOD.GetValue() == '':
            chck.append(4)
        if self.cmbCor.GetValue() == '':
            chck.append(5)
        if (self.ERWtxt.GetValue() == '' and self.Smlstxt.GetValue() == ''):
            chck.append(6)
        if self.StrsFact.GetValue() == '':
            chck.append(7)
        return chck

    def DataWarn(self, chck):
        bxnames = ['Allowable Stress Sd', 'Allowable Stress Sh',
                   'Design Pressure Pd', 'Pipe OD', 'Corrosion Allowance',
                   'Pipe Quality factor E', 'Temp Compensation factor Y']
        missingstr = ''
        for num in chck:
            missingstr = missingstr + bxnames[num-1] + '\n'
        MsgBx = wx.MessageDialog(self, 'Value(s) needed for;\n' + missingstr,
                                 'Missing Data', wx.OK | wx.ICON_HAND)

        MsgBx_val = MsgBx.ShowModal()
        MsgBx.Destroy()
        if MsgBx_val == wx.ID_OK:
            return False

    def ValHydroData(self):
        chckH = []
        if self.Stresstxt.GetValue() == '':
            chckH.append(1)
        if self.Hydrotxt.GetValue() == '':
            chckH.append(2)
        if self.DPtxt.GetValue() == '':
            chckH.append(3)
        return chckH

    def OnFlange(self, evt):
        FlgRatg(self)

    def OnClose(self, evt):
        self.GetParent().Enable(True)   # add for child form
        self.__eventLoop.Exit()        # add for child form
        self.Destroy()


class FlgRatg(wx.Frame):
    def __init__(self, parent):
        ttl = 'Flange Pressure and Temperature Rating'
        super(FlgRatg, self).__init__(parent, title=ttl,
                                      size=(660, 750),
                                      style=wx.DEFAULT_FRAME_STYLE &
                                      ~ (wx.RESIZE_BORDER |
                                         wx.MAXIMIZE_BOX |
                                         wx.MINIMIZE_BOX))

        self.Bind(wx.EVT_CLOSE, self.OnClose)

        self.parent = parent
        self.tempbxs = []
        self.pressbxs = []
        self.datax = []
        self.datay = []
        self.lctrls = []

        self.pnl = FlgRatgPnl(self)
        self.InitUI()

    def InitUI(self):
        self.mainSizer = wx.BoxSizer(wx.VERTICAL)

        sizer_1 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_1.Add(self.pnl, 1, wx.EXPAND)

        self.tempszr = wx.BoxSizer(wx.HORIZONTAL)
        self.presszr = wx.BoxSizer(wx.HORIZONTAL)
        self.txttemp = wx.StaticText(self, label="Temperature Data:")
        self.txtpress = wx.StaticText(self, label="Pressure Data:")
        self.tempszr.Add(self.txttemp, 0, wx.ALIGN_LEFT |
                         wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 15)
        self.presszr.Add(self.txtpress, 0, wx.ALIGN_LEFT |
                         wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 45)

        btnszr = wx.BoxSizer(wx.HORIZONTAL)
        self.cmbPipe = wx.ComboCtrl(self, size=(90, -1), style=wx.CB_READONLY)
        self.Bind(wx.EVT_COMBOBOX_CLOSEUP, self.OnSelectP, self.cmbPipe)
        self.Bind(wx.EVT_TEXT, self.OnSelectP, self.cmbPipe)
        self.cmbPipe.SetPopupControl(ListCtrlComboPopup(
            'PipeMaterialSpec', showcol=1, PupQuery='', lctrls=self.lctrls))
        self.tmp = wx.TextCtrl(self, name='temp', style=wx.TE_PROCESS_ENTER)
        self.tmp.SetHint('Temperature')
        self.tmp.Enable(False)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnClosePt, self.tmp)
        self.prs = wx.TextCtrl(self, name='press', style=wx.TE_PROCESS_ENTER)
        self.prs.SetHint('Pressure')
        self.prs.Enable(False)
        self.Bind(wx.EVT_TEXT_ENTER, self.OnClosePt, self.prs)
        lblpt = wx.StaticText(self, label='\t<--\nMaximum\n\t-->')
        self.b1 = wx.Button(self, label="Exit")
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b1)
        btnszr.Add(self.cmbPipe, 0, wx.LEFT | wx.RIGHT | wx.ALIGN_CENTER, 25)
        btnszr.Add(self.tmp, 0, wx.ALIGN_CENTER)
        btnszr.Add(lblpt, 0, wx.ALIGN_CENTER_VERTICAL)
        btnszr.Add(self.prs, 0, wx.ALIGN_CENTER)
        btnszr.Add(self.b1, 0, wx.LEFT | wx.ALIGN_CENTER, 25)

        lblszr = wx.BoxSizer(wx.HORIZONTAL)
        lbltmp = wx.StaticText(self, label='Temperature')
        lblprs = wx.StaticText(self, label='Pressure')
        lblszr.Add(lbltmp, 0, wx.ALIGN_TOP)
        lblszr.Add((65, 10))
        lblszr.Add(lblprs, 0, wx.ALIGN_TOP)

        self.mainSizer.Add(sizer_1, 0, wx.EXPAND)
        self.mainSizer.Add((50, 25))
        self.mainSizer.Add(self.tempszr, 0,
                           wx.EXPAND | wx.ALIGN_LEFT)
        self.mainSizer.Add((50, 25))
        self.mainSizer.Add(self.presszr, 0, wx.ALIGN_LEFT | wx.BOTTOM, 20)
        self.mainSizer.Add(btnszr, 0, wx.ALIGN_CENTER)
        self.mainSizer.Add(lblszr, 0, wx.ALIGN_CENTER)

        self.SetSizer(self.mainSizer)
        self.Layout()

        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def Update_pnl(self):
        self.pnl.draw(self.datax, self.datay)

    def OnSelectP(self, evt):
        code = self.cmbPipe.GetValue()
        data1 = []
        if code != '':
            qry = ('''SELECT Material_Spec_ID FROM PipeMaterialSpec
                    WHERE Pipe_Material_Spec = "''' +
                   self.cmbPipe.GetValue() + '"')
            id = Dbase().Dsqldata(qry)[0][0]
            qry = ('''SELECT Temperature, Pressure FROM PressTempTables
                    WHERE Specification_Number = ''' + str(id))
            data1 = Dbase().Dsqldata(qry)
        if data1 != []:
            sortdata = sorted(data1, key=lambda tup: tup[0])
            self.datax, self.datay = map(list, zip(*sortdata))
            if len(self.datax) <= 10:
                nr = 10
            else:
                nr = len(self.datax)
            self.SetSize(66*nr, 750)
            self.Update_pnl()
            self.pnl.axes.set_title("Material Code " + code)
            self.tblfill()
            self.prs.Enable()
            self.tmp.Enable()
        else:
            self.pnl.axes.set_title("Material Code not setup for " + code)
            self.tblfill()
            n = 0
            for item in self.pressbxs:
                item.ChangeValue('')
                self.tempbxs[n].ChangeValue('')
                n += 1

    def tblfill(self):
        for n in range(len(self.tempbxs)):
            self.tempbxs[n].Destroy()
            self.pressbxs[n].Destroy()
        self.tempbxs = []
        self.pressbxs = []
        self.tempszr.Clear()
        self.presszr.Clear()
        self.txtpress.SetLabel('')
        self.txttemp.SetLabel('')
        self.txttemp = wx.StaticText(self, label="Temperature Data:")
        self.txtpress = wx.StaticText(self, label="Pressure Data:")
        self.tempszr.Add(self.txttemp, 0, wx.ALIGN_LEFT |
                         wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 15)
        self.presszr.Add(self.txtpress, 0, wx.ALIGN_LEFT |
                         wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 45)
        for n in range(len(self.datax)):
            tempbx = wx.TextCtrl(self, value=str(self.datax[n]), size=(50, 33),
                                 style=wx.TE_READONLY | wx.TE_RIGHT)
            pressbx = wx.TextCtrl(self, value=str(self.datay[n]),
                                  size=(50, 33),
                                  style=wx.TE_READONLY | wx.TE_RIGHT)
            self.tempszr.Add(tempbx, 0, wx.ALIGN_LEFT)
            self.presszr.Add(pressbx, 0, wx.ALIGN_LEFT)
            self.tempbxs.append(tempbx)
            self.pressbxs.append(pressbx)
        self.Layout()

    def OnClosePt(self, evt):
        # call closept calculate the value coressponding to the specified value
        # indicate the value of the point and the box in which it was specified
        maxvl = self.closept(evt.GetEventObject().GetValue(),
                             evt.GetEventObject().GetName())
        # place the calculated value in the proper text box
        if evt.GetEventObject().GetName() == 'temp':
            self.prs.ChangeValue(str(maxvl))
        else:
            self.tmp.ChangeValue(str(maxvl))

    def closept(self, pt, axs):
        # find closest number to pt
        pt = int(pt)
        if axs == 'temp':
            dt = self.datax
            # if the point specified is an existing data point
            # just return corresponding value
            if pt in dt:
                indx = dt.index(pt)
                maxvl = self.datay[indx]
                return maxvl
            # if the specified point is outside the range
            # of data do not extrapolate
            if pt < min(self.datax) or pt > max(self.datax):
                maxvl = 'none'
                return maxvl
        else:
            dt = self.datay
            # if the point specified is an existing data point
            # just return corresponding value
            if pt in dt:
                indx = dt.index(pt)
                maxvl = self.datax[indx]
                return maxvl
            # if the specified point is outside the range
            # of data do not extrapolate
            if pt < min(self.datay) or pt > max(self.datay):
                maxvl = 'none'
                return maxvl

        nrpt = min(dt, key=lambda x: abs(x - pt))
        # find the index value of the near point
        indx1 = dt.index(nrpt)
        # set the index value of the point to the other side of your number
        if nrpt > pt:
            indx2 = indx1 - 1
        else:
            indx2 = indx1 + 1

        # set up the coefficents needed
        a = self.datay[indx2] - self.datay[indx1]
        b = self.datax[indx2] - self.datax[indx1]
        c = (self.datax[indx2]*self.datay[indx1] -
             self.datax[indx1]*self.datay[indx2])
        # based on the line between the two points being linear
        # calculate point of intersection
        if axs == 'temp':
            maxvl = a*pt/b + c/b
        else:
            maxvl = pt*b/a - c/a
        return round(maxvl, 1)

    def OnClose(self, evt):
        self.GetParent().Enable(True)   # add this line for child
        self.__eventLoop.Exit()     # add this line for child
        self.Destroy()


class FlgRatgPnl(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        self.figure = Figure()
        self.axes = self.figure.add_subplot(111)
        self.axes.grid(color='b', linestyle='-', linewidth=.2)
        self.canvas = FigureCanvas(self, -1, self.figure)
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.canvas, 1, wx.LEFT | wx.TOP | wx.GROW)
        self.SetSizer(self.sizer)
        self.Fit()

    def draw(self, datax, datay):
        temp = datax
        press = datay
        self.axes.clear()
        self.axes.grid(color='b', linestyle='-', linewidth=.2)
        self.axes.set_xticks(temp)
        self.axes.set_yticks(press)
        self.axes.set_ylabel('Pressure (psig)')
        self.axes.set_xlabel('Temperature (F)')
        self.axes.plot(temp, press, marker='.')
        self.canvas.draw()


class BldTrvlSht(wx.Frame):
    '''Routine to build form and populate grid'''
    def __init__(self, parent, model=None):

        self.parent = parent
        self.model = model
        self.Lvl2tbl = 'InspectionTravelSheet'
        self.NoteStr = []

        self.ComCode = ''
        self.PipeMtrSpec = ''

        if self.Lvl2tbl.find("_") != -1:
            frmtitle = (self.Lvl2tbl.replace("_", " "))
        else:
            frmtitle = (' '.join(re.findall('([A-Z][a-z]*)', self.Lvl2tbl)))

        super(BldTrvlSht, self).__init__(parent, title=frmtitle,
                                         size=(1200, 900),
                                         style=wx.DEFAULT_FRAME_STYLE &
                                         ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX |
                                           wx.MINIMIZE_BOX | wx.CLOSE_BOX))

        self.InitUI()

    def InitUI(self):
        model1 = None
        font1 = wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.FONTWEIGHT_BOLD)

        link_fld = 'Timing'
        self.frg_tbl = 'TrvlShtTime'
        self.frg_fld = 'TimeID'
        frgn_col = 'Timing'
        self.columnames = ['ID', 'Timing', 'Note']

        realnames = []
        datatable_str = ''

        # here we get the information needed in the report
        # and for the SQL from the Lvl2 table
        # and determine report column width based either
        # on data type or specified field size
        n = 0
        ColWdth = []
        for item in Dbase().Dcolinfo(self.Lvl2tbl):
            col_wdth = ''
        # check to see if field length is specified if
        # so use it to set grid col width
            for s in re.findall(r'\d+', item[2]):
                if s.isdigit():
                    col_wdth = int(s)
                    ColWdth.append(col_wdth)
            realnames.append(item[1])
            if item[5] == 1:
                self.pk_Name = item[1]
                self.pk_col = n
                if 'INTEGER' in item[2]:
                    self.autoincrement = True
                    if col_wdth == '':
                        ColWdth.append(6)
                # include the primary key and table name into SELECT statement
                datatable_str = (datatable_str + self.Lvl2tbl + '.'
                                 + self.pk_Name + ',')
            # need to make frgn_fld column noneditable in DVC
            elif 'INTEGER' in item[2] or 'FLOAT' in item[2]:
                if col_wdth == '' and 'FLOAT' in item[2]:
                    ColWdth.append(10)
                elif col_wdth == '':
                    ColWdth.append(6)
            elif 'BLOB' in item[2]:
                if col_wdth == '':
                    ColWdth.append(30)
            elif 'TEXT' in item[2] or 'BOOLEAN' in item[2]:
                if col_wdth == '':
                    ColWdth.append(10)
            elif 'DATE' in item[2]:
                if col_wdth == '':
                    ColWdth.append(10)

            # get first Lvl2 datatable column name in item[1]
            # check to see if name is lvl2 primary key or lvl1 linked field
            # if they are not then add tablename and
            # datafield to SELECT statement
            if item[1] != link_fld and item[1] != self.pk_Name:
                datatable_str = (datatable_str + ' ' + self.Lvl2tbl +
                                 '.' + item[1] + ',')
            elif item[1] == link_fld:
                datatable_str = (datatable_str + ' ' + self.frg_tbl +
                                 '.' + frgn_col + ',')

            n += 1

        self.realnames = realnames
        self.ColWdth = ColWdth

        datatable_str = datatable_str[:-1]

        DsqlLvl2 = ('SELECT ' + datatable_str + ' FROM ' + self.Lvl2tbl
                    + ' INNER JOIN ' + self.frg_tbl)
        DsqlLvl2 = (DsqlLvl2 + ' ON ' + self.Lvl2tbl + '.' + link_fld
                    + ' = ' + self.frg_tbl + '.' + self.frg_fld)

        # specify data for upper dvc with all the notes in databasse
        self.DsqlLvl2 = DsqlLvl2

        self.data = Dbase().Dsqldata(self.DsqlLvl2)

        self.data1 = []

        # Create the dataview for the commocity property notes (lower table)
        self.dvc1 = dv.DataViewCtrl(self, wx.ID_ANY, wx.DefaultPosition,
                                    wx.Size(500, 150),
                                    style=wx.BORDER_THEME
                                    | dv.DV_ROW_LINES
                                    | dv.DV_VERT_RULES
                                    | dv.DV_HORIZ_RULES
                                    | dv.DV_MULTIPLE
                                    )

        self.dvc1.SetMinSize = (wx.Size(100, 200))
        self.dvc1.SetMaxSize = (wx.Size(500, 400))

        if model1 is None:
            self.model1 = DataMods(self.Lvl2tbl, self.data1)
        else:
            self.model1 = model1
        self.dvc1.AssociateModel(self.model1)

        # specify which listbox column to display in the combobox
        self.showcol = int

        # Create a dataview control for all the database notes (upper table)
        self.dvc = dv.DataViewCtrl(self, wx.ID_ANY, wx.DefaultPosition,
                                   wx.Size(500, 300),
                                   style=wx.BORDER_THEME
                                   | dv.DV_ROW_LINES
                                   | dv.DV_VERT_RULES
                                   | dv.DV_HORIZ_RULES
                                   | dv.DV_MULTIPLE
                                   )

        self.dvc.SetMinSize = (wx.Size(100, 200))
        self.dvc.SetMaxSize = (wx.Size(500, 400))

    # if autoincrement is false then the data can be sorted based on ID_col
        if self.autoincrement == 0:
            self.data.sort(key=lambda tup: tup[self.pk_col])

    # use the sorted data to load the dataviewlistcontrol
        if self.model is None:
            self.model = DataMods(self.Lvl2tbl, self.data)
        self.dvc.AssociateModel(self.model)

        n = 0
        for colname in self.columnames:
            col_mode = dv.DATAVIEW_CELL_INERT
            self.dvc.AppendTextColumn(colname, n,
                                      width=wx.LIST_AUTOSIZE_USEHEADER,
                                      mode=col_mode)

            self.dvc1.AppendTextColumn(colname, n,
                                       width=wx.LIST_AUTOSIZE_USEHEADER,
                                       mode=col_mode)
            n += 1

        # make columns not sortable and but reorderable.
        n = 0
        for c in self.dvc.Columns:
            c.Sortable = False
            # make the category column sortable
            if n == 1:
                c.Sortable = True
            c.Reorderable = True
            c.Resizeable = True
            n += 1

        # change to not let the ID col be moved.
        self.dvc.Columns[(self.pk_col)].Reorderable = False
        self.dvc.Columns[(self.pk_col)].Resizeable = False

        # set the Sizer property (same as SetSizer)
        self.Sizer = wx.BoxSizer(wx.VERTICAL)

        # develope the comboctrl and attach popup list
        self.cmb1 = wx.ComboCtrl(self, pos=(10, 10), size=(200, -1))
        self.Bind(wx.EVT_TEXT, self.OnSelect, self.cmb1)
        self.cmb1.SetHint(frgn_col)
        self.showcol = 1
        self.popup = ListCtrlComboPopup(self.frg_tbl, showcol=self.showcol)

        self.cmbsizer = wx.BoxSizer(wx.HORIZONTAL)

        # add a button to call main form to search combo list data
        self.b6 = wx.Button(self, label="Restore Data")
        self.Bind(wx.EVT_BUTTON, self.OnRestore, self.b6)
        self.cmbsizer.Add(self.b6, 0, wx.ALIGN_LEFT)

        self.cmb1.SetPopupControl(self.popup)
        self.cmbsizer.Add(self.cmb1, 0, wx.ALIGN_TOP)

        # add a button to call main form to search combo list data
        self.b5 = wx.Button(self, label="<= Search Data")
        self.Bind(wx.EVT_BUTTON, self.OnSearch, self.b5)
        self.cmbsizer.Add(self.b5, 0, wx.BOTTOM, 15)

        self.b9 = wx.Button(self, label='Add Note(s) to\nTravel Sheet Report')
        self.Bind(wx.EVT_BUTTON, self.OnSelectDone, self.b9)
        self.cmbsizer.Add(20, 10)
        self.cmbsizer.Add(self.b9, 0, wx.BOTTOM, 15)

        self.addlbl = wx.StaticText(self, -1, style=wx.ALIGN_CENTER_HORIZONTAL)
        txt = '   Complete Listing of All Inspection Travel Sheet Notes'
        self.addlbl.SetLabel(txt)
        self.addlbl.SetForegroundColour((255, 0, 0))
        self.addlbl.SetFont(font1)
        self.Sizer.Add(self.addlbl, 0, wx.ALIGN_LEFT)
        self.Sizer.Add((10, 20))
        self.Sizer.Add(self.cmbsizer, 0, wx.ALIGN_CENTER)
        self.Sizer.Add(self.dvc, 1, wx.EXPAND)

        self.cmbsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        # add a button to call main form to search combo list data
        self.b10 = wx.Button(self, label="Restore Data")
        self.Bind(wx.EVT_BUTTON, self.OnRstrDVC1, self.b10)
        self.cmbsizer2.Add(self.b10, 0)

        # develope the comboctrl and attach popup list
        self.cmb10 = wx.ComboCtrl(self, pos=(10, 10), size=(200, -1))
        self.Bind(wx.EVT_TEXT, self.OnSelect, self.cmb10)
        self.cmb10.SetHint('Note Category')
        self.showcol = 1
        self.popup = ListCtrlComboPopup(self.frg_tbl, showcol=self.showcol)
        self.cmb10.SetPopupControl(self.popup)
        self.cmbsizer2.Add(self.cmb10, 0, wx.ALIGN_TOP)

        # add a button to call main form to search combo list data
        self.b11 = wx.Button(self, label="<= Search Data")
        self.Bind(wx.EVT_BUTTON, self.OnSrchDVC1, self.b11)
        self.cmbsizer2.Add(self.b11, 0)

        self.b8 = wx.Button(self, id=2, label="Print Inspection\nTravel Sheet")
        self.Bind(wx.EVT_BUTTON, self.PrintFile, self.b8)
        self.cmbsizer2.Add(self.b8, 0)

        self.b12 = wx.Button(
            self, label='Remove Inspection\nNote from List')
        self.Bind(wx.EVT_BUTTON, self.OnRmvNote, self.b12)
        self.cmbsizer2.Add((30, 10))
        self.cmbsizer2.Add(self.b12, 0)

        self.b4 = wx.Button(self, label="Exit")
        self.b4.SetForegroundColour('red')
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.b4)
        self.cmbsizer2.Add((60, 10))
        self.cmbsizer2.Add(self.b4, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        self.Sizer.Add((10, 20))
        self.Sizer.Add(self.dvc1, 1, wx.EXPAND)
        self.Sizer.Add((10, 15))
        self.Sizer.Add(self.cmbsizer2, 0, wx.ALIGN_CENTER)
        self.Sizer.Add((10, 20))

        self.CenterOnParent()
        self.GetParent().Enable(False)
        self.Show(True)
        self.__eventLoop = wx.GUIEventLoop()
        self.__eventLoop.Run()

    def OnClose(self, evt):
        self.GetParent().Enable(True)
        self.__eventLoop.Exit()
        # Dbase().close_database()
        self.Destroy()

    def PrintFile(self, evt):
        import Report_Lvl2

        colwdths = [5, 20, 70, 5, 5]
        colnames = ['QC', 'Timing', 'Note', 'Hold', 'Pass']
        ttl = 'Inspection Travel Sheet'
        rptdata = []

        # confirm there is data to print
        if self.data1 == []:
            NoData = wx.MessageDialog(
                None, 'No Items Selected to Print', 'Error', wx.OK |
                wx.ICON_EXCLAMATION)
            NoData.ShowModal()
            return
        else:
            for item in self.data1:
                item = list(item)
                item[0] = ''
                item.append('')
                item.append('')
                rptdata.append(item)

        filename = self.ReportName()

        Report_Lvl2.Report(self.Lvl2tbl, rptdata, colnames,
                           colwdths, filename, ttl).create_pdf()

    def ReportName(self):

        saveDialog = wx.FileDialog(self, message='Save Report as PDF.',
                                   wildcard='PDF (*.pdf)|*.pdf',
                                   style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        if saveDialog.ShowModal() == wx.ID_CANCEL:
            filename = ''
        filename = saveDialog.GetPath()
        if filename.find(".pdf") == -1:
            filename = filename + '.pdf'
        saveDialog.Destroy()

        if not filename:
            exit()

        return filename

    def OnSelectDone(self, evt):
        # add note(s) selected from upper table to lower table
        # rowID is table index value
        NoteIds = []
        items = self.dvc.GetSelections()
        for item in items:
            row = self.model.GetRow(item)
            rowID = self.model.GetValueByRow(row, self.pk_col)
            if int(rowID) in [self.data1[i][0]
                              for i in range(len(self.data1))]:
                continue
            NoteIds.append(rowID)
        if self.NoteStr == []:
            self.NoteStr = str(tuple(NoteIds))
        else:
            self.NoteStr = self.NoteStr[:-1] + ', ' + str(tuple(NoteIds))[1:]

        self.NoteStr = "".join(self.NoteStr.split())
        if len(NoteIds) <= 1:
            self.NoteStr = self.NoteStr[:-2] + ')'

        self.Dsql1 = (self.DsqlLvl2 + ' WHERE ' + self.pk_Name +
                      ' IN ' + self.NoteStr)
        self.data1 = Dbase().Dsqldata(self.Dsql1)
        self.model1 = DataMods(self.Lvl2tbl, self.data1)
        self.dvc1.AssociateModel(self.model1)
        self.dvc1.Refresh

    def OnRmvNote(self, evt):
        # remove selected note(s) from Commodity property
        # rowID is table index value
        NoteIds = []
        items = self.dvc1.GetSelections()
        for item in items:
            row = self.model1.GetRow(item)
            rowID = self.model1.GetValueByRow(row, self.pk_col)
            NoteIds.append(rowID)

        NoteLst = self.ConvertStr_Lst(self.NoteStr)

        for i in NoteIds:
            NoteLst.remove(i)

        self.NoteStr = str(tuple(NoteLst))
        if self.NoteStr[-2] == ',':
            self.NoteStr = self.NoteStr[:-2] + ')'
        elif self.NoteStr == '()':
            self.NoteStr = None

        if self.NoteStr is not None:
            self.Dsql1 = (self.DsqlLvl2 + ' WHERE ' + self.pk_Name +
                          ' IN ' + self.NoteStr)
            self.data1 = Dbase().Dsqldata(self.Dsql1)
        else:
            self.data1 = []
        self.model1 = DataMods(self.Lvl2tbl, self.data1)
        self.dvc1.AssociateModel(self.model1)
        self.dvc1.Refresh

    def ConvertStr_Lst(self, vals):
        lst = vals.split("'")
        newlst = []
        for i in range(len(lst)):
            if lst[i].isdigit():
                newlst.append(lst[i])
        return newlst

    def OnSrchDVC1(self, evt):
        # collect feign table info
        frgn_info = Dbase().Dtbldata(self.Lvl2tbl)
        field = frgn_info[0][4]
        frg_tbl1 = frgn_info[0][2]

        # do search of string value from combobox
        # equal to value in the self.frg_tbl
        Shcol = Dbase().Dcolinfo(frg_tbl1)[self.showcol][1]
        ShQry = ("SELECT " + field + " FROM " + frg_tbl1 + " WHERE " +
                 Shcol + " LIKE '%" + self.cmb10.GetValue() +
                 "%' COLLATE NOCASE")
        ShQryVal = str(Dbase().Dsqldata(ShQry)[0][0])
        # append the found frgn_fld to the original data grid SQL and
        # find only records equal to the combo selection
        ShQry = (self.Dsql1 + ' AND ' + frg_tbl1 + '.' +
                 field + ' = "' + ShQryVal + '"')
        OSdata = Dbase().Search(ShQry)
        # if nothing is found show blank grid
        if OSdata is False:
            OSdata = []
        self.model1 = DataMods(self.Lvl2tbl, OSdata)
        self.dvc1.AssociateModel(self.model1)
        self.dvc1.Refresh

    def OnRstrDVC1(self, evt):
        ORdata1 = Dbase().Restore(self.Dsql1)
        self.cmb10.ChangeValue('')
        self.cmb10.SetHint(self.frg_fld)
        self.model1 = DataMods(self.Lvl2tbl, ORdata1)
        self.dvc1._AssociateModel(self.model1)
        self.dvc1.Refresh

    def OnSearch(self, evt):
        # collect feign table info
        frgn_info = Dbase().Dtbldata(self.Lvl2tbl)
        field = frgn_info[0][4]
        self.frg_tbl = frgn_info[0][2]

        # do search of string value from combobox equal
        # to value in the self.frg_tbl
        Shcol = Dbase().Dcolinfo(self.frg_tbl)[self.showcol][1]
        ShQuery = ('SELECT ' + field + ' FROM ' + self.frg_tbl +
                   " WHERE " + Shcol + " LIKE '%" + self.cmb1.GetValue()
                   + "%' COLLATE NOCASE")
        ShQueryVal = str(Dbase().Dsqldata(ShQuery)[0][0])
        # append the found frgn_fld to the original data grid SQL and
        # find only records equal to the combo selection
        ShQuery = (self.DsqlLvl2 + ' WHERE ' + self.frg_tbl + '.' +
                   field + ' = "' + ShQueryVal + '"')

        OSdata = Dbase().Search(ShQuery)
        # if nothing is found show blank grid
        if OSdata is False:
            OSdata = []
        self.model = DataMods(self.Lvl2tbl, OSdata)
        self.dvc.AssociateModel(self.model)
        self.dvc.Refresh
    #    self.b2.Enable()

    def OnSelect(self, evt):
        # self.b2.Enable()
        txt = ('''To complete adding a new record click "Add Row".\nThen
                edit data by double click on the cell.''')
        self.editlbl.SetLabel(txt)
        self.editlbl.SetForegroundColour((255, 0, 0))
        self.Sizer.Layout()

    def OnRestore(self, evt):
        self.ORdata = Dbase().Restore(self.DsqlLvl2)
        self.cmb1.ChangeValue('')
        self.cmb1.SetHint(self.frg_fld)
        self.model = DataMods(self.Lvl2tbl, self.ORdata)
        self.dvc._AssociateModel(self.model)
        self.dvc.Refresh


class Conduit:
    # used to print the tubing and piping informatio for a commodity
    def __init__(self, condtbl, ComdPrtyID, ComCode, MtrSpc,
                 filename, PipeCode=None):

        self.condtbl = condtbl
        self.ComdPrtyID = ComdPrtyID
        self.filename = filename

        # specify a title for the report if table name
        # is not to be used
        if PipeCode == 'None':
            ttl = ComCode + '-' + MtrSpc
        else:
            ttl = PipeCode
        self.ttl = ' for ' + ttl

    def CondBld(self):
        import Rpt_Conduit_All

        Colnames = []
        Colwdths = []
        rptdata = []
        # default values for column width, headers and title
        for tblrpt in self.condtbl:
            if tblrpt == 'Piping':
                Colns = [('Pipe ID', 'Pipe Material Specification'),
                         ('Minimum\nPipe OD', 'Maximum\nPipe OD',
                          'Pipe\nMaterial', 'Pipe\nSchedule'),
                         ('Minimum\nNipple OD', 'Maximum\nNipple OD',
                          'Nipple\nMaterial', 'Nipple\nSchedule',
                          'End\nConnection')]

                wdths = [(10, 10), (10, 10, 20, 10), (10, 10, 20, 10, 10)]
                pipedata = self.tbldata(tblrpt, 'PipingID', 'ComponentPack_ID')
                rptdt = self.ReportdataPipe(pipedata)

            if tblrpt == 'Tubing':
                Colns = [('Tube ID', 'Notes'),
                         ('Tube\nSize', 'Tube Wall\nThickness',
                          'Tube\nMaterial'),
                         ('Valve\nType', 'Valve Body\nMaterial',
                          'Valve\nManufacture', 'Model\nNumber')]

                wdths = [(10, 40), (10, 10, 20), (15, 20, 20, 15)]
                tubedata = self.tbldata(tblrpt, 'TubeID', 'Tube_ID')
                rptdt = self.ReportdataTube(tubedata)

            rptdata.append(rptdt)
            Colnames.append(Colns)
            Colwdths.append(wdths)

        Rpt_Conduit_All.Report(self.condtbl, rptdata, Colnames,
                               Colwdths, self.ttl, self.filename).create_pdf()

        if rptdata != []:
            return False

    def tbldata(self, tblrpt, field, field2):
        query = ('SELECT ' + field2 +
                 ' FROM PipeSpecification WHERE Commodity_Property_ID = '
                 + str(self.ComdPrtyID))
        chk = Dbase().Dsqldata(query)
        # get the information pipe specification related to the commodity
        if chk != []:  # if the item is set up with a
            # pipespec then get the items information
            StartQry = chk[0][0]
            if StartQry is not None:
                qry = ('SELECT * FROM ' + tblrpt + ' WHERE '
                       + field + ' = ' + str(StartQry))
                tbldata = Dbase().Dsqldata(qry)
                return tbldata

    def ReportdataTube(self, tubedata):
        rptdt = []
        if tubedata is not None:
            for segn in range(0, len(tubedata)):
                data1 = [(tubedata[segn][i]) for i in [0, 18]]
                data2 = [x for x in list(tubedata[segn][1:9])
                         if x is not None]
                data3 = [x for x in list(tubedata[segn][10:17])
                         if x is not None]

                n = 0
                for i in data2:
                    # first column numbers
                    if n in [m*3 for m in range(0, 3)]:
                        qry = ("SELECT OD FROM TubeSize WHERE SizeID = "
                               + str(i))
                    # second column numbers
                    elif n in [m*3+1 for m in range(0, 3)]:
                        qry = ("SELECT Wall FROM TubeWall WHERE WallID = "
                               + str(i))
                    # third column numbers
                    elif n in [m*3+2 for m in range(0, 3)]:
                        qry = ('''SELECT Material FROM TubeMaterial
                                WHERE ID = ''' + str(i))
                    data2[n] = str(Dbase().Dsqldata(qry)[0][0])
                    n += 1

                m = 0
                for i in data3:
                    # first column numbers
                    if m in [p*4 for p in range(0, 2)]:
                        qry = ('''SELECT Tube_Valve_Body_Material FROM
                                TubeValveMatr WHERE TVID = ''' + str(i))
                        data3[m] = str(Dbase().Dsqldata(qry)[0][0])
                    # second column numbers
                    elif m in [p*4+1 for p in range(0, 2)]:
                        data3[m] = i
                    # third column numbers
                    elif m in [p*4+2 for p in range(0, 2)]:
                        data3[m] = i
                    # fourth column numbers
                    elif m in [p*4+3 for p in range(0, 2)]:
                        data3[m] = i

                    m += 1

                rptdata2 = []
                rptdata3 = []
                for n in range(0, 3):
                    if data2 != []:
                        rptdata2.append(tuple(data2[n*3:(n+1)*3]))
                for m in range(0, 2):
                    if data3 != []:
                        rptdata3.append(tuple(data3[m*4:(m+1)*4]))

                rptdt.append(data1)
                rptdt.append(rptdata2)
                rptdt.append(rptdata3)

            return rptdt

    def ReportdataPipe(self, pipedata):
        rptdt = []
        if pipedata is not None:
            for segn in range(0, len(pipedata)):
                data1 = list(pipedata[segn][0:2])
                data2 = [x for x in list(pipedata[segn][2:18])
                         if x is not None]
                data3 = [x for x in list(pipedata[segn][18:28])
                         if x is not None]

                # determine the matr spec designation used for all tables
                query = ('''SELECT Pipe_Material_Spec FROM PipeMaterialSpec
                        WHERE Material_Spec_ID = ''' + str(data1[1]))
                data1[1] = Dbase().Dsqldata(query)[0][0]

                n = 0
                for i in data2:
                    # first column numbers
                    if n in [m*4 for m in range(0, 4)]:
                        qry = ("SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = "
                               + str(i))
                    # second column numbers
                    elif n in [m*4+1 for m in range(0, 4)]:
                        qry = ("SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = "
                               + str(i))
                    # third column numbers
                    elif n in [m*4+2 for m in range(0, 4)]:
                        qry = ('''SELECT Pipe_Material FROM PipeMaterial
                                WHERE ID = ''' + str(i))
                    # fourth column numbers
                    elif n in [m*4+3 for m in range(0, 4)]:
                        qry = ('''SELECT Pipe_Schedule FROM PipeSchedule
                                WHERE ID = ''' + str(i))
                    data2[n] = str(Dbase().Dsqldata(qry)[0][0])
                    n += 1

                m = 0
                for i in data3:
                    # first column numbers
                    if m in [p*5 for p in range(0, 2)]:
                        qry = ("SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = "
                               + str(i))
                    # second column numbers
                    elif m in [p*5+1 for p in range(0, 2)]:
                        qry = ("SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = "
                               + str(i))
                    # third column numbers
                    elif m in [p*5+2 for p in range(0, 2)]:
                        qry = ('''SELECT Pipe_Material FROM PipeMaterial
                                WHERE ID = ''' + str(i))
                    # fourth column numbers
                    elif m in [p*5+3 for p in range(0, 2)]:
                        qry = ('''SELECT Pipe_Schedule FROM PipeSchedule
                                WHERE ID = ''' + str(i))
                    # fifth column numbers
                    elif m in [p*5+4 for p in range(0, 2)]:
                        qry = ('''SELECT Connection FROM EndConnects
                                WHERE EndID = ''' + str(i))
                    data3[m] = str(Dbase().Dsqldata(qry)[0][0])
                    m += 1

                rptdata2 = []
                rptdata3 = []
                for n in range(0, 4):
                    if data2 != []:
                        rptdata2.append(tuple(data2[n*4:(n+1)*4]))
                for m in range(0, 2):
                    if data3 != []:
                        rptdata3.append(tuple(data3[m*5:(m+1)*5]))

                rptdt.append(data1)
                rptdt.append(rptdata2)
                rptdt.append(rptdata3)
            return rptdt


class Fabricate:
    def __init__(self, fabtbl, ComdPrtyID, ComCode, MtrSpc,
                 filename, PipeCode=None):

        self.fabtbl = fabtbl
        self.data = []
        self.ComdPrtyID = ComdPrtyID
        self.PipeCode = PipeCode
        self.ComCode = ComCode
        self.MtrSpc = MtrSpc
        self.filename = filename

    def FabBld(self):
        import Rpt_Fab_All

        Colnames = []
        Colwdths = []
        RptData = []
        exclud = []
        newtbl = []
        # go through each of the tables selected on the main form
        for tblrpt in self.fabtbl:
            self.frgtbl2 = None

            # add pipecode or comd code plus mtrspec to the title
            if self.PipeCode == 'None':
                ttl = self.ComCode + '-' + self.MtrSpc
            else:
                ttl = self.PipeCode

            # add table information depending on the selected table
            if tblrpt == 'InspectionPacks':
                IDcol = 'Inspection'
                frgtbl1 = 'FluidCategory'
                cmb1fld = 'Fluid_ID'
                cmb1val = 'Designation'
                RptColWdth = [6, 20, 30, 30]
                RptColNames = ['ID', 'Fluid Category',
                               'Enhanced Inspection', 'Notes']
            elif tblrpt == 'GasketPacks':
                IDcol = 'Gasket'
                frgtbl1 = 'ANSI_Rating'
                cmb1fld = 'Rating_Designation'
                cmb1val = 'ANSI_Class'
                RptColWdth = [6, 10, 30, 30]
                RptColNames = ['ID', 'ANSI Class', 'Description', 'Notes']
            elif tblrpt == 'PaintSpec':
                IDcol = 'PaintSpec'
                frgtbl1 = 'PaintPrep'
                cmb1fld = 'PaintPrepID'
                cmb1val = 'Sandblast_Spec'
                frgtbl2 = 'PaintColors'
                cmb2fld = 'ColorCodeID'
                cmb2val = 'Color'
                RptColWdth = [6, 15, 10, 10, 15, 15, 30]
                RptColNames = ['ID', 'Surface Prep', 'Base Coat',
                               'Final Coat', 'Color', 'Tagging', 'Notes']

            query = ('SELECT ' + IDcol +
                     '''_ID FROM PipeSpecification WHERE
                      Commodity_Property_ID = '''
                     + str(self.ComdPrtyID))
            chk = Dbase().Dsqldata(query)
            # check to see fab table has been specified in pipespec table
            # under commodity property
            if chk != []:
                StartQry = chk[0][0]
                # get the fab table information based on ID in pipespec table
                if StartQry is not None:
                    qry = ('SELECT * FROM ' + tblrpt + ' WHERE '
                           + IDcol + 'ID = ' + str(StartQry))
                    self.data = Dbase().Dsqldata(qry)
                    newtbl.append(tblrpt)

                    newlst = []
                    for i in self.data:
                        newlst = list(i)
                        frgID = i[1]
                        # look up the text string corresponding to
                        # the ID values in the foreign tables
                        if type(frgID) == str:
                            qry = ('SELECT ' + cmb1val + ' FROM ' + frgtbl1 +
                                   ' WHERE ' + cmb1fld + ' = "' + frgID + '"')
                        else:
                            qry = ('SELECT ' + cmb1val + ' FROM ' + frgtbl1 +
                                   ' WHERE ' + cmb1fld + ' = ' + str(frgID))

                        newtxt1 = Dbase().Dsqldata(qry)[0][0]
                        newlst[1] = newtxt1

                        if tblrpt == 'PaintSpec':
                            if type(frgID) == str:
                                qry = ('SELECT ' + cmb2val + ' FROM '
                                       + frgtbl2 + ' WHERE ' + cmb2fld +
                                       ' = "' + frgID + '"')
                            else:
                                qry = ('SELECT ' + cmb2val + ' FROM '
                                       + frgtbl2 + ' WHERE ' + cmb2fld +
                                       ' = ' + str(frgID))

                            newtxt2 = Dbase().Dsqldata(qry)[0][0]

                            newlst[4] = newtxt2
                        # add the string value from the foreign
                        # tables to the rptdta
                        RptData.append(tuple(newlst))

                    Colnames.append(RptColNames)
                    Colwdths.append(RptColWdth)
                else:
                    exclud.append(tblrpt)
                    # exclud = list(set(self.fabtbl).difference(set(newtbl)))
            else:
                exclud.append(tblrpt)

        Rpt_Fab_All.Report(newtbl, RptData, Colnames,
                           Colwdths, self.filename, exclud,
                           ttl).create_pdf()

        if RptData != []:
            return False


class ValvesRpt:
    def __init__(self, vlvtbl, ComdPrtyID, ComCode, MtrSpc,
                 filename, PipeCode=None):

        self.vlvtbl = vlvtbl
        self.ComdPrtyID = ComdPrtyID
        self.PipeCode = PipeCode
        self.ComCode = ComCode
        self.MtrSpc = MtrSpc
        self.filename = filename

    def VlvBld(self):
        import Rpt_Vlvs_All

        colms = []
        colnames = []
        titl = []
        rptdata = []
        VlvIDs = []
        new_list = []
        exclud = []
#        vlvlst = ['BallValve', 'GateValve', 'GlobeValve', 'PlugValve',
#                  'ButteflyValve', 'PistonCheckValve', 'SwingCheckValve']

        for vlv in self.vlvtbl:
            colwdth = [6, 25, 10, 10, 10, 10, 10, 15, 9, 9, 25,
                       9, 9, 30]
            DsqlVlv = ('''SELECT v.ID, v.Valve_Code, a.ANSI_class, b.Pipe_OD,
                    c.Pipe_OD, d.Connection, e.Material_Type, f.Body_Material,
                    g.Wedge_Material, ''')

            if vlv == 'BallValve':
                DsqlVlv = (DsqlVlv + '''h.Stem_Material, i.Packing_Material,
                            j.Body_Type, k.Seat_Material, l.Porting, v.Notes
                        \n FROM ''')
                self.tblID = 'Ball_Valve_ID'
                self.PrtyTbl = 'BallProperty'
                self.PrtyID = 'BAID'
                colwdth.insert(13, 10)
            elif vlv == 'GateValve':
                DsqlVlv = (DsqlVlv + '''h.Stem_Material, i.Packing_Material,
                            j.Bonnet_Type, k.Seat_Material, l.Wedge_Type,
                            m.Porting, v.Notes \n FROM ''')
                self.tblID = 'Gate_Valve_ID'
                self.PrtyTbl = 'GateProperty'
                self.PrtyID = 'GTID'
                colwdth.insert(13, 9)
                colwdth.insert(13, 9)
            elif vlv == 'GlobeValve':
                DsqlVlv = (DsqlVlv + '''h.Stem_Material, i.Packing_Material,
                            j.Bonnet_Type, k.Seat_Material,
                            v.Notes \n FROM ''')
                self.tblID = 'Globe_Valve_ID'
                self.PrtyTbl = 'GlobeProperty'
                self.PrtyID = 'GBID'
            elif vlv == 'PlugValve':
                DsqlVlv = (DsqlVlv + '''h.Stem_Material, i.Packing_Material,
                            j.Body_Type, k.Sleeve_Material, l.Porting, v.Notes
                            \n FROM ''')
                self.tblID = 'Plug_Valve_ID'
                self.PrtyTbl = 'PlugProperty'
                self.PrtyID = 'PGID'
                colwdth.insert(13, 10)
            elif vlv == 'SwingCheckValve':
                DsqlVlv = (DsqlVlv + '''h.Wedge_Material, i.Packing_Material,
                            j.Bonnet_Type, k.Seat_Material,
                            v.Notes \n FROM ''')
                self.tblID = 'Swing_ID'
                self.PrtyTbl = 'SwingProperty'
                self.PrtyID = 'SWID'
            elif vlv == 'PistonCheckValve':
                DsqlVlv = (DsqlVlv + '''h.Spring_Material, i.Packing_Material,
                            j.Bonnet_Type, k.Seat_Material,
                            v.Notes \n FROM ''')
                self.tblID = 'Piston_Check_ID'
                self.PrtyTbl = 'PistonProperty'
                self.PrtyID = 'PSID'
            elif vlv == 'ButterflyValve':
                DsqlVlv = ('''SELECT v.ID, v.Valve_Code, a.ANSI_class, b.Pipe_OD,
                            c.Pipe_OD, d.Body_Type, e.Material_Type,
                            f.Body_Material, g.Disc_Material, h.Shaft_Material,
                            i.Packing_Material, j.Seat_Material,
                            v.Notes \n FROM ''')
                self.tblID = 'Butterfly_Valve_ID'
                self.PrtyTbl = 'ButterflyProperty'
                self.PrtyID = 'BUID'
                del colwdth[-3]

            DsqlVlv = DsqlVlv + vlv + ' v' + '\n'

            # the index fields are returned from the PRAGMA
            # statement tbldata in reverse order so the alpha
            # characters need to count up not down
            n = 0
            tbldata = Dbase().Dtbldata(vlv)
            for element in tbldata:
                alpha = chr(96-n+len(tbldata))
                DsqlVlv = (DsqlVlv + ' INNER JOIN ' + element[2] + ' '
                           + alpha + ' ON v.' + element[3] + ' = ' +
                           alpha + '.' + element[4] + '\n')
                n += 1

            Valves = ('SELECT ' + self.tblID + ' FROM ' + self.PrtyTbl +
                      ' WHERE Commodity_Property_ID = ' + str(self.ComdPrtyID))

            VlvIDs = Dbase().Dsqldata(Valves)

            if VlvIDs != []:

                ValveIDs = tuple([i[0] for i in VlvIDs])
                if len(ValveIDs) == 1:
                    self.vlvstrg = '(' + str(ValveIDs[0]) + ')'
                else:
                    self.vlvstrg = str(ValveIDs)
                self.DsqlVlv = DsqlVlv + ' WHERE v.ID IN ' + self.vlvstrg
                data = Dbase().Dsqldata(self.DsqlVlv)

                columnames = []
                Col_Names = Dbase(vlv).ColNames()
                for item in Col_Names:
                    columnames.append(item.replace(' ', '\n'))
                columnames[2] = 'Rating\nDesignation'
                columnames[3] = 'Minimum\nPipe Size'
                columnames[4] = 'Maximum\nPipe Size'
                columnames[6] = 'Matr Type\nDesignation'

                if self.PipeCode == 'None':
                    ttl = self.ComCode + '-' + self.MtrSpc
                else:
                    ttl = self.PipeCode
                frmname = (' '.join(re.findall('([A-Z][a-z]*)', vlv)))
                ttl = frmname + ' for ' + ttl

                new_list.append(vlv)
                colnames.append(columnames)
                colms.append(colwdth)
                titl.append(ttl)
                rptdata.append(data)
                exclud = list(set(self.vlvtbl).difference(set(new_list)))

            else:
                if len(self.vlvtbl) == 1:
                    exclud = self.vlvtbl
                else:
                    if vlv not in exclud:
                        exclud.append(vlv)

        Rpt_Vlvs_All.Report(new_list, rptdata, colms, colnames,
                            self.filename, exclud, titl).create_pdf()

        if rptdata != []:
            return False


class FittingsRpt:
    def __init__(self, ftgtbl, ComdPrtyID, ComCode, MtrSpc,
                 filename, PipeCode=None):

        self.ftgtbl = ftgtbl
        self.ComdPrtyID = ComdPrtyID
        self.filename = filename
        self.data = []

        # specify a title for the report if table name
        # is not to be used
        if PipeCode == 'None':
            ttl = ComCode + '-' + MtrSpc
        else:
            ttl = PipeCode
        self.ttl = ' for ' + ttl

    def FitngBld(self):
        import Rpt_Lvl4_All

        Colnames = []
        Colwdths = []
        rptdata = []
        newtbl = []
        exclud = []
        for tblrpt in self.ftgtbl:
            query = ('SELECT ' + tblrpt[:-1] +
                     '''_ID FROM PipeSpecification WHERE
                      Commodity_Property_ID = '''
                     + str(self.ComdPrtyID))
            chk = Dbase().Dsqldata(query)
            # Confirm if fitting item is setup in the PipeSpec tbl
            if chk != []:
                StartQry = chk[0][0]
                # if item exists in PipeSpec tbl use returned ID to
                # find properties in the corresponding fitting table
                if StartQry is not None:
                    qry = (
                        'SELECT * FROM ' + tblrpt + ' WHERE '
                        + tblrpt[:-1] + 'ID = ' + str(StartQry))
                    self.data = Dbase().Dsqldata(qry)
                    # call the function to get the data for the report,
                    # data,column names,column widths
                    rptdt, Colns, wdths = self.ReportData(tblrpt)
                    newtbl.append(tblrpt)
                    rptdata.append(rptdt)
                    Colnames.append(Colns)
                    Colwdths.append(wdths)
                    # exclud = list(set(self.ftgtbl).difference(set(newtbl)))
                else:
                    exclud.append(tblrpt)
            else:
                exclud.append(tblrpt)

        Rpt_Lvl4_All.Report(newtbl, rptdata, Colnames,
                            Colwdths, self.filename, exclud,
                            self.ttl).create_pdf()

        if rptdata != []:
            return False

    def ReportData(self, tblrpt):
        rptdata = []
        # common table column lables
        lbl_txt = ['Min. Pipe\nDiameter', 'Max. Pipe\nDiameter',
                   'Flange\nStyle', 'Flange\nFace', 'Material',
                   'Schedule\nClass']

        if tblrpt == 'Fittings':
            lbl_txt = ['Min. Pipe\nDiameter', 'Max. Pipe\nDiameter',
                       'End\nConnections', 'Material',
                       'Schedule\nClass']
            Colnames = [('ID', 'Fitting Code', 'Material\nSpecification'),
                        tuple(lbl_txt)]
            Colwdths = [(6, 40, 10), (10, 10, 20, 20, 10)]
        elif tblrpt == 'Flanges':
            Colnames = [('ID', 'Flange Code', 'Material\nSpecification'),
                        tuple(lbl_txt)]
            Colwdths = [(6, 40, 10), (10, 10, 20, 20, 10, 10)]
        elif tblrpt == 'OrificeFlanges':
            Colnames = [('ID', 'Orifice\nFlange Code',
                        'Material\nSpecification'),
                        tuple(lbl_txt)]
            Colwdths = [(6, 40, 10), (10, 10, 20, 20, 10, 10)]
        elif tblrpt == 'Unions':
            Colnames = [('ID', 'Union Code', 'Material\nSpecification'),
                        tuple(lbl_txt)]
            Colwdths = [(6, 40, 10), (10, 10, 20, 20, 10, 10)]
        elif tblrpt == 'OLets':
            Colnames = [('ID', 'Olet Code', 'Material\nSpecification'),
                        tuple(lbl_txt)]
            Colwdths = [(6, 40, 10), (10, 10, 15, 10, 20, 10)]
        elif tblrpt == 'GrooveClamps':
            lbl_txt = ['Min. Pipe\nDiameter', 'Max. Pipe\nDiameter',
                       'Schedule', 'Groove Type', 'Style',
                       'Seal\nMaterial', 'Clamp\nMaterial']
            Colnames = [('ID', 'Groove\nClamp Code',
                         'Material\nSpecification'),
                        tuple(lbl_txt)]
            Colwdths = [(6, 40, 10), (10, 10, 10, 10, 20, 20, 20)]

        numcols = len(Colwdths[1])

        for segn in range(len(self.data)):
            data1 = list(self.data[segn][0:3])
            data2 = list(self.data[segn][3:])
            numrows = len(data2)//numcols

            # determine the matr spec designation used for all tables
            query = ('''SELECT Pipe_Material_Spec FROM PipeMaterialSpec
                    WHERE Material_Spec_ID = ''' + str(data1[2]))
            data1[2] = Dbase().Dsqldata(query)[0][0]

            n = 0
            for i in data2:
                if i is None:
                    break
                # first column numbers
                if n in [m*numcols for m in range(0, numrows)]:
                    # [0,1,5,6,10,11,15,16]:
                    query3 = ("SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = "
                              + str(i))
                # second column numbers
                elif n in [m*numcols+1 for m in range(0, numrows)]:
                    query3 = ("SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = "
                              + str(i))
                # third column numbers
                elif n in [m*numcols+2 for m in range(0, numrows)]:
                    # [2,7,12,17]:
                    if tblrpt in ['Fittings', 'Unions']:
                        query3 = ('''SELECT Connection FROM EndConnects
                                   WHERE EndID = ''' + str(i))
                        if i in [2, 4, 5]:
                            self.rtg_tbl = 'ForgeClass'
                            self.mtr_tbl = 'ForgedMaterial'
                        elif i == 1:
                            self.rtg_tbl = 'ANSI_Rating'
                            self.mtr_tbl = 'ForgedMaterial'
                        elif i == 3:
                            self.rtg_tbl = 'PipeSchedule'
                            self.mtr_tbl = 'ButtWeldMaterial'
                        elif i == 6:
                            self.rtg_tbl = 'PipeSchedule'
                            self.mtr_tbl = 'PipeMaterial'
                    elif tblrpt in ['Flanges', 'OrificeFlanges']:
                        query3 = ('''SELECT Style_Type FROM FlangeStyle WHERE
                                StyleID = ''' + str(i))
                    elif tblrpt == 'OLets':
                        query3 = (
                            'SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = '
                            + str(i))
                    elif tblrpt == 'GrooveClamps':
                        query3 = ('''SELECT Pipe_Schedule FROM PipeSchedule
                                WHERE ID = ''' + str(i))
                # fourth column numbers
                elif n in [m*numcols+3 for m in range(0, numrows)]:
                    # [3,8,13,18]:
                    if tblrpt == 'Fittings':
                        tbl_name = self.mtr_tbl
                        if tbl_name == 'ForgedMaterial':
                            query3 = (
                                'SELECT Forged_Material FROM ' +
                                tbl_name + " WHERE ID = " + str(i))
                        elif tbl_name == 'ButtWeldMaterial':
                            query3 = ('SELECT Butt_Weld_Material FROM ' +
                                      tbl_name + " WHERE ID = " + str(i))
                        elif tbl_name == 'PipeMaterial':
                            query3 = ('SELECT Pipe_Material FROM ' + tbl_name +
                                      " WHERE ID = " + str(i))
                    elif tblrpt in ['Flanges', 'OrificeFlanges']:
                        query3 = ('''SELECT Face_Style FROM FlangeFace
                                   WHERE FaceID = ''' + str(i))
                    elif tblrpt == 'Unions':
                        query3 = ('''SELECT Seat_Material FROM SeatMaterial
                                   WHERE SeatMaterialID = ''' + str(i))
                    elif tblrpt == 'OLets':
                        query3 = (
                            "SELECT Style FROM OLetStyle WHERE StyleID = "
                            + str(i))
                    elif tblrpt == 'GrooveClamps':
                        query3 = ('''SELECT Groove FROM ClampGroove
                                   WHERE ClampGrooveID = ''' + str(i))
                # fifth column numbers
                elif n in [m*numcols+4 for m in range(0, numrows)]:
                    # [4,9,14,19]:
                    if tblrpt == 'Fittings':
                        tbl_name = self.rtg_tbl
                        if tbl_name == 'ANSI_Rating':
                            query3 = ('SELECT ANSI_Class FROM ' + tbl_name +
                                      ' WHERE Rating_Designation = ' + str(i))
                        elif tbl_name == 'ForgeClass':
                            query3 = ('SELECT Forged_Class FROM ' + tbl_name +
                                      ' WHERE ClassID = ' + str(i))
                        elif tbl_name == 'PipeSchedule':
                            query3 = ('SELECT Pipe_Schedule FROM ' + tbl_name +
                                      " WHERE ID = " + str(i))
                    elif tblrpt in ['Flanges', 'OrificeFlanges',
                                    'Unions', 'OLets']:
                        query3 = ('''SELECT Forged_Material FROM
                                   ForgedMaterial WHERE ID = ''' + str(i))
                    elif tblrpt == 'GrooveClamps':
                        query3 = ('''SELECT Style FROM GrooveClampVendor
                                   WHERE VendorID = ''' + str(i))
                # sixth column numbers
                elif n in [m*numcols+5 for m in range(0, numrows)]:
                    if tblrpt in ['Flanges', 'OrificeFlanges']:
                        query3 = ('''SELECT ANSI_Class FROM ANSI_Rating
                                   WHERE Rating_Designation = ''' + str(i))
                    elif tblrpt == 'Unions':
                        query3 = ('''SELECT Forged_Class FROM ForgeClass
                                   WHERE ClassID = ''' + str(i))
                    elif tblrpt == 'OLets':
                        query3 = ("SELECT Weight FROM OLetWt WHERE OLetWtID = "
                                  + str(i))
                    elif tblrpt == 'GrooveClamps':
                        query3 = ('''SELECT GasketSealMaterial FROM GasketSealMaterial
                                   WHERE SealID = ''' + str(i))
                # seventh column numbers
                elif n in [m*numcols+6 for m in range(0, numrows)]:
                    if tblrpt == 'GrooveClamps':
                        query3 = ('''SELECT Forged_Material FROM
                                   ForgedMaterial WHERE ID = ''' + str(i))

                data2[n] = str(Dbase().Dsqldata(query3)[0][0])

                n += 1

            rptdata.append(tuple(data1))
            rptdata.append(tuple(data2[:n]))

        return (rptdata, Colnames, Colwdths)


class FastenersRpt:
    def __init__(self, fsttbl, ComdPrtyID, ComCode, MtrSpc,
                 filename, PipeCode=None):

        self.ComdPrtyID = ComdPrtyID
        self.filename = filename
        self.fsttbl = fsttbl[0]

        if PipeCode is not None:
            ttl = ComCode + '-' + MtrSpc
        else:
            ttl = PipeCode
        self.ttl = ' for ' + ttl

    def FstBld(self):
        import Report_Lvl1

        rptdata = []
        Colwdths = [20, 20]
        Colnames = ['Bolt\nMaterial', 'Nut\nMaterial']

        query = ('''SELECT Fastener_ID FROM PipeSpecification
                  WHERE Commodity_Property_ID = ''' + str(self.ComdPrtyID))
        chk = Dbase().Dsqldata(query)
        if chk != []:
            StartQry = chk[0][0]
            if StartQry is not None:
                qry = ('SELECT * FROM ' + self.fsttbl
                       + ' WHERE FastenerID = ' + str(StartQry))
                self.data = Dbase().Dsqldata(qry)

                data2 = []

                Colwdths = [20, 20]
                Colnames = ['Bolt\nMaterial', 'Nut\nMaterial']

                for segn in range(len(self.data)):
                    data1 = [x for x in list(self.data[segn][1:])
                             if x is not None]

                    for n in range(4):
                        if n <= len(data1)-1:
                            if n in (0, 2):
                                qry = ('''SELECT Bolt_Material FROM BoltMaterial
                                        WHERE BoltID = ''' + str(data1[n]))
                                data2.append(Dbase().Dsqldata(qry)[0][0])
                            else:
                                qry = ('''SELECT Nut_Material FROM NutMaterial
                                        WHERE NutID = ''' + str(data1[n]))
                                data2.append(Dbase().Dsqldata(qry)[0][0])
                                rptdata.append(tuple(data2))
                                data2 = []

        Report_Lvl1.Report(self.fsttbl, rptdata, Colnames,
                           self.filename, self.ttl, Colwdths).create_pdf()

        if rptdata != []:
            return False


class WeldingRpt:
    def __init__(self, wldtbl, ComdPrtyID, ComCode, MtrSpc,
                 filename, PipeCode=None):

        self.ComdPrtyID = ComdPrtyID
        self.filename = filename
        self.wldtbl = wldtbl[0]
        self.subdata = []

        if PipeCode == 'None':
            ttl = ComCode + '-' + MtrSpc
        else:
            ttl = PipeCode
        self.ttl = 'Weld Requirements for ' + ttl

    def WldBld(self):
        import Report_Weld

        rptdata = []

        Colnames = [('Process', 'Notes'),
                    ('Weld\nProcedure', 'Process', 'Weld\nFiller',
                     'Filler\nGroup', 'Approved\nThickness', 'Material',
                     'Position', 'Position\nDescription',
                     'Welder Qualification\nCertificate Notes', 'Notes')]

        Colwdths = [(10, 20),
                    (8, 8, 10, 8, 10, 8, 8, 20, 15, 15)]

        qry = ('SELECT ' + self.wldtbl +
               '_ID FROM PipeSpecification WHERE Commodity_Property_ID = '
               + str(self.ComdPrtyID))
        chk = Dbase().Dsqldata(qry)
        if chk != []:
            StartQry = chk[0][0]
            if StartQry is not None:
                qry = ('SELECT * FROM ' + self.wldtbl + ' WHERE '
                       + self.wldtbl + 'ID = ' + str(StartQry))
                self.data = Dbase().Dsqldata(qry)
                # range of procedureID fields is 2 to 5
                for n in range(2, 6):
                    SubSQL = ('''SELECT * FROM WeldProcedures
                               WHERE ProcedureID = "'''
                              + str(self.data[0][n]) + '"')
                    if Dbase().Dsqldata(SubSQL) != []:
                        self.subdata.append(Dbase().Dsqldata(SubSQL)[0])

                for seg in self.data:
                    rptdata3 = []
                    rptdata1 = []
                    qry = (
                        'SELECT Process FROM WeldProcesses WHERE ProcessID = '
                        + str(seg[1]))
                    rptdata1.append(Dbase().Dsqldata(qry)[0][0])
                    rptdata1.append(seg[6])

                    rptdata3.append(tuple(rptdata1))

                    for m in seg[2:6]:
                        if m is not None:
                            qry = ('''SELECT * FROM WeldProcedures
                                    WHERE ProcedureID = ''' + str(m))

                            suddt = Dbase().Dsqldata(qry)
                            rptdata2 = []
                            rptdata2.append(suddt[0][1])
                            qry = ('''SELECT Process FROM WeldProcesses
                                    WHERE ProcessID = ''' + str(suddt[0][5]))
                            rptdata2.append(Dbase().Dsqldata(qry)[0][0])
                            qry = ('''SELECT Metal_Spec FROM
                                    WeldFiller WHERE ID = ''' +
                                   str(suddt[0][2]))
                            rptdata2.append(Dbase().Dsqldata(qry)[0][0])
                            qry = ('''SELECT Filler_Group FROM WeldFillerGroup
                                    WHERE ID = ''' + str(suddt[0][7]))
                            rptdata2.append(Dbase().Dsqldata(qry)[0][0])
                            qry = ('''SELECT Thickness FROM WeldQualifyThickness
                                    WHERE ID = ''' + str(suddt[0][3]))
                            rptdata2.append(Dbase().Dsqldata(qry)[0][0])
                            qry = ('''SELECT Material_Group FROM WeldMaterialGroup
                                    WHERE ID = ''' + str(suddt[0][6]))
                            rptdata2.append(Dbase().Dsqldata(qry)[0][0])
                            qry = ('''SELECT Position, Description FROM WeldQualifyPosition
                                    WHERE ID = ''' + str(suddt[0][8]))
                            tmp = Dbase().Dsqldata(qry)
                            rptdata2.append(tmp[0][0])
                            rptdata2.append(tmp[0][1])

                            rptdata2.append(suddt[0][4])
                            rptdata2.append(suddt[0][9])
                            rptdata3.append(tuple(rptdata2))

                    rptdata.append(rptdata3)

        Report_Weld.Report(rptdata, Colnames,
                           Colwdths, self.filename, self.ttl).create_pdf()

        if rptdata != []:
            return False


class InsulRpt:
    def __init__(self, insultbl, ComdPrtyID, ComCode, MtrSpc,
                 filename, PipeCode=None):

        self.ComdPrtyID = ComdPrtyID
        self.filename = filename
        self.insultbl = insultbl[0]

        if PipeCode == 'None':
            ttl = ComCode + '-' + MtrSpc
        else:
            ttl = PipeCode
        self.ttl = ttl

    def InsulBld(self):
        import Report_Insul

        rptdata = []

        Colnames = [('ID', 'Insulation Code', 'Jacketing', 'Surface Prep',
                     'Insulation Class', 'Adhesive', 'Sealer', 'Note'),
                    ('Min. Dia', 'Max. Dia', 'Material', 'Thickness')]

        Colwdths = [(5, 35, 10, 15, 15, 15, 15, 30), (10, 10, 15, 10)]

        query = ('SELECT ' + self.insultbl +
                 '_ID FROM PipeSpecification WHERE Commodity_Property_ID = '
                 + str(self.ComdPrtyID))
        chk = Dbase().Dsqldata(query)
        if chk != []:
            StartQry = chk[0][0]
            if StartQry is not None:
                self.DsqlFtg = ('SELECT * FROM ' + self.insultbl + ' WHERE '
                                + self.insultbl + 'ID = ' + str(StartQry))
                self.data = Dbase().Dsqldata(self.DsqlFtg)

                for segn in range(0, len(self.data)):
                    # specify the general data for the insulation
                    data1 = []
                    data1.append(str(self.data[segn][0]))  # Insulation ID
                    data1.append(str(self.data[segn][31]))  # Insulation code

                    n = 0
                    for i in self.data[segn][25:30]:
                        if i is None:
                            data1.append('N/A')
                            continue
                        else:
                            if n == 0:
                                qry = ('''SELECT JacketMatr, Jacket_Thk FROM InsulationJacket
                                        WHERE JacketID = ''' + str(i))
                                data1.append(Dbase().Dsqldata(qry)[0][0])
                            if n == 1:
                                qry = ('''SELECT Notes FROM PaintPrep
                                        WHERE PaintPrepID = ''' + str(i))
                                data1.append(Dbase().Dsqldata(qry)[0][0])
                            if n == 2:
                                qry = ('''SELECT Insulation_Class FROM InsulationClass
                                        WHERE ClassID = ''' + str(i))
                                data1.append(Dbase().Dsqldata(qry)[0][0])
                            if n == 3:
                                qry = ('''SELECT Adhesive, Vendor FROM InsulationAdhesive
                                        WHERE AdhesiveID = ''' + str(i))
                                data1.append(Dbase().Dsqldata(qry)[0][0])
                            if n == 4:
                                qry = ('''SELECT Sealer, Vendor FROM InsulationSealer
                                        WHERE SealerID = ''' + str(i))
                                data1.append(Dbase().Dsqldata(qry)[0][0])
                        n += 1
                    data1.append(str(self.data[segn][30]))  # Insulation Note

                    # generate the data for the various diameter ranges
                    qry = []
                    qry.append(
                        'SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = ')
                    qry.append(
                        'SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID = ')
                    qry.append('''SELECT Material FROM InsulationMaterial
                                WHERE MatrID = ''')
                    qry.append('''SELECT Thickness FROM InsulationThickness
                                WHERE ThkID = ''')

                    data2 = []
                    data2.extend(range(len(self.data[segn][31].split('.'))-5))

                    for s in range(1, 5):
                        for n in [m*4+s for m in range(0, 6)]:
                            ID = self.data[segn][n]
                            if ID is None:
                                break
                            data2[n] = (str(Dbase().Dsqldata(qry[s-1] +
                                        str(ID))[0][0]))
                    del data2[0]

                    rptdata.append(tuple(data1))
                    rptdata.append(tuple(data2))

        Report_Insul.Report(self.insultbl, rptdata, Colnames, Colwdths,
                            self.filename, self.ttl).create_pdf()

        if rptdata != []:
            return False


class BranchRpt:
    def __init__(self, rpttal, ComdPrtyID, ComCode, MtrSpc,
                 Ends, filename, PipeCode=None):

        self.ComdPrtyID = ComdPrtyID
        self.filename = filename
        self.Ends = Ends

        if PipeCode == 'None':
            ttl = ComCode + '-' + MtrSpc
        else:
            ttl = PipeCode
        self.ttl = ttl

    def BrchBld(self):
        import Report_Branch

        colsz = 0
        rows = []

        query = ('''SELECT * FROM PipeSpecification WHERE
                 Commodity_Property_ID = ''' + str(self.ComdPrtyID))
        chk = Dbase().Dsqldata(query)
        if chk != []:    # commodity code exists in pipespec
            CurrentID = chk[0][18]  # get the current branch criteria ID
            if CurrentID is not None:  # ID of branch criteria exists
                # for commodity code
                qry = ('''SELECT * FROM BranchCriteria WHERE
                                BranchID = ''' + str(CurrentID))
                self.data = Dbase().Dsqldata(qry)

                Dia = []
                PltAxis = []

                rMin = self.data[0][1]
                qry = ("SELECT PipeOD_ID FROM Pipe_OD WHERE Pipe_OD = '" +
                       rMin + "'")
                rMinID = Dbase().Dsqldata(qry)[0][0]

                rMax = self.data[0][2]
                qry = ("SELECT PipeOD_ID FROM Pipe_OD WHERE Pipe_OD = '" +
                       rMax + "'")
                rMaxID = Dbase().Dsqldata(qry)[0][0]

                qry = ("SELECT Pipe_OD FROM Pipe_OD WHERE PipeOD_ID >= "
                       + str(rMinID) + ' AND PipeOD_ID <= ' + str(rMaxID))
                Axis = Dbase().Dsqldata(qry)

                for item in Axis:
                    PltAxis.append(item[0])
                for D in PltAxis:
                    Dia.append(eval(D[:-1].replace('-', '+')))

                nRTr = eval(self.data[0][3][:-1].replace('-', '+'))
                nFgb = eval(self.data[0][5][:-1].replace('-', '+'))
                nBWr1 = eval(self.data[0][9][:-1].replace('-', '+'))
                nBWr2 = eval(self.data[0][8][:-1].replace('-', '+'))
                nRTDif = eval(self.data[0][4])
                nFgDif = eval(self.data[0][6])
                nBWDif = eval(self.data[0][10])

                row = []
                rows = []
                Rindex = 0
                for r in Dia:
                    Bindex = 0
                    row = []
                    for b in Dia:
                        if b <= r:
                            # equal tee
                            if b == r:
                                row.append('ET')
                            # Reducing Tee determined by run size
                            elif (r <= nRTr) or (Rindex-Bindex) <= nRTDif:
                                row.append('RT')
                            # O-Let fittings
                            elif (b <= nFgb and (Rindex-Bindex) >= nFgDif):
                                row.append('OL')
                            # Set-On w/wo Repad
                            elif (self.data[0][7] is True and b >= nFgb and
                                  (Rindex - Bindex) >= nBWDif and
                                  nBWr1 >= r >= nBWr2):
                                row.append('BO')
                            # Sweep Outlet
                            elif (self.data[0][11] is True and
                                  r > nBWr1 and b > nFgb):
                                row.append('SW')
                            else:
                                row.append('Eng')
                            Bindex += 1
                        else:
                            continue
                    row.insert(0, PltAxis[Rindex])
                    Rindex += 1
                    rows.append(row)

                PltAxis.insert(0, 'Run^')

                rows.append(PltAxis)

                colsz = 560//len(PltAxis)

        Report_Branch.Report(rows, colsz, self.Ends, self.filename, self.ttl)

        if rows != []:
            return False


class SpcNotesRpt:
    def __init__(self, rpttbl, ComdPrtyID, ComCode, MtrSpc,
                 filename, PipeCode=None):

        self.rpttbl = rpttbl
        self.ComdPrtyID = ComdPrtyID
        self.PipeCode = PipeCode
        self.ComCode = ComCode
        self.MtrSpc = MtrSpc
        self.filename = filename

    def SpcNtsBld(self):
        import Rpt_SpcNotes_All

        ttl = []
        Colnames = []
        Colwdths = []
        rptdata = []

        for tbl in self.rpttbl:

            if tbl == 'Notes':
                ttlname = 'Commodity Notes'
                colwdths = [5, 20, 80]
                link_fld = 'Category_ID'
                tbl_ID = 'NoteID'
                self.frg_tbl = 'NoteCategory'
                self.frg_fld = 'CategoryID'
                frgn_col = 'Category'
                columnames = ['ID', 'Category', 'Note']
                spec_col = 19
            elif tbl == 'Specials':
                ttlname = 'Specialty Items'
                colwdths = [5, 10, 40, 20, 10, 15]
                link_fld = 'Item_Type_ID'
                tbl_ID = 'SpecialsID'
                self.frg_tbl = 'SpecialItems'
                self.frg_fld = 'SpecialTypeID'
                frgn_col = 'Item_Type'
                columnames = ['ID', 'Item Type', 'Description', 'Notes',
                              'Vendor', 'Part Number']
                spec_col = 17

            # specify a title for the report if table name is not to be used
            if self.PipeCode == 'None':
                rpt_ttl = self.ComCode + '-' + self.MtrSpc
            else:
                rpt_ttl = self.PipeCode
            ttl.append(ttlname + ' for ' + rpt_ttl)

            datatable_str = ''

            # if there is a specified ComdPrtyID then show dvc1
            if self.ComdPrtyID is not None:
                query = ('''SELECT * FROM PipeSpecification WHERE
                        Commodity_Property_ID = ''' + str(self.ComdPrtyID))
                chk = Dbase().Dsqldata(query)
                # the commodity property exists in pipe secification
                # and has Notes assigned
                if chk != []:
                    # this is the list of the IDs for the Notes
                    self.NoteStr = chk[0][spec_col]
                    # if there are notes assinged the data to dvc1
                    if self.NoteStr is not None:
                        ColInfo = Dbase().Dcolinfo(tbl)
                        for item in ColInfo:
                            # get first Lvl2 datatable column name in item[1]
                            # check to see if name is lvl2 primary key
                            # or lvl1 linked field if they are not then
                            #  add tablename and datafield to SELECT statement
                            if item[1] == tbl_ID:
                                datatable_str = (datatable_str + ' ' + tbl +
                                                 '.' + item[1] + ',')
                            elif item[1] != link_fld and item[1] != tbl_ID:
                                datatable_str = (datatable_str + ' ' + tbl +
                                                 '.' + item[1] + ',')
                            elif item[1] == link_fld:
                                datatable_str = (datatable_str + ' ' +
                                                 self.frg_tbl + '.' +
                                                 frgn_col + ',')

                        datatable_str = datatable_str[:-1]

                        DsqlLvl2 = ('SELECT ' + datatable_str + ' FROM ' + tbl
                                    + ' INNER JOIN ' + self.frg_tbl)
                        DsqlLvl2 = (DsqlLvl2 + ' ON ' + tbl + '.' + link_fld
                                    + ' = ' + self.frg_tbl + '.'
                                    + self.frg_fld)

                        self.DsqlLvl2 = DsqlLvl2
                        self.Dsql1 = (self.DsqlLvl2 + ' WHERE ' + tbl_ID +
                                      ' IN ' + self.NoteStr)
                        self.data1 = Dbase().Dsqldata(self.Dsql1)
                        rptdata.append((self.data1))

            Colnames.append(columnames)
            Colwdths.append(colwdths)

        Rpt_SpcNotes_All.Report(self.rpttbl, rptdata, Colnames,
                                Colwdths, self.filename, ttl).create_pdf()

        if rptdata != []:
            return False


class DataMods(dv.DataViewIndexListModel):
    '''CLASS TO HANDLE THE CHANGES TO THE TABLE DATA'''
    def __init__(self, frmtbl, data):
        dv.DataViewIndexListModel.__init__(self, len(data))
        self.data = data
        self.frmtbl = frmtbl
        self.edit_data = []

# This method is called when the user edits a data item in the view.
    def SetValueByRow(self, value, row, col):
        self.edit_data = [row, col, value]
        return True

# This method is called to provide the data object for a particular row,col
    def GetValueByRow(self, row, col):
        if (len(str(self.data[row][col]))) > 70:
            new_strg = self.wrap_paragraphs(str(self.data[row][col]), 70)
            new_strg = "".join(new_strg)
            return new_strg
        else:
            return str(self.data[row][col])

# Report how many columns this model provides data for.
    def GetColumnCount(self):
        if self.data == []:
            return 0
        return len(self.data[0])

# Specify the data type for a column
    def GetColumnType(self, col):
        return "string"

# Report the number of rows in the model
    def GetCount(self):
        return len(self.data)

# Called to check if non-standard attributes should be used
# in the cell at (row, col)
    def GetAttrByRow(self, row, col, attr):
        if col == 0:
            attr.SetColour('blue')
            attr.SetBold(True)
            return True
        return False

    def dedent(self, text):

        if text.startswith('\n'):
            # text starts with blank line, don't ignore the first line
            return textwrap.dedent(text)

        # split first line
        splits = text.split('\n', 1)
        if len(splits) == 1:
            # only one line
            return textwrap.dedent(text)

        first, rest = splits
        # dedent everything but the first line
        rest = textwrap.dedent(rest)
        return '\n'.join([first, rest])

    def wrap_paragraphs(self, text, ncols):

        paragraph_re = re.compile(r'\n(\s*\n)+', re.MULTILINE)
        text = self.dedent(text).strip()
        # every other entry is space
        paragraphs = paragraph_re.split(text)[::2]
        out_ps = []
        indent_re = re.compile(r'\n\s+', re.MULTILINE)

        for p in paragraphs:
            # presume indentation that survives dedent
            # is meaningful formatting,
            # so don't fill unless text is flush.
            if indent_re.search(p) is None:
                # wrap paragraph
                p = textwrap.fill(p, ncols)
            out_ps.append(p)
        return out_ps

    # copy then delete the row of data
    def DeleteRow(self, row, delPK_name):
        try:
            DtQueryVal = self.data[row][0]
            Dbase().TblDelete(self.frmtbl, DtQueryVal, delPK_name)
            del self.data[row]
            # notify the view(s) using this model that it has been removed
            self.RowDeleted(row)
        except sqlite3.IntegrityError:
            wx.MessageBox('This ' + delPK_name +
                          ''' is associated with other tables and cannot
                           be deleted!''', "Cannot Delete",
                          wx.OK | wx.ICON_INFORMATION)

    def EditCell(self, pk_Name, pk_Col):

        # use the self.edit_data from SetValueByRow to
        # specify location of edited
        # cell and new value

        # specify the colume of table which contains the ID
        colID = pk_Col
        # value changed in column extracted from SetValueByRow
        value = self.edit_data[2]
        colChgNum = self.edit_data[1]

        # get the table ID value in edited row
        rowID = self.GetValueByRow(self.edit_data[0], colID)
        return (rowID, colID, value, colChgNum)

    # insert a blank row
    def AddRow(self):
        # notify views
        self.RowAppended()
        return


class ListCtrlComboPopup(wx.ComboPopup):

    '''CLASS TO HANDLE THE CHANGES TO THE COMBO POPUP LIST'''
    def __init__(self, tbl_name, PupQuery='',
                 cmbvalue='', showcol=0, lctrls=None):

        wx.ComboPopup.__init__(self)
        self.tbl_name = tbl_name
        self.PupQuery = PupQuery
        self.cmbvalue = cmbvalue
        self.showcol = showcol
        self.lctrls = lctrls
        self.name_list = []

        if self.PupQuery == '':
            self.PupQuery = 'SELECT * FROM ' + self.tbl_name
            for item in Dbase().Dcolinfo(self.tbl_name):
                self.name_list.append(item[1])
        else:    # get last of required column names for the combo box list
            spot = self.PupQuery.index('FROM', 0, len(self.PupQuery))
            tblnms = self.PupQuery[6:spot]
            tblnms = tblnms.replace(self.tbl_name+'.', '')
            tblnms = tblnms.strip(' ')
            if tblnms == '*':
                for item in Dbase().Dcolinfo(self.tbl_name):
                    self.name_list.append(item[1])
            else:
                self.name_list = tblnms.split(',')

    def AddItem(self, txt):
        self.lc.InsertItem(self.lc.GetItemCount(), txt)

    def OnLeftDown(self, evt):
        item, flags = self.lc.HitTest(evt.GetPosition())
        if item == -1:
            return
        if item >= 0:
            self.lc.Select(item, on=0)
            self.curitem = item
        self.value = self.curitem
        self.Dismiss()

    # This is called immediately after construction finishes.  You can
    # use self.GetCombo if needed to get to the ComboCtrl instance.
    def Init(self):
        self.value = -1
        self.curitem = -1

    # Create the popup child control.  Return true for success.
    def Create(self, parent):
        self.index = 0

        self.lc = wx.ListCtrl(parent, size=wx.DefaultSize,
                              style=wx.LC_REPORT | wx.BORDER_SUNKEN
                              | wx.LB_SINGLE)

        # this looks up the data for
        # the individual tables representing each combo box
        self.lc.Bind(wx.EVT_LEFT_DOWN, self.OnLeftDown)
        if self.lctrls is not None:
            self.lctrls.append(self.lc)

        LstColWdth = []
        for item in Dbase().Dcolinfo(self.tbl_name):
            if item[1] in self.name_list:
                lstcol_wdth = ''
                # check to see if field length is specified if so
                # use it to set grid col width
                for s in re.findall(r'\d+', item[2]):
                    if s.isdigit():
                        lstcol_wdth = int(s)
                        LstColWdth.append(lstcol_wdth)

                if 'INTEGER' in item[2] or 'FLOAT' in item[2]:
                    if lstcol_wdth == '' and 'FLOAT' in item[2]:
                        LstColWdth.append(10)
                    elif lstcol_wdth == '' and 'INTEGER' in item[2]:
                        LstColWdth.append(6)
                elif 'BLOB' in item[2]:
                    if lstcol_wdth == '':
                        LstColWdth.append(30)
                elif 'TEXT' in item[2]:
                    if lstcol_wdth == '':
                        LstColWdth.append(10)
                elif 'DATE' in item[2]:
                    if lstcol_wdth == '':
                        LstColWdth.append(10)

        n = 0
        for info in Dbase().Dcolinfo(self.tbl_name):
            info = list(info)
            colname = info[1]
            if colname in self.name_list:
                if colname.find("ID", -2) != -1:
                    colname = "ID"
                elif colname.find("_") != -1:
                    colname = colname.replace("_", " ")
                else:
                    colname = (' '.join(re.findall('([A-Z][a-z]*)', colname)))
                self.lc.InsertColumn(n, colname)

                lstcolwdth = LstColWdth[n]*9
                if (len(colname)*9) > lstcolwdth:
                    lstcolwdth = len(colname)*9

                self.lc.SetColumnWidth(n, lstcolwdth)
                n += 1

        index = 0
        for values in Dbase().Dsqldata(self.PupQuery):
            col = 0
            for value in values:
                if col == 0:
                    self.lc.InsertItem(index, str(value))
                else:
                    self.lc.SetItem(index, col, str(value))
                col += 1
            index += 1
        return True

    # Return the widget that is to be used for the popup
    def GetControl(self):
        return self.lc

    # Return a string representation of the current item.
    def GetStringValue(self):
        if self.value == -1:
            return
        return self.lc.GetItemText(self.value, self.showcol)

    # Called immediately after the popup is shown
    def OnPopup(self):
        # this provides the original combox value
        # if self.value >= 0:
        #    print('OnPopUp',self.lc.GetItemText(self.value))
        wx.ComboPopup.OnPopup(self)

    # Called when popup is dismissed
    def OnDismiss(self):
        # this provides the new combo value
        # print('OnDismiss',self.lc.GetItemText(self.value))
        wx.ComboPopup.OnDismiss(self)

    def PaintComboControl(self, dc, rect):
        # This is called to custom tube in the combo control itself
        # (ie. not the popup).  Default implementation draws value as
        # string.
        wx.ComboPopup.PaintComboControl(self, dc, rect)

    # Return final size of popup. Called on every popup, just prior to OnPopup.
    # minWidth = preferred minimum width for window
    # prefHeight = preferred height. Only applies if > 0,
    # maxHeight = max height for window, as limited by screen size
    #   and should only be rounded down, if necessary.
    def GetAdjustedSize(self, minWidth, prefHeight, maxHeight):
        minWidthNew = 0
        for cl in range(0, self.lc.GetColumnCount()):
            minWidthNew = self.lc.GetColumnWidth(cl) + minWidthNew
        minWidth = minWidthNew
        return wx.ComboPopup.GetAdjustedSize(self, minWidth, 150, 800)

    # Return true if you want delay the call to Create until the popup
    # is shown for the first time. It is more efficient, but note that
    # it is often more convenient to have the contrminWidth
    # immediately.
    # Default returns false.
    def LazyCreate(self):
        return wx.ComboPopup.LazyCreate(self)


class Dbase(object):
    '''DATABASE CLASS HANDLER'''
    # this initializes the database and opens the specified table
    def __init__(self, frmtbl=None):
        # this sets the path to the database and needs
        # to be changed accordingly
        self.frmtbl = frmtbl

    def Dcolinfo(self, table):
        # sequence for items in colinfo is column number, column name,
        # data type(size), not null, default value, primary key
        cursr.execute("PRAGMA table_info(" + table + ");")
        colinfo = cursr.fetchall()
        return colinfo

    def Dtbldata(self, table):
        # this will provide the foreign keys and their related tables
        # unknown,unknown,Frgn Tbl,Parent Tbl Link fld,
        # Frgn Tbl Link fld,action,action,default
        cursr.execute('PRAGMA foreign_key_list(' + table + ')')
        tbldata = list(cursr.fetchall())
        return tbldata

    def Dsqldata(self, DataQuery):
        # provides the actual data from the table
        cursr.execute(DataQuery)
        sqldata = cursr.fetchall()
        return sqldata

    def TblDelete(self, table, val, field):
        '''Call the function to delete the values in
        the grid into the database error trapping will occure
        in the call def delete_data'''

        if type(val) != str:
            DeQuery = ("DELETE FROM " + table + " WHERE "
                       + field + " = " + str(val))
        else:
            DeQuery = ("DELETE FROM " + table + " WHERE "
                       + field + " = '" + val + "'")
        cursr.execute(DeQuery)
        self.db.commit()

    def TblEdit(self, UpQuery, data_strg=None):
        if data_strg is None:
            cursr.execute(UpQuery)
        else:
            cursr.execute(UpQuery, data_strg)
        self.db.commit()

        # determine the required column width, name and primary status
    def Fld_Size_Type(self):
        # specified field type or size
        values = []
        auto_incr = True
#        New_ID = ''
        ColWdth = []
        n = 0

        # collect all the table information needed to build the grid
        # colinfo includes schema for each column: column number, name,
        # field type & length, is null , default value, is primary key
        for item in self.Dcolinfo(self.frmtbl):
            col_wdth = ''
        # check to see if field length is specified if so
        # use it to set grid col width
            for s in re.findall(r'\d+', item[2]):
                if s.isdigit():
                    col_wdth = int(s)
                    ColWdth.append(col_wdth)
        # check if it is a text string and primary key if it is then
        # it is not auto incremented develope a string of data based
        # on record field type for adding new row
            if item[5] == 1:
                pk = item[1]
                pk_col = n
        # if the primary key is not an interger then assign a text
        # value and indicate it is not auto incremented
                if 'INTEGER' not in item[2]:
                    values.append('Required')
                    auto_incr = False
        # it must be an integer and will be auto incremeted,
        # New_ID is assigned in AddRow routine
                else:
                    values.append('New_ID')
                    if col_wdth == '':
                        ColWdth.append(6)
        # remaining steps assing values to not
        # null fields otherwise leave empty
            elif 'INTEGER' in item[2] or 'FLOAT' in item[2]:
                if item[3]:
                    values.append(0)
                else:
                    values.append('')
                if col_wdth == '' and 'FLOAT' in item[2]:
                    ColWdth.append(10)
                elif col_wdth == '':
                    ColWdth.append(6)
            elif 'BLOB' in item[2]:
                if item[3]:
                    values.append('Required')
                else:
                    values.append('')
                if col_wdth == '':
                    ColWdth.append(30)
            elif 'TEXT' in item[2] or 'BOOLEAN' in item[2]:
                if item[3]:
                    values.append('Required')
                else:
                    values.append('')
                if col_wdth == '':
                    ColWdth.append(10)
            elif 'DATE' in item[2]:
                i = datetime.datetime.now()
                today = ("%s-%s-%s" % (i.month, i.day, i.year))
                if item[3]:
                    values.append(today)
                if col_wdth == '':
                    ColWdth.append(10)
            n += 1

        # the variables in FldInfo are;database column name for ID,
        # database number of ID column, if ID is autoincremented
        # (imples interger or stg), list of database specified column
        # width, a list of database column names,
        # a list of values to insert into none null fields
        FldInfo = [pk, pk_col, auto_incr, ColWdth, values]
        return FldInfo

    def ColNames(self):
        colnames = []
        for item in self.Dcolinfo(self.frmtbl):
            # modify the column names to remove
            # underscore and seperate words
            colname = item[1]
            if colname.find("ID", -2) != -1:
                colname = "ID"
            elif colname.find("_") != -1:
                colname = colname.replace("_", " ")
            else:
                colname = (' '.join(re.findall('([A-Z][a-z]*)', colname)))
            colnames.append(colname)
        return colnames

    def Search(self, ShQuery):
        cursr.execute(ShQuery)
        data = cursr.fetchall()
        if data == []:
            return False
        else:
            return data

    def Restore(self, RsQuery):
        cursr.execute(RsQuery)
        data = cursr.fetchall()
        cursr.close()
        return data


if __name__ == '__main__':

    app = wx.App()
    frm = StrUpFrm()
    frm.Show()
    frm.CenterOnScreen()
    app.MainLoop()
