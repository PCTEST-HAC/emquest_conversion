import csv
import os
from threading import Thread
import shutil
import pandas as pd
import numpy as np
import openpyxl
import wx
import wx.lib.agw.multidirdialog as mdd
from wx.lib.pubsub import Publisher  # pip install PyPubSub
import wx.grid


class conversionThread(Thread):
    def __init__(self, dirlist, checkpointbool):
        Thread.__init__(self)
        self.dirlist = dirlist
        self.checkpointbool = checkpointbool
        self.start()

    def run(self):
        # Convert output files and compile in a copy of the template excel sheet.
        # Get list of files for a specific power LTE band
        if len(self.dirlist) <= 0:
            print("No directories selected - Please add directories to the list on the right.")
            return
        for filedir in self.dirlist:
            powerfiles = [f for f in os.listdir(filedir) if f.endswith('.csv')]
            lteband = filedir.split('\\')[-1]

            # Open EMQuest Conversion Sheet
            convsheetfname = 'EMQuest conversion Rev A'
            srcfile = 'H:\Transfer-SAR-MD\Chang\emquest\EMQuest conversion Rev A.xlsx'
            dstfile = 'H:\Transfer-SAR-MD\Chang\emquest\EMQuest conversion Rev A ' + lteband + '.xlsx'
            shutil.copyfile(srcfile, dstfile)  # Make a copy of the template file to compile output conversions

            xlswriter = pd.ExcelWriter(dstfile, engine='openpyxl')
            convsheets = openpyxl.load_workbook(dstfile)
            xlswriter.book = convsheets
            xlswriter.sheets = dict((ws.title, ws) for ws in convsheets.worksheets)

            # Traverse output files, extract power values
            for file in powerfiles:
                with open(filedir + "\\" + file, newline='') as csvfile:
                    print("Current file: " + file)
                    filedescription = file[:-4].split(' ')
                    bandwidth = float(filedescription[1])
                    filenum = int(filedescription[3])
                    csvreader = csv.reader(csvfile)  # CSV file iterator

                    # Copy power outputs into a list
                    csvfile.seek(0)  # resetting the iterator to the beginning of the csv file
                    next(csvreader)  # Skipping the file and column headers
                    next(csvreader)
                    powvals = list()
                    for row in csvreader:
                        try:
                            powvals.append(float(row[1]))
                        except ValueError:
                            continue
                    # print("Row Count: " + str(len(powvals)))
                    df = pd.DataFrame({'col': powvals})

                    # Finding the correct index in the excel sheet
                    sheet_name = ''
                    startcol = filenum
                    startrow = 2  # By default, on channel 1
                    # Sheet name
                    sheetlist = list(xlswriter.sheets.keys())
                    for sheetname in sheetlist:
                        if bandwidth == 1.4:
                            str_bandwidth = "%.1f" % bandwidth
                        else:
                            str_bandwidth = str(int(bandwidth))
                        if str_bandwidth in sheetname:
                            sheet_name = sheetname
                    # Start row
                    if filenum % 3 == 0:  # 100% RB
                        if len(powvals) == 5:
                            startrow = 52
                        elif len(powvals) == 1:
                            startrow = 28
                    else:  # 1% and 50% RB
                        if len(powvals) == 15:
                            startrow = 52
                        elif len(powvals) == 3 or len(powvals) == 6:
                            startrow = 28

                    # print("Col: " + chr(startcol+97).upper() + " - Row: " + str(startrow))
                    # print("")
                    df.to_excel(xlswriter, sheet_name=sheet_name, header=None, index=False, startcol=startcol,
                                startrow=startrow)
                    xlswriter.save()
            print('Output transcription for ' + lteband + ' complete. Power values can be found in ' + dstfile)
            if self.checkpointbool:
                self.showcheckpoint(dstfile)
        wx.CallAfter(Publisher().sendMessage, "finished", "Thread Complete")

    def showcheckpoint(self, dstfilepath):
        dlg = wx.MessageBox("'" + self.lteband + "'" + " conversion complete.\nCheck power excel sheet?",
                            "Band Checkpoint.", wx.YES_NO | wx.ICON_INFORMATION)
        if dlg == wx.YES:
            os.startfile(dstfilepath)


class CheckpointDialog(wx.Dialog):
    def __init__(self, parent, title):
        super(CheckpointDialog, self).__init__(parent, title=title, size=(250,150))
        panel = wx.Panel(self)
        self.btn = wx.Button(panel, wx.ID_OK, label="Open Conversion Sheet")


class ConversionFrame(wx.Frame):
    """ Description of class. """
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, title=title, size=(1000, 840))

        self.savetxt = wx.TextCtrl(self)
        self.savebutton = wx.Button(self, -1, "Select Save Directory...")
        self.Bind(wx.EVT_BUTTON, self.save, self.savebutton)
        self.templatetxt = wx.TextCtrl(self)
        self.templatebutton = wx.Button(self, -1, "Find Template File...")
        # self.Bind(wx.EVT_BUTTON, self.SelectTemplateFile, self.templatebutton)
        self.perbandcheckbox = wx.CheckBox(self, label="Enable Checkpoints")
        self.perbandcheckbox.SetValue(False)

        self.dirgrid = wx.grid.Grid(self, -1, size=(10, 20))
        self.dirgrid.CreateGrid(20, 1)
        self.dirgrid.HideColLabels()
        self.dirgrid.HideRowLabels()
        self.dirgrid.SetRowSize(0, 5)
        self.dirgrid.SetColSize(0, 120)
        attr = wx.grid.GridCellAttr()
        attr.SetReadOnly(True)
        self.dirgrid.SetColAttr(0, attr)
        self.dirgrid.Bind(wx.grid.EVT_GRID_RANGE_SELECT, self.onDragSelection)
        self.dirgrid.Bind(wx.grid.EVT_GRID_SELECT_CELL, self.onSingleSelection)

        self.openbutton = wx.Button(self, -1, "Add")
        self.Bind(wx.EVT_BUTTON, self.add, self.openbutton)
        self.removebutton = wx.Button(self, -1, "Remove")
        self.Bind(wx.EVT_BUTTON, self.remove, self.removebutton)
        self.clearbutton = wx.Button(self, -1, "Clear")
        self.Bind(wx.EVT_BUTTON, self.clear, self.clearbutton)

        self.runbutton = wx.Button(self, -1, "Run")
        self.Bind(wx.EVT_BUTTON, self.run, self.runbutton)

        Publisher().subscribe(self.postconversion, "finished")

        self.dirlist = []  # List of directories with output files of interest
        self.savedir = ''
        self.templatefilename = ''
        self.topselect = None
        self.bottomselect = None
        self.lteband = ''

        # Sizers/Layout
        self.mainvertsizer = wx.BoxSizer(wx.VERTICAL)
        self.mainhorizsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.settingssizer = wx.BoxSizer(wx.VERTICAL)
        self.dirsizer = wx.BoxSizer(wx.VERTICAL)
        self.dirbtnsizer = wx.BoxSizer(wx.HORIZONTAL)

        self.dirbtnsizer.Add(self.openbutton, proportion=5, flag=wx.EXPAND)
        self.dirbtnsizer.Add(self.removebutton, proportion=5, flag=wx.EXPAND)
        self.dirbtnsizer.Add(self.clearbutton, proportion=5, flag=wx.EXPAND)
        self.dirsizer.Add(wx.StaticText(self, label="Selected Output Files to Convert"))
        self.dirsizer.Add(self.dirgrid, proportion=10, flag=wx.EXPAND)
        self.dirsizer.Add(self.dirbtnsizer, proportion=1, flag=wx.EXPAND | wx.ALIGN_BOTTOM)

        self.settingssizer.Add(wx.StaticText(self, label="Save Directory"))
        self.settingssizer.Add(self.savetxt, proportion=1, flag=wx.EXPAND)
        self.settingssizer.Add(self.savebutton, proportion=1, flag=wx.EXPAND)
        self.settingssizer.AddStretchSpacer(prop=1)
        self.settingssizer.Add(wx.StaticText(self, label="Template File"))
        self.settingssizer.Add(self.templatetxt, proportion=1, flag=wx.EXPAND)
        self.settingssizer.Add(self.templatebutton, proportion=1, flag=wx.EXPAND)
        self.settingssizer.AddStretchSpacer(prop=1)
        self.settingssizer.Add(self.perbandcheckbox, proportion=1, flag=wx.EXPAND)
        self.settingssizer.AddStretchSpacer(prop=1)
        self.settingssizer.Add(self.runbutton, proportion=2, flag=wx.EXPAND | wx.ALIGN_BOTTOM)

        self.mainhorizsizer.AddStretchSpacer(prop=1)
        self.mainhorizsizer.Add(self.settingssizer, proportion=8)
        self.mainhorizsizer.AddStretchSpacer(prop=1)
        self.mainhorizsizer.Add(self.dirsizer, proportion=8)
        self.mainhorizsizer.AddStretchSpacer(prop=1)
        self.mainvertsizer.AddStretchSpacer(prop=1)
        self.mainvertsizer.Add(self.mainhorizsizer, proportion=10, flag=wx.EXPAND)
        self.mainvertsizer.AddStretchSpacer(prop=1)

        self.SetSizer(self.mainvertsizer)
        self.SetAutoLayout(True)
        self.mainhorizsizer.Fit(self)
        self.Show(True)

    def save(self, e):
        """ Select directory to save converted files. """
        dlg = mdd.MultiDirDialog(self, "Select directory to save converted files",
                                 defaultPath='H:\Transfer-SAR-MD\Chang\emquest',
                                 agwStyle=mdd.DD_DIR_MUST_EXIST | mdd.DD_NEW_DIR_BUTTON)
        if dlg.ShowModal() == wx.ID_OK:
            self.savedir = dlg.GetPaths()
            self.savetxt.write(self.savedir[0])
        dlg.Destroy()

    #def SelectTemplateFile(self, e):
    #    """ Select EMQuest conversion file. """
    #    dlg = wx.FileDialog(self, "Select EMQuest conversion file",
    #                        #defaultPath='H:\Transfer-SAR-MD\Chang\emquest',
    #                        style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
    #    if dlg.ShowModal() == wx.ID_OK:
    #        self.templatefilename = dlg.GetDirectory() + dlg.GetFilename()
    #    dlg.Destroy()
    #    self.templatetxt.write(self.templatefilename)

    def onSingleSelection(self, e):
        self.topselect = (e.GetRow(), e.GetCol())
        self.bottomselect = self.topselect
        print(self.topselect)
        print(self.bottomselect)
        e.Skip()

    def onDragSelection(self, e):
        if self.dirgrid.GetSelectionBlockTopLeft():
            self.topselect = self.dirgrid.GetSelectionBlockTopLeft()[0]
            self.bottomselect = self.dirgrid.GetSelectionBlockBottomRight()[0]
            print(self.topselect)
            print(self.bottomselect)

    def add(self, e):
        """ Select directory/ies to open. """
        dlg = mdd.MultiDirDialog(self, "Choose a directory/directories to open",
                                 defaultPath='H:\Transfer-SAR-MD\Chang\emquest',
                                 agwStyle=mdd.DD_MULTIPLE | mdd.DD_DIR_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:
            for filedir in dlg.GetPaths():
                drive = filedir.split(':')[0][-1] + ':'
                folderspath = filedir.split(':')[1]
                path = drive + folderspath[folderspath.find('\\'):]
                if path not in self.dirlist:
                    self.dirlist.append(path)
        dlg.Destroy()
        self.populatedirlist()

    def remove(self, e):
        if self.topselect is None:
            return
        else:
            for i in range(self.topselect[0], self.bottomselect[0]+1):
                band = self.dirgrid.GetCellValue(i, 0)
                for filedir in self.dirlist:
                    if band in filedir:
                        self.dirlist.remove(filedir)
            self.dirgrid.ClearSelection()
            self.topselect = None
            self.bottomselect = None
            self.populatedirlist()
            return

    def clear(self, e):
        self.dirlist = []
        self.dirgrid.ClearGrid()
        pass

    def run(self, e):
        # Convert output files and compile in a copy of the template excel sheet.
        # Get list of files for a specific power LTE band
        conversionThread(self.dirlist, self.perbandcheckbox.GetValue())
        return

    def postconversion(self, e):
        print("boi")

    def populatedirlist(self):
        self.dirgrid.ClearGrid()
        for ind, filedir in enumerate(self.dirlist):
            self.dirgrid.SetCellValue(ind, 0, filedir.split('\\')[-1])

    def OnExit(self, e):
        self.Close(True)


app = wx.App(False)
frame = ConversionFrame(None, 'EMQuest Output Converter')
app.MainLoop()
