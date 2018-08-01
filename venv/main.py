import csv
import os
import threading
import shutil
import pandas as pd
import numpy as np
import openpyxl
import wx
import wx.lib.agw.multidirdialog as mdd
import wx.grid

# TODO: We could potentially add checkpoints directly onto the list/wx.Grid itself??

class CheckpointDialog(wx.Dialog):
    """
    (OUTDATED)
    GUI wx.Dialog for checking if user wants to check the output file between LTE band conversions.
    """
    def __init__(self, parent, title):
        """
        Constructor for the CheckpointDialog class.
        :param parent: Parent object of the wx.Dialog.
        :param title: The title of the wx.Dialog (to be shown above the menu bar).
        """
        super(CheckpointDialog, self).__init__(parent, title=title, size=(250, 150))
        panel = wx.Panel(self)
        self.btn = wx.Button(panel, wx.ID_OK, label="Open Conversion Sheet")


class ConversionFrame(wx.Frame):
    """
    GUI wx.Frame hosting all main GUI elements and conversion scripts.
    """
    def __init__(self, parent, title):
        """
        Initializes a new frame for the EMQuest Conversion Program's GUI.
        :param parent: Parent object of the frame.
        :param title: Title of the program, to be displayed above the menu bar.
        """
        wx.Frame.__init__(self, parent, title=title, size=(1000, 840))

        self.dirlist = []  # List of directories with output files of interest
        self.savedir = ''
        self.templatefilename = ''
        self.topselect = None
        self.bottomselect = None
        self.lteband = ''
        self.success = []
        self.failure = []

        # GUI Item IDs
        save_id = 110
        temp_id = 111
        run_id = 112
        add_id = 113
        remove_id = 114
        clear_id = 115
        close_id = 116
        help_id = 117

        # Accelerator Table/Shortcut Keys
        self.accel_tbl = wx.AcceleratorTable([(wx.ACCEL_CTRL, ord('s'), save_id), (wx.ACCEL_CTRL, ord('t'), temp_id),
                                              (wx.ACCEL_CTRL, ord('r'), run_id), (wx.ACCEL_CTRL, ord('a'), add_id),
                                              (wx.ACCEL_CTRL, ord('b'), clear_id), (wx.ACCEL_CTRL, ord('h'), help_id)])
        self.SetAcceleratorTable(self.accel_tbl)

        self.savetxt = wx.TextCtrl(self)
        self.savebutton = wx.Button(self, save_id, "Select Save Directory...")
        self.Bind(wx.EVT_BUTTON, self.save, self.savebutton)
        self.templatetxt = wx.TextCtrl(self)
        self.templatetxt.WriteText('H:\Transfer-SAR-MD\Chang\emquest\EMQuest conversion Rev A.xlsx')
        self.templatebutton = wx.Button(self, temp_id, "Find Template File...")
        self.Bind(wx.EVT_BUTTON, self.SelectTemplateFile, self.templatebutton)
        self.perbandcheckbox = wx.CheckBox(self, label="Enable Checkpoints")
        self.perbandcheckbox.SetValue(False)

        self.dirgrid = wx.grid.Grid(self, -1, size=(10, 250))
        self.dirgrid.CreateGrid(25, 1)
        self.dirgrid.HideColLabels()
        self.dirgrid.HideRowLabels()
        self.dirgrid.SetRowSize(0, 5)
        self.dirgrid.SetColSize(0, 250)
        attr = wx.grid.GridCellAttr()
        attr.SetReadOnly(True)
        self.dirgrid.SetColAttr(0, attr)
        # self.dirgrid.Bind(wx.grid.EVT_GRID_RANGE_SELECT, self.onDragSelection)
        # self.dirgrid.Bind(wx.grid.EVT_GRID_SELECT_CELL, self.onSingleSelection)

        self.openbutton = wx.Button(self, add_id, "Add")
        self.Bind(wx.EVT_BUTTON, self.add, self.openbutton)
        # self.removebutton = wx.Button(self, -1, "Remove")   # TODO: Temporarily removing the remove button
        # self.Bind(wx.EVT_BUTTON, self.remove, self.removebutton)  # TODO: Temporarily removing the remove button
        self.clearbutton = wx.Button(self, clear_id, "Clear")
        self.Bind(wx.EVT_BUTTON, self.clear, self.clearbutton)

        self.runbutton = wx.Button(self, run_id, "Run")
        self.Bind(wx.EVT_BUTTON, self.startconversion, self.runbutton)

        # Sizers/Layout
        self.mainvertsizer = wx.BoxSizer(wx.VERTICAL)
        self.mainhorizsizer = wx.BoxSizer(wx.HORIZONTAL)
        self.settingssizer = wx.BoxSizer(wx.VERTICAL)
        self.dirsizer = wx.BoxSizer(wx.VERTICAL)
        self.dirbtnsizer = wx.BoxSizer(wx.HORIZONTAL)

        # Static Lines and Boxes
        self.filestbox = wx.StaticBoxSizer(wx.VERTICAL, self)
        self.horizstline = wx.StaticLine(self, wx.ID_ANY, style=wx.LI_VERTICAL)

        self.dirbtnsizer.Add(self.openbutton, proportion=5, flag=wx.EXPAND)
        # self.dirbtnsizer.Add(self.removebutton, proportion=5, flag=wx.EXPAND)  # TODO: Temp. removing the remove button
        self.dirbtnsizer.Add(self.clearbutton, proportion=5, flag=wx.EXPAND)
        text = wx.StaticText(self, label="Selected Output Files to Convert")

        text.SetFont(wx.Font(9, wx.DECORATIVE, wx.NORMAL, wx.BOLD))
        text2 = wx.StaticText(self, label="Filename Format: yyyy.mm.dd.hh.mm.ss_B<LTE band> <BW> MHz \n" +
                                          "<data column> <RB offset> RB <modulation type>.csv")
        text3 = wx.StaticText(self, label="Example: 2018.06.08.12.45.49_B41 20 MHz 05 50 RB 16QAM.csv")
        self.dirsizer.Add(text, proportion=0)
        self.dirsizer.Add(text2, proportion=0)
        self.dirsizer.Add(text3, proportion=0)
        self.dirsizer.Add(self.dirgrid, proportion=10, flag=wx.EXPAND | wx.ALL)
        self.dirsizer.Add(self.dirbtnsizer, proportion=1, flag=wx.EXPAND | wx.ALIGN_BOTTOM)
        text = wx.StaticText(self, label="Save Directory")
        text.SetFont(wx.Font(9, wx.DECORATIVE, wx.NORMAL, wx.BOLD))
        self.filestbox.Add(text)
        self.filestbox.Add(self.savetxt, proportion=1, flag=wx.EXPAND)
        self.filestbox.Add(self.savebutton, proportion=1, flag=wx.EXPAND)
        self.filestbox.AddStretchSpacer(prop=1)
        text = wx.StaticText(self, label="Template File")
        text.SetFont(wx.Font(9, wx.DECORATIVE, wx.NORMAL, wx.BOLD))
        self.filestbox.Add(text)
        self.filestbox.Add(self.templatetxt, proportion=1, flag=wx.EXPAND)
        self.filestbox.Add(self.templatebutton, proportion=1, flag=wx.EXPAND)
        self.settingssizer.Add(self.filestbox, proportion=5, flag=wx.EXPAND | wx.ALL)
        self.settingssizer.AddStretchSpacer(prop=1)
        self.settingssizer.Add(self.perbandcheckbox, proportion=1, flag=wx.EXPAND)
        self.settingssizer.Add(self.runbutton, proportion=2, flag=wx.EXPAND | wx.ALIGN_BOTTOM)
        self.mainhorizsizer.AddStretchSpacer(prop=1)
        self.mainhorizsizer.Add(self.settingssizer, proportion=8)
        self.mainhorizsizer.Add(self.horizstline, proportion=0, flag=wx.ALL | wx.EXPAND, border=5)
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
        """
        Opens a wx Dialog to select directory to save converted output files.
        :param e: Event handler.
        :return: Nothing.
        """
        with wx.DirDialog(self, "Select directory to save converted files", style=wx.DD_DIR_MUST_EXIST) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                self.savetxt.Clear()
                self.savetxt.WriteText(dlg.GetPath())

    def SelectTemplateFile(self, e):
        """
        Opens a wx Dialog to select the output conversion template file.
        :param e: Event handler.
        :return: Nothing.
        """
    #    """ Select EMQuest conversion file. """
        with wx.FileDialog(self, "Select EMQuest conversion file", style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                self.templatetxt.Clear()
                self.templatetxt.SetValue(dlg.GetPath())

    def onSingleSelection(self, e):
        """
        Registers clicks on the wx.Grid (cells), determines current coordinates (saves them on self.topselect).
        :param e: Event handler.
        :return: Nothing.
        """
        self.topselect = (e.GetRow(), e.GetCol())
        self.bottomselect = self.topselect
        e.Skip()

    def onDragSelection(self, e):
        """
        Registers drags/multi-cell selections on the wx.Grid (cells), determines the top & bottom cell coordinates.
        :param e: Event handler.
        :return: Nothing.
        """
        self.topselect = self.dirgrid.GetSelectionBlockTopLeft()
        if self.topselect:
            self.topselect = self.topselect[0]
            self.bottomselect = self.dirgrid.GetSelectionBlockBottomRight()[0]
            print("Coords: " + str(self.topselect) + " " + str(self.bottomselect))
            return
        cells = self.dirgrid.GetSelectedCells()
        min = self.dirgrid.GetNumberRows()
        max = -1
        for cell in cells:
            if cell[0] > max:
                max = cell[0]
                self.bottomselect = cell[0]
            if cell[0] < min:
                min = cell[0]
                self.topselect = cell[0]
        print("Coords: " + str(self.topselect) + " " + str(self.bottomselect))

    def add(self, e):
        """
        Adds input file(s) to the list of files to convert.
        :param e: Event handler.
        :return: Nothing.
        """
        with wx.FileDialog(self, "Select EMQuest conversion file(s)", defaultDir='H:\Transfer-SAR-MD\Chang\emquest',
                            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE,
                            wildcard="CSV files (.csv)|*.csv") as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                for path in dlg.GetPaths():
                    if path not in self.dirlist:
                        self.dirlist.append(path)
        self.populatedirlist()

    def add_folder(self, e):
        """
        Adds a directory of input file(s) to the list of file(s) to convert.
        :param e: Event handler.
        :return: Nothing.
        """
        with mdd.MultiDirDialog(self, "Choose a directory/directories to open",
                                 defaultPath='H:\Transfer-SAR-MD\Chang\emquest',
                                 agwStyle=mdd.DD_MULTIPLE | mdd.DD_DIR_MUST_EXIST) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                for filedir in dlg.GetPaths():
                    drive = filedir.split(':')[0][-1] + ':'
                    folderspath = filedir.split(':')[1]
                    path = drive + folderspath[folderspath.find('\\'):]
                    if path not in self.dirlist:
                        self.dirlist.append(path)
        self.populatedirlist()

    def remove(self, e):
        """
        Removes selected files from the list of files to convert.
        :param e: Event handler.
        :return: Nothing.
        """
        if self.topselect is None:
            return
        else:
            for i in range(self.topselect[0], self.bottomselect[0]+1):
                print(i)
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
        """
        Clears all files from the list of files to convert. Also clears out the wx.Grid to reflect this.
        :param e: Event handler.
        :return: Nothing.
        """
        self.dirlist = []
        self.dirgrid.ClearGrid()
        pass

    def startconversion(self, e):
        """
        Starts the thread for the conversion script. Disables the GUI while the thread is running.
        :param e: Event handler.
        :return: Nothing.
        """
        if self.savetxt.GetValue() == '':
            with wx.MessageDialog(self, "No save directory selected.\nPlease select a directory on the top left.",
                                  style=wx.OK | wx.ICON_WARNING | wx.CENTER) as dlg:
                dlg.ShowModal()
            return
        elif not self.dirlist:
            with wx.MessageDialog(self, "No files selected.\nPlease select file(s) to convert.",
                                  style=wx.OK | wx.ICON_WARNING | wx.CENTER) as dlg:
                dlg.ShowModal()
            return
        self.disablegui()
        threading.Thread(target=self.runconversion, args=(e,)).start()

    def runconversion(self, e):
        """
        Converts and compiles the data from the input files into a single excel sheet based on the template file
        and in the save directory specified.
        :param e: Event handler.
        :return: Nothing.
        """
        self.success = []
        self.failure = []
        if len(self.dirlist) <= 0:
            with wx.MessageDialog(self, "No files selected.\nPlease add directories to the list on the right.",
                                  style=wx.OK | wx.ICON_ERROR | wx.CENTER) as dlg:
                dlg.ShowModal()
            return
        for file in self.dirlist:
            lteband = file.split('_')[1].split(' ')[0]
            # Open EMQuest Conversion Sheet
            # convsheetfname = 'EMQuest conversion Rev A'
            # Check if this is a valid directory
            if not os.path.exists(self.savetxt.GetValue()):
                self.errormsg("Error: no such save directory exists: '%s'" % self.savetxt.GetValue())
                self.enablegui()
                return
            if not os.path.exists(self.templatetxt.GetValue()):
                self.errormsg("Error: no such template file exists: '%s" % self.templatetxt.GetValue())
                self.enablegui()
                return

            # Check if export file already exists
            dstfile = self.savetxt.GetValue() + '\\EMQuest conversion Rev A LTE ' + lteband + '.xlsx'
            if not os.path.exists(dstfile):
                shutil.copyfile(self.templatetxt.GetValue(), dstfile)  # Make copy of template file for output conv.
            xlswriter = pd.ExcelWriter(dstfile, engine='openpyxl')
            convsheets = openpyxl.load_workbook(dstfile)
            xlswriter.book = convsheets
            xlswriter.sheets = dict((ws.title, ws) for ws in convsheets.worksheets)
            with open(file, newline='') as csvfile:
                filedescription = file[:-4].split('_')[1].split(' ')
                bandwidth = float(filedescription[1])
                filenum = int(filedescription[3])
                csvreader = csv.reader(csvfile)  # CSV file iterator
                # Copy power outputs into a list
                csvfile.seek(0)  # resetting the iterator to the beginning of the csv file
                filedetails = str(next(csvreader))  # Skipping the file and column headers
                if not ('Test Method:      Communication Tester Frequency Response Measurement' in filedetails):
                    print("Wrong file format")
                    self.failure.append(file)
                    continue
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
                    else:
                        print("Wrong file format - missing or superfluous number of power values.")
                        self.failure.append(file)
                        continue
                else:  # 1% and 50% RB
                    if len(powvals) == 15:
                        startrow = 52
                    elif len(powvals) == 3 or len(powvals) == 6:
                        startrow = 28
                    else:
                        print("Wrong file format - missing or superfluous number of power values.")
                        self.failure.append(file)
                        continue
                df.to_excel(xlswriter, sheet_name=sheet_name, header=None, index=False, startcol=startcol,
                            startrow=startrow)
                xlswriter.save()  # Save the changes onto the excel file
                self.success.append(file)  # File successfully added - add to list to display at the end
                print('Output transcription for ' + file + ' complete. Power values added to ' + dstfile)
        endmsg = "Conversions completed.\n"
        endmsg += "Files converted successfully:\n"
        if self.success:
            for file in self.success:
                endmsg += file + "\n"
        else:
            endmsg += "None.\n"
        endmsg += "\nFiles converted unsuccessfully:\n"
        if self.failure:
            for file in self.failure:
                endmsg += file + "\n"
        else:
            endmsg += "None.\n"
        with wx.MessageDialog(self, endmsg, style=wx.ICON_INFORMATION | wx.OK | wx.CENTER) as dlg:
            dlg.ShowModal()
        self.enablegui()

    def runconversion_folder(self, e):
        """
        (OUTDATED)
        Converts and compiles the data from the input files in the specified directories into a single excel sheet
        based on the template file and in the save directory specified.
        :param e: Event handler.
        :return: Nothing.
        """
        # Convert output files and compile in a copy of the template excel sheet.
        # Get list of files for a specific power LTE band
        if len(self.dirlist) <= 0:
            print("No directories selected - Please add directories to the list on the right.")
            return
        for filedir in self.dirlist:
            powerfiles = [f for f in os.listdir(filedir) if f.endswith('.csv')]
            self.lteband = filedir.split('\\')[-1]

            # Open EMQuest Conversion Sheet
            # convsheetfname = 'EMQuest conversion Rev A'
            srcfile = 'H:\Transfer-SAR-MD\Chang\emquest\EMQuest conversion Rev A.xlsx'
            dstfile = 'H:\Transfer-SAR-MD\Chang\emquest\EMQuest conversion Rev A ' + self.lteband + '.xlsx'
            shutil.copyfile(self.templatetxt.GetValue(), dstfile)  # Make copy of template file for output conversions

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
                    filedetails = next(csvreader)  # Skipping the file and column headers
                    print("Filedetails: " + filedetails)
                    next(csvreader)
                    powvals = list()
                    for row in csvreader:
                        try:
                            powvals.append(float(row[1]))
                        except ValueError:
                            print("Henlo")
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
            print('Output transcription for ' + self.lteband + ' complete. Power values can be found in ' + dstfile)
            if self.perbandcheckbox.GetValue():
                self.showcheckpoint(dstfile)
            self.enablegui()

    def disablegui(self):
        """
        Disables all GUI elements while the conversion thread is running.
        :return: Nothing.
        """
        self.runbutton.Enable(False)
        self.perbandcheckbox.Enable(False)
        self.templatetxt.Enable(False)
        self.savetxt.Enable(False)
        self.savebutton.Enable(False)
        self.templatebutton.Enable(False)
        self.openbutton.Enable(False)
        # self.removebutton.Enable(False)
        self.clearbutton.Enable(False)

    def enablegui(self):
        """
        Enables all GUI elements while the conversion thread has completed.
        :return: Nothing.
        """
        self.perbandcheckbox.Enable(True)
        self.templatetxt.Enable(True)
        self.savetxt.Enable(True)
        self.savebutton.Enable(True)
        self.templatebutton.Enable(True)
        self.openbutton.Enable(True)
        # self.removebutton.Enable(True)
        self.clearbutton.Enable(True)
        self.runbutton.Enable(True)

    def showcheckpoint(self, dstfilepath):
        """
        (OUTDATED)
        Opens a wx.Dialog asking if we want to check the output file once all input files for a given band have been
        completed.
        :param dstfilepath: The output file's destination path.
        :return: Nothing.
        """
        with wx.MessageBox("'" + self.lteband + "'" + " conversion complete.\nCheck power excel sheet?",
                            "Band Checkpoint.", wx.YES_NO | wx.ICON_INFORMATION) as dlg:
            if dlg == wx.YES:
                os.startfile(dstfilepath)

    def populatedirlist(self):
        """
        Prints the list of files to convert onto the wx.Grid on the GUI.
        :return: Nothing.
        """
        self.dirgrid.ClearGrid()
        for ind, filedir in enumerate(self.dirlist):
            try:
                self.dirgrid.SetCellValue(ind, 0, filedir.split('\\')[-1])
            except wx._core.wxAssertionError:
                self.dirgrid.AppendRows(self.dirgrid.GetNumberRows())
                self.dirgrid.SetCellValue(ind, 0, filedir.split('\\')[-1])
                #return

    def OnExit(self, e):
        """
        Function run on GUI exit.
        :param e: Event handler.
        :return: Nothing.
        """
        self.Close(True)

    def errormsg(self, errmsg):
        """
        Shows an error message as a wx.Dialog.
        :param errmsg: String error message to show in the message dialog.
        :return: Nothing
        """
        with wx.MessageDialog(self, errmsg, style=wx.OK | wx.ICON_ERROR | wx.CENTER) as dlg:
            dlg.ShowModal()


conversionProgram = wx.App(False)
frame = ConversionFrame(None, 'EMQuest Output Converter')
conversionProgram.MainLoop()
