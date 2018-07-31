import csv
import os
import shutil
import pandas as pd
import numpy as np
import openpyxl
import wx
import wx.lib.agw.multidirdialog as MDD

class conversionFrame(wx.Frame):
    """ Description of class. """
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, title=title, size=(200, 100))
        #self.control = wx.TextCtrl(self, style=wx.TE_MULTILINE)
        self.openbutton = wx.Button(self, -1, "Select Directory/ies")
        self.Bind(wx.EVT_BUTTON, self.OnOpen, self.openbutton)

        self.dirlist = []  # List of directories with output files of interest

        self.Show(True)

    def OnOpen(self, e):
        """ Select directory/ies to open. """
        self.dirname = ''
        dlg = MDD.MultiDirDialog(self, "Choose a directory/directories to open",
                                 defaultPath='H:\Transfer-SAR-MD\Chang\emquest',
                                 agwStyle=MDD.DD_MULTIPLE|MDD.DD_DIR_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:
            self.dirlist = dlg.GetPaths()
            print(self.dirlist)
        dlg.Destroy()

    def OnExit(self, e):
        self.CLose(True)


app = wx.App(False)
frame = conversionFrame(None, 'Small Editor')
app.MainLoop()

exit()


def coord2ind(coord):
    numstartindex = 0
    for ind, char in enumerate(coord):
        if char.isnumeric():
            numstartindex = ind
            break
    print(numstartindex)
    chars = coord[:numstartindex]
    nums = coord[numstartindex:]
    return [ord(chars.lower())-97, int(nums)-1]


# Reference Points in EMQuest Conversion Sheet
# 1 Channel
ch1_qpsk = 'B29'
ch1_16qam = 'E29'
ch1_64qam = 'H29'
# 3 Channel
ch3_qpsk = 'B3'
ch3_16qam = 'E3'
ch3_64qam = 'H3'
# 5 Channel
ch5_qpsk = 'B53'
ch5_16qam = 'E53'
ch5_64qam = 'H53'

# Get list of files for a specific power LTE band
filerootdir = 'H:\Transfer-SAR-MD\Chang\emquest'
lteband = 'LTE B4 max'
subdir = filerootdir + '\\' + lteband
powerfiles = [f for f in os.listdir(subdir) if f.endswith('.csv')]

# Open EMQuest Conversion Sheet
convsheetfname = 'EMQuest conversion Rev A'
srcfile = filerootdir + '\\' + convsheetfname + '.xlsx'
dstfile = filerootdir + '\\' + convsheetfname + ' ' + lteband + '.xlsx'
shutil.copyfile(srcfile, dstfile)

xlswriter = pd.ExcelWriter(dstfile, engine='openpyxl')
convsheets = openpyxl.load_workbook(dstfile)
xlswriter.book = convsheets
xlswriter.sheets = dict((ws.title, ws) for ws in convsheets.worksheets)

# Traverse output files, extract power values
for file in powerfiles:
    with open(subdir + "\\" + file, newline='') as csvfile:
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
        print("Row Count: " + str(len(powvals)))
        df = pd.DataFrame({'col': powvals})

        # Finding the correct index in the excel sheet
        sheet_name = ''
        startcol = 1  # By default, QPSK 1 RB
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
        if filenum % 3 == 0 :  # 100% RB
            if len(powvals) == 5:
                startrow = 52
            elif len(powvals) == 1:
                startrow = 28
        else:  # 1% and 50% RB
            if len(powvals) == 15:
                startrow = 52
            elif len(powvals) == 3 or len(powvals) == 6:
                startrow = 28
        # Start column
        startcol = filenum

        # print("Col: " + chr(startcol+97).upper() + " - Row: " + str(startrow))
        # print("")
        df.to_excel(xlswriter, sheet_name=sheet_name, header=None, index=False, startcol=startcol, startrow=startrow)
        xlswriter.save()

print('Output transcription complete. Power values can be found in ' + dstfile)
