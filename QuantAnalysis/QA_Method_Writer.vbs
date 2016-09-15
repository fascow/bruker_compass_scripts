Option Explicit

' QuantAnalysis Method Writer
'
' Version 3
'
' Take a list of (Compound Names,) PIs and RTs from an Excel file to
' create a Bruker Compass QuantAnalysis Method file by GUI scripting
' Tested under Windows 7 with Excel 2010 and QuantAnalysis 2.1 + 2.2
' this script has to be in the same directory as the Excel file
' there should only be one Excel file in this folder
' the sheetname has to be provided before using this script
' the data has to be sorted in a special way, see example file
' the script can take care of the sorting procedure
' the data shall be stored as values, no Excel formulas
' the method file will be saved in the same folder
' the MS data files have to be in the same folder as well
' it is advised to use copies and not you original files

' IMPORTANT
' This script assumes that QuantAnalysis opens in Nested or Work Table view
' This script assumes that you use the standard table layout of QuantAnalysis
' Your Windows environment shall use dots as decimal separator (not commas)

' TODO
' set retention time window to "1" if no value or no column (-1) specified
' make more values accessible to the user of the script (e.g. EIC mass width, smoothing width)
' if there are no data files, don't ask user to switch to nested or work table view

Dim sheetname, cd, filepath, methodfile ' strings
Dim methodpath, compname
Dim filelist, pilist                    ' lists
Dim objFSO, objFolder, objFile          ' file system objects
Dim xl, Shell                           ' app opjects
Dim wb, ws                              ' Excel objects
Dim result                              ' MsgBox result
Dim row, initrow, count, pipos, i       ' integers
Dim LastRow, LastCol
Dim namecol, picol, rtcol, rtwcol       ' integers: column numbers
Dim pivalue, prevpi, rtvalue, rtwin     ' double
Dim delayS, delayM, delayL              ' integers: delays
Dim sorted, keepchanges, debug          ' booleans

'##############################'
'######### SET VALUES #########'

' where to find you table? There shall be only one Excel file in the folder, where you run this script
sheetname   = "Sheet3"
sorted      = False ' is your sheet already sorted for QuantAnalysis? set to False if you want this script to sort your data for QuantAnalyis
keepchanges = False ' do you want to save a sorted copy of your sheet?

' how to name the method file?
methodfile = "QAmethod"

' in which columns to find the values?
' the header (first row) will be ignored, this script will only look for the column numbers specified here
namecol = 1 ' set this to -1 if you have no compound names or 
            ' simply want the legacy version: compound name = pivalue_1 , pivalue_2, ...
picol   = 2
rtcol   = 4
rtwcol  = 5

' how slow is you computer? specify delays in milliseconds
delayS = 50
delayM = 100
delayL = 500

'##############################'
'##############################'

' Script starts here

debug = False

' find xlsx file in current directory
Set objFSO = CreateObject("Scripting.FileSystemObject")
cd = objFSO.GetParentFolderName(WScript.ScriptFullName)
Set objFolder = objFSO.GetFolder(cd)
Set filelist = CreateObject("System.Collections.ArrayList")
For Each objFile In objFolder.Files
    'only proceed if there is an extension on the file.
    If (InStr(objFile.Name, ".") > 0) Then
        'If the file's extension is "xlsx", echo the path to it.
        If (LCase(Mid(objFile.Name, InStrRev(objFile.Name, "."))) = ".xlsx") Then
            'Wscript.Echo objFile.Path
            filelist.Add objFile.Path
        End If
    End If
Next

'Wscript.Echo filelist.Count
'Wscript.Echo filelist.Item(0)
'wscript.echo join(filelist.ToArray(), ", ")

If filelist.Count = 0 Then
    MsgBox "No Excel file was found!", _
           vbOKOnly+vbCritical+vbSystemModal, _
           "QuantAnalysis method writer"
    Wscript.Quit
Else ' TODO add warning if multiple Excel files were found in the folder ##########################
    filepath = filelist.Item(0)
End If

methodpath = cd + "\" + methodfile + ".m"
'MsgBox methodpath
'Wscript.Quit

If (objFSO.FolderExists(methodpath)) Then ' the Bruker methods are actually folders!
    MsgBox "The specified method file already exists!", _
           vbOKOnly+vbCritical+vbSystemModal, _
           "QuantAnalysis method writer"
    Wscript.Quit
End If

Set xl = CreateObject("Excel.Application")
Set wb = xl.Workbooks.Open(filepath) ' set path to your file
xl.Visible = debug
xl.DisplayAlerts = debug 

If Not SheetExists(sheetname) Then 
    MsgBox "The specified Excel worksheet was not found!", _
           vbOKOnly+vbCritical+vbSystemModal, _
           "QuantAnalysis method writer"
    wb.Close False
    xl.Quit
    Wscript.Quit
End If 

If wb.Worksheets.Count > 1 Then
    If SheetExists(sheetname + " copy") or SheetExists(sheetname + " sorted") Then
        MsgBox "Sheet with a similar name (""" & sheetname & " copy/sorted"") " _
               & "already exists. Aborting...", _
               vbOKOnly+vbCritical+vbSystemModal, _ 
               "QuantAnalysis method writer"
        wb.Close False
        xl.Quit
        Wscript.Quit
    End If
    'TODO if only one sheet, no sheetname needs to be provided? have to skip the very first SheetExists(), but may lead to processing the wrong sheet/file
End If

Set ws = wb.Sheets(sheetname) ' set name of your worksheet

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sorting of Excel sheet for QuantAnalysis
' first a copy is created, which is used for temporary sorting (... copy)
' a second copy with the finally sorted values is created (... sorted)
' the first copy will always be deleted
' the second's fate depends on user's choice (keepchanges)

If Not sorted Then

Function SheetExists(sheetname2)
    Dim cws1
    SheetExists = False
    For Each cws1 In wb.Worksheets
        If cws1.Name = sheetname2 Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function

Sub DeleteSheet(sheetname3)
    Dim cws2
    For Each cws2 In wb.Worksheets
        If cws2.Name = sheetname3 Then
            cws2.Delete
            Exit Sub
        End If
    Next
End Sub

Sub NumToText()
    ' convert the selected numbers to text by changing cell format
    ' http://www.mrexcel.com/forum/excel-questions/30232-convert-numbers-text-visual-basic-applications.html
    Dim cell
    For Each cell In xl.Selection
        If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then
            Dim Temp
            Temp = cell.Value 'cell.Text
            cell.ClearContents
            cell.NumberFormat = "@"
            cell.Value = CStr(Temp)
        End If
    Next
End Sub

Dim DataBlock
Dim SheetOne, SheetTwo
Dim Basename '= Sheetname
Dim Frows 'row where the first block of data ends
Dim Dest, DestTwo 'where to put the first and second block of data

' Excel VBA Constants must be manually defined in VBS
' http://www.smarterdatacollection.com/Blog/?p=374
Const xlCellTypeLastCell = 11
Const xlCellTypeVisible = 12
Const xlYes = 1
Const xlNo = 2
Const xlAscending = 1
Const xlSortNormal = 0

Set SheetOne = wb.Sheets(sheetname)
Basename = SheetOne.Name
'MsgBox Basename & ": " & wb.Sheets.Count

' copy the original sheet, for safety, copy will be deleted at the end
' http://stackoverflow.com/a/22771845/3852788
SheetOne.Copy , wb.Sheets(wb.Sheets.Count)
wb.ActiveSheet.Name = Basename & " copy"
Set SheetOne = wb.ActiveSheet

Set SheetTwo = wb.Sheets.Add(, wb.Sheets(wb.Sheets.Count))
SheetTwo.Name = Basename + " sorted"
Set Dest = SheetTwo.Cells(1, 1) 'this is where we'll put the filtered data

SheetOne.Activate

' identify the "data block" range, which we'll apply .autofilter to
With SheetOne
    'http://www.excel-inside.de/vba-loesungen/zellen-a-bereiche/337-letzte-zeile-letzte-spalte-und-letzte-zelle-per-vba-ermitteln
    LastCol = .UsedRange.SpecialCells(xlCellTypeLastCell).Column
    LastRow = .UsedRange.SpecialCells(xlCellTypeLastCell).Row
    Set DataBlock = .Range(.Cells(1, 1), .Cells(LastRow, LastCol))
End With

'MsgBox "LastRow: " & CStr(LastRow) & "       LastCol: " & CStr(LastCol)

' change PI from Number to Text format, otherwise the sorting does not work properly
With SheetOne
    .Range(.Cells(2, picol), .Cells(LastRow, picol)).Select
    Call NumToText
    .Range("A1").Select
End With

' sort based on column PI value
' Sort(Key1, Order1, Key2, Type, Order2, Key3, Order3, Header, OrderCustom, MatchCase, Orientation, SortMethod, DataOption1, DataOption2, DataOption3)
DataBlock.Sort DataBlock.Cells(1, picol), xlAscending, , , , , , xlYes, , , , , xlSortNormal

' first copy all rows containing multiple PI values
With DataBlock
    'AutoFilter Field, Criteria1, Operator, Criteria2, VisibleDropDown
    .AutoFilter picol, "=*;*"
    'copy the still-visible cells to sheet 2
    .SpecialCells(xlCellTypeVisible).Copy Dest
End With

' turn off the autofilter
With SheetOne
    .AutoFilterMode = False
    If .FilterMode = True Then .ShowAllData
End With

' second copy all rows containing single PI values
Frows = SheetTwo.UsedRange.SpecialCells(xlCellTypeLastCell).Row
'MsgBox "Frows: " & CStr(Frows)
Set DestTwo = SheetTwo.Cells(Frows + 1, 1) 'this is where we'll put the filtered data

' apply the autofilter to column with PI values
With DataBlock
    .AutoFilter picol, "<>*;*"
    'copy the still-visible cells to sheet 2
    .SpecialCells(xlCellTypeVisible).Copy DestTwo
End With

' turn off the autofilter
With SheetOne
    .AutoFilterMode = False
    If .FilterMode = True Then .ShowAllData
End With

' delete second header
SheetTwo.Rows(Frows + 1).EntireRow.Delete

' delete copy of original
DeleteSheet(Basename & " copy")

' end of sorting
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' save changes
If keepchanges Then
    wb.Save
End If

' use the sorted sheet for QuantAnalysis
sheetname = sheetname + " sorted"
Set ws = wb.Sheets(sheetname)

End If ' end of "If Not sorted Then"

initrow = 2 ' the first row contains the header

' LastRow and LastCol were determined before sorting

If debug Then
    MsgBox "Success!"
    wb.Close False
    xl.Quit

    Set xl = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set objFile = Nothing
    Set filelist = Nothing

    Wscript.Quit
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' open QuantAnalysis
'' click on "Method", "Next"
'' select "EIC", "All MS"
'' set Width value, Polarity

' check if QuantAnaylsis is already running and warn user if true
Dim Process, strObject, strProcess, IsProcessRunning
Const strComputer = "." 
strProcess = "QuantAnalysis.exe"
IsProcessRunning = False
strObject = "winmgmts://" & strComputer
For Each Process in GetObject(strObject).InstancesOf("win32_process")
    If UCase(Process.name) = UCase(strProcess) Then
        MsgBox "QuantAnalysis is already running! It might be safer to check" &_
               " it, save your work and close the application before using " &_
               "this script. Will abort now...", _
               vbOKOnly+vbCritical+vbSystemModal, _
               "QuantAnalysis method writer"
        wb.Close False
        xl.Quit
        Wscript.Quit
    End If
Next

result = MsgBox("The data will be extracted from the following Excel file: " &_
                vbCrLf & filepath & vbCrLf &_
                "The specified sheet (" & sheetname & ") contains " &_
                CStr(LastRow - 1) & " rows of data." & vbCrLf & vbCrLf &_
                "The method will be saved as: " & vbCrLf & methodpath &_
                vbCrLf & vbCrLf & "Now starting QuantAnalysis...", _
                vbOKCancel+vbInformation+vbSystemModal, _
                "QuantAnalysis method writer")
If result = vbCancel Then
    wb.Close False
    xl.Quit
    Wscript.Quit
End If

' open QuantAnaylsis
Set Shell = WScript.CreateObject("WScript.Shell")
Shell.Run("QuantAnalysis")
WScript.Sleep 6000 '3000

' open "Method..."
Shell.SendKeys "%{M}" ' alt+M
Shell.SendKeys "{DOWN 2}"
Shell.SendKeys "{ENTER}"
' click "Next"
Shell.SendKeys "{ENTER}"

' select EIC
Shell.SendKeys "{TAB}"
Shell.SendKeys "{DOWN}"
' select All MS
Shell.SendKeys "{TAB}"
Shell.SendKeys "{DOWN}"
' select mass error in +-Dalton
Shell.SendKeys "{TAB 2}"
Shell.SendKeys "{DOWN 2}" ' select +-0.3
' set Polarity to "negative"
Shell.SendKeys "{TAB}"
Shell.SendKeys "{DOWN 2}"

' go back prior to entering the first PI mass
Shell.SendKeys "+{TAB 2}"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' collect all unique PI values from the Excel file and write them into QA
Set pilist = CreateObject("System.Collections.ArrayList")
prevpi = 0
For row = initrow to lastrow
    pivalue = ws.Cells(row, picol).Value
    'If (pivalue <> prevpi) Then ' Does not work?
    If StrComp(pivalue, prevpi, 1) <> 0 Then 
        pilist.Add pivalue
        '''''''''''''''''''''''''''''''''''''''''''''
        
        '' generate EIC list
        ''   enter Masses
        ''   click "Add"
        
        Shell.SendKeys ws.Cells(row,picol).Value
        ' click "Add"
        Shell.SendKeys "{TAB 5}"
        Shell.SendKeys "{ENTER}"
        ' go back to mass input
        Shell.SendKeys "+{TAB 5}"
        
        '''''''''''''''''''''''''''''''''''''''''''''
        prevpi = pivalue
    End If
Next

'Wscript.Echo pilist.Count
'Wscript.Echo join(pilist.ToArray(), "| ")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' click "Next"
'' click "Next" to skip ISTD page

' click "Next"
Shell.SendKeys "{ENTER}"
' skip ISTD
Shell.SendKeys "{ENTER}"
' set focus to enter first CompoundName
Shell.SendKeys "{TAB}"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' enter all retention times
prevpi = ws.Cells(initrow,picol).Value
count = 1
For row = initrow to lastrow
    If (namecol > 0) Then
        compname = ws.Cells(row,namecol).Value
    Else
        compname = ""
    End If
    pivalue = ws.Cells(row,picol).Value
    rtvalue = ws.Cells(row,rtcol).Value
    rtwin   = ws.Cells(row,rtwcol).Value
    pipos   = pilist.IndexOf(pivalue, 0)
    If pivalue = prevpi Then
    
        '''''''''''''''''''''''''''''''''''''''''''
    
        '' enter every single compound
        ''   select EIC from list according to pi
        ''   set CompoundName, RT, RTwindow
        ''   click "Add"
    
        ' CompoundName; have always one digit after decimal point and no grouping of digits
        If (namecol > 0) Then
            Shell.SendKeys compname
        Else
            If IsNumeric(pivalue) Then 
                Shell.SendKeys CStr(FormatNumber(pivalue, 1,,, 0)) + "_" + CStr(count)
            Else
                Shell.SendKeys CStr(pivalue) + "_" + CStr(count)
            End If
        End If
        ' do not change EIC
        Shell.SendKeys "{TAB 2}"
        ' enter RT
        Shell.SendKeys "{TAB 2}"
        Shell.SendKeys CStr(rtvalue)
        ' enter RTwindow
        Shell.SendKeys "{TAB}"
        Shell.SendKeys CStr(rtwin)
        ' click "Add"
        Shell.SendKeys "{TAB 2}"
        Shell.SendKeys "{ENTER}"
        WScript.Sleep delayM ' wait for QA to respond otherwise irreproducible errors do occur
        ' go back to CompoundName field and clear it
        Shell.SendKeys "+{TAB 7}"
        Shell.SendKeys "{CLEAR}"
        
        '''''''''''''''''''''''''''''''''''''''''''
        
        count = count + 1
    Else
        count = 1
        '''''''''''''''''''''''''''''''''''''''''''
    
        ' CompoundName
        If (namecol > 0) Then
            Shell.SendKeys compname
        Else
            If IsNumeric(pivalue) Then 
                Shell.SendKeys CStr(FormatNumber(pivalue, 1,,, 0)) + "_" + CStr(count)
            Else
                Shell.SendKeys CStr(pivalue) + "_" + CStr(count)
            End If
        End If
        ' select EIC from list
        Shell.SendKeys "{TAB 2}"
        Shell.SendKeys "{DOWN}"
        ' enter RT
        Shell.SendKeys "{TAB 2}"
        Shell.SendKeys CStr(rtvalue)
        ' enter RTwindow
        Shell.SendKeys "{TAB}"
        Shell.SendKeys CStr(rtwin)
        ' click "Add"
        Shell.SendKeys "{TAB 2}"
        Shell.SendKeys "{ENTER}"
        WScript.Sleep delayM ' wait for QA to respond otherwise irreproducible errors do occur
        ' go back to CompoundName field and clear it
        Shell.SendKeys "+{TAB 7}"
        Shell.SendKeys "{CLEAR}"
        
        '''''''''''''''''''''''''''''''''''''''''''
        
        count = count + 1
        prevpi = pivalue
    End If
Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' click "Next"
'' check "Apply to all chromatograms"
'' set Smoothing width to 3
'' click "Finish"
'' save Method

' click "Next"
Shell.SendKeys "{ENTER}"

' enter smoothing width
Shell.SendKeys "{TAB 7}"
Shell.SendKeys "3"
' apply to all chromatograms
Shell.SendKeys "{TAB 3}"
Shell.SendKeys "{+}" ' check checkbox; {SPACE} to toggle; {-} to uncheck
' click "Finish"
Shell.SendKeys "{ENTER}"

WScript.Sleep delayL

' save method file
Shell.SendKeys "%{M}" ' alt+M
Shell.SendKeys "{DOWN}"
Shell.SendKeys "{ENTER}"

WScript.Sleep (delayL * 2)

' enter complete path with filename
Shell.SendKeys "%{M}" ' alt+M
Shell.SendKeys "{CLEAR}"
Shell.SendKeys methodpath ' extension is not nessessary
Shell.SendKeys "{ENTER}", True ' if the file exists you will be prompted
' but the script does not wait for your answer to the prompt
' thus it's best to delete the method with the same name beforehand
' or use a second script file to populate the work table
' or use a dialog that you have to accept before it start's populating the table

WScript.Sleep (delayL * 2)

result = MsgBox("Now the Work Table will be populated." &_
                vbCrLf & vbCrLf &_
                "Please, switch to Nested or Work Table view!", _
                vbOKCancel+vbInformation+vbSystemModal, _
                "QuantAnalysis method writer")
If result = vbCancel Then
    wb.Close False
    xl.Quit
    Wscript.Quit
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' set focus to Table
'' enter filenames
'' set some fields to "1"

' populate work table
' the QA WorkTable is very slow responding, thus we need to wait alot

' go to work table
Shell.SendKeys "+{TAB 5}"
WScript.Sleep delayS
' go to cell A1
Shell.SendKeys "{DOWN}"
WScript.Sleep delayS
Shell.SendKeys "{RIGHT}"
WScript.Sleep delayS

'Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set objFolder = objFSO.GetFolder(cd)
For Each objFile In objFolder.SubFolders
    'only proceed if there is an extension on the file.
    If (InStr(objFile.Name, ".") > 0) Then
        'If the file's extension is "xlsx", echo the path to it.
        If (LCase(Mid(objFile.Name, InStrRev(objFile.Name, "."))) = ".d") Then
            'Wscript.Echo objFile.Path
            'Wscript.Echo objFile.Name
            
            '''''''''''''''''''''''''''''''''''''
            
            ' enter filename
            Shell.SendKeys CStr(objFile.Name)
            WScript.Sleep delayS
            ' enter Inj.Vol.
            Shell.SendKeys "{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "1"
            WScript.Sleep delayS
            ' enter Dil.Factor
            Shell.SendKeys "{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "1"
            WScript.Sleep delayS
            ' enter Inj.No.
            Shell.SendKeys "{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "1"
            WScript.Sleep delayS
            ' go to first cell of next line
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delayS
            Shell.SendKeys "{DOWN}"
            WScript.Sleep delayS
            
            '''''''''''''''''''''''''''''''''''''
            
        End If
    End If
Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

MsgBox "Dont't forget to enter sample names and to " &_
        "save the batch file before processing!", _
        vbOKOnly+vbExclamation+vbSystemModal, _
        "QuantAnalysis method writer"

' manually add sample names (important for plotting in R)
' save Batch file before Process
' manual process and check and save under different name

' close everything, otherwise you cannot open the file in Excel
' you can see that is still open, when you change the visibilty to true
'xl.visible = TRUE
wb.Close False ' do not save any changes to the xlsx file
xl.Quit

Set Shell = Nothing
Set xl = Nothing
Set ws = Nothing
Set wb = Nothing
Set objFSO = Nothing
Set objFolder = Nothing
Set objFile = Nothing
Set filelist = Nothing
Set pilist = Nothing

Wscript.Quit
