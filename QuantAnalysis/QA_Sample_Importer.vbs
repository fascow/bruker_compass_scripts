Option Explicit

' QuantAnalysis Method Writer
'
' SHORT VERSION: only populates the QA work table (17.02.2015)
'
' Take a list of (Compund Names,) PIs and RTs from an Excel file to
' create a Bruker Compass QuantAnalysis Method file
' Tested with Windows 7 and QuantAnalysis 2.1
' the *.d data files have to be in the same folder as well
' it is advised to use copies and not your original files

' IMPORTANT
' This script assumes that QuantAnalysis opens in Nested or Work Table view
' This script assumes that you use the standard table layout of QuantAnalysis

Dim sheetname, cd, filepath, methodfile, methodpath, compname ' strings
Dim filelist, pilist                    ' lists
Dim objFSO, objFolder, objFile          ' file system objects
Dim xl, Shell                           ' app opjects
Dim wb, ws                              ' Excel objects
Dim result                              ' MsgBox result
Dim row, initrow, lastrow, count, pipos, i ' integers
Dim namecol, picol, rtcol, rtwcol		' integers: column numbers
Dim pivalue, prevpi, rtvalue, rtwin     ' double
Dim delayS, delayM, delayL				' integers: delays

'##############################'
'######### SET VALUES #########'

REM ' where to find you table? There shall be only one excel file in the folder, where you run this script
REM sheetname  = "300 sorted"

REM ' how to name the method file?
REM methodfile = "OG-library_QA-Method_300_20150217" '"QAmethod"

REM ' in which columns to find the values?
REM ' the header (first row) will be ignored, this script will only look for the column numbers specified here
REM namecol = 1 ' set this to -1 if you have no compound names or 
			REM ' simply want the legacy version: compound name = pivalue_1 , pivalue_2, ...
REM picol   = 2
REM rtcol   = 3
REM rtwcol  = 4

' how slow is you computer? specify delays in milliseconds
delayS = 50 '20
delayM = 100 '50
delayL = 500

'##############################'
'##############################'

' get date of today
REM Dim dt
REM dt = now
REM 'output format: yyyymmddHHnn
REM wscript.echo ((year(dt)*100 + month(dt))*100 + day(dt))*10000 + hour(dt)*100 + minute(dt)

' find xlsx file in current directory
Set objFSO = CreateObject("Scripting.FileSystemObject")
cd = objFSO.GetParentFolderName(WScript.ScriptFullName)
Set objFolder = objFSO.GetFolder(cd)
REM Set filelist = CreateObject("System.Collections.ArrayList")
REM For Each objFile In objFolder.Files
	REM 'only proceed if there is an extension on the file.
	REM If (InStr(objFile.Name, ".") > 0) Then
		REM 'If the file's extension is "xlsx", echo the path to it.
		REM If (LCase(Mid(objFile.Name, InStrRev(objFile.Name, "."))) = ".xlsx") Then
			REM 'Wscript.Echo objFile.Path
			REM filelist.Add objFile.Path
		REM End If
	REM End If
REM Next

'Wscript.Echo filelist.Count
'Wscript.Echo filelist.Item(0)
'wscript.echo join(filelist.ToArray(), ", ")

REM If filelist.Count = 0 Then
	REM MsgBox "No Excel file was found!", vbOKOnly+vbCritical+vbSystemModal, "QuantAnalysis method writer"
	REM Wscript.Quit
REM Else
	REM filepath = filelist.Item(0)
REM End If

REM methodpath = cd + "\" + methodfile + ".m"
REM 'MsgBox methodpath
REM 'Wscript.Quit

REM If (objFSO.FolderExists(methodpath)) Then ' the Bruker methods are actually folders!
	REM MsgBox "The specified method file already exists!", vbOKOnly+vbCritical+vbSystemModal, "QuantAnalysis method writer"
	REM Wscript.Quit
REM End If

REM Set xl = CreateObject("Excel.Application")
REM Set wb = xl.Workbooks.Open(filepath) 'Set path to your file

REM Dim SheetExists
REM SheetExists = False
REM For Each ws In wb.Worksheets
    REM If ws.Name = sheetname Then
        REM SheetExists = True
        REM Exit For
    REM End If
REM Next 

REM If Not SheetExists Then 
	REM MsgBox "The specified Excel worksheet was not found!", vbOKOnly+vbCritical+vbSystemModal, "QuantAnalysis method writer"
	REM wb.Close False
	REM xl.Quit
	REM Wscript.Quit
REM End If 

REM Set ws = wb.Sheets(sheetname) 'Set name of your worksheet

REM ' find the last used row on the worksheet
REM ' unfortunately the first variant only works when you are inside Excel?
REM With ws
    REM If xl.WorksheetFunction.CountA(.Cells) <> 0 Then
        REM lastrow = .Cells.Find(What:="*", _
                      REM After:=.Range("A1"), _
                      REM Lookat:=xlPart, _
                      REM LookIn:=xlFormulas, _
                      REM SearchOrder:=xlByRows, _
                      REM SearchDirection:=xlPrevious, _
                      REM MatchCase:=False).Row
    REM Else
        REM lastrow = 1
    REM End If
REM End With

REM initrow = 2 ' the first row contains the header
REM ' column C contains the retention times and is filled completely
REM ' thus we can use it to count the total number of rows

REM With ws
    REM lastrow = .Range("A" & .Rows.Count).End(-4162).Row
REM End With

REM 'Wscript.Echo lastrow

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' open QuantAnalysis
'' click on "Method", "Next"
'' select "EIC", "All MS"
'' set Width value, Polarity

' check if QuantAnaylsis is already running and warn user if true
Dim Process, strObject, strProcess, IsProcessRunning
Const strComputer = "." 
strProcess = "QuantAnalysis.exe"
IsProcessRunning = False
strObject   = "winmgmts://" & strComputer
For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
If UCase( Process.name ) = UCase( strProcess ) Then
        MsgBox "QuantAnalysis is already running! It might be safer to check it" _
		+", save your work and close the application before using this script. Will abort now...", _
		vbOKOnly+vbCritical+vbSystemModal, "QuantAnalysis method writer"
		wb.Close False
		xl.Quit
		Wscript.Quit
    End If
Next

REM result = MsgBox ("The data will be extracted from the following Excel file: " _
				REM + vbCrLf + filepath + vbCrLf + "The specified sheet (" _
				REM + sheetname + ") contains " + CStr(lastrow-1) + " rows of data." _
				REM + vbCrLf + vbCrLf + "The method will be saved as: " + vbCrLf _
				REM + methodpath + vbCrLf + vbCrLf + "Now starting QuantAnalysis...", _
				REM vbOKCancel+vbInformation+vbSystemModal, "QuantAnalysis method writer")
REM If result = vbCancel Then
	REM wb.Close False
	REM xl.Quit
	REM Wscript.Quit
REM End If

' open QuantAnaylsis
Set Shell = WScript.CreateObject("WScript.Shell")
Shell.Run("QuantAnalysis")
WScript.Sleep 5000

REM ' open "Method..."
REM Shell.SendKeys "%{M}" ' alt+M
REM Shell.SendKeys "{DOWN 2}"
REM Shell.SendKeys "{ENTER}"
REM ' click "Next"
REM Shell.SendKeys "{ENTER}"

REM ' select EIC
REM Shell.SendKeys "{TAB}"
REM Shell.SendKeys "{DOWN}"
REM ' select All MS
REM Shell.SendKeys "{TAB}"
REM Shell.SendKeys "{DOWN}"
REM ' select mass error in +-Dalton
REM Shell.SendKeys "{TAB 2}"
REM Shell.SendKeys "{DOWN 2}" ' select +-0.3
REM ' set Polarity to "negative"
REM Shell.SendKeys "{TAB}"
REM Shell.SendKeys "{DOWN 2}"

REM ' go back prior to entering the first PI mass
REM Shell.SendKeys "+{TAB 2}"

REM ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


REM ' collect all unique PI values from the Excel file and write them into QA
REM Set pilist = CreateObject("System.Collections.ArrayList")
REM prevpi = 0
REM For row = initrow to lastrow
	REM pivalue = ws.Cells(row,picol).Value
	REM If pivalue <> prevpi Then
		REM pilist.Add pivalue
		REM '''''''''''''''''''''''''''''''''''''''''''''
		
		REM '' generate EIC list
		REM ''   enter Masses
		REM ''   click "Add"
		
		REM Shell.SendKeys ws.Cells(row,picol).Value
		REM ' click "Add"
		REM Shell.SendKeys "{TAB 5}"
		REM Shell.SendKeys "{ENTER}"
		REM WScript.Sleep delayS ' normally not needed, but essential in OG data set as of January 2015
		REM ' go back to mass input
		REM Shell.SendKeys "+{TAB 5}"
		
		REM '''''''''''''''''''''''''''''''''''''''''''''
		REM prevpi = pivalue
	REM End If
REM Next

REM 'Wscript.Echo pilist.Count
REM 'Wscript.Echo join(pilist.ToArray(), "| ")

REM ''''''''''''''''''''''''''''''''''''''''''''''

REM '' click "Next"
REM '' click "Next" to skip ISTD page

REM ' click "Next"
REM Shell.SendKeys "{ENTER}"
REM ' skip ISTD
REM Shell.SendKeys "{ENTER}"
REM ' set focus to enter first CompoundName
REM Shell.SendKeys "{TAB}"

REM ''''''''''''''''''''''''''''''''''''''''''''''

REM ' enter all retention times
REM prevpi = ws.Cells(initrow,picol).Value
REM count = 1
REM For row = initrow to lastrow
	REM If (namecol > 0) Then
		REM compname = ws.Cells(row,namecol).Value
	REM Else
		REM compname = ""
	REM End If
	REM pivalue = ws.Cells(row,picol).Value
	REM rtvalue = ws.Cells(row,rtcol).Value
	REM rtwin   = ws.Cells(row,rtwcol).Value
	REM pipos   = pilist.IndexOf(pivalue, 0)
	REM If pivalue = prevpi Then
	
		REM '''''''''''''''''''''''''''''''''''''''''''
	
		REM '' enter every single compound
		REM ''   select EIC from list according to pi
		REM ''   set CompoundName, RT, RTwindow
		REM ''   click "Add"
	
		REM ' CompoundName; have always one digit after decimal point and no grouping of digits
		REM If (namecol > 0) Then
			REM Shell.SendKeys compname
		REM Else
			REM If IsNumeric(pivalue) Then 
				REM Shell.SendKeys CStr(FormatNumber(pivalue, 1,,, 0)) + "_" + CStr(count)
			REM Else
				REM Shell.SendKeys CStr(pivalue) + "_" + CStr(count)
			REM End If
		REM End If
		REM ' do not change EIC
		REM Shell.SendKeys "{TAB 2}"
		REM ' enter RT
		REM Shell.SendKeys "{TAB 2}"
		REM Shell.SendKeys CStr(rtvalue)
		REM ' enter RTwindow
		REM Shell.SendKeys "{TAB}"
		REM Shell.SendKeys CStr(rtwin)
		REM ' click "Add"
		REM Shell.SendKeys "{TAB 2}"
		REM Shell.SendKeys "{ENTER}"
		REM WScript.Sleep delayM ' wait for QA to respond otherwise irreproducible errrors do occur
		REM ' go back to CompoundName field and clear it
		REM Shell.SendKeys "+{TAB 7}"
		REM Shell.SendKeys "{CLEAR}"
		
		REM '''''''''''''''''''''''''''''''''''''''''''
		
		REM count = count + 1
	REM Else
		REM count = 1
		REM '''''''''''''''''''''''''''''''''''''''''''
	
		REM ' CompoundName
		REM If (namecol > 0) Then
			REM Shell.SendKeys compname
		REM Else
			REM If IsNumeric(pivalue) Then 
				REM Shell.SendKeys CStr(FormatNumber(pivalue, 1,,, 0)) + "_" + CStr(count)
			REM Else
				REM Shell.SendKeys CStr(pivalue) + "_" + CStr(count)
			REM End If
		REM End If
		REM ' select EIC from list
		REM Shell.SendKeys "{TAB 2}"
		REM Shell.SendKeys "{DOWN}"
		REM ' enter RT
		REM Shell.SendKeys "{TAB 2}"
		REM Shell.SendKeys CStr(rtvalue)
		REM ' enter RTwindow
		REM Shell.SendKeys "{TAB}"
		REM Shell.SendKeys CStr(rtwin)
		REM ' click "Add"
		REM Shell.SendKeys "{TAB 2}"
		REM Shell.SendKeys "{ENTER}"
		REM WScript.Sleep delayM ' wait for QA to respond otherwise irreproducible errrors do occur
		REM ' go back to CompoundName field and clear it
		REM Shell.SendKeys "+{TAB 7}"
		REM Shell.SendKeys "{CLEAR}"
		
		REM '''''''''''''''''''''''''''''''''''''''''''
		
		REM count = count + 1
		REM prevpi = pivalue
	REM End If
REM Next

REM '''''''''''''''''''''''''''''''''''''''''''''''

REM '' click "Next"
REM '' check "Apply to all chromatograms"
REM '' set Smoothing width to 3
REM '' click "Finish"
REM '' save Method

REM ' click "Next"
REM Shell.SendKeys "{ENTER}"

REM ' enter smoothing width
REM Shell.SendKeys "{TAB 7}"
REM Shell.SendKeys "3"
REM ' apply to all chromatograms
REM Shell.SendKeys "{TAB 3}"
REM Shell.SendKeys "{+}" ' check checkbox; {SPACE} to toggle; {-} to uncheck
REM ' click "Finish"
REM Shell.SendKeys "{ENTER}"

REM WScript.Sleep delayL

REM ' save method file
REM Shell.SendKeys "%{M}" ' alt+M
REM Shell.SendKeys "{DOWN}"
REM Shell.SendKeys "{ENTER}"

REM WScript.Sleep delayL

REM ' enter complete path with filename
REM Shell.SendKeys "%{M}" ' alt+M
REM Shell.SendKeys "{CLEAR}"
REM Shell.SendKeys methodpath ' extension is not nessessary
REM Shell.SendKeys "{ENTER}", True ' if the file exists you will be prompted
REM ' but the script does not wait for your answer to the prompt
REM ' thus it's best to delete the methode with the same name beforehand
REM ' or use a second script file to populate the work table
REM ' or use a dialog that you have to accept before it start's populating the table

WScript.Sleep (delayL * 2)

result = MsgBox ("Now the Work Table will be populated." _
				+ vbCrLf + vbCrLf _
				+ "Please, switch to Nested or Work Table view!", _
				vbOKCancel+vbInformation+vbSystemModal, _
				"QuantAnalysis method writer")
If result = vbCancel Then
	wb.Close False
	xl.Quit
	Wscript.Quit
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''

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
			' go to B1
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''

MsgBox "Dont't forget to enter sample names and to " + _
	   "save the batch file before processing!", _
	   vbOKOnly+vbExclamation+vbSystemModal, _
	   "QuantAnalysis method writer"

' manually add sample names (important for plotting in R)
' save Batch file before Process
' manual process and check and save under different name

' close everything, otherwise you cannot open the file in Excel
' you can see that is still open, when you chang the visibilty to true
'xl.visible = TRUE
REM wb.Close False ' do not save any changes to the xlsx file
REM xl.Quit

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