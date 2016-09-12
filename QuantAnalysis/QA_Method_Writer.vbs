Option Explicit

' QuantAnalysis Method Writer
'
' Take a list of PIs and RTs from an Excel file to create a
' Bruker Compass QuantAnalysis Method file by GUI scripting
' Tested under Windows 7 with Excel 2010 and Quantanalysis 2.1
' this script has to be in the same directory as the excel file
' there should only be one excel file in this folder
' the sheetname has to be provided before using his script
' the data has to be sorted in a special way, see example file
' the method file will be saved in the same folder
' the *.d data files have to be in the same folder as well
' it is advised to use copies and not you original files

' IMPORTANT
' This script assumes that QuantAnalysis opens in Nested or Work Table view
' If this is not the case, you have to switch to one of these views
' Then close the application and try to run the script again

Dim sheetname, cd, filepath, methodfile, methodpath ' strings
Dim filelist, pilist                                ' lists
Dim objFSO, objFolder, objFile                      ' file system objects
Dim xl, Shell                                       ' app opjects
Dim wb, ws                                          ' Excel objects
Dim result                                          ' MsgBox result
Dim row, initrow, lastrow, count, pipos, i          ' integers
Dim pivalue, prevpi, rtvalue, rtwin                 ' double

'##############################'
'######### SET VALUES #########'

sheetname  = "Sheet1"
methodfile = "QAmethod"

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
Else
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
Set wb = xl.Workbooks.Open(filepath) 'Set path to your file

Dim SheetExists
SheetExists = False
For Each ws In wb.Worksheets
    If ws.Name = sheetname Then
        SheetExists = True
        Exit For
    End If
Next 

If Not SheetExists Then 
	MsgBox "The specified Excel worksheet was not found!", _
        vbOKOnly+vbCritical+vbSystemModal, _
        "QuantAnalysis method writer"
	Wscript.Quit
End If 

Set ws = wb.Sheets(sheetname) 'Set name of your worksheet

' find the last used row on the worksheet
' unfortunately the first variant only works when you are inside Excel?
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

initrow = 2 ' the first row contains the header
' column C contains the retention times and is filled completely
' thus we can use it to count the total number of rows

With ws
    lastrow = .Range("C" & .Rows.Count).End(-4162).Row
End With

'Wscript.Echo lastrow

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
		Wscript.Quit
    End If
Next

result = MsgBox ("The data will be extracted from the following Excel file: " _
				+ vbCrLf + filepath + vbCrLf + "The specified sheet (" _
				+ sheetname + ") contains " + CStr(lastrow-1) + " rows of data." _
				+ vbCrLf + vbCrLf + "The method will be saved as: " + vbCrLf _
				+ methodpath + vbCrLf + vbCrLf + "Now starting QuantAnalysis...", _
				vbOKCancel+vbInformation+vbSystemModal, "QuantAnalysis method writer")
If result = vbCancel Then
	Wscript.Quit
End If

' open QuantAnaylsis
Set Shell = WScript.CreateObject("WScript.Shell")
Shell.Run("QuantAnalysis")
WScript.Sleep 3000

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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' collect all unique PI values from the Excel file and write them into QA
Set pilist = CreateObject("System.Collections.ArrayList")
prevpi = 0
For row = initrow to lastrow
	pivalue = ws.Cells(row,1).Value
	If pivalue <> prevpi Then
		pilist.Add pivalue
		'''''''''''''''''''''''''''''''''''''''''''''
		
		'' generate EIC list
		''   enter Masses
		''   click "Add"
		
		Shell.SendKeys ws.Cells(row,1).Value
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

''''''''''''''''''''''''''''''''''''''''''''''

'' click "Next"
'' click "Next" to skip ISTD page

' click "Next"
Shell.SendKeys "{ENTER}"
' skip ISTD
Shell.SendKeys "{ENTER}"
' set focus to enter first CompoundName
Shell.SendKeys "{TAB}"

''''''''''''''''''''''''''''''''''''''''''''''

' enter all retention times
prevpi = ws.Cells(initrow,1).Value
count = 1
For row = initrow to lastrow
	pivalue = ws.Cells(row,1).Value
	rtvalue = ws.Cells(row,3).Value
	rtwin   = ws.Cells(row,4).Value
	pipos   = pilist.IndexOf(pivalue, 0)
	If pivalue = prevpi Then
	
		'''''''''''''''''''''''''''''''''''''''''''
	
		'' enter every single compound
		''   select EIC from list according to pi
		''   set CompoundName, RT, RTwindow
		''   click "Add"
	
		' CompoundName; have always one digit after decimal point and no grouping of digits
		If IsNumeric(pivalue) Then 
			Shell.SendKeys CStr(FormatNumber(pivalue, 1,,, 0)) + "_" + CStr(count)
		Else
			Shell.SendKeys CStr(pivalue) + "_" + CStr(count)
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
		WScript.Sleep 50 ' wait for QA to respond otherwise irreproducible errors do occur
		' go back to CompoundName field and clear it
		Shell.SendKeys "+{TAB 7}"
		Shell.SendKeys "{CLEAR}"
		
		'''''''''''''''''''''''''''''''''''''''''''
		
		count = count + 1
	Else
		count = 1
		'''''''''''''''''''''''''''''''''''''''''''
	
		' CompoundName
		If IsNumeric(pivalue) Then 
			Shell.SendKeys CStr(FormatNumber(pivalue, 1,,, 0)) + "_" + CStr(count)
		Else
			Shell.SendKeys CStr(pivalue) + "_" + CStr(count)
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
		WScript.Sleep 50 ' wait for QA to respond otherwise irreproducible errors do occur
		' go back to CompoundName field and clear it
		Shell.SendKeys "+{TAB 7}"
		Shell.SendKeys "{CLEAR}"
		
		'''''''''''''''''''''''''''''''''''''''''''
		
		count = count + 1
		prevpi = pivalue
	End If
Next

'''''''''''''''''''''''''''''''''''''''''''''''

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

WScript.Sleep 500

' save method file
Shell.SendKeys "%{M}" ' alt+M
Shell.SendKeys "{DOWN}"
Shell.SendKeys "{ENTER}"

WScript.Sleep 500

' enter complete path with filename
Shell.SendKeys "%{M}" ' alt+M
Shell.SendKeys "{CLEAR}"
Shell.SendKeys methodpath ' extension is not nessessary
Shell.SendKeys "{ENTER}", True ' if the file exists you will be prompted
' but the script does not wait for your answer to the prompt
' thus it's best to delete the methode with the same name beforehand
' or use a second script file to populate the work table
' or use a dialog that you have to accept before it start's populating the table

WScript.Sleep 1000

result = MsgBox ("Now the Work Table will be populated." _
				+ vbCrLf + vbCrLf _
				+ "Please, switch to Nested or Work Table view!", _
				vbOKCancel+vbInformation+vbSystemModal, _
				"QuantAnalysis method writer")
If result = vbCancel Then
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
WScript.Sleep 20
' go to cell A1
Shell.SendKeys "{DOWN}"
WScript.Sleep 20
Shell.SendKeys "{RIGHT}"
WScript.Sleep 20

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
			WScript.Sleep 20
			' enter Inj.Vol.
			Shell.SendKeys "{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "1"
			WScript.Sleep 20
			' enter Dil.Factor
			Shell.SendKeys "{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "1"
			WScript.Sleep 20
			' enter Inj.No.
			Shell.SendKeys "{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "1"
			WScript.Sleep 20
			' go to B1
			Shell.SendKeys "+{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "+{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "+{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "+{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "+{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "+{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "+{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "+{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "+{TAB}"
			WScript.Sleep 20
			Shell.SendKeys "{DOWN}"
			WScript.Sleep 20
			
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
