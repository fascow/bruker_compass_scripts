Option Explicit

' DataAnalysis EIC Export
'
' Extract the list of EICs from a DataAnalysis file and save it in an Excel table

' IMPORTANT
' Start this script by double clicking it in Windows Explorer (do not save in DA method)
' This script assumes that DataAnalysis is running and only one file is active (checked)
' Excel has to be closed before starting the script
' A new Excel file will be opened and you need to save it manually
' You cannot use your PC while the script is running
' It is not possible to distinguish the kind nor state of the EIC
'   E.g. kind: MS or MSn; state: checked or not

' TODO
' Implement a way of automatically saving the Excel file
' Use a different approach: Save values in an array, and later save it to the Excel file

Dim xl, Shell                           ' app opjects
Dim wb, ws                              ' Excel objects
Dim row, lastvalue, newvalue, delay     ' integers
Dim result                              ' MsgBox result

result = MsgBox ("Starting to extract EICs...", _
    vbOKCancel+vbInformation+vbSystemModal, _
    "DataAnalysis EIC Export")
If result = vbCancel Then
    Wscript.Quit
End If

' Start Excel, open a new workbook
Set xl = CreateObject("Excel.Application")
xl.Visible = True
Set wb = xl.Workbooks.Add()
Set ws = wb.Worksheets(1)
' three sheets are there by default, so no new sheet is added

delay = 400 ' millisconds, maybe increase the value on a slower machine
row = 1
lastvalue = 0
newvalue = 1

Set Shell = WScript.CreateObject("WScript.Shell")
WScript.Sleep delay
Shell.AppActivate("Compass DataAnalysis")
WScript.Sleep delay
Shell.SendKeys "{F7}"
Shell.SendKeys "+{TAB}"
Shell.SendKeys "{DOWN 3}" ' skip BPC and TIC
Shell.SendKeys "{TAB}" ' because the first time 5*TAB is required

Do While True
    ' Copy and go to next
    Shell.SendKeys "{TAB 4}"
    Shell.SendKeys "^{c}"
    Shell.SendKeys "+{TAB 5}"
    Shell.SendKeys "{DOWN}"
    WScript.Sleep delay ' nessessary
    
    ' Write to Excel file
    Shell.AppActivate("Microsoft Excel")
    ws.Cells(row, 1).Select
    Shell.SendKeys "^{v}"
    WScript.Sleep 100 ' nessessary delay, otherwise newvalue will be empty
    newvalue = ws.Cells(row, 1).Value
    'MsgBox "Last: " + CStr(lastvalue) + "New: " + CStr(newvalue)
    If lastvalue = newvalue Then
        ws.Cells(row, 1).Value = ""
        'WScript.Sleep 100
        Shell.AppActivate("Compass DataAnalysis")
        WScript.Sleep 100
        ' on last, to close Chromatogram-window
        Shell.SendKeys "{ESC}"
        Exit Do
    Else 
        lastvalue = newvalue
        row = row + 1
    End If
    
    Shell.AppActivate("Compass DataAnalysis")
    WScript.Sleep 100
Loop

WScript.Sleep 100 ' Needed, without the delay the MsgBox is not displayed
MsgBox "Finished!", _
    vbOKOnly+vbInformation+vbSystemModal, _
    "DataAnalysis EIC Export"
Set Shell = Nothing
Wscript.Quit
