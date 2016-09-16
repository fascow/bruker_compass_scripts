Option Explicit

' QuantAnalysis Sample Importer
'
' Short Variant of Method Writer, this scipt only populates the QuantAnalysis 
' work table with all the Bruker LCMS data files in the current directory
'
' The *.d data files have to be in the same folder as the script
' Has been successfully tested with QuantAnalysis 2.1 + 2.2 on Windows 7 (64 bit)

' IMPORTANT
' This script assumes that QuantAnalysis opens in Nested or Work Table view
' This script assumes that you use the standard table layout of QuantAnalysis

Dim cd                         ' strings
Dim objFSO, objFolder, objFile ' file system objects
Dim Shell                      ' app opjects
Dim result                     ' MsgBox result
Dim delay                      ' integers

'##############################'
'######### SET VALUES #########'

' how slow is you computer? specify delays in milliseconds
delay = 50

'##############################'
'##############################'

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
strObject = "winmgmts://" & strComputer
For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
If UCase( Process.name ) = UCase( strProcess ) Then
        MsgBox "QuantAnalysis is already running! It might be safer to check" &_
               " it, save your work and close the application before using " &_
               "this script. Will abort now...", _
               vbOKOnly+vbCritical+vbSystemModal, "QuantAnalysis sample importer"
        Wscript.Quit
    End If
Next

' open QuantAnaylsis
Set Shell = WScript.CreateObject("WScript.Shell")
Shell.Run("QuantAnalysis")
WScript.Sleep 5000

result = MsgBox("Now the Work Table will be populated." &_
                vbCrLf + vbCrLf &_
                "Please, switch to Nested or Work Table view!", _
                vbOKCancel+vbInformation+vbSystemModal, _
                "QuantAnalysis sample importer")
If result = vbCancel Then
    Wscript.Quit
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''

'' set focus to Table
'' enter filenames
'' set some fields to "1"

' populate work table
' the QA WorkTable is very slow responding, thus we need to wait a lot

' go to work table
Shell.SendKeys "+{TAB 5}"
WScript.Sleep delay
' go to cell A1
Shell.SendKeys "{DOWN}"
WScript.Sleep delay
Shell.SendKeys "{RIGHT}"
WScript.Sleep delay


' find all *.d directories in current directory
Set objFSO = CreateObject("Scripting.FileSystemObject")
cd = objFSO.GetParentFolderName(WScript.ScriptFullName)
Set objFolder = objFSO.GetFolder(cd)
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
            WScript.Sleep delay
            ' enter Inj.Vol.
            Shell.SendKeys "{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "1"
            WScript.Sleep delay
            ' enter Dil.Factor
            Shell.SendKeys "{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "1"
            WScript.Sleep delay
            ' enter Inj.No.
            Shell.SendKeys "{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "1"
            WScript.Sleep delay
            ' go to B1
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "+{TAB}"
            WScript.Sleep delay
            Shell.SendKeys "{DOWN}"
            WScript.Sleep delay
            
            '''''''''''''''''''''''''''''''''''''
            
        End If
    End If
Next

'''''''''''''''''''''''''''''''''''''''''''''''''''''

MsgBox "Dont't forget to enter sample names and to " &_
       "save the batch file before processing!", _
       vbOKOnly+vbExclamation+vbSystemModal, _
       "QuantAnalysis sample importer"

' manually add sample names (important for plotting in R)
' save Batch file before Process
' manual process and check and save under different name

Set Shell = Nothing
Set objFSO = Nothing
Set objFolder = Nothing
Set objFile = Nothing

Wscript.Quit
