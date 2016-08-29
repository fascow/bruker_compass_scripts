Option Explicit

' Library Spectra Export (aka Spectra Liberator)
'
' Opens the specified DataAnalysis LibraryEditor spectra library
' Saves every spectrum to a .spectrum file in a folder ending with "_Xpec"
' Exported spectra contain pure text data, that can be further processed
' Tested under Windows 7 (64bit), LibraryEditor 4.2 (Build 391, 64bit)

' IMPORTANT
' Start this script by double clicking it in Windows Explorer
' The library is assumed to be located in the same folder as the script itself

' TODO
' if more than one spectrum is contained in a compound, only the first one is exported
' check if some additional error checking is needed somewhere
' check if we can somehow facilitate clip.exe, since this would 
'   allow us to keep track of the current compound

Dim Shell                                  ' Shell objects
Dim objFSO, objFolder, objArchive, objFile ' FSO objects
Dim filelist                               ' Array
Dim libraryname, filepath, cd              ' Strings
Dim count, first, skip, i                  ' Integers
Dim delayS, delayM, delayL				   ' Integers


'##############################'
'######### SET VALUES #########'

' Where to find your library, needs to be in the same folder as this script
libraryname = "My spectra library"

' Which spectra do you want to save?
first  = 1
count  = 0 ' let this be zero to ask the user interactively

' How slow is you computer? specify delays in milliseconds
delayS = 20
delayM = 100 ' increase M and L in case the script fails
delayL = 400

'##############################'
'##############################'

' Main script starts here

' Get current folder path and find specified library
Set objFSO = CreateObject("Scripting.FileSystemObject")
cd = objFSO.GetParentFolderName(WScript.ScriptFullName)
If Not objFSO.FolderExists(cd + "\" + libraryname) Then
    MsgBox "Library was not found!", _
      vbOKOnly+vbCritical+vbSystemModal, _
      "Library Spectra Export"
	Wscript.Quit
Else
    Set objFolder = objFSO.GetFolder(cd + "\" + libraryname)
    ' just for testing
    REM If Not objFSO.FolderExists(objFolder + "_Xpec") Then
        REM Set objArchive = objFSO.CreateFolder(objFolder + "_Xpec")
    REM Else
        REM Set objArchive = objFSO.GetFolder(objFolder + "_Xpec")
    REM End If
    ' create folder for extracted spectra (Xpec)
    If objFSO.FolderExists(objFolder + "_Xpec") Then
        MsgBox "Spectra archive (Xpec) already exists! Will abort now!", _
          vbOKOnly+vbCritical+vbSystemModal, _
          "Library Spectra Export"
        Wscript.Quit
    Else
        Set objArchive = objFSO.CreateFolder(objFolder + "_Xpec")
    End If
End If

' Find the .mlb file in the Library
Set filelist = CreateObject("System.Collections.ArrayList")
For Each objFile In objFolder.Files
	' Only proceed if there is an extension on the file.
	If (InStr(objFile.Name, ".") > 0) Then
		' If the file's extension is "mlb", echo the path to it.
		If (LCase(Mid(objFile.Name, InStrRev(objFile.Name, "."))) = ".mlb") Then
			'Wscript.Echo objFile.Path
			filelist.Add objFile.Path
		End If
	End If
Next

If filelist.Count = 0 Then
	MsgBox "No .mlb file was found in the library!", _
      vbOKOnly+vbCritical+vbSystemModal, _
      "Library Spectra Export"
	Wscript.Quit
Else
    'MsgBox CStr(filelist.Count) + ": " + filelist.Item(0)
	filepath = filelist.Item(0)
End If

' Check if LibraryEditor is already running and warn user if true
Dim Process, strObject, strProcess, IsProcessRunning
Const strComputer = "." 
strProcess = "LibraryEditor.exe"
IsProcessRunning = False
strObject   = "winmgmts://" & strComputer
For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
If UCase( Process.name ) = UCase( strProcess ) Then
        MsgBox "LibraryEditor is already running! It might be safer to check it" _
		+", save your work and close the application before using this script. Will abort now...", _
		vbOKOnly+vbCritical+vbSystemModal, "Library Spectra Export"
		Wscript.Quit
    End If
Next

' Open LibraryEditor
Set Shell = WScript.CreateObject("WScript.Shell")
Shell.Run("LibraryEditor")
WScript.Sleep 1000

' Open library file
Shell.SendKeys "^{o}" ' ctrl+O
Shell.SendKeys " " + filepath ' enter an additional letter first to avoid loss of first letter
Shell.SendKeys "{ENTER}"
WScript.Sleep 1000

' Determine the number of compounds to export
If count = 0 Then
    ' view properties
    Shell.SendKeys "%{ENTER}" ' alt+ENTER
    ' ask user to enter the number of compounds
    count = InputBox("How many compounds are in this library?", _
                     "Library Spectra Export")
    Shell.SendKeys "{ENTER}"
End If

WScript.Sleep 2000

If first > 1 Then
    skip = first - 1 
    Shell.SendKeys "%{RIGHT skip}"
End If

' Loop over compounds and export spectra
For i = first to count
    Shell.SendKeys "+{RIGHT 4}" ' select compound number; max 4 digits compound numbers
    Shell.SendKeys "^{c}" ' copy to clipboard alternative: alt+E alt+C
    WScript.Sleep delayL
    Shell.SendKeys "%{S}" ' spectrum menu
    WScript.Sleep delayM
    'Shell.SendKeys "%{E}" ' export spectrum; selects Edit menu, so do it with arrow keys
    Shell.SendKeys "{DOWN}"
    Shell.SendKeys "{ENTER}" ' export spectrum
    WScript.Sleep delayM ' wait 100 ms or enter two characters before starting to avoid loss of first letters
    'Shell.SendKeys "  " ' would be faster, but does not always work
    ' Enter file name for spectrum
    Shell.SendKeys objArchive + "\"
    Shell.SendKeys "^{v}"
    WScript.Sleep delayM
    Shell.SendKeys ".spectrum"
    Shell.SendKeys "{ENTER}" ' save
    WScript.Sleep delayM
    ' Go to next compound
    Shell.SendKeys "%{RIGHT}" ' alt+RIGHT
    WScript.Sleep delayL
Next

Set Shell = Nothing

Set objFSO = Nothing
Set objFolder = Nothing
Set objArchive = Nothing
Set objFile = Nothing

Wscript.Quit
