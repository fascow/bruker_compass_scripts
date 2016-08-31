option explicit 
 
' DataAnalysis EIC import and FindAutoMSn 
' 
' Take a list of precursor ions (PI) from an Excel file to 
' import into a Bruker DataAnalysis file 
' and create compounds by means of FindAutoMSn 
' Tested under Windows 7 with Excel 2010 and DataAnalysis 4.2 
' This script has needs to be stored in the DA method,  
' it is also possible to just load the script from a method into the current method 
' The Excel file should be in the same folder as the analysis files 
' there should only be one excel file in this folder 
' The sheetname has to be provided before using this script 
' It is advised to use copies and not you original files 
 
' IMPORTANT 
' The script needs to be stored in each DA method to be used
'   Do not start this script from Windows Explorer
' It is possible to just open the script from another DA method file 
' But be careful! Always test with copies of your original data! 
' You should not work in the Excel file while running the script, Save and Close it beforehand 
 
' TODO 
' should work with different analysis files using different Excel files  
' since it always searches for one Excel file in the folder that contains the analysis file, needs to be tested! 
' close Excel properly in case of error, should be fine now
 
Dim sheetname, cd, filepath             ' strings 
Dim filelist                            ' lists 
Dim objFSO, objFolder, objFile          ' file system objects 
Dim xl, Shell                           ' app opjects 
Dim wb, ws                              ' Excel objects 
Dim result                              ' MsgBox result 
Dim row, initrow, lastrow, picol        ' integers 
Dim pivalue, prevpi, width              ' strings containing floating point numbers 
Dim EIC                                 ' ChromatogramDefinition Object 
Dim Start_RT, End_RT, i                 ' integers 
Dim EnableNewEICs, DeleteExistingEICs   ' booleans 
Dim autoeic, autofind, autosave, enableTIC, disableTIC  ' booleans 
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
''' SETTINGS 
 
' Where to find the EICs in the Excel file? 
sheetname = "DA_EIC_LIST" 
picol   = 2 ' column containing precursor ion values for EIC 
initrow = 2 ' if the first row contains a header set this value to 2
 
' EIC handling 
disableTIC         = True ' disable TIC 
autoeic            = True ' Shall EICs from Excel file be imported into current Analysis? 
DeleteExistingEICs = True ' Shall all the existing EICs be deleted? 
width              = 0.3  ' EIC uncertainty in Dalton (results in EIC+-width) 
EnableNewEICs      = True ' Shall the newly generated EICs be enabled?
 
' Compound generation settings 
autofind = True ' Shall compounds be generated? old ones will be deleted first! 
Start_RT = 5  ' minutes
End_RT   = 70 ' minutes
 
' Autosave the anaylsis file 
autosave = True 
 
''' END OF SETTINGS 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
 
 
' Main script starts here 
 
'Analysis.save  
 
' Add EICs to Analysis 
If DeleteExistingEICs Then 
    For i = Analysis.Chromatograms.Count to 1 step -1  
        If Analysis.Chromatograms(i).Definition.Type = daEICChromType Then  
            Analysis.Chromatograms.DeleteChromatogram i  
        End If  
    Next  
End If 
 
' Disable TIC 
If disableTIC Then 
    For i = Analysis.Chromatograms.Count to 1 step -1 
        If Analysis.Chromatograms(i).Definition.Type = daTICChromType Then 
            Analysis.Chromatograms(i).Enable False 
        End If 
    Next 
' un-comment this section if you want to enable TIC with setting disableTIC = False
'Else 
'    For i = Analysis.Chromatograms.Count to 1 step -1 
'        If Analysis.Chromatograms(i).Definition.Type = daTICChromType Then 
'            Analysis.Chromatograms(i).Enable True 
'        End If 
'    Next 
End If 

' TIC toggle
'enableTIC = Not disableTIC
'For i = Analysis.Chromatograms.Count to 1 step -1 
'    If Analysis.Chromatograms(i).Definition.Type = daTICChromType Then 
'        Analysis.Chromatograms(i).Enable enableTIC 
'    End If 
'Next 

If autoeic Then 
    ' Find xlsx file in current directory 
    Set objFSO = CreateObject("Scripting.FileSystemObject") 
    cd = objFSO.GetParentFolderName(Analysis.Path) 
    Set objFolder = objFSO.GetFolder(cd) 
    Set filelist = CreateObject("System.Collections.ArrayList") 
    For Each objFile In objFolder.Files 
        ' Only proceed if there is an extension on the file. 
        If (InStr(objFile.Name, ".") > 0) Then 
            'If the file's extension is "xlsx", save the fiel path in list
            If (LCase(Mid(objFile.Name, InStrRev(objFile.Name, "."))) = ".xlsx") Then 
                filelist.Add objFile.Path 
            End If 
        End If 
    Next 
     
    If filelist.Count = 0 Then 
        MsgBox "No Excel file was found!", _
            vbOKOnly+vbCritical+vbSystemModal, _
            "DataAnalysis script" 
        Form.Close 
    Else 
        ' Pick the first Excel file
        filepath = filelist.Item(0) 
        'MsgBox filepath 
    End If 
     
    Set xl = CreateObject("Excel.Application") 
    Set wb = xl.Workbooks.Open(filepath) ' Set path to your file 
     
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
            "DataAnalysis script" 
        wb.Close False 
        xl.Quit 
        Form.Close 
    End If  
     
    Set ws = wb.Sheets(sheetname) ' Set name of your worksheet 
     
    
    ' Column A needs to be filled completely
    ' then we can use it to count the total number of rows  
    With ws 
        lastrow = .Range("A" & .Rows.Count).End(-4162).Row 
    End With 
     
    'result = MsgBox ("The data will be extracted from the following Excel file: " _ 
    '                + vbCrLf + filepath + vbCrLf + "The specified sheet (" _ 
    '                + sheetname + ") contains " + CStr(lastrow-1) + " rows of data.", _ 
    '                vbOKCancel+vbInformation+vbSystemModal, "DataAnalysis script") 
    'If result = vbCancel Then 
    '    wb.Close False 
    '    xl.Quit 
    '    Form.Close 
    'End If 
     
    ' Collect all unique PI values from the Excel file and add them to current Analysis 
    prevpi = "" 
    For row = initrow to lastrow 
        pivalue = ws.Cells(row,picol).Value 
        'If (pivalue <> prevpi) Then ' Does not work 
        If StrComp(pivalue, prevpi, 1) <> 0 Then  
     
            Set EIC = CreateObject("DataAnalysis.EICChromatogramDefinition") 
     
            'EIC.Name_ = pivalue ' Does not work, but is not needed anyway
            EIC.Range = pivalue 
            EIC.WidthLeft = width 
            EIC.WidthRight = width 
            EIC.Polarity = daNegative 
            EIC.MSFilter.Type = daMSFilterMS  
            EIC.ScanMode = daScanModeAll 
            EIC.BackgroundType = daBgrdTypeNone 
            EIC.Color = daBlack ' Would random colors be possible? 
    
            Analysis.Chromatograms.AddChromatogram EIC 
    
            Analysis.Chromatograms(Analysis.Chromatograms.Count).Enable EnableNewEICs 
    
            Set EIC = Nothing 
    
            prevpi = pivalue 
        End If 
    Next 
    wb.Close False 
    xl.Quit 
End If
 
' FindAutoMSn 
If autofind Then 
    ' Delete all compounds 
    ' In case there are alredy some compounds, they would not be overridden by FindAutoMSn 
    For i = Analysis.Compounds.Count to 1 step -1 
        Analysis.Compounds.DeleteCompound i 
    Next 
 
    ' FindAutoMSn to find compounds 
    ' Clear all selected time ranges beforehand 
    ' since it is possible to have multiple selections and it would not work as expected 
    Analysis.ClearChromatogramRangeSelections 
    Analysis.AddChromatogramRangeSelection Start_RT, End_RT, 0, 0 
    Analysis.FindAutoMSn 
End IF 
 
If autosave Then 
    Analysis.save  
End If 

Form.Close
