Attribute VB_Name = "ImportExcelFiles"
Option Explicit


' --------------------------------------------------------------------------------------------------------------------------------------------
'      Global Variables and Constants
' --------------------------------------------------------------------------------------------------------------------------------------------

Dim gbDEBUG As Boolean

Public gbOBJFSO As Scripting.FileSystemObject
Public gbSCRTEXT As Scripting.TextStream

Public Const MS_MODULENAME = "Bare_Metal.xlm"

Sub Main()

    Dim wkbCurrent As Workbook, wkbOut As Workbook
    Dim shtData As Worksheet
    Dim rngFileList As Range, currentCell As Range
    Dim lngLastRow As Long, lngBreakCount As Long
    Dim lngBareMetalOutRow As Long, lngVMServerRow As Long
    Dim strCurrentBuildFile As String
    Dim clsApp As clsAppMnemonicInfo
    Dim clsPtB As clsPtBStatus
    
    Set wkbOut = ThisWorkbook
    Set shtData = wkbOut.Worksheets("Files")
    Set clsApp = New clsAppMnemonicInfo
    Set clsPtB = New clsPtBStatus
    
    shtData.Activate
    
    lngLastRow = shtData.UsedRange.Rows.count
    lngBreakCount = 1

    Set rngFileList = shtData.Range("A2:A" & lngLastRow)
    lngCurrentOutRow = wkbOut.Worksheets("Composite List").UsedRange.Rows.count + 1
    
    For Each currentCell In rngFileList
        strCurrentBuildFile = currentCell.Value
        
        If shtData.Cells(currentCell.Row, 2).Value <> "x" Then
            Application.StatusBar = "Working on " & strCurrentBuildFile
            
            If strCurrentBuildFile <> "" Then
                Set wkbCurrent = Workbooks.Open(Filename:=strCurrentBuildFile, UpdateLinks:=False, ReadOnly:=True)
                Set clsApp = GetApplicationSummaryInfo(wkbCurrent)
                Set clsPtB = GetPtBStatus(wkbOut, clsApp.Mnemonic)
                lngVMServerRow = GetVMServerInfo(wkbCurrent)
                'lngBareMetalOutRow = BareMetal(wkbOut.Name, strCurrentBuildFile, "Bare Metal", lngBareMetalOutRow)
            End If
                    
            lngBreakCount = lngBreakCount + 1
            If lngBreakCount Mod 25 = 0 Then
                wkbOut.Save
                If lngBreakCount Mod 250 = 0 Then
                    RestoreExcel
                End If
                SpeedUpExcel
            End If
            shtData.Cells(currentCell.Row, 2).Value = "x"
            shtData.Cells(currentCell.Row, 2).Select

        End If
            
    Next currentCell
    
    RestoreExcel

End Sub

Function GetApplicationSummaryInfo(wkbCurrent As Workbook) As clsAppMnemonicInfo

    Dim shtCurrent As Worksheet
    Dim clsApp As clsAppMnemonicInfo
    
    Set shtCurrent = wkbCurrent.Sheets("Summary")
    Set clsApp = New clsAppMnemonicInfo
    
    clsApp.Mnemonic = shtCurrent.Cells(2, 2).Value
    clsApp.Name = shtCurrent.Cells(3, 2).Value
    clsApp.Description = shtCurrent.Cells(4, 2).Value
    clsApp.LOB = shtCurrent.Cells(5, 2).Value
    clsApp.LOBCIO = shtCurrent.Cells(6, 2).Value
    clsApp.LOBContactGroup = shtCurrent.Cells(7, 2).Value
    clsApp.LOBSupported = shtCurrent.Cells(8, 2).Value
    clsApp.LifecycleStatus = shtCurrent.Cells(9, 2).Value
    clsApp.DataClassification = shtCurrent.Cells(10, 2).Value
    clsApp.CBSTier = shtCurrent.Cells(11, 2).Value
    clsApp.WebAppClass = shtCurrent.Cells(12, 2).Value
    clsApp.DevType = shtCurrent.Cells(13, 2).Value
    clsApp.RTO = shtCurrent.Cells(14, 2).Value
    clsApp.RPO = shtCurrent.Cells(15, 2).Value
    clsApp.SOX = shtCurrent.Cells(16, 2).Value
    clsApp.SOC1 = shtCurrent.Cells(17, 2).Value
    clsApp.OFAC = shtCurrent.Cells(18, 2).Value
    clsApp.AvgConcurrentUsers = shtCurrent.Cells(19, 2).Value & " " & shtCurrent.Cells(19, 3).Value
    clsApp.PeakConcurrentUsers = shtCurrent.Cells(20, 2).Value & " " & shtCurrent.Cells(20, 3).Value
    clsApp.AvgROTrans = shtCurrent.Cells(21, 2).Value & " " & shtCurrent.Cells(21, 3).Value
    clsApp.AvgRWTrans = shtCurrent.Cells(22, 2).Value & " " & shtCurrent.Cells(22, 3).Value
    clsApp.AvgTotalTrans = shtCurrent.Cells(23, 2).Value & " " & shtCurrent.Cells(23, 3).Value
    clsApp.Mayor = shtCurrent.Cells(24, 2).Value
    clsApp.DomainArch = shtCurrent.Cells(25, 2).Value
    clsApp.IntegrationAnalyst = shtCurrent.Cells(26, 2).Value
    clsApp.ReleaseAnalyst = shtCurrent.Cells(27, 2).Value
    clsApp.MigrationManager = shtCurrent.Cells(28, 2).Value
    clsApp.WaveName = shtCurrent.Cells(29, 2).Value
    clsApp.WaveDate = shtCurrent.Cells(30, 2).Value
    
    'MsgBox clsApp.Name
    
    Set GetApplicationSummaryInfo = clsApp
End Function

Function GetPtBStatus(wkbData As Workbook, strAppMnemonic As String) As clsPtBStatus

    Dim clsPtB As clsPtBStatus
    Dim shtData As Worksheet, shtOut As Worksheet
    Dim strFirstFound As String
    Dim rngFound As Range, rngSearch As Range
    
    Set shtData = wkbData.Worksheets("PtB")
    Set clsPtB = New clsPtBStatus
       
    Set rngSearch = shtData.Range("B:B")
    Set rngFound = rngSearch.Find(WHAT:=strAppMnemonic, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)

    If Not rngFound Is Nothing Then
        strFirstFound = rngFound.Address
        Do
            If shtData.Cells(rngFound.Row(), 2).Value <> "" Then
                clsPtB.Mnemonic = strAppMnemonic
                clsPtB.MigrationWave = shtData.Cells(rngFound.Row(), 3).Value
                clsPtB.DDRComplete = shtData.Cells(rngFound.Row(), 4).Value
                clsPtB.LogicalDesignComplete = shtData.Cells(rngFound.Row(), 5).Value
                clsPtB.PtBComplete = shtData.Cells(rngFound.Row(), 6).Value
                
            End If
            Exit Do 'Match found

            Set rngFound = rngSearch.FindNext(After:=rngFound)
            If rngFound.Address = strFirstFound Then Set rngFound = Nothing
        Loop Until rngFound Is Nothing
    End If

    Set GetPtBStatus = clsPtB
End Function

Function GetVMServerInfo(wkbCurrent As Workbook) As Long
    
End Function

Function BareMetal(strOutFile As String, strCurrentBuildFile As String, strOutSheet As String, lngCurrentOutRow As Long) As Long

    Dim wkbCurrent As Workbook, wkbOut As Workbook
    Dim shtData As Worksheet, shtOut As Worksheet, shtCurrent As Worksheet
    Dim lngLastRow As Long, lngCurrentRow As Long, lngCurrentColumn As Long
    Dim strApp As String, strWave As String
    
    Set wkbOut = Workbooks(strOutFile)
    Set shtOut = wkbOut.Worksheets(strOutSheet)
    Set wkbCurrent = Workbooks.Open(Filename:=strCurrentBuildFile, UpdateLinks:=False, ReadOnly:=True)
    Set shtCurrent = wkbCurrent.Sheets("Bare Metal")
        
    lngLastRow = shtCurrent.UsedRange.Rows.count
    'shtCurrent.Range("A2:AV" & lngLastRow).Copy _
    '    Destination:=shtOut.Range("B" & lngCurrentOutRow)
    
    For lngCurrentRow = 4 To lngLastRow
        Application.StatusBar = "Working on " & strCurrentBuildFile & " - Line " & lngCurrentRow
        If shtCurrent.Cells(lngCurrentRow, 2) <> "" And shtCurrent.Cells(lngCurrentRow, 3) <> "" Then
            strApp = Left(shtCurrent.Cells(2, 1).Value, 3)
            strWave = Right(shtCurrent.Cells(2, 1).Value, (Len(shtCurrent.Cells(2, 1).Value) - 5))
            
            shtOut.Cells(lngCurrentOutRow, 1).Value = strCurrentBuildFile
            shtOut.Cells(lngCurrentOutRow, 2).Value = strApp
            shtOut.Cells(lngCurrentOutRow, 3).Value = strWave
            
            For lngCurrentColumn = 1 To ColumnLetterToNumber("AV")
                shtOut.Cells(lngCurrentOutRow, (lngCurrentColumn + 3)).Value = shtCurrent.Cells(lngCurrentRow, lngCurrentColumn).Value
            Next lngCurrentColumn
            lngCurrentOutRow = lngCurrentOutRow + 1
        End If
        

    Next lngCurrentRow
    
    wkbCurrent.Close SaveChanges:=False
    ProcessFile = lngCurrentOutRow
End Function



' --------------------------------------------------------------------------------------------------------------------------------------------
'      Test Routines
' --------------------------------------------------------------------------------------------------------------------------------------------
Sub test_AppClass()

    Dim clsApp As clsAppMnemonicInfo
    Set clsApp = New clsAppMnemonicInfo
    
    clsApp.Name = "Test App"
    clsApp.Mnemonic = "TST"
    
    MsgBox clsApp.Mnemonic & " - " & clsApp.Name

End Sub

Sub test_GetPtBStatus()

    Dim wkbCurrent As Workbook, wkbOut As Workbook
    Dim shtData As Worksheet
    Dim rngFileList As Range, currentCell As Range
    Dim lngLastRow As Long, lngBreakCount As Long
    Dim lngBareMetalOutRow
    Dim strCurrentBuildFile As String
    Dim clsPtB As clsPtBStatus
    
    Set wkbOut = ThisWorkbook
    Set shtData = wkbOut.Worksheets("Files")
    Set clsPtB = New clsPtBStatus
    
    Set clsPtB = GetPtBStatus(wkbOut, "DOQ")

    MsgBox clsPtB.Mnemonic & " - " & clsPtB.MigrationWave & " - " & clsPtB.DDRComplete & " - " & clsPtB.LogicalDesignComplete & " - " & clsPtB.PtBComplete
    

End Sub

Sub test_GetApplicationSummaryInfo()
    Dim wkbCurrent As Workbook, wkbOut As Workbook
    Dim shtData As Worksheet
    Dim rngFileList As Range, currentCell As Range
    Dim lngLastRow As Long, lngBreakCount As Long
    Dim lngBareMetalOutRow
    Dim strCurrentBuildFile As String
    Dim clsApp As clsAppMnemonicInfo
    
    Set wkbOut = ThisWorkbook
    Set shtData = wkbOut.Worksheets("Files")
    Set clsApp = New clsAppMnemonicInfo
    
    strCurrentBuildFile = "C:\local_data\Corporate Services - Underwood\CPX\CPX Build Sheet -V6.1.xlsx"
    
    Set wkbCurrent = Workbooks.Open(Filename:=strCurrentBuildFile, UpdateLinks:=False, ReadOnly:=True)
    Set clsApp = GetApplicationSummaryInfo(wkbCurrent)

    MsgBox clsApp.Mnemonic & " - " & clsApp.Name
End Sub

' --------------------------------------------------------------------------------------------------------------------------------------------
'      Utility Routines
' --------------------------------------------------------------------------------------------------------------------------------------------
Public Function WorksheetExists(ByVal WorkbookName As String, ByVal WorksheetName As String) As Boolean
    
    Dim wkb As Workbook
    Dim sht As Worksheet
    
    Set wkb = Workbooks.Open(Filename:=WorkbookName, UpdateLinks:=False, ReadOnly:=True)

    WorksheetExists = False
        
    For Each sht In wkb.Worksheets
        If sht.Name = WorksheetName Then WorksheetExists = True
    Next sht
    
    If WorksheetExists = False Then
        wkb.Close SaveChanges:=False
    End If
    
End Function

Sub HideRawDataSheets()

    Dim CurrentSheetName As String
    Dim sht As Worksheet
    
    Application.DisplayAlerts = False
    CurrentSheetName = ActiveSheet.Name

    ' this hides the sheet but users will be able
    ' to unhide it using the Excel UI
    Worksheets("Files").Visible = xlSheetHidden
    
    ' this hides the sheet so that it can only be made visible using VBA
    'sheet.Visible = xlSheetVeryHidden


End Sub


Sub MakeAllTablesRegularRanges()
  
  Dim sht As Worksheet
  Dim lo As ListObject
  
  For Each sht In Worksheets
    For Each lo In sht.ListObjects
      lo.Unlist
    Next
  Next
End Sub

Sub UnhideAllSheetsCount()
    Dim sht As Worksheet
    Dim count As Integer
 
    count = 0
 
    For Each sht In ActiveWorkbook.Worksheets
        If sht.Visible <> xlSheetVisible Then
            sht.Visible = xlSheetVisible
            count = count + 1
        End If
    Next sht
 
    If count > 0 Then
        MsgBox count & " worksheets have been unhidden.", vbOKOnly, "Unhiding worksheets"
    Else
        MsgBox "No hidden worksheets have been found.", vbOKOnly, "Unhiding worksheets"
    End If
End Sub

Public Function ColumnLetterToNumber(ByVal strColChar As String, _
                  Optional ByVal strWshName As String = "") _
                           As Integer
    
    Const FUNCTION_NAME As String = "ColumnLetterToNumber"

    Dim intStartNumber As Integer
    
    If Len(strColChar) = 1 Then
       If strWshName <> "" Then _
          ColumnLetterToNumber = Worksheets(strWshName).Range(strColChar & "1").Column
       If strWshName = "" Then ColumnLetterToNumber = Range(strColChar & "1").Column
    Else
       intStartNumber = Range((Left(strColChar, 1)) & "1").Column
       ColumnLetterToNumber = ((intStartNumber)) * 26 * (Len(strColChar) - 1) + _
                                       ColumnLetterToNumber(Right(strColChar, Len(strColChar) - 1))
    End If
    If gbDEBUG = False Then Exit Function
    

End Function

Public Function ColumnNumberToLetter(ByVal intColumnNumber As Integer) As String

    Const FUNCTION_NAME As String = "ColumnNumberToLetter"

    On Error GoTo ErrorHandler
   
    Dim strStartLetter As String

    Select Case intColumnNumber
       Case 0:         ColumnNumberToLetter = Chr(90)
       Case Is <= 26:  ColumnNumberToLetter = Chr(intColumnNumber + 64)
       Case Else
          If intColumnNumber Mod 26 = 0 Then
             ColumnNumberToLetter = Chr(64 + (intColumnNumber / 26) - 1) & ColumnNumberToLetter(intColumnNumber Mod 26)
             Exit Function
          End If
          
          strStartLetter = Chr(Int(intColumnNumber / 26) + 64)
          ColumnNumberToLetter = strStartLetter & ColumnNumberToLetter(intColumnNumber Mod 26)
          
    End Select
    
    If gbDEBUG = False Then Exit Function
    
ErrorHandler:
    Call ErrorHandling(MS_MODULENAME & ":" & FUNCTION_NAME, 1, _
         "return the corresponding letter for the column number " & _
         "'" & intColumnNumber & "'.")
         
End Function

Public Sub ErrorHandling(ByVal strRoutineName As String, _
                         ByVal strErrorNumber As String, _
                         ByVal strErrorDescription As String)

    Dim strMessage As String
    
    strMessage = strErrorNumber & " - " & strErrorDescription
    
    Call MsgBox(strMessage, vbCritical, strRoutineName & " - Error")
    Call LogFileWriteError(strRoutineName, strMessage)
    
End Sub

Public Function LogFileWriteError(ByVal strRoutineName As String, _
                                    ByVal strMessage As String)
                                    
    Const FUNCTION_NAME As String = "LogFileWriteError"

    On Error GoTo ErrorHandler
    
    Dim strText As String
      
    If (gbOBJFSO Is Nothing) Then
        Set gbOBJFSO = New FileSystemObject
    End If
    
    If (gbSCRTEXT Is Nothing) Then
       If (gbOBJFSO.FileExists("C:\temp\" & MS_MODULENAME & ".log") = False) Then
          Set gbSCRTEXT = gbOBJFSO.OpenTextFile("C:\temp\" & MS_MODULENAME & ".log", IOMode.ForWriting, True)
       Else
          Set gbSCRTEXT = gbOBJFSO.OpenTextFile("C:\temp\" & MS_MODULENAME & ".log", IOMode.ForAppending)
       End If
    End If
    
    strText = strText & "" & vbCrLf
    strText = strText & Format(Date, "dd mmm yyyy") & "-" & Time() & vbCrLf
    strText = strText & " " & strRoutineName & vbCrLf
    strText = strText & " " & strMessage & vbCrLf
    
    gbSCRTEXT.WriteLine strText
    gbSCRTEXT.Close
    Set gbSCRTEXT = Nothing
    
    Exit Function
    
ErrorHandler:
    Set gbSCRTEXT = Nothing
    Call MsgBox("Unable to write to log file", vbCritical, "LogFileWriteError")
    
End Function


Public Sub StringToArray(ByVal strText As String, _
                       ByRef vArrayName As Variant, _
                       Optional ByVal strSeparateChar As String = ";")
                       
    Const SUBROUTINE_NAME As String = "LogFileWriteError"

    On Error GoTo ErrorHandler
    
    Dim strNextEntry As String
    Dim intArrayCount As Integer
    Dim intNumberOfCharacters As Integer
    
    On Error GoTo ErrorHandler
    
    intNumberOfCharacters = Str_CharsNoOf(strText, strSeparateChar)
    
    If intNumberOfCharacters > 0 Then
       ReDim vArrayName(intNumberOfCharacters, 1)
       intArrayCount = 1
       
       Do While Len(strText) > 0
          If InStr(1, strText, sSeperateChar) = 0 Then
             strNextEntry = strText
             strText = ""
          Else
             strNextEntry = Left(strText, InStr(1, strText, strSeparateChar) - 1)
             strText = Right(strText, Len(strText) - Len(strNextEntry) - 1)
          End If
          vArrayName(intArrayCount, 1) = strNextEntry
          intArrayCount = intArrayCount + 1
       Loop
    End If
    
    If gbDEBUG = False Then Exit Sub
    
ErrorHandler:
    Call ErrorHandling(MS_MODULENAME & ":" & SUBROUTINE_NAME, 1, _
         "split up the string" & vbCrLf & strText & vbCrLf & _
         "separated by " & strSeparateChar & " and place them in an array")

End Sub

Sub SpeedUpExcel()
    ' Resume Normal Operations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.CalculateBeforeSave = True

End Sub


Sub RestoreExcel()
    ' Resume Normal Operations
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayAlerts = True
    Application.CalculateBeforeSave = True
End Sub


