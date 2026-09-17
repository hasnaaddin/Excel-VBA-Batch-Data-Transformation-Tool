Attribute VB_Name = "POPULATE"

Option Explicit

Private Sub SaveAmendedCopy(ByVal wbTarget As Workbook)

    Dim macroPath As String
    Dim amendedPath As String
    Dim timeStamp As String
    Dim newFileName As String

    macroPath = ThisWorkbook.Path
    If macroPath = "" Then Exit Sub

    amendedPath = macroPath & "\Amended"

    '  SAFE folder check (does NOT reset Dir)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(amendedPath) Then
        fso.CreateFolder amendedPath
    End If

    timeStamp = Format(Now, "ddmm_hhnnss")
    newFileName = "_Amended_" & timeStamp & "_" & wbTarget.Name

    wbTarget.SaveCopyAs amendedPath & "\" & newFileName

End Sub


Public Sub PopulateMyColumn()

    '  SILENT / SAFE MODE
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    Dim folderPath As String
    folderPath = SelectFolder()
    If folderPath = "" Then GoTo Cleanup

    Call ProcessAllFilesInFolder(folderPath)
    
    
 'Print COMPLETION; Message
    MsgBox "The macro has finished processing all selected files." & vbCrLf & vbCrLf & _
           "Please check the 'Amended' folder for the output files.", _
           vbInformation, _
           "Process Complete"


Cleanup:
    '  RESTORE EXCEL STATE
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.StatusBar = False

End Sub



Private Sub ProcessAllFilesInFolder(ByVal folderPath As String)

    Dim files As Collection
    Set files = New Collection

    Dim fileName As String
    fileName = Dir(folderPath & "\*.xls")

    Do While fileName <> ""
        files.Add fileName
        fileName = Dir
    Loop

    Dim totalFiles As Long
    totalFiles = files.Count
    If totalFiles = 0 Then Exit Sub

    Dim i As Long
    For i = 1 To totalFiles

        If Not ProcessSingleWorkbook( _
                folderPath & "\" & files(i), _
                i, _
                totalFiles _
           ) Then
            Exit For
        End If

    Next i

End Sub



Private Function ProcessSingleWorkbook( _
    ByVal filePath As String, _
    ByVal fileIndex As Long, _
    ByVal totalFiles As Long _
) As Boolean


    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = Workbooks.Open( _
        fileName:=filePath, _
        UpdateLinks:=0 _
    )


    On Error Resume Next
    Set ws = wb.Worksheets(GetControlValue("TargetWS"))
    On Error GoTo 0

    If ws Is Nothing Then
        wb.Close False
        Exit Function
    End If
    

    Call PopulateWorksheet(ws)
    
    Call SaveAmendedCopy(wb)

    wb.Close SaveChanges:=False
    
    
Dim resp As VbMsgBoxResult

resp = MsgBox( _
    "File " & fileIndex & " of " & totalFiles & vbCrLf & vbCrLf & _
    "Processed:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
    "Do you want to continue?", _
    vbYesNo + vbQuestion, _
    "Continue Processing?" _
)

If resp = vbNo Then
    ProcessSingleWorkbook = False
    Exit Function
End If

ProcessSingleWorkbook = True


End Function



Private Sub PopulateWorksheet(ByVal targetWS As Worksheet)

    Dim ctrlWS As Worksheet
    Set ctrlWS = ThisWorkbook.Worksheets("Control")

    Dim headerRow As Long
    Dim dataStartRow As Long

    headerRow = GetControlValue("TargetHeaderRowStart")
    dataStartRow = GetControlValue("TargetDataRowStart")

    Dim convertFormulaToValue As Boolean
    convertFormulaToValue = _
        (UCase(Trim(GetControlValue("CONVERT_FORMULA_TO_VALUE"))) = "Y")

    Dim targetColNameCol As Long
    targetColNameCol = GetControlColumn("TargetColName")
    If targetColNameCol = 0 Then Exit Sub

    Dim lastCtrlRow As Long
    lastCtrlRow = ctrlWS.Cells(ctrlWS.Rows.Count, targetColNameCol).End(xlUp).Row

    Dim formulaCols As Collection
    Set formulaCols = New Collection

    Dim r As Long
    Dim targetColName As String
    Dim dataType As String

    For r = 6 To lastCtrlRow

        targetColName = Trim(ctrlWS.Cells(r, targetColNameCol).value)
        If targetColName <> "" Then

            dataType = Trim(ctrlWS.Cells(r, GetControlColumn("DataType")).value)

            Call PopulateSingleColumn( _
                targetWS, _
                targetColName, _
                ctrlWS.Cells(r, GetControlColumn("DEFAULT_VALUE")).value, _
                dataType, _
                ctrlWS.Cells(r, GetControlColumn("COPY_FROM_COL")).value, _
                ctrlWS.Cells(r, GetControlColumn("OVERWRITE_YN")).value, _
                headerRow, _
                dataStartRow _
            )

            If UCase(dataType) = "FORMULA" Then
                formulaCols.Add targetColName
            End If

        End If

    Next r

    ' Convert formula columns to values (existing logic)
    If convertFormulaToValue And formulaCols.Count > 0 Then
        ConvertFormulaColumnsToValues targetWS, formulaCols, headerRow, dataStartRow
    End If

End Sub



Public Function GetControlValue(ByVal SearchName As String) As Variant
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim col As Long

    Set ws = ThisWorkbook.Worksheets("Control")

    ' Find last used column in row 5
    lastCol = ws.Cells(5, ws.Columns.Count).End(xlToLeft).Column

    ' Loop through row 5 to find the match
    For col = 1 To lastCol
        If Trim(ws.Cells(5, col).value) = Trim(SearchName) Then
            GetControlValue = ws.Cells(6, col).value
            Exit Function
        End If
    Next col

    ' If no match found
    GetControlValue = CVErr(xlErrNA)
End Function

Private Function SelectFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Target Files"
        If .Show = -1 Then
            SelectFolder = .SelectedItems(1)
        Else
            SelectFolder = ""
        End If
    End With
End Function

Public Function GetControlColumn(ByVal headerName As String) As Long

    Dim ws As Worksheet
    Dim lastCol As Long
    Dim col As Long

    Set ws = ThisWorkbook.Worksheets("Control")

    ' Find last used column in row 5
    lastCol = ws.Cells(5, ws.Columns.Count).End(xlToLeft).Column

    ' Look for the header name in row 5
    For col = 1 To lastCol
        If Trim(ws.Cells(5, col).value) = Trim(headerName) Then
            GetControlColumn = col
            Exit Function
        End If
    Next col

    ' Not found
    GetControlColumn = 0

End Function

Private Sub PopulateSingleColumn( _
    ByVal targetWS As Worksheet, _
    ByVal targetColName As String, _
    ByVal defaultValue As Variant, _
    ByVal dataType As String, _
    ByVal copyFromColName As String, _
    ByVal overwriteYN As String, _
    ByVal headerRow As Long, _
    ByVal dataStartRow As Long _
)

    Dim baseValue As Variant


    Dim targetCol As Long
    targetCol = FindHeaderColumn(targetWS, headerRow, targetColName)
    If targetCol = 0 Then Exit Sub

    Dim copyFromCol As Long
    If Trim(copyFromColName) <> "" Then
        copyFromCol = FindHeaderColumn(targetWS, headerRow, Trim(copyFromColName))
        If copyFromCol = 0 Then Exit Sub
    End If
    


    Dim lastDataRow As Long
    lastDataRow = targetWS.Cells(targetWS.Rows.Count, 1).End(xlUp).Row
    If lastDataRow < dataStartRow Then Exit Sub

    Dim overwriteMode As String
    overwriteMode = UCase(Trim(overwriteYN))
    
    If overwriteMode = "C" Then
        baseValue = targetWS.Cells(dataStartRow, targetCol).value
        copyFromCol = 0   ' defensive: C ignores helper logic
    End If


    Dim r As Long
    For r = dataStartRow To lastDataRow

        Select Case overwriteMode

            Case "Y"   ' overwrite all
                Call ApplyPopulationRule _
                    (targetWS, r, targetCol, dataType, defaultValue, copyFromCol)

            Case "N"   ' blanks only
                If IsEmpty(targetWS.Cells(r, targetCol)) Then
                    Call ApplyPopulationRule _
                        (targetWS, r, targetCol, dataType, defaultValue, copyFromCol)
                End If

            Case "E"   ' existing only
                If Not IsEmpty(targetWS.Cells(r, targetCol)) Then
                    Call ApplyPopulationRule _
                        (targetWS, r, targetCol, dataType, defaultValue, copyFromCol)
                End If
                
            Case "C" 'For C = copy-down logic, capture the first-row value once
                If r > dataStartRow Then
                    With targetWS.Cells(r, targetCol)
                        .NumberFormat = "@"
                        .value = baseValue
                    End With
                End If



        End Select

    Next r

End Sub

Private Sub ApplyPopulationRule( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal targetCol As Long, _
    ByVal dataType As String, _
    ByVal defaultValue As Variant, _
    ByVal copyFromCol As Long _
)

    ' IF this column copies from a helper

If copyFromCol > 0 Then

    With ws.Cells(rowNum, targetCol)
        .NumberFormat = "@"
        .value = ws.Cells(rowNum, copyFromCol).value
    End With

    ' Clear temporary/helper source after copying
    ws.Cells(rowNum, copyFromCol).ClearContents

    Exit Sub

End If


'  Otherwise normal behaviour
Select Case UCase(Trim(dataType))

    Case "FORMULA"
    
        Dim formulaText As String
        Dim originalValue As String
    
        formulaText = CStr(defaultValue)
    
        ' Capture the existing target value before overwriting it
        ' Double any quotes so the value remains valid inside the Excel formula
        originalValue = Replace(CStr(ws.Cells(rowNum, targetCol).value), """", """""")
    
        ' Replace dynamic placeholders
        formulaText = Replace(formulaText, "{VALUE}", originalValue)
        formulaText = Replace(formulaText, "{ROW}", CStr(rowNum))
        formulaText = Replace(formulaText, "__WBNAME__", ThisWorkbook.Name)
    
        ' Write as a normal A1-style Excel formula
        ws.Cells(rowNum, targetCol).NumberFormat = "General"
        ws.Cells(rowNum, targetCol).Formula2 = formulaText


    Case Else
        ws.Cells(rowNum, targetCol).value = _
            ApplyDataType(defaultValue, dataType)

End Select

End Sub


Private Function FindHeaderColumn( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal headerName As String _
) As Long

    Dim c As Long
    Dim lastCol As Long
    Dim cellText As String
    Dim searchText As String

    searchText = NormalizeHeader(headerName)
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        cellText = NormalizeHeader(ws.Cells(headerRow, c).value)

        If cellText = searchText Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c

    FindHeaderColumn = 0

End Function
Private Function NormalizeHeader(ByVal txt As String) As String
    Dim s As String
    s = LCase(CStr(txt))
    s = Replace(s, Chr(160), " ")   ' non-breaking space
    s = Application.WorksheetFunction.Trim(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")   ' collapse double spaces
    Loop
    NormalizeHeader = s
End Function



Private Function ApplyDataType(value As Variant, dataType As String) As Variant
    Select Case UCase(dataType)
        Case "TEXT"
            ApplyDataType = CStr(value)
        Case "NUMBER"
            ApplyDataType = CDbl(value)
        Case "DATE"
            ApplyDataType = CDate(value)
        Case "FORMULA"
            ApplyDataType = CStr(value)
        Case Else
            ApplyDataType = value
    End Select
End Function


Private Sub ConvertFormulaColumnsToValues( _
    ByVal ws As Worksheet, _
    ByVal formulaCols As Collection, _
    ByVal headerRow As Long, _
    ByVal dataStartRow As Long _
)

    Dim colName As Variant
    Dim colNum As Long
    Dim lastDataRow As Long

    lastDataRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastDataRow < dataStartRow Then Exit Sub

    For Each colName In formulaCols

        colNum = FindHeaderColumn(ws, headerRow, CStr(colName))
        If colNum > 0 Then
            
            With ws.Range(ws.Cells(dataStartRow, colNum), ws.Cells(lastDataRow, colNum))
                .NumberFormat = "@"   '  Force TEXT format
                .value = .value      '  Now preserves leading zeros
            End With

        End If

    Next colName

End Sub





