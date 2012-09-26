Attribute VB_Name = "JavaCallBacks"
Option Explicit

Private Sub Reset()
    Application.ScreenUpdating = True
End Sub

'
' Java Callback Function
'
Public Function GetQuickTab(rng As Range, row1 As Long, row2 As Long, col1 As Long, col2 As Long, visIn As Long) As String
    Dim vis As Long
    vis = visIn
    
    Dim tvisA As Areas
    Set tvisA = Nothing
    
    Dim visr As Range
    
    If vis > 0 Then
        On Error Resume Next
        Set visr = rng.SpecialCells(xlCellTypeVisible)
        
        If Not visr Is rng Then
            Set tvisA = visr.Areas
            If tvisA.Count = 1 Then
                'If tvisA.Address = rng.Address Then    ' this evaluates to true unexpectedly, so we force operands to string, then compare
                Dim rngAddr As String
                rngAddr = tvisA.Address
                Dim visAddr As String
                visAddr = rng.Address
                If rngAddr = visAddr Then
                    Set tvisA = Nothing
                End If
            End If
        End If
        On Error GoTo 0
    End If
    
    Dim ans As String
    ans = ""
    
    Dim r As Long
    For r = row1 To row2
        If vis = 2 Then
            If r = row1 Then
                ' Header needs to be visible regardless
                ans = ans & GetQuickRowX(rng, Nothing, r, r, col1, col2) & vbCr
            Else
                ' We can't use the metadata's visibility directly because often
                ' the metadata rows are themselves hidden, and yet must be seen by the tool
                ' Take metadata cells only if it's column header is visible
                ' This supports the notion that user might hide a column
                ' that is referenced by another column, without getting a reference error
                ' because the header will be visible to match the reference, but its
                ' metadata will be unseen.
                ans = ans & GetQuickRowX(rng, tvisA, r, row1, col1, col2) & vbCr
            End If
        Else
            ' Take cells according to their (own) visiblity
            ans = ans & GetQuickRowX(rng, tvisA, r, r, col1, col2) & vbCr
        End If
    Next
    GetQuickTab = ans
    
    'OpenLogForAppend
    'Print #gLogFile, "===="
    'Print #gLogFile, "rng = " & rng.Parent.Parent.Name & "." & rng.Parent.Name & ":" & rng.Address
    'Print #gLogFile, ans
    'Print #gLogFile, "----"
    'CloseLog
End Function

#If False Then
Public Function GetQuickRow(rng As Range, row As Long, col1 As Long, col2 As Long) As String
    GetQuickRow = GetQuickRowX(rng, Nothing, row, col1, col2)
End Function

Public Function GetQuickCell(rng As Range, row As Long, col As Long)
    GetQuickCell = rng(row, col)
End Function

    Public Function GetQuickCol(rng As Range, row1 As Long, row2 As Long, col As Long, col2 As Long) As String
        Dim ans As String
        ans = ""
        For c = col1 To col2
            ans = ans & GetQuickCell(rng, row, c)
        Next
        GetQuickRow = ans
    End Function
#End If

Private Function GetQuickRowX(rng As Range, tvisA As Areas, row As Long, visRow As Long, col1 As Long, col2 As Long) As String
    Dim ans As String
    ans = ""
    
    Dim c As Long
    For c = col1 To col2
        Dim vis As Boolean
        If tvisA Is Nothing Then
            vis = True
        Else
            vis = CellVisible(rng(visRow, c), tvisA)
        End If
        If vis Then
            ans = ans & GetQuickCellX(rng, row, c) & vbTab
        Else
            ans = ans & vbTab
        End If
    Next
    GetQuickRowX = ans
End Function

Private Function GetQuickCellX(rng As Range, row As Long, col As Long)
    Dim v As String
    v = rng(row, col)
    v = Trim(v)
    v = Replace(v, "\", "\3")
    v = Replace(v, vbTab, "\2")
    v = Replace(v, vbCr, "\1")
    GetQuickCellX = v
End Function

Private Sub Test()
    Dim wks As Worksheet
    Set wks = ThisWorkbook.Sheets("MetaModel")
    Dim rng As Range
    Set rng = wks.ListObjects(1).Range
    Dim rc As Long
    rc = rng.Rows.Count
    Dim cc As Long
    cc = rng.Columns.Count
    Dim ans As String
    ans = GetQuickTab(rng, 1, rc, 1, cc, False)
    Dim l As Long
    l = Len(ans)
End Sub

'
' Java Callback Function
'
Public Function HasInfoTable(wkb As Workbook) As Range
    Dim ans As Range
    Set ans = Nothing
    Dim wks As Worksheet
    For Each wks In wkb.Worksheets
        If wks.Visible = xlSheetVisible Then
            If wks.Cells(1, 1) = "QuickRDA Build Table" Then
                'Need to find range of build table, starting from 2,1
                Set ans = FindBuildTable(wks)
            Else
                Dim lo As ListObject
                For Each lo In wks.ListObjects
                    If lo.Name = "QGraphSpec" Then
                        Set ans = lo.Range
                        GoTo 99
                    End If
                Next
            End If
        End If
    Next
    
99: Set HasInfoTable = ans
End Function

'
' Java Callback Function
'
Public Function GetSourceUnitTable(wks As Worksheet) As Range
    Dim tabR As Range
    Dim tcnt As Long
    tcnt = 0
    Dim tbl As ListObject
    For Each tbl In wks.ListObjects
        Set tabR = tbl.Range
        If IsDeclarativeTable(tabR) Then
            tcnt = tcnt + 1
            Exit For
            'conceptMgr.SetProvenanceInfo wks.Parent.Name, wks.Name, tabR.row
            'Application.StatusBar = "Working on workbook: " & wks.Parent.Name & ", worksheet: " & wks.Name
            'BuildGraphFromRangeObject tabR, highlightR
        End If
    Next
    If tcnt = 0 Then
        Dim marker As Boolean
        Set tabR = FindDeclarativeSheet(wks, marker)
        If marker Or IsDeclarativeTable(tabR) Then
            'aa = tabR.Parent.Name
            'conceptMgr.SetProvenanceInfo wks.Parent.Name, wks.Name, tabR.row
            'Application.StatusBar = "Working on workbook: " & wks.Parent.Name & ", worksheet: " & wks.Name
            'BuildGraphFromRangeObject tabR, highlightR
        Else
            Set tabR = Nothing
        End If
    End If
    Set GetSourceUnitTable = tabR
End Function

'
' Java Callback Function
'
Public Function FindTableRangeOnSheet(wks As Worksheet) As Range
    Dim ans As Range
    Set ans = Nothing
    
    Dim cnt As Long
    cnt = 0
   
    Dim lo As ListObject
    For Each lo In wks.ListObjects
        cnt = cnt + 1
        Set ans = lo.Range
R14: Next
    
    If cnt <> 1 Then
        Set ans = Nothing
    End If
    
    Set FindTableRangeOnSheet = ans
    Exit Function

E14: ErrorProc
    Resume R14
End Function

'
' Java Callback Function
'
Public Sub SetReportColor(r As Range, clr As Long)
    With r.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = clr
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub

Public Function Test4() As Long
    Test4 = 44
End Function

Public Function SetReport(wks As Worksheet, asz As Long, rn As String, dn As String, filePath As String, wkbName As String) As Long
    Dim r As Range
    Set r = wks.Cells
    With r
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Set r = wks.Range(wks.Cells(1, 1), wks.Cells(50, 1))
    r.Font.Bold = True
    r.Font.Italic = True
    
    wks.Name = rn & " Report"
    wks.Columns(1).ColumnWidth = asz
    wks.Columns(2).ColumnWidth = 72
    
    wks.Columns(3).ColumnWidth = 11
    wks.Columns(4).ColumnWidth = 11
    wks.Columns(5).ColumnWidth = 11
    
    wks.Cells(1, 1).Value = "Report Name"
    wks.Cells(1, 2).Value = dn
    
    wks.Cells(2, 1).Value = "File Path"
    wks.Cells(2, 2).Value = filePath
    
    wks.Cells(3, 1).Value = "Diagram Name"
    wks.Cells(3, 2).Value = wkbName
    
    wks.Cells(4, 1).Value = "Date/Time"
    wks.Cells(4, 2).Value = Now()
    
    SetReport = 6
End Function

Public Sub SetDropDownSources(wkb As Workbook, bulkTab As String)
    Dim wks As Worksheet
    
    Dim btName As String
    btName = ""
    
    ' First clear all validations on sheet, except the build table, if present
    On Error Resume Next
    btName = HasInfoTable(wkb).Parent.Name
    For Each wks In wkb.Worksheets
        If wks.Name <> btName Then
            wks.Cells.Validation.Delete
        End If
    Next
    On Error GoTo 0
    
    ' Now fill in DropDowns sheet
    Set wks = FindDDSheet(wkb)
    wks.Cells.Clear

    Dim ln As Long
    ln = Len(bulkTab)
    
    Dim ix As Long
    ix = 1
    
    Dim iy As Long
    iy = 1
    
    Dim r As Long
    r = 1
    
    Dim c As Long
    c = 1
    
    Do
        iy = InStr(ix, bulkTab, vbTab)
        If iy = 0 Then
            Exit Do
        End If
        
        wks.Cells(r, c).Value = Mid(bulkTab, ix, iy - ix)
        r = r + 1
                
        If Mid(bulkTab, iy + 1, 1) = vbCr Then
            r = 1
            c = c + 1
            iy = iy + 1
        End If
        
        If iy >= ln Then
            Exit Do
        End If
        
        ix = iy + 1
    Loop
    
    
End Sub

Public Sub SetValidationTarget(hdrR As Range, wkb As Workbook, targetColNum As Long, sourceColIndex As Long, sourceLength As Long)
    Dim wks As Worksheet
    Set wks = FindDDSheet(wkb)
    
    If sourceLength > 0 Then
        Dim xr As Range
        Set xr = wks.Range(wks.Cells(2, sourceColIndex), wks.Cells(sourceLength + 1, sourceColIndex))
        
        Dim f As String
        
        If xr.Rows.Count >= 1 Then
            'Set xr = xr.Offset(1).Resize(xr.Rows.Count - 1)
            f = "='" & xr.Parent.Name & "'!" & xr.Address(True, True)
            'Bug fix for Excel 2007
            If xr.Rows.Count = 1 Then
                f = f & ":" & xr.Address(True, True)
                'Set xr = xr.Resize(xr.Rows.Count + 1)
                'f = "='" & xr.Parent.Name & "'!" & xr.Address(True, True)
            End If
        End If
        
        Dim r As Range
        Set r = hdrR.Columns(targetColNum).Rows(1 + kTableToFirstRowOffset)
        'aa = r.Address
        
        Dim q As Range
        Set q = r.Parent.Columns(hdrR.Column + targetColNum - 1)
        'bb = q.Address
        Set q = q.Rows(q.Rows.Count)
        'bb = q.Address
        
        Set r = Range(r, q)
        'aa = r.Address
        
        SetValidation r, f
    End If
End Sub


