Attribute VB_Name = "UtilitiesCommon"
Option Explicit

Public Const kTableColNameRow = 1
Public Const kTableTypeRow = 2
Public Const kTableFormulaRow = 3
Public Const kTableToFirstRowOffset = 3

Public Const QuickRDAOutputFolderName = "QuickRDA"
Public Const QuickRDALinkbackFolderName = "links"

'From Module BuildDeclarative
Public Const gDeclarativePink = 14408946
Public Const gOtherPink = 14474738

'From Module Dropdowns
Public Const gDropDownSheetName = ".QDropDowns"

Public gInitialized As Boolean
Public gMMWKB As Workbook
Public gMMWKS As Worksheet
Public gAppInstallPath As String
Public gQuickRDATEMPPath As String


Function InitializeCommon() As Boolean
    If gInitialized Then
        InitializeCommon = False
        Exit Function
    End If
    
    gInitialized = True
    
    Set gMMWKB = ThisWorkbook
    gAppInstallPath = gMMWKB.path
    gQuickRDATEMPPath = GetQuickRDATEMPDirectory()
    Set gMMWKS = gMMWKB.Worksheets("MetaModel")
    
    InitializeCommon = True
End Function


'File Utilities
Function GetQuickRDATEMPDirectory() As String
    Dim theFolderName As String
    theFolderName = QuickRDAOutputFolderName
    
    Dim thePath As String
    
    thePath = Environ("TEMP")
    If thePath = "" Then
        thePath = Environ("TMP")
    End If
    
    If thePath = "" Then
        thePath = gAppInstallPath
        theFolderName = "Output"
    End If
    
    Dim ans As String
    ans = CreateDirectory(thePath, theFolderName)
    
    GetQuickRDATEMPDirectory = ans
End Function

Function CreateDirectory(filePath As String, fileName As String) As String
    Dim d As String
    d = filePath & "\" & fileName
    If dir(d, vbDirectory) = "" Then
        On Error GoTo E1
        MkDir d
    End If
    CreateDirectory = d
    Exit Function
E1: CreateDirectory = ""
End Function

Public Sub StartQuickRDA4j(func As String)
    'Run batch file to generate .svg and bring up viewer (browser)
    Dim retVal As Variant
    Dim ff As String
    ff = """" & gAppInstallPath & "\StartQuickRDA4j.bat"" """ & func & """ """ & ThisWorkbook.path & """ """ & ThisWorkbook.Name & """"
    'ff = "java -jar """ & gAppInstallPath & "\QuickRDA.jar"" """ & parm & """"
    retVal = Shell(ff, vbReadOnly)
End Sub

Function IsDeclarativeTable(hdrR As Range) As Boolean
    Dim ans As Boolean
    ans = False
    
    If Not hdrR Is Nothing Then
        'ok so now using the pink color as qualifier... great;)
        If hdrR.Rows.Count > kTableToFirstRowOffset Then
            Dim r As Long
            For r = 2 To 3
                Dim c As Long
                For c = 1 To hdrR.Columns.Count
                    Dim iclr As Long
                    iclr = hdrR(r, c).Interior.Color
                    If iclr <> gDeclarativePink And iclr <> gOtherPink Then
                        GoTo 99
                    End If
                Next
            Next
            ans = True
        End If
    End If
    
99: IsDeclarativeTable = ans
End Function

Function FindDeclarativeSheet(wks As Worksheet, marker As Boolean) As Range
    Set FindDeclarativeSheet = FindRangeOnWorksheet(wks, marker, "QuickRDA")
End Function

Function FindBuildTable(wks As Worksheet) As Range
    Dim marker As Boolean
    Set FindBuildTable = FindRangeOnWorksheet(wks, marker, "QuickRDA Build Table")
End Function

Function FindRangeOnWorksheet(wks As Worksheet, marker As Boolean, match As String) As Range
    Dim kTableStartRow As Integer
    kTableStartRow = 1
    Dim kTableStartCol As Integer
    kTableStartCol = 1
    
    Dim ans As Range
    Set ans = Nothing
    
    'consider fixed form worksheet
    Dim rLast As Long
    Dim cLast As Long

    Dim tabR As Range
    Set tabR = wks.Cells(kTableStartRow, kTableStartCol)
    If tabR(1, 1) = match Then
        marker = True
        kTableStartRow = kTableStartRow + 1
        Set tabR = wks.Cells(kTableStartRow, kTableStartCol)
    Else
        marker = False
    End If
    
    'aa = tabR.Address
    
    If tabR.Cells(1, 2) = "" Then
        cLast = 1   'avoid bug that the else here doesn't work for a one column table...
    Else
        Dim colR As Range
        Set colR = tabR.End(xlToRight)
        'aa = tabR.Address
        'bb = tabR.Parent.Name
        cLast = colR.Column + colR.Columns.Count - 1
    End If
    
    If cLast > 16000 Then
        cLast = 1
        
        While tabR.Cells(1, cLast) <> ""
            cLast = cLast + 1
        Wend
    End If
    
    If cLast < 100 Then
        rLast = kTableStartRow + 3

        Dim rn As Long
        rn = FirstVisibleColInRow(wks.Rows(kTableStartRow))
        If rn = 0 Then
            GoTo X9
        End If
        
        Dim rLastVisible As Long
        rLastVisible = FindLastVisibleRowInColumn(wks.Columns(rn))
        
        Dim c As Long
        For c = kTableStartCol To cLast
            Set tabR = wks.Cells(kTableStartRow + 2, c)
            While True
                Set tabR = tabR.End(xlDown)
                'aa = tabR.Address
                Dim rLastThisCol As Long
                rLastThisCol = tabR.row + tabR.Rows.Count - 1
                If rLastThisCol >= rLastVisible Then
                    GoTo L1
                End If
                If rLastThisCol > rLast Then
                    rLast = rLastThisCol
                End If
            Wend
L1:     Next
    
        Set ans = wks.Range(wks.Cells(kTableStartRow, kTableStartCol), wks.Cells(rLast, cLast))
        'aa = tabR.Address
        'bb = tabR.Parent.Name
    End If

X9: Set FindRangeOnWorksheet = ans
End Function

Function FirstVisibleColInRow(aRow As Range) As Long
    Dim r As Range
    On Error Resume Next
    Set r = aRow.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If r Is Nothing Then
        FirstVisibleColInRow = 0
    Else
        'aa = r.Address
        Set r = r.Areas(1)
        'aa = r.Address
        FirstVisibleColInRow = r.Column
    End If
End Function

Function FindLastVisibleRowInColumn(aCol As Range) As Long
    Dim r As Range
    Set r = aCol.SpecialCells(xlCellTypeVisible)
    'aa = r.Address
    Dim a As Areas
    Set a = r.Areas
    Set r = a(a.Count)
    'aa = r.Address
    FindLastVisibleRowInColumn = r.row + r.Rows.Count - 1
End Function

Public Sub ErrorProc()
End Sub

Public Function CellVisible(rc As Range, a As Areas) As Boolean
    CellVisible = False
    On Error GoTo 99
    Dim r As Range
    For Each r In a
        'aa = r.Address
        Dim xx As Range
        Set xx = Intersect(rc, r)
        If Not xx Is Nothing Then
            CellVisible = True
            Exit Function
        End If
    Next
99:
End Function

Function FindDDSheet(wkb As Workbook) As Worksheet
    On Error GoTo L8
    
    Dim wks As Worksheet
    Set wks = wkb.Sheets(gDropDownSheetName)
    
    If False Then
        On Error GoTo 0
L8:     Resume R8
R8:     Set wks = wkb.Sheets.Add()
        
        wks.Name = gDropDownSheetName
        wks.Visible = xlSheetHidden
    End If
    
    Set FindDDSheet = wks
End Function

Sub SetValidation(r As Range, f As String)
    Dim v As Validation
    Set v = r.Validation
    On Error GoTo E1
    v.Delete
E1: Resume R1
R1: On Error GoTo 0
    If f <> "" Then
        v.Add xlValidateList, xlValidAlertInformation, , f
        v.IgnoreBlank = True ' was false
        v.ShowError = False
        
        'v.Add xlValidateList, xlValidAlertStop, , f
        'v.ShowError = True
        'v.ErrorTitle = "Bad Value for this column"
        'v.ErrorMessage = "Use the dropdown to choose valid values for this cell"
    
    End If
End Sub



