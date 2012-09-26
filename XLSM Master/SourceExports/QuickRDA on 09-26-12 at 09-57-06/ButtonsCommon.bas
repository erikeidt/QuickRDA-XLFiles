Attribute VB_Name = "ButtonsCommon"
Option Explicit

Public Sub A7_InstallAddIn()
    Dim twkb As Workbook
    Set twkb = Nothing
    
Retry:

    Dim wkb As Workbook
    Dim ans As String
    Dim a As AddIn
 
    'On Error GoTo E99
    
    Dim adin As AddIn
    Set adin = Nothing
    
    Dim nm As String
#If False Then
    If InStr(1, ThisWorkbook.Name, ".xlsm") > 0 Then
        nm = "QuickRDA.xlam"
    Else
        nm = ThisWorkbook.Name
    End If
    
    For Each a In Application.AddIns
        If a.Name = nm Then
            Set adin = a
            Exit For
        End If
    Next
#Else
    nm = ThisWorkbook.BuiltinDocumentProperties("title")
    For Each a In Application.AddIns
        If a.title = nm Then
            Set adin = a
            Exit For
        End If
    Next
#End If
    '
    'Interesting Cases:
    '   not installed and can install & enable
    '   not installed and cannot install (is .xlsm not .xlam)
    '   installed and same version: can enable/disable
    '   installed and different version:
    '       can help uninstall
    '
    
    Dim msg As String
    msg = "This button is in:" & vbNewLine & GetThisNameAndVer() & vbNewLine
        
    If adin Is Nothing Then
        msg = msg & "No QuickRDA Add-In is currently installed (as either enabled or disabled)." & vbNewLine & vbNewLine
        If InStr(1, ThisWorkbook.Name, ".xlam") > 0 Then
            msg = msg & "Would you like to install this Add-In?"
            ans = MsgBox(msg, vbOKCancel, "Install?")
            If ans = vbOK Then
                
                Set twkb = Application.Workbooks.Add
                Set adin = Application.AddIns.Add(ThisWorkbook.FullName, False)
                If Not twkb Is Nothing Then
                    twkb.Close
                    Set twkb = Nothing
                End If
                
                adin.Installed = True
            End If
        Else
            msg = msg & "The is the workbook version, not an Excel Add-In." & vbNewLine & "To install, please use the .xlam with the Install button."
            MsgBox msg, vbOKOnly, "Not Installable"
        End If
    Else
        If adin.FullName = ThisWorkbook.FullName Then
            If adin.Installed Then
                msg = msg & "This QuickRDA Add-In is installed and enabled." & vbNewLine & vbNewLine
                msg = msg & "Would you like to disable this add-in?"
                ans = MsgBox(msg, vbOKCancel, "Disable?")
                If ans = vbOK Then
                    adin.Installed = False
                'Else
                '    msg = "FYI: to uninstall, rename or remove the folder or file." & vbNewLine
                '    msg = msg & "Then use Excel's Add-In Manager to locate " & adin.Name & vbNewLine
                '    msg = msg & "It will prompt you to uninstall the add-in."
                '    MsgBox msg, , "Help"
                End If
            Else
                msg = msg & "This QuickRDA Add-In is installed but disabled." & vbNewLine & vbNewLine
                msg = msg & "Would you like to enable this Excel Add-In?"
                ans = MsgBox(msg, vbOKCancel, "Enable?")
                If ans = vbOK Then
                    adin.Installed = True
                End If
            End If
        Else
            Set wkb = Nothing
            If adin.Installed Then
                On Error Resume Next
                Set wkb = Application.Workbooks(adin.Name)
                On Error GoTo 0
            End If
            
            If Not wkb Is Nothing Then
                msg = msg & "Another QuickRDA Add-In is installed and enabled, and is:" & vbNewLine
                msg = msg & " (" & adin.FullName & ")" & vbNewLine & vbNewLine
                msg = msg & Application.Run(wkb.Name & "!GetThisNameAndVer") & vbNewLine
            Else
                msg = msg & "Another QuickRDA is installed but disabled, at:" & vbNewLine & adin.FullName & vbNewLine
                msg = msg & vbTab & "and its version is unknown." & vbNewLine & vbNewLine
            End If
            
            msg = msg & "You have two choices, either:" & vbNewLine & vbNewLine
            msg = msg & "  (1) Remove the other version, then uninstall via Excel's Add-In Manager, or " & vbNewLine & vbNewLine
            msg = msg & "  (2) Copy the contents of this folder:" & vbNewLine & "      " & ThisWorkbook.path & vbNewLine & "    into this folder:" & vbNewLine & "      " & adin.path & vbNewLine & vbNewLine
            msg = msg & "For help with (1), click Yes, otherwise click No."
            ans = MsgBox(msg, vbYesNo, "Switch Version")
            
            If ans = vbYes Then
                Set twkb = Application.Workbooks.Add
                If adin.Installed Then
                    adin.Installed = False
                    'Set wkb = Application.Workbooks(adin.Name)
                    'If Not wkb Is Nothing Then
                    '    wkb.Close
                    'End If
                End If
                
                Dim oldName As String
                oldName = adin.FullName
                
                Dim newName As String
                newName = oldName & ".old"
                
                Name oldName As newName
                
                msg = "In the upcoming dialog box, select the row for " & adin.title & " and answer Yes to delete from the list; then click Ok."
                MsgBox msg, vbOKOnly, "Instruction"
                
                If twkb Is Nothing Then
                    Set twkb = Application.Workbooks.Add
                End If
                
                Application.Dialogs(xlDialogAddinManager).Show
                
                Name newName As oldName
                
                If Not twkb Is Nothing Then
                    twkb.Close
                    Set twkb = Nothing
                End If
                
                GoTo Retry
            End If
        End If
    End If
    
    If Not twkb Is Nothing Then
        twkb.Close
        Set twkb = Nothing
    End If

Exit Sub

E99:
    Resume R99
R99:
    MsgBox "Unknown error in installation or removal.  Use Excel's Add-In Manager instead.", , "Unknown Error"
    If Not twkb Is Nothing Then
        twkb.Close
        Set twkb = Nothing
    End If
End Sub

Private Function GetThisNameAndVer() As String
    Dim msg As String
    msg = ThisWorkbook.BuiltinDocumentProperties("title")
    msg = "  " & msg & " version " & ThisWorkbook.Names("QuickRDA_Version_Number").RefersToRange.Cells(1, 1) & " from:"
    'If InStr(1, ThisWorkbook.Name, ".xlam") > 0 Then
    '    msg = msg & " (an Excel Add-In)"
    'Else
    '    msg = msg & " (an Excel workbook)"
    'End If
    msg = msg & vbNewLine & "  " & ThisWorkbook.FullName & vbNewLine
    GetThisNameAndVer = msg
End Function

'
' Invoked by Excel
'
Public Sub Auto_Add()
    MsgBox GetThisNameAndVer() & vbNewLine & vbTab & vbTab & "is now installed and enabled." & vbNewLine & vbNewLine & _
            "To record the installation, quit all running copies of Excel (!)" & vbNewLine & vbNewLine & _
            "(Note: to disable this add-in, use this Add-In's Pushpin button, or use Excel's Add-In Manager.)", , "Installed"
End Sub

'
' Invoked by Excel
'
Public Sub Auto_Remove()
    MsgBox GetThisNameAndVer() & vbNewLine & vbTab & vbTab & "is now disabled." & vbNewLine & vbNewLine & _
            "(Note: to re-enable this add-in, use Excel's Add-In Manager, or use the Pushpin button after launching the .xlam file.)", , "Uninstalled"
End Sub
