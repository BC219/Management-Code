' In module: SheetUtilities
' Force explicit declaration of all variables in this module.
Option Explicit

' =========================================================================================
' === CONTROLLER SUBROUTINE TO UNLOCK ROWS AND APPLY VALIDATION
' =========================================================================================
Sub UnlockNextInputRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim unlockRange As Range, dataValidationRange As Range
    
    
    
    Set ws = ActiveSheet
    If Not IsSheetAllowed(ws) Then
        MsgBox "This operation is not allowed on sheet '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo Cleanup_Unlock
    
    ws.Unprotect PASSWORD:=SHEET_PASSWORD
    
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    ws.Range("D" & lastRow + 1 & ":D" & lastRow + 20).Value = GetType(ws.codeName)
    
    Set unlockRange = ws.Range("E" & lastRow + 1 & ":K" & lastRow + 20)
    
    ' Step 1: Unlock and Format the Range
    unlockRange.Locked = False
    unlockRange.Interior.Color = RGB(235, 241, 222)
    
    ' Step 2: Call the dedicated subroutine to apply validation
    Call ApplyDataValidation(unlockRange)
    
Cleanup_Unlock:
    If Err.Number <> 0 Then
        MsgBox "An operation failed. Please check password or data validation setup.", vbCritical
    Else
        ws.Protect PASSWORD:=SHEET_PASSWORD, UserInterfaceOnly:=True
        Application.ScreenUpdating = True
        MsgBox "Unlocked range " & unlockRange.Address(False, False) & " and applied data validation.", vbInformation
    End If
    
End Sub


' =========================================================================================
' === PRIVATE HELPER SUBROUTINE FOR DATA VALIDATION
' =========================================================================================
Private Sub ApplyDataValidation(ByVal targetRange As Range)
    Dim ws As Worksheet, wsData As Worksheet
    Dim lastRow As Long
    Dim validationRange As Range, firstColumn As Range
    Dim allowedSheets As Variant
    Dim s As Variant
    Dim isAllowed As Boolean
    Dim tblApp As ListObject

    Set ws = targetRange.Worksheet
    Set wsData = Sheet_Data ' "sheet_data" is a CodeName already attached in the VBE
    Set tblApp = wsData.ListObjects("tbl_Application")
    
    If Not tblApp Is Nothing Then
        If Not tblApp.ListColumns(1).DataBodyRange Is Nothing Then
            ' Apply validation
        Else
            MsgBox "No data found in the first column of tbl_Application.", vbExclamation
        End If
    Else
        MsgBox "Table 'tbl_Application' not found on Sheet Data.", vbCritical
    End If
    
    Set firstColumn = tblApp.ListColumns(1).DataBodyRange
    
    Set validationRange = Intersect(targetRange, ws.Columns("E"))
    
    If Not validationRange Is Nothing Then
        With validationRange.Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="=" & GetRangeAddressAbsolute(firstColumn)
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
        Debug.Print "Validation on Column E with source tbl_Application(1)"
    End If

End Sub

Function GetRangeAddressAbsolute(ByVal rng As Range) As String
    GetRangeAddressAbsolute = "'" & rng.Parent.Name & "'!" & rng.Address
End Function