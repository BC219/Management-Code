' A3_Datavalidation Module

Private nextStartRow As Long ' Declared at module level

' Returns a unique list of values (as array) from column 2 where column 1 matches the key
Public Function GetValidationList(tbl As ListObject, keyValue As String) As Variant
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim tblData As Variant ' Declare array variable to store table data

    ' Read the entire DataBodyRange into the array
    If Not tbl.DataBodyRange Is Nothing Then
        tblData = tbl.DataBodyRange.Value
    Else
        ' Handle case where DataBodyRange is empty
        GetValidationList = Empty
        Exit Function
    End If

    With tbl
        ' Loop through the array instead of ListRows
        For i = 1 To UBound(tblData, 1)
            ' Compare value from column 1 of the array with keyValue
            If CStr(tblData(i, 1)) = keyValue Then
                ' Add value from column 2 of the array to the dictionary
                dict(CStr(tblData(i, 2))) = 1
            End If
        Next i
    End With

    If dict.Count > 0 Then
        GetValidationList = dict.Keys
    Else
        GetValidationList = Empty
    End If
End Function

' Applies validation directly from a VBA array (no need to write to sheet)
Public Sub ApplyValidationFromArray(rng As Range, listArr As Variant)
    Dim tmpName As String
    Dim wsHelper As Worksheet ' Declare a worksheet variable for Sheet_helper

    ' Initialize nextStartRow if it's 0 (first run or after reset)
    If nextStartRow = 0 Then nextStartRow = 1

    ' Ensure Sheet_helper exists and get a reference to it
    On Error Resume Next
    Set wsHelper = Sheet_helper ' Use Sheets collection by name
    On Error GoTo 0

    If wsHelper Is Nothing Then
        MsgBox "Error: The helper sheet 'Sheet_helper' was not found. Please ensure it exists and its CodeName is 'Sheet_helper'.", vbCritical
        Exit Sub
    End If

    tmpName = "valList_" & rng.Address(False, False)

    ' Delete existing named range if it exists
    On Error Resume Next
    ThisWorkbook.Names(tmpName).Delete
    On Error GoTo 0

    ' Create a temporary range to store values for validation (on a hidden sheet)
    With wsHelper ' Use the worksheet variable
        ' Clear old content in the area where the new list will be written
        ' This ensures no old data remains if the new list is shorter
        .Cells(nextStartRow, 1).Resize(UBound(listArr) + 1, 1).ClearContents

        ' Write the list to the next available block of rows
        .Cells(nextStartRow, 1).Resize(UBound(listArr) + 1, 1).Value = Application.Transpose(listArr)

        ' Define the named range to refer to this specific block of rows
        ThisWorkbook.Names.Add Name:=tmpName, RefersTo:=.Cells(nextStartRow, 1).Resize(UBound(listArr) + 1, 1)

        ' Update nextStartRow for the next function call
        nextStartRow = nextStartRow + UBound(listArr) + 1
    End With

    With rng.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & tmpName
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
End Sub

' Public procedure to reset the module-level variable and clear Sheet_helper
Public Sub ResetValidationHelper()
    nextStartRow = 1 ' Reset the module-level variable to 1
    On Error Resume Next
    ThisWorkbook.Sheets("Sheet_helper").Cells.ClearContents ' Clear all content on Sheet_helper
    On Error GoTo 0
End Sub