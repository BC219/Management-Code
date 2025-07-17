Sub ExportVbaModulesToSheet()
    Dim ws As Worksheet
    Dim vbComp As Object
    Dim moduleCode As String
    Dim rowIndex As Long

    ' --- Create or select the sheet to export data ---
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ExportedCode")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "ExportedCode"
    End If
    On Error GoTo 0

    ' --- Clear old data ---
    ws.Cells.Clear
    rowIndex = 1

    ' --- Loop through each module ---
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 1 Or vbComp.Type = 2 Or vbComp.Type = 3 Then ' Only get Module, Class, Form
            ' Check if the module contains code
            If vbComp.CodeModule.CountOfLines > 0 Then
                moduleCode = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
                
                ' Write module name
                ws.Cells(rowIndex, 1).Value = "Module - " & vbComp.Name
                rowIndex = rowIndex + 1

                ' Write script in the next row
                ws.Cells(rowIndex, 1).Value = moduleCode
                rowIndex = rowIndex + 1
            End If
        End If
    Next vbComp

    ' --- Format the sheet ---
    ws.Columns("A").AutoFit

    ' --- Completion notification ---
    MsgBox "VBA code export complete! The code has been copied to the sheet 'ExportedCode'.", vbInformation
End Sub
