Sub UnlockAllSheetsDebug()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        On Error Resume Next
        ws.Unprotect PASSWORD:=SHEET_PASSWORD
        On Error GoTo 0
    Next ws
End Sub


' This Sub forcefully re-enables events for the entire Excel application.
Public Sub ForceEnableEvents()
    Application.EnableEvents = True
    MsgBox "Application.EnableEvents has been set to TRUE."
End Sub
