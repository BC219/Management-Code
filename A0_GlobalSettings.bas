' In module: GlobalSettings
' Force explicit declaration of all variables in this module.
Option Explicit

' ========================================================================
' === GLOBAL PROJECT SETTINGS
' === Public constants are accessible from any module in this project.
' ========================================================================

' --- Sheet Structure Settings ---
Public Const START_CELL As String = "D5"
Public Const RESULT_COL As String = "L"
Public Const RESULT_COL_WIDTH As Long = 3 ' Number of result collumns  need to record
Public Const FORMAT_COL_FIRST As String = "C" ' First column of format range
Public Const FORMAT_COL_LAST As String = "N"   ' Last column of format range
Public Const CLEAR_COL_LAST As String = "N"    ' For Clear format
Public Const EXTRA_ROW_BUFFER As Long = 50


Public Const DOE_SHEET_CODENAME As String = "Sheet_DOE"

' --- Sheet Permissions ---
' A comma-separated list of sheet CodeNames where the script is allowed to run.
' IMPORTANT: Do NOT add spaces around the commas.
' Example: "Sheet1,SheetData,SheetReport"
Public Const ALLOWED_SHEET_CODENAMES As String = "Sheet_SOP,Sheet_QCP,Sheet_IQC,Sheet_QR,Sheet_CPHP,Sheet_BoxLBL,Sheet_Process,Sheet_CS,Sheet_DOE"

' --- General Settings ---
Public Const DICT_OBJECT As String = "Scripting.Dictionary"

' --- Security Settings ---
' !!! IMPORTANT: Change your password here !!!
Public Const SHEET_PASSWORD As String = "YourPasswordHere"

' === SUPPORT FUNCTIONS FOR USE IN OTHER MODULES ===
Public Function GetColumnNumber(colLetter As String) As Long
    GetColumnNumber = Range(colLetter & "1").Column
End Function

Public Function GetColumnLetter(colNum As Long) As String
    GetColumnLetter = Split(Cells(1, colNum).Address(True, False), "$")(0)
End Function

Public Function GetType(ByVal codeName As String) As String
    Select Case codeName
        Case "Sheet_SOP":     GetType = "SOP"
        Case "Sheet_QCP":     GetType = "QCP"
        Case "Sheet_IQC":     GetType = "IQC"
        Case "Sheet_QR":      GetType = "QRC"
        Case "Sheet_CPHP":    GetType = "CHC"
        Case "Sheet_BoxLBL":  GetType = "BOX"
        Case "Sheet_Process": GetType = "PRO"
        Case "Sheet_CS":      GetType = "C/S"
        Case "Sheet_DOE":     GetType = "DOE"
        Case Else:            GetType = "" ' default
    End Select
End Function


Public Function GetSheetByCodeName(codeNameToFind As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.codeName = codeNameToFind Then
            Set GetSheetByCodeName = ws
            Exit Function
        End If
    Next ws
    
End Function