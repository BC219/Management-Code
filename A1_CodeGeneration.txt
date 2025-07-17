' Force explicit declaration of all variables in this module.
Option Explicit

' =========================================================================================
' === MAIN SUBROUTINE: GENERATE CODES AND LOCK DOWN THE ENTIRE SHEET
' =========================================================================================
Sub GenerateManagementCodes_WithProtection()
    ' --- Variable Declaration ---
    Dim ws As Worksheet
    Dim data As Variant, results() As Variant
    Dim colC_Data As Variant
    Dim dictUnique As Object, dictMaxDocNo As Object, dictVersion As Object
    Dim lastRow As Long
    Dim isDOE As Boolean
    Dim dataRange As Range, alignRange As Range
    Dim i As Long, isSuccess As Boolean

    ' --- Get the currently active sheet
    Set ws = ThisWorkbook.ActiveSheet
    ' ===============================================================================
    ' === CHECK EXECUTION PERMISSION ON THE CURRENT SHEET (NEW)
    ' ===============================================================================
    If Not IsSheetAllowed(ws) Then
        MsgBox "This script is not permitted to run on sheet '" & ws.Name & "'.", vbExclamation, "Operation Denied"
        Exit Sub ' Stop the script immediately
    End If
    ' ===============================================================================

    ' --- Initial Setup ---
    isSuccess = False
    Application.ScreenUpdating = False

    ' --- 1. UNPROTECT SHEET ---
    On Error GoTo CleanUp
    ws.Unprotect PASSWORD:=SHEET_PASSWORD

    ' --- The rest of your code remains unchanged ---
    ' --- For example:
    
    ' --- 2. FIND LAST ROW ---
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    If lastRow < 5 Then GoTo CleanUp
    Set dataRange = ws.Range(FORMAT_COL_FIRST & "5:" & FORMAT_COL_LAST & lastRow)

    ' --- 3. ADD SEQUENTIAL NUMBERS TO COLUMN C ---
    ReDim colC_Data(1 To lastRow - 4, 1 To 1)
    For i = 1 To lastRow - 4
        colC_Data(i, 1) = i
    Next i
    ws.Range("C5:C" & lastRow).Value = colC_Data

    ' --- 4. LOAD AND PROCESS DATA ---
    isDOE = (ws.codeName = DOE_SHEET_CODENAME)
    data = ws.Range(START_CELL).Resize(lastRow - 4, 11).Value
    If Not ValidateData(data) Then GoTo CleanUp
    
    ReDim results(1 To UBound(data), 1 To 3)
    Set dictUnique = CreateObject(DICT_OBJECT)
    Set dictMaxDocNo = CreateObject(DICT_OBJECT)
    Set dictVersion = CreateObject(DICT_OBJECT)
    
    ' ======================= LOGIC FIX: RECALCULATE ALL ROWS =======================
    ' Call the corrected subroutine that processes the entire dataset every time.
    Call ProcessData_RecalculateAll(data, results, dictUnique, dictMaxDocNo, dictVersion, isDOE)
    ' ===============================================================================
    
    ' --- 5. WRITE RESULTS AND APPLY FORMATTING ---
    ws.Range(RESULT_COL & "5").Resize(UBound(results), RESULT_COL_WIDTH).Value = results
    ws.Range(FORMAT_COL_FIRST & "5:" & FORMAT_COL_LAST & lastRow + EXTRA_ROW_BUFFER).Interior.ColorIndex = xlNone
    dataRange.Borders.LineStyle = xlContinuous
    Set alignRange = ws.Range(FORMAT_COL_FIRST & "5:" & FORMAT_COL_LAST & lastRow)
    With alignRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    '--------------
        ' --- NEW: CLEAR ALL UNUSED ROWS (columns C to M)
    ' This clears the content and formatting of all rows below the last used row.
    If lastRow < ws.Rows.Count Then
        ws.Range(FORMAT_COL_FIRST & lastRow + 1 & ":" & CLEAR_COL_LAST & ws.Rows.Count).Clear
    End If
    '--------------
    
    ' --- Set success flag to True before finishing ---
    isSuccess = True

' --- 6. CLEANUP, RE-PROTECT, AND SHOW FINAL MESSAGE ---
CleanUp:
    ' === Only re-protect sheet if code generation was successful ===
    Application.ScreenUpdating = True
    If isSuccess Then
        ws.Cells.Locked = True
        ws.Protect PASSWORD:=SHEET_PASSWORD
        MsgBox "Code generation complete." & vbCrLf & _
               "Formatting (numbers, colors, borders, alignment) applied successfully.", vbInformation
    ElseIf Err.Number <> 0 Then
        MsgBox "An error occurred: " & Err.Description, vbCritical
    End If

End Sub


' =========================================================================================
' === HELPER FUNCTIONS
' =========================================================================================

Private Function ValidateData(ByVal data As Variant) As Boolean
    ' ... This function remains unchanged ...
    Dim i As Long
    ValidateData = True
    For i = 1 To UBound(data)
        If Trim(data(i, 2)) = "" Or Trim(data(i, 3)) = "" Or Trim(data(i, 4)) = "" Or Trim(data(i, 5)) = "" Then
            MsgBox "Error: Row " & (i + 4) & " has missing data in columns E, F, G, or H!", vbCritical
            ValidateData = False
            Exit Function
        End If
    Next i
End Function

' ======================= LOGIC FIX: Renamed and rewritten Sub ========================
Private Sub ProcessData_RecalculateAll(ByVal data As Variant, ByRef results() As Variant, _
                                       ByVal dictUnique As Object, ByVal dictMaxDocNo As Object, ByVal dictVersion As Object, _
                                       ByVal isDOE As Boolean)

    Dim i As Long
    Dim docKey As String, keyEFGH As String
    Dim maxDocNo As Long, versionNo As Long
    Dim tblApp As ListObject
    Dim lookupRange As Range
    Dim lookupResult As Variant
    Dim ws As Worksheet, wsData As Worksheet
    Dim yearShort As String

    ' Define active worksheet and lookup range (O5:P-lastrow)
    Set ws = ActiveSheet
    Set wsData = Sheet_Data
    Set tblApp = wsData.ListObjects("tbl_Application")
    'Set lookupRange = ws.Range("P5:Q" & ws.Cells(ws.Rows.Count, "P").End(xlUp).Row)
    ' Loop through each row of data
    For i = 1 To UBound(data)
        Debug.Print "Processing row: " & (i + 4)
        Dim dVal As String, eVal As String, fVal As String, gVal As String, hVal As String
        dVal = data(i, 1)
        eVal = data(i, 2)
        fVal = data(i, 3)
        gVal = data(i, 4)
        hVal = data(i, 5)

        keyEFGH = eVal & "|" & fVal & "|" & gVal & "|" & hVal
        
        ' --- Lookup with error handling ---
        lookupResult = Application.Vlookup(CStr(eVal), tblApp.DataBodyRange, 2, False)

        ' If lookup fails, set a default value
        If IsError(lookupResult) Then
            MsgBox "Application lookup failed at row " & (i + 4) & ": '" & eVal & "' not found in table tbl_Application on Sheet Data.", vbCritical, "Lookup Error"
            Err.Raise vbObjectError + 513, "ProcessData_RecalculateAll", "Lookup value '" & eVal & "' not found. Process halted."
        End If

        ' Determine document key based on isDOE logic
        If isDOE Then
            docKey = eVal & "|" & fVal & "|" & hVal
        Else
            docKey = eVal & "|" & fVal
        End If

        ' Handle column K (Document Number)
        If dictUnique.Exists(keyEFGH) Then
            results(i, 1) = dictUnique(keyEFGH)
        Else
            maxDocNo = dictMaxDocNo(docKey) + 1
            dictMaxDocNo(docKey) = maxDocNo
            results(i, 1) = maxDocNo
            dictUnique.Add keyEFGH, maxDocNo
        End If

        ' Handle column L (Version Number)
        versionNo = dictVersion(keyEFGH) + 1
        dictVersion(keyEFGH) = versionNo
        results(i, 2) = versionNo

        ' Generate column M using the new format
        yearShort = Right(fVal, 2)
        If isDOE Then
            results(i, 3) = lookupResult & "-" & dVal & "-" & hVal & "-" & yearShort & "-" & Format(results(i, 1), "000") & "-" & Format(results(i, 2), "00")
        Else
            results(i, 3) = lookupResult & "-" & dVal & "-" & yearShort & "-" & Format(results(i, 1), "000") & "-" & Format(results(i, 2), "00")
        End If
    Next i
End Sub

' =========================================================================================
' === HELPER FUNCTION: CHECK IF THE SHEET IS ALLOWED TO RUN THE SCRIPT
' =========================================================================================

' In your main code module
Public Function IsSheetAllowed(ByVal ws As Worksheet) As Boolean
    ' This function checks if the provided worksheet's CodeName
    ' exists in the global list of allowed CodeNames defined in GlobalSettings.

    ' --- Create an array by splitting the global constant string
    Dim allowedNamesArray As Variant
    allowedNamesArray = Split(ALLOWED_SHEET_CODENAMES, ",")

    ' --- Loop control variable
    Dim allowedName As Variant

    ' --- Default to not allowed
    IsSheetAllowed = False

    ' --- Loop through the array of allowed names
    For Each allowedName In allowedNamesArray
        ' Compare the sheet's CodeName against the current item from the array
        ' Trim() is used to handle any accidental whitespace.
        If StrComp(ws.codeName, Trim(CStr(allowedName)), vbTextCompare) = 0 Then
            IsSheetAllowed = True      ' Set result to True if a match is found
            Exit Function              ' Exit immediately to improve performance
        End If
    Next allowedName
End Function