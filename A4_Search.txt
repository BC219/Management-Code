Option Explicit

Sub OptimizedFilterDataFromMultipleSheets()
    Dim wsMain As Worksheet
    Dim combinedData As Variant
    Dim filteredData As Variant
    Dim filterCriteria As Variant
    Dim headers As Variant
    Dim outputRange As Range
    Dim startTime As Double
    Dim showLatestVersionOnly As Boolean ' NEW variable
    
    ' Start measuring execution time
    startTime = Timer
    
    ' Turn off screen updating for speed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set wsMain = Sheet_Search
    
    ' --- NEW: Get CheckBox value (corrected for Form Control CheckBox) ---
    ' On Error Resume Next ' Keep this if you want to gracefully handle missing checkbox
    showLatestVersionOnly = (wsMain.CheckBoxes("chkLatestVersionOnly").Value = xlOn)
    ' On Error GoTo CleanUp ' Re-enable if needed after debugging
    
    Debug.Print "DEBUG: CheckBox 'chkLatestVersionOnly' Value: " & showLatestVersionOnly
    ' --- END NEW ---

    ' Collect data from sheets into an array
    combinedData = CombineDataToArray()
    
    If IsEmpty(combinedData) Then
        MsgBox "No data found to filter!", vbExclamation
        GoTo CleanUp
    End If
    
    ' Get filter criteria from range B7:E7
    filterCriteria = wsMain.Range("B7:E7").Value2
    headers = wsMain.Range("B6:E6").Value2
    
    ' Perform data filtering in the array
    filteredData = FilterArrayData(combinedData, filterCriteria, headers)
    
    ' --- NEW: Apply Latest Version filter if checked ---
    If showLatestVersionOnly Then
        ' Determine column indices dynamically based on the last column of the filtered data
        ' Assuming Full Code is the last column and Version is the second to last column
        Dim dynamicFullCodeColIndex As Long
        Dim dynamicVersionColIndex As Long
        
        ' UBound(filteredData, 2) gives the upper bound of the second dimension (columns)
        dynamicFullCodeColIndex = UBound(filteredData, 2) ' Last column in filteredData array
        dynamicVersionColIndex = UBound(filteredData, 2) - 1 ' Second to last column in filteredData array
        
        filteredData = GetLatestVersions(filteredData, dynamicFullCodeColIndex, dynamicVersionColIndex)
    End If
    ' --- END NEW ---

    ' Clear old results and output new results
    Set outputRange = wsMain.Range("B10")
    ClearPreviousResults wsMain, outputRange
    
    If Not IsEmpty(filteredData) Then
        outputRange.Resize(UBound(filteredData, 1), UBound(filteredData, 2)).Value2 = filteredData
        FormatResults wsMain, outputRange, UBound(filteredData, 1), UBound(filteredData, 2)
    End If
    
CleanUp:
    ' Restore Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Data filtering completed in " & Format(Timer - startTime, "0.00") & " seconds!" & vbCrLf & _
       "Found " & IIf(IsEmpty(filteredData), 0, UBound(filteredData, 1) - 1) & " matching records.", vbInformation
End Sub

Function CombineDataToArray() As Variant
    Dim sheetNames As Variant
    Dim ws As Worksheet
    Dim tempData As Variant
    Dim combinedArray() As Variant
    Dim totalRows As Long
    Dim totalCols As Long
    Dim currentRow As Long
    Dim i As Long, j As Long, k As Long
    Dim sheetName As Variant
    Dim isFirstSheet As Boolean
    
    ' Configure list of sheets to combine
    sheetNames = Split(ALLOWED_SHEET_CODENAMES, ",")
    
    ' Calculate total number of rows and columns needed
    totalRows = 1 ' Header row
    totalCols = 0
    isFirstSheet = True
    
    For Each sheetName In sheetNames
        On Error Resume Next
        Set ws = GetSheetByCodeName(CStr(sheetName))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            With ws
                If .Cells(3, 4).Value <> "" Then
                    Dim lastDataRow As Long, lastHeaderCol As Long
                    ' Find the last row of data in column 4 (assuming data starts from col 4)
                    lastDataRow = .Cells(.Rows.Count, 4).End(xlUp).Row
                    ' Find the last column of the header (row 3)
                    lastHeaderCol = .Cells(3, .Columns.Count).End(xlToLeft).Column ' Using the user's confirmed working line
                    
                    If lastDataRow >= 5 Then ' Ensure there is actual data (starting from row 5)
                        If isFirstSheet Then
                            ' Calculate totalCols based on desired columns (e.g., lastHeaderCol - 3 if A-C are excluded)
                            totalCols = lastHeaderCol - 3 ' Assuming columns A, B, C are not needed
                            If totalCols < 1 Then totalCols = 1 ' Ensure totalCols is at least 1
                            totalRows = totalRows + (lastDataRow - 4) ' Add data rows (from row 5)
                            isFirstSheet = False
                        Else
                            totalRows = totalRows + (lastDataRow - 4) ' Add data rows (from row 5), excluding header
                        End If
                    End If
                End If
            End With
        End If
        Set ws = Nothing
    Next sheetName
    
    If totalRows <= 1 Then
        CombineDataToArray = Empty
        Exit Function
    End If
    
    ' Initialize result array
    ReDim combinedArray(1 To totalRows, 1 To totalCols)
    currentRow = 1
    isFirstSheet = True
    
    ' Combine data from sheets
    For Each sheetName In sheetNames
        On Error Resume Next
        Set ws = GetSheetByCodeName(CStr(sheetName))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            With ws
                If .Cells(3, 4).Value <> "" Then
                    Dim currentDataLastRow As Long, currentDataLastCol As Long
                    Dim headerData As Variant
                    Dim actualData As Variant
                    
                    ' Find the last row of data in column 4
                    currentDataLastRow = .Cells(.Rows.Count, 4).End(xlUp).Row
                    ' Find the last column of the header (row 3)
                    currentDataLastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column
                    
                    If currentDataLastRow >= 5 Then ' Ensure there is actual data (starting from row 5)
                        ' Get header row (row 3)
                        headerData = .Range(.Cells(3, 4), .Cells(3, currentDataLastCol)).Value2
                        ' Get actual data (from row 5 to currentDataLastRow)
                        actualData = .Range(.Cells(5, 4), .Cells(currentDataLastRow, currentDataLastCol)).Value2
                        
                        If isFirstSheet Then
                            ' Copy header for the first sheet
                            For j = 1 To UBound(headerData, 2)
                                combinedArray(currentRow, j) = headerData(1, j)
                            Next j
                            currentRow = currentRow + 1
                            
                            ' Copy actual data for the first sheet
                            For i = 1 To UBound(actualData, 1)
                                For j = 1 To UBound(actualData, 2)
                                    combinedArray(currentRow, j) = actualData(i, j)
                                Next j
                                currentRow = currentRow + 1
                            Next i
                            isFirstSheet = False
                        Else
                            ' Copy actual data for subsequent sheets (no header)
                            For i = 1 To UBound(actualData, 1)
                                For j = 1 To UBound(actualData, 2)
                                    combinedArray(currentRow, j) = actualData(i, j)
                                Next j
                                currentRow = currentRow + 1
                            Next i
                        End If
                    End If
                End If
            End With
        End If
        Set ws = Nothing
    Next sheetName
    combinedArray(1, 5) = "Content"
    CombineDataToArray = combinedArray
End Function

Function FilterArrayData(sourceData As Variant, criteria As Variant, headers As Variant) As Variant
    '--- VERSION 2.0 - Filters by matching header names ---
    Dim filteredArray() As Variant
    Dim sourceRows As Long, sourceCols As Long
    Dim filteredRows As Long
    Dim i As Long, j As Long, k As Long
    Dim matchRow As Boolean
    Dim criteriaCount As Long
    Dim columnIndexMap() As Long ' Array to map criteria columns to data columns
    
    sourceRows = UBound(sourceData, 1)
    sourceCols = UBound(sourceData, 2)
    
    ' Count non-empty criteria
    criteriaCount = 0
    For j = 1 To UBound(criteria, 2)
        If Len(Trim(CStr(criteria(1, j)))) > 0 Then
            criteriaCount = criteriaCount + 1
        End If
    Next j
    
    ' If no criteria, return all data
    If criteriaCount = 0 Then
        FilterArrayData = sourceData
        Exit Function
    End If
    
    ' --- NEW: Create the Column Index Map ---
    ReDim columnIndexMap(1 To UBound(headers, 2))
    Dim critHeader As String
    Dim dataHeader As String
    Dim isFound As Boolean
    
    For j = 1 To UBound(headers, 2) ' Loop through filter headers (e.g., from B6:E6)
        critHeader = Trim(CStr(headers(1, j)))
        isFound = False
        
        If Len(critHeader) > 0 Then
            ' Find this header in the source data's header row
            For k = 1 To sourceCols
                dataHeader = Trim(CStr(sourceData(1, k)))
                If UCase(critHeader) = UCase(dataHeader) Then
                    columnIndexMap(j) = k ' Map filter column 'j' to data column 'k'
                    isFound = True
                    Exit For ' Found match, move to next header
                End If
            Next k
        End If
        ' If a header in the filter area is not found in the data, its map index will be 0
    Next j
    
    ' --- First Pass: Count matching rows ---
    filteredRows = 1 ' Header row
    For i = 2 To sourceRows
        matchRow = True
        For j = 1 To UBound(criteria, 2)
            Dim targetDataCol As Long
            targetDataCol = columnIndexMap(j)
            
            ' Check if criterion exists and its header was found in the data
            If Len(Trim(CStr(criteria(1, j)))) > 0 And targetDataCol > 0 Then
                Dim criteriaValue As String
                Dim cellValue As String
                
                criteriaValue = Trim(CStr(criteria(1, j)))
                cellValue = Trim(CStr(sourceData(i, targetDataCol))) ' Use the mapped column index
                
                If Not IsMatch(cellValue, criteriaValue) Then
                    matchRow = False
                    Exit For
                End If
            End If
        Next j
        
        If matchRow Then filteredRows = filteredRows + 1
    Next i
    
    ' --- Second Pass: Build the filtered array ---
    If filteredRows = 1 Then
        ' Only header, no data matched
        ReDim filteredArray(1 To 1, 1 To sourceCols)
        For j = 1 To sourceCols
            filteredArray(1, j) = sourceData(1, j)
        Next j
    Else
        ReDim filteredArray(1 To filteredRows, 1 To sourceCols)
        
        ' Copy header
        For j = 1 To sourceCols
            filteredArray(1, j) = sourceData(1, j)
        Next j
        
        ' Copy matching data
        k = 2
        For i = 2 To sourceRows
            matchRow = True
            For j = 1 To UBound(criteria, 2)
                Dim checkTargetCol As Long
                checkTargetCol = columnIndexMap(j)
                
                If Len(Trim(CStr(criteria(1, j)))) > 0 And checkTargetCol > 0 Then
                    Dim checkCriteria As String
                    Dim checkCell As String
                    checkCriteria = Trim(CStr(criteria(1, j)))
                    checkCell = Trim(CStr(sourceData(i, checkTargetCol))) ' Use the mapped column index
                    
                    If Not IsMatch(checkCell, checkCriteria) Then
                        matchRow = False
                        Exit For
                    End If
                End If
            Next j
            
            If matchRow Then
                For j = 1 To sourceCols
                    filteredArray(k, j) = sourceData(i, j)
                Next j
                k = k + 1
            End If
        Next i
    End If
    
    FilterArrayData = filteredArray
End Function

Function IsMatch(cellValue As String, criteriaValue As String) As Boolean
    ' Supports wildcard and exact match
    If InStr(criteriaValue, "*") > 0 Or InStr(criteriaValue, "?") > 0 Then
        IsMatch = (cellValue Like criteriaValue)
    Else
        IsMatch = (UCase(cellValue) = UCase(criteriaValue))
    End If
End Function

Function GetAllSheetsExcept(excludeSheet As String) As Variant
    Dim sheetList() As String
    Dim ws As Worksheet
    Dim counter As Long
    
    counter = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> excludeSheet Then
            ReDim Preserve sheetList(counter)
            sheetList(counter) = ws.Name
            counter = counter + 1
        End If
    Next ws
    
    GetAllSheetsExcept = sheetList
End Function

Sub FormatResults(ws As Worksheet, outputRange As Range, numRows As Long, numCols As Long)
    Dim resultRange As Range
    Set resultRange = outputRange.Resize(numRows, numCols)
    
    With resultRange
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Format header
    With outputRange.Resize(1, numCols)
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Auto-fit columns
    ' resultRange.Columns.AutoFit
    
    ' Format column 8 as date
    With resultRange.Columns(8)
        .NumberFormat = "dd/mm/yyyy" ' Or other date format as required
    End With
End Sub

Sub ClearPreviousResults(ws As Worksheet, outputRange As Range)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim clearRange As Range
    
    lastCol = outputRange.End(xlToRight).Column
    With ws
        lastRow = .Cells(.Rows.Count, outputRange.Column).End(xlUp).Row
        
        If lastRow >= outputRange.Row And lastCol >= outputRange.Column Then
            Set clearRange = .Range(outputRange, .Cells(lastRow, lastCol))
            clearRange.Clear
        End If
    End With
End Sub


' NEW FUNCTION: Filters an array to keep only the latest version for each base code
Function GetLatestVersions(ByVal sourceData As Variant, ByVal fullCodeColIndex As Long, ByVal versionColIndex As Long) As Variant
    Dim dictLatest As Object
    Set dictLatest = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, j As Long
    Dim baseCode As String
    Dim currentVersion As Long
    Dim storedVersion As Long
    Dim currentDataRow() As Variant ' Declare as dynamic array
    Dim headerRow As Variant
    
    ' Handle empty sourceData or only header row
    If IsEmpty(sourceData) Then
        GetLatestVersions = Empty ' Return Empty if sourceData is Empty
        Exit Function
    End If

    If UBound(sourceData, 1) < 2 Then ' Only header or no data rows
        GetLatestVersions = sourceData ' Return as is if no data rows
        Exit Function
    End If
    
    ' Store header row
    ReDim headerRow(1 To UBound(sourceData, 2))
    For j = 1 To UBound(sourceData, 2)
        headerRow(j) = sourceData(1, j)
    Next j

    ' Loop through data rows (skip header)
    For i = 2 To UBound(sourceData, 1)
        ' --- NEW: Error handling for data extraction ---
        ' Check if the values in the code or version columns are errors
        If IsError(sourceData(i, fullCodeColIndex)) Or IsError(sourceData(i, versionColIndex)) Then
            Debug.Print "Skipping row " & i & " due to error value in code or version column."
            GoTo NextRow ' Skip this row if data is erroneous
        End If

        ' Extract base code and version from the current row
        baseCode = ExtractBaseCode(CStr(sourceData(i, fullCodeColIndex)))
        
        ' Ensure version is a valid number
        On Error Resume Next ' Temporarily disable error handling for CLng conversion
        currentVersion = CLng(sourceData(i, versionColIndex))
        If Err.Number <> 0 Then ' If an error occurred during CLng conversion
            Debug.Print "Skipping row " & i & " due to invalid version number: " & sourceData(i, versionColIndex) & " (Error: " & Err.Description & ")"
            Err.Clear ' Clear the error
            On Error GoTo 0 ' Re-enable error handling
            GoTo NextRow ' Skip this row if version is not a number
        End If
        On Error GoTo 0 ' Re-enable error handling
        ' --- END NEW ---
        
        ' Store the entire row for comparison
        ' ReDim currentDataRow for each row to ensure correct sizing
        ReDim currentDataRow(1 To UBound(sourceData, 2))
        For j = 1 To UBound(sourceData, 2)
            currentDataRow(j) = sourceData(i, j)
        Next j
        
        ' --- DEBUG PRINT: Log current row data ---
        Debug.Print "Processing Row " & i & ":"
        Debug.Print "  Full Code (Col " & fullCodeColIndex & "): " & sourceData(i, fullCodeColIndex)
        Debug.Print "  Base Code: " & baseCode
        Debug.Print "  Current Version (Col " & versionColIndex & "): " & currentVersion
        
        If dictLatest.Exists(baseCode) Then
            ' Get stored version for this base code
            storedVersion = CLng(dictLatest(baseCode)(versionColIndex))
            
            ' --- DEBUG PRINT: Log comparison ---
            Debug.Print "  Base Code '" & baseCode & "' already exists. Stored Version: " & storedVersion
            
            ' If current version is higher, update dictionary
            If currentVersion > storedVersion Then
                dictLatest(baseCode) = currentDataRow
                Debug.Print "  UPDATED: New version " & currentVersion & " is higher than " & storedVersion & ". Storing new row."
            Else
                Debug.Print "  SKIPPED: Current version " & currentVersion & " is not higher than " & storedVersion & "."
            End If
        Else
            ' Add new base code and its row
            dictLatest.Add baseCode, currentDataRow
            Debug.Print "  ADDED: New Base Code '" & baseCode & "' with Version " & currentVersion & "."
        End If
NextRow: ' Label for GoTo
    Next i
    
    ' Convert dictionary values (rows) back to a 2D array
    Dim resultArr() As Variant
    Dim dictKeys As Variant
    Dim k As Long
    
    If dictLatest.Count > 0 Then
        ReDim resultArr(1 To dictLatest.Count + 1, 1 To UBound(sourceData, 2))
        
        ' Add header row to result array
        For j = 1 To UBound(sourceData, 2)
            resultArr(1, j) = headerRow(j)
        Next j
        
        k = 2 ' Start filling from the second row (after header)
        For Each dictKeys In dictLatest.Keys
            Dim rowData As Variant
            rowData = dictLatest(dictKeys)
            For j = 1 To UBound(rowData)
                resultArr(k, j) = rowData(j)
            Next j
            k = k + 1
        Next dictKeys
        GetLatestVersions = resultArr
    Else
        ' If no latest versions found (e.g., no matching base codes), return only header
        ReDim resultArr(1 To 1, 1 To UBound(sourceData, 2))
        For j = 1 To UBound(sourceData, 2)
            resultArr(1, j) = headerRow(j)
        Next j
        GetLatestVersions = resultArr
    End If
End Function

' NEW FUNCTION: Extracts the base code by removing the last version number
Function ExtractBaseCode(ByVal fullCode As String) As String
    ' Logic: Remove the last 3 characters from the full code string
    ' Example: "WEA-QCP-21-001-01" -> "WEA-QCP-21-001"
    If Len(fullCode) >= 3 Then
        ExtractBaseCode = Left(fullCode, Len(fullCode) - 3)
    Else
        ExtractBaseCode = fullCode ' If string is too short, return original
    End If
End Function



Sub SetupOptimizedFilterCriteria()
    Dim ws As Worksheet
    Set ws = Sheet_Search
    
    ' Set header for filter condition
    ws.Range("B6").Value = "Item"
    ws.Range("C6").Value = "Application"
    ws.Range("D6").Value = "Year"
    ws.Range("E6").Value = "Model"
    
    ' Format
    With ws.Range("B6:E6")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 240)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With ws.Range("B7:E7")
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(255, 255, 240)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Create button
    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Range("G6").Left, ws.Range("G6").Top, 120, 35)
    btn.Caption = "Search Management code"
    btn.OnAction = "OptimizedFilterDataFromMultipleSheets"
    
    ' Add manual
    ws.Range("B3").Value = "Enter filter conditions (supports * and ?):"
    ws.Range("B4").Value = "Example: Model - ""Fresh*"" instead ""Fresh6 L"" & ""Fresh 8 L"": * for any sequence of characters, ? for any single character"
    ws.Range("B3").Font.Bold = True
    ws.Range("B3:B4").Font.ColorIndex = 5
    ws.Range("B3:B4").Font.Italic = True

    ' --- NEW: Add CheckBox for Latest Version Only ---
    Dim chkBox As CheckBox
    On Error Resume Next ' In case checkbox already exists
    Set chkBox = ws.CheckBoxes("chkLatestVersionOnly")
    On Error GoTo 0

    If chkBox Is Nothing Then
        Set chkBox = ws.CheckBoxes.Add(ws.Range("B8").Left, ws.Range("B8").Top, 150, 20)
        chkBox.Name = "chkLatestVersionOnly"
    End If
    chkBox.Caption = "Only show latest version"
    chkBox.Value = xlOff ' Default to unchecked
    ' --- END NEW --
    
    MsgBox "Optimized version set! Support wildcards (* and ?) in filter conditions.", vbInformation
End Sub
