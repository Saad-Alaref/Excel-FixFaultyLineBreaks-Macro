Attribute VB_Name = "Module1"
Sub FixFaultyLineBreaksPreservingSentenceBreaks()
    Dim ws As Worksheet
    Dim targetColumns As Variant
    Dim lastRow As Long
    Dim col As Variant
    Dim i As Long
    Dim cellContent As String
    Dim lines As Variant
    Dim reconstructedText As String
    Dim j As Long
    Dim currentLine As String
    Dim lastChar As String
    Dim sentenceTerminators As String
    
    ' === Configuration ===
    ' Set the worksheet name
    Set ws = ThisWorkbook.Sheets("Sheet1") ' <-- Change "Sheet1" to your actual sheet name
    
    ' Specify the columns to process (e.g., "A", "B", "D")
    targetColumns = Array("F", "H") ' <-- Modify this array with your target columns
    
    ' Define sentence-ending punctuation
    sentenceTerminators = ".!?"
    
    ' === End of Configuration ===
    
    ' Disable text wrapping for all cells in the main data sheet
    ws.Cells.WrapText = False
    
    ' Initialize Logging
    Dim logSheet As Worksheet
    Dim logTable As ListObject
    Dim logSheetName As String
    logSheetName = "ChangeLog" ' You can change this name if desired
    
    On Error Resume Next
    Set logSheet = ThisWorkbook.Sheets(logSheetName)
    On Error GoTo 0
    
    ' If the log sheet does not exist, create it
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        logSheet.Name = logSheetName
        ' Add headers to the log sheet
        With logSheet
            .Range("A1").Value = "Timestamp"
            .Range("B1").Value = "Sheet Name"
            .Range("C1").Value = "Column"
            .Range("D1").Value = "Row"
            .Range("E1").Value = "Original Content"
            .Range("F1").Value = "New Content"
        End With
        
        ' Convert headers to table
        Set logTable = logSheet.ListObjects.Add(xlSrcRange, logSheet.Range("A1:F1"), , xlYes)
        logTable.Name = "ChangeLogTable"
        logTable.TableStyle = "TableStyleMedium2"
        
        ' No need to add a blank row; new entries will be added within the table
    Else
        ' If the table already exists, set the logTable variable
        On Error Resume Next
        Set logTable = logSheet.ListObjects("ChangeLogTable")
        On Error GoTo 0
        If logTable Is Nothing Then
            ' If table doesn't exist, create it
            With logSheet
                .Range("A1").Value = "Timestamp"
                .Range("B1").Value = "Sheet Name"
                .Range("C1").Value = "Column"
                .Range("D1").Value = "Row"
                .Range("E1").Value = "Original Content"
                .Range("F1").Value = "New Content"
            End With
            Set logTable = logSheet.ListObjects.Add(xlSrcRange, logSheet.Range("A1:F1"), , xlYes)
            logTable.Name = "ChangeLogTable"
            logTable.TableStyle = "TableStyleMedium2"
        End If
        ' Disable text wrapping for the ChangeLog table
        logSheet.Cells.WrapText = False
    End If
    
    ' Initialize Statistics
    Dim statsSheet As Worksheet
    Dim statsTable As ListObject
    Dim statsSheetName As String
    statsSheetName = "Statistics" ' You can change this name if desired
    
    On Error Resume Next
    Set statsSheet = ThisWorkbook.Sheets(statsSheetName)
    On Error GoTo 0
    
    ' If the statistics sheet does not exist, create it
    If statsSheet Is Nothing Then
        Set statsSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        statsSheet.Name = statsSheetName
        ' Add headers to the statistics sheet
        With statsSheet
            .Range("A1").Value = "Metric"
            .Range("B1").Value = "Value"
        End With
        
        ' Define predefined metrics
        Dim metricsArray As Variant
        metricsArray = Array("Total Cells Processed", "Total Cells Modified", _
                             "F - Cells Processed", "H - Cells Processed", _
                             "F - Cells Modified", "H - Cells Modified")
        
        ' Populate the metrics in the statistics sheet
        For i = LBound(metricsArray) To UBound(metricsArray)
            statsSheet.Cells(i + 2, "A").Value = metricsArray(i)
            statsSheet.Cells(i + 2, "B").Value = 0 ' Initialize with 0
        Next i
        
        ' Convert to table
        Set statsTable = statsSheet.ListObjects.Add(xlSrcRange, statsSheet.Range("A1:B7"), , xlYes)
        statsTable.Name = "StatisticsTable"
        statsTable.TableStyle = "TableStyleLight9"
        
        ' Disable text wrapping for the Statistics table
        statsSheet.Cells.WrapText = False
    Else
        ' If the table already exists, set the statsTable variable
        On Error Resume Next
        Set statsTable = statsSheet.ListObjects("StatisticsTable")
        On Error GoTo 0
        If statsTable Is Nothing Then
            ' If table doesn't exist, create it
            With statsSheet
                .Range("A1").Value = "Metric"
                .Range("B1").Value = "Value"
                ' Define predefined metrics
                Dim metricsArray2 As Variant
                metricsArray2 = Array("Total Cells Processed", "Total Cells Modified", _
                                      "F - Cells Processed", "H - Cells Processed", _
                                      "F - Cells Modified", "H - Cells Modified")
                
                ' Populate the metrics in the statistics sheet
                For i = LBound(metricsArray2) To UBound(metricsArray2)
                    .Cells(i + 2, "A").Value = metricsArray2(i)
                    .Cells(i + 2, "B").Value = 0 ' Initialize with 0
                Next i
            End With
            Set statsTable = statsSheet.ListObjects.Add(xlSrcRange, statsSheet.Range("A1:B7"), , xlYes)
            statsTable.Name = "StatisticsTable"
            statsTable.TableStyle = "TableStyleLight9"
        End If
        ' Disable text wrapping for the Statistics table
        statsSheet.Cells.WrapText = False
    End If
    
    ' Initialize Statistics Variables
    Dim totalCellsProcessed As Long
    Dim totalCellsModified As Long
    Dim processedF As Long
    Dim processedH As Long
    Dim modifiedF As Long
    Dim modifiedH As Long
    
    totalCellsProcessed = 0
    totalCellsModified = 0
    processedF = 0
    processedH = 0
    modifiedF = 0
    modifiedH = 0
    
    ' Start processing
    Application.ScreenUpdating = False ' Improve performance and prevent screen flickering
    
    For Each col In targetColumns
        ' Find the last used row in the current column
        lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        
        For i = 2 To lastRow ' Starts from row 2; assuming row 1 has headers
            cellContent = ws.Cells(i, col).Value
            totalCellsProcessed = totalCellsProcessed + 1
            If col = "F" Then
                processedF = processedF + 1
            ElseIf col = "H" Then
                processedH = processedH + 1
            End If
            
            If InStr(cellContent, vbLf) > 0 Or InStr(cellContent, vbCr) > 0 Then
                ' Store the original content for logging
                Dim originalContent As String
                originalContent = cellContent
                
                ' Normalize line breaks to vbLf
                cellContent = Replace(cellContent, vbCrLf, vbLf)
                cellContent = Replace(cellContent, vbCr, vbLf)
                
                ' Split the cell content into lines
                lines = Split(cellContent, vbLf)
                
                reconstructedText = ""
                
                For j = LBound(lines) To UBound(lines)
                    currentLine = Trim(lines(j))
                    
                    ' If not the first line, decide whether to prepend a space or a line break
                    If j > LBound(lines) Then
                        ' Get the last character of the previous line
                        Dim prevLineLastChar As String
                        prevLineLastChar = Right(Trim(lines(j - 1)), 1)
                        
                        ' Check if the previous line ends with a sentence terminator
                        If InStr(sentenceTerminators, prevLineLastChar) > 0 Then
                            ' Preserving the line break (as it likely separates sentences)
                            reconstructedText = reconstructedText & vbLf & currentLine
                        Else
                            ' Replacing faulty line break with a space
                            reconstructedText = reconstructedText & " " & currentLine
                        End If
                    Else
                        ' First line, add as is
                        reconstructedText = currentLine
                    End If
                Next j
                
                ' Replace multiple spaces with a single space
                Do While InStr(reconstructedText, "  ") > 0
                    reconstructedText = Replace(reconstructedText, "  ", " ")
                Loop
                
                ' Trim any leading or trailing spaces
                reconstructedText = Trim(reconstructedText)
                
                ' Check if the reconstructed text is different from the original
                If reconstructedText <> originalContent Then
                    ' Update the cell with the cleaned content
                    ws.Cells(i, col).Value = reconstructedText
                    totalCellsModified = totalCellsModified + 1
                    If col = "F" Then
                        modifiedF = modifiedF + 1
                    ElseIf col = "H" Then
                        modifiedH = modifiedH + 1
                    End If
                    
                    ' Log the change by adding a new row to the ChangeLog table
                    With logTable
                        .ListRows.Add
                        With .ListRows(.ListRows.Count).Range
                            .Cells(1, 1).Value = Now ' Timestamp
                            .Cells(1, 2).Value = ws.Name ' Sheet Name
                            .Cells(1, 3).Value = col ' Column
                            .Cells(1, 4).Value = i ' Row
                            .Cells(1, 5).Value = originalContent ' Original Content
                            .Cells(1, 6).Value = reconstructedText ' New Content
                        End With
                    End With
                End If
            End If
        Next i
    Next col
    
    ' Update Statistics Table
    With statsTable
        ' Update "Total Cells Processed" (first data row)
        .ListRows(1).Range(2).Value = totalCellsProcessed
        ' Update "Total Cells Modified" (second data row)
        .ListRows(2).Range(2).Value = totalCellsModified
        ' Update "F - Cells Processed" (third data row)
        .ListRows(3).Range(2).Value = processedF
        ' Update "H - Cells Processed" (fourth data row)
        .ListRows(4).Range(2).Value = processedH
        ' Update "F - Cells Modified" (fifth data row)
        .ListRows(5).Range(2).Value = modifiedF
        ' Update "H - Cells Modified" (sixth data row)
        .ListRows(6).Range(2).Value = modifiedH
    End With
    
    Application.ScreenUpdating = True ' Re-enable screen updating
    
    ' Display Statistics in a Message Box
    Dim statsMessage As String
    statsMessage = "Processing Completed!" & vbCrLf & vbCrLf
    statsMessage = statsMessage & "Total Cells Processed: " & totalCellsProcessed & vbCrLf
    statsMessage = statsMessage & "Total Cells Modified: " & totalCellsModified & vbCrLf & vbCrLf
    statsMessage = statsMessage & "Modifications per Column:" & vbCrLf
    statsMessage = statsMessage & " - Column F: " & modifiedF & " modifications" & vbCrLf
    statsMessage = statsMessage & " - Column H: " & modifiedH & " modifications" & vbCrLf
    
    statsMessage = statsMessage & vbCrLf & _
                   "A detailed log is available in the '" & logSheetName & "' sheet." & vbCrLf & _
                   "Statistics are recorded in the '" & statsSheetName & "' sheet."
    
    MsgBox statsMessage, vbInformation, "Process Completed"
End Sub


