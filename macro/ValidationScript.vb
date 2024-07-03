Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim ws As Worksheet
    Dim isValid As Boolean
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    isValid = ValidateSheet(ws)
    
    If isValid Then
        Cancel = True
        MsgBox "Save operation cancelled due to validation errors. Please refer 'ValidationErrors' sheet", vbCritical
    End If
End Sub

Function ValidateSheet(ws As Worksheet) As Boolean
    Dim LastRow As Long
    Dim i As Long
    Dim validateErrors() As String
        
    'Column names
    Dim snColumn As Integer
    Dim ipColumn As Integer
    Dim hostNameColumn As Integer
    Dim rowValidate As Boolean
    Dim errorColumn As String
    Dim cell As Range
    
    snColumn = 4
    ipColumn = 6
    hostNameColumn = 7
    
    errorCount = 0
    errorColumn = ""
    
    ' Get the last row with data in column A (assuming column A always has data in each row)
    'LastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    Dim snLastRow As Long
    Dim ipLastRow As Long
    
    snLastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    ipLastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    LastRow = FindLargestRow(snLastRow, ipLastRow, 0, 0)
    
    If LastRow = 1 Then
        'MsgBox "Last Row of Column D = " & LastRow
        ReDim Preserve validateErrors(errorCount)
        validateErrors(errorCount) = "D2 cannot be empty;"
        errorCount = errorCount + 1
        ValidateSheet = True
    End If
    
    For i = 2 To LastRow
     
        ' Validate that column D is not empty
        If ws.Cells(i, snColumn).Value = "" Then
            'MsgBox "Row " & i & ": Column D cannot be empty", vbExclamation
            ReDim Preserve validateErrors(errorCount)
            validateErrors(errorCount) = "D" & i & " cannot be empty;"
            errorCount = errorCount + 1
           
            
            ' Set the cell to be highlighted
            'errorColumn = "D" & i
            'Set cell = ws.Range(errorColumn)
            'cell.Interior.color = RGB(255, 0, 0)
            'cell.Borders.color = RGB(0, 0, 0)
                
            ValidateSheet = True
            'Exit For
        End If
        
        ' Validate that column F contains a date
        If ws.Cells(i, ipColumn).Value = "" Then
            'MsgBox "Row " & i & ": Column F cannot be empty", vbExclamation
            ReDim Preserve validateErrors(errorCount)
            validateErrors(errorCount) = "F" & i & " cannot be empty;"
            errorCount = errorCount + 1
            
            ValidateSheet = True
            'Exit For
        End If
      
    Next i
    
    If errorCount > 0 Then
        Dim errorMessage As String
        errorMessage = Join(validateErrors, vbCrLf)
        ValidationErrorSheet (errorMessage)
        'MsgBox "Validation Errors:" & vbCrLf & errorMessage, vbExclamation
    Else
        MsgBox "All validations passed", vbInformation
    End If
    
    ' If no errors are found, return True
End Function
Function ValidationErrorSheet(errorMessage As String) As Boolean
    Dim wsName As String
    Dim wsExists As Boolean
    Dim ws As Worksheet
    
    wsName = "ValidationErrors"
    wsExists = False

    'Check if the "ValidationErrors" worksheet already exists
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = wsName Then
            wsExists = True
            Exit For
        End If
    Next ws
    
    ' If the sheet does not exist, add a new worksheet
    If Not wsExists Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = wsName
    Else
        Set ws = ThisWorkbook.Worksheets(wsName)
    End If
    
    ws.Columns(1).ClearContents
    
    ws.Cells(1, 1).Value = "Errors"
    
    ' Split the errorMessage by semicolon
    errorMessages = Split(errorMessage, ";")
    
    For i = LBound(errorMessages) To UBound(errorMessages)
        ws.Cells(ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1, 1).Value = Trim(errorMessages(i))
    Next i

    ws.Activate
    
End Function

Function FindLargestRow(num1 As Long, num2 As Long, num3 As Long, num4 As Long) As Long
    Dim largest As Long
    
    ' Assume the first number is the largest initially
    largest = num1
    
    If num2 > largest Then
        largest = num2
    End If
    
    If num3 > largest Then
        largest = num3
    End If
    
    If num4 > largest Then
        largest = num4
    End If
    
    FindLargestRow = largest
End Function