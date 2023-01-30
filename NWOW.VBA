'https://github.com/MattR2718/StudentFormResponseFormatter

Sub processName(id As String, headings As Variant, numRows As Long, numCols As Long)

    'Get num rows for that student
    Dim numRowsWithId As Long
    numRowsWithId = 0
    For i = 2 To numRows
        If Cells(i, 6) = id Then
            numRowsWithId = numRowsWithId + 1
        End If
    Next i
    
    'MsgBox numRowsWithId & " Rows with Id " & id
    
    'Get all rows for inputted id and store rows as arrays in student
    Dim student() As Variant
    ReDim student(1 To numRowsWithId)
    Dim studentRow() As Variant
    ReDim studentRow(1 To numCols)
    Dim counter As Long
    counter = 1
    For i = 2 To numRows
        If Cells(i, 6) = id Then
            For j = 1 To numCols
                studentRow(j) = Cells(i, j)
            Next j
            student(counter) = studentRow
            counter = counter + 1
        End If
    Next i
    
    
    'Create new sheet with name of id and exit if name already exists
    For Each sht In ThisWorkbook.Worksheets
        If sht.name = CStr(id) Then
            'MsgBox "ERROR SHEET ALREADY EXISTS WITH NAME: " & CStr(id)
            Exit Sub
        End If
    Next sht
    Sheets.Add(After:=Sheets(Sheets.Count)).name = id
    Sheets(CStr(id)).Cells(1, 1).Value = "Name: " & CStr(student(1)(6))
    
    'Add headings to first column in sheet
    Dim currRow As Integer
    currRow = 5
    For i = 5 To numCols
        If Not (i = 6) Then
            If headings(i) = "Name" Then
                Sheets(CStr(id)).Cells(currRow - 3, 1).Value = "Teacher Name"
            Else
                Sheets(CStr(id)).Cells(currRow - 3, 1).Value = headings(i)
            End If
            currRow = currRow + 1
        End If
    Next i
    
    'Add student data to sheet
    For i = 1 To numRowsWithId
        Sheets(CStr(id)).Cells(1, i + 1).Value = student(i)(7)
        currRow = 5
        For j = 5 To numCols
            If Not (j = 6) Then
                Sheets(CStr(id)).Cells(currRow - 3, i + 1).Value = student(i)(j)
                currRow = currRow + 1
            End If
        Next j
    Next i
    
    'Fit columns to size of text
    ActiveSheet.UsedRange.EntireColumn.AutoFit
    
    'Set border around all used cells
    Dim iRange As Range
    Dim iCells As Range
    Set iRange = ThisWorkbook.ActiveSheet.UsedRange
    For Each iCells In iRange
        iCells.BorderAround _
                LineStyle:=xlContinuous, _
                Weight:=xlThin
    Next iCells
    
    
End Sub

Sub Format()

    'Get number of rows and columns in sheet
    Dim numRows As Long
    Dim numCols As Long
    numRows = Cells(Rows.Count, "A").End(xlUp).Row
    numCols = Cells(1, Columns.Count).End(xlToLeft).Column
    'MsgBox "Number of rows: " & numRows
    'MsgBox "Number of columns: " & numCols
    
    'Get all headings in sheet
    Dim headings() As Variant
    ReDim headings(1 To numCols)
    For i = 1 To numCols
        headings(i) = Cells(1, i)
    Next i
    
    'Get student id
    Dim name As String
    name = InputBox("Enter Student Name (Leave Blank To Run On All): ")
    
    'counter for rows with no student name
    Dim numNoStudentName As Long
    numNoStudentName = 0
    
    If Len(name) = 0 Then
        'Check whether includes last item
        For n = 2 To numRows
            'Student Name is in column 6
            'If student name is empty then will throw error as cannot create sheet with empty name
            If IsEmpty(Cells(n, 6)) Then
                numNoStudentName = numNoStudentName + 1
            Else
                processName Cells(n, 6), headings, numRows, numCols
            End If
        Next n
    Else
        processName name, headings, numRows, numCols
    End If
    
    MsgBox "Number of Empty Student Names: " & numNoStudentName

End Sub
