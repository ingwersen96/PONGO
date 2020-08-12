Attribute VB_Name = "Range_Preparation"
Option Explicit

Sub Range_Prep(ByRef ref As String, _
               ByRef rng As Variant, _
               ByRef Varnames As Variant, _
               ByRef Fault As Boolean)
    
    '*************************************************************************
    '**  Subroutine assembles a single return array from a (possibly) non-  **
    '**  contiguous range reference.                                        **
    '**    Arguments:                                                       **
    '**      ref = a string that contains the range reference to be         **
    '**            analyzed (may be non contiguous)                         **
    '**      rng = variant that returns the range specified in 'ref'        **
    '**      Varnames - a return array that contains the variable names     **
    '**      Fault - boolean equals 1 if the procedure encounters an error  **
    '**  Note:  The procedure assumes that the range begins on the first    **
    '**         row of the worksheet.  The procedure also assumes the range **
    '**         has a non-numeric variable name in the first row.           **
    '**  Developed by Chad A. Rankin 2007 chad_rankin@hotmail.com           **
    '*************************************************************************
    
Dim TestVal As Range, FullString As Variant, j As Long
Dim i As Long, partstring As Variant, temp As Variant
Dim Str1 As Variant, Str2 As Variant, MaxRows As Long
Dim Cols As Long, cnt As Variant, Rows As Long
Dim wks As String, k As Long

    On Error GoTo EndProc
    
    With Application
        'Test for non-contiguous range reference
        
        'If the name of the worksheet contains a single apostrophe, _
         there are spaces in the name.  The apostrophe must be added _
         back prior to removing the sheet name from the reference.
        If InStr(1, ref, "'") > 0 Then
            wks = "'" & .ActiveSheet.Name & "'!"
        Else
            wks = .ActiveSheet.Name & "!"
        End If
        temp = Replace(ref, wks, "")        'Remove "Sheet1!" from string
        temp = Replace(temp, "$", "")       'Remove "$" from string
        FullString = Split(temp, ",")       'String reference
        
        'Determine the number of columns in the total range including _
         non-contiguous ranges
        'Used to dimension return (Data) array later
        
        For i = 0 To UBound(FullString, 1)
            Cols = Cols + .Range(FullString(i)).Columns.Count
        Next i
        
        'Determine rows by assigning a range object to a range on the worksheet
        Set TestVal = .Range(FullString(0))
        
        'Count number of rows (version 2003: 65536)
        MaxRows = .Range("A:A").Rows.Count
        'Find the last row of the data range (uses only the first column)
        If TestVal.Rows.Count = .ActiveSheet.Cells.Rows.Count Then
            Rows = TestVal(MaxRows, 1).End(xlUp).Row  'Start at bottom and go up
        Else
            Rows = TestVal.Rows.Count
        End If
        
        'Dimension the return matrix and variable name array
        ReDim rng(1 To Rows - 1, 1 To Cols), Varnames(1 To Cols)
        
        cnt = 1     'First column in data matrix
        
        'Loop once for each non-contiguous range
        For i = 0 To UBound(FullString, 1)
            
            'Break the first range reference into two parts (column refs)
            partstring = Split(FullString(i), ":")
            
            'Get first part of range reference without row numbers
            '(Val function does not recognize zeros as numbers)
            Str1 = Split(partstring(0), StrReverse(Val(StrReverse(partstring(0)))))
            Str1 = Str1(0)
            
            'Get second part of range reference without row numbers
            Str2 = Split(partstring(1), StrReverse(Val(StrReverse(partstring(1)))))
            Str2 = Str2(0)
            
            'Assemble the entire reference with row numbers
            temp = Str1 & 1 & ":" & Str2 & Rows
            
            'Set range object equal to range on worksheet
            Set TestVal = .Range(temp)
            
            'Iterate through the columns and add to data matrix
            For k = 1 To .Range(temp).Columns.Count
                'Check for numeric variable names
                If IsNumeric(TestVal(1, k)) Then
                    'No variable names
                    GoTo EndProc
                Else
                    'Load variable names
                    Varnames(cnt) = TestVal(1, k)
                    'Load return matrix
                    For j = 1 To Rows - 1
                        rng(j, cnt) = TestVal(j + 1, k)
                    Next j
                    cnt = cnt + 1
                End If
            Next k
        Next i
        
    End With
    
    'No errors encountered
    Exit Sub

'Error encountered
EndProc:
    Fault = True
    Err.Raise 515, "Range_Prep", "Invalid range reference."
End Sub

Sub Remove_Missing(ByRef Data As Variant, _
                   ByRef Mis_Obs As Variant, _
                   Optional Fault As Boolean)
    
    '*************************************************************************
    '**  Subroutine removes rows from the input matrix that contain missing **
    '**  values. The returned matrix is redimensioned.                      **
    '**  Arguments:                                                         **
    '**    Data = input array of data from which rows with missing          **
    '**              values are removed                                     **
    '**    Mis_Obs = return array containing the observation numbers that   **
    '**              removed.                                               **
    '**  Developed by Chad A. Rankin 2007 chad_rankin@hotmail.com           **
    '*************************************************************************
    
Dim temp As Variant, hold As Variant, MisObs As Variant
Dim Missing As Long, cnt As Long, i As Long, j As Long, k As Long
Dim n As Long, p As Long, Indx As Variant, p1 As Long

    On Error GoTo EndProc
    
    'Count missing cells in the data matrix -> use as upper bound estimate of _
                                               rows to delete.
    n = UBound(Data, 1)
    p = UBound(Data, 2)
    p1 = p + 1
    Missing = n * p - Application.Count(Data)
    
    If Missing > 0 Then
        'Dimension temporary matrix and return missing obs vector
        ReDim temp(1 To n, 1 To p), MisObs(1 To Missing), Indx(1 To n)
        
        cnt = 0
        Missing = 0
        For i = 1 To n
            cnt = cnt + 1
            For j = 1 To p
                If IsEmpty(Data(i, j)) Or Not IsNumeric(Data(i, j)) Then
                    Missing = Missing + 1
                    MisObs(Missing) = i
                    cnt = cnt - 1
                    'Remove entire row
                    Exit For
                End If
            Next j
            If j = p1 Then Indx(cnt) = i
        Next i
        
        'Redimension the a temporary array to equal the ouput array
        ReDim hold(1 To cnt, 1 To p), Mis_Obs(1 To n - cnt)
        
        'Dump the matrix with deleted rows into the resized array
        For i = 1 To cnt
            k = Indx(i)
            For j = 1 To p
                hold(i, j) = Data(k, j)
            Next j
        Next i
        
        'Dump the missing obs into the correctly dimensioned array
        For i = 1 To Missing
            Mis_Obs(i) = MisObs(i)
        Next i
        
        'Destroy Data and replace with the reconstructed data matrix
        Data = hold
    End If
    
    'No errors encountered
    Exit Sub

'Error encountered
EndProc:
    Fault = True
    Err.Raise 516, "Remove_Missing", "Error in removal of missing observations."
End Sub

