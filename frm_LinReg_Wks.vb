''
''
''
''                                                             ......
''                                                      .............
''                                               ....................
''                                        ...........................
''                                 ..................................
''                          ....................
''                   .............
''            .........
''        .....
''
''        &&&&&&&&&&&&&&&&&&   &&&&&&&&&&         &&&&&&&&&&
''        &&&&&&&&&&&&&&&&&&&&  &&&&&&&&&&      &&&&&&&&&&
''        &&&&&&&&&&&&&&&&&&&&&   &&&&&&&&&&   &&&&&&&&&&
''        &&&&&&&&&                &&&&&&&&&& &&&&&&&&&
''        &&&&&&&&&                  &&&&&&&&&&&&&&&&&
''        &&&&&&&&&&&&&&&&&&&&        &&&&&&&&&&&&&&
''        &&&&&&&&&&&&&&&&&&&&          &&&&&&&&&&&
''        &&&&&&&&&                      &&&&&&&&&
''        &&&&&&&&&                      &&&&&&&&&
''        &&&&&&&&&&&&&&&&&&&&&&&&       &&&&&&&&&
''        &&&&&&&&&&&&&&&&&&&&&&&&       &&&&&&&&&
''        &&&&&&&&&&&&&&&&&&&&&&&&       &&&&&&&&&
''
''
''====================================================================================

Option Explicit

Private Sub graph_Btn_Click()
    
    Call chartopen.chartopen
    
End Sub

Private Sub SheetsLbo_Change()
Application.ScreenUpdating = False
Application.Calculation = xlManual
Application.DisplayAlerts = False
Application.EnableEvents = False

On Error GoTo errhandler

Dim NumericCols As Variant
    
    'Checks for numeric columns
    If Not SheetsLbo.Value = "" Then
        NumericCols = Column_Check.Column_Check(SheetsLbo.Value)
    End If
    
    'Adds All Numeric Columns to the listbox
    frm_LinReg_Wks.VariableList.List = NumericCols
    
    GoTo endmacro

errhandler:

   Call bCentralErrorHandler( _
                  "frm_LinReg_Wks", _
                  "SheetsLbo_Change", , True)
    
endmacro:
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.DisplayAlerts = True
Application.EnableEvents = True
End Sub

Private Sub UserForm_Activate()
Application.ScreenUpdating = False
Application.Calculation = xlManual
Application.DisplayAlerts = False
Application.EnableEvents = False

On Error GoTo errhandler

Dim ws As Variant
    
    For Each ws In Worksheets
        SheetsLbo.AddItem ws.Name
    Next ws
        
    GoTo endmacro

errhandler:

   Call bCentralErrorHandler( _
                  "frm_LinReg_Wks", _
                  "UserForm_Activate", , True)
    
endmacro:
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.DisplayAlerts = True
Application.EnableEvents = True
End Sub

Private Sub ADD_Y_Click()
Dim i As Integer, j As Integer

    '***************************************************************************
    '* The purpose of this subroutine is to identify and select a response     *
    '* variable.  The VariableList listbox must have at least two entries      *
    '* (at least a simple regression) and the response variable must not have  *
    '* already been chosen.  Mutliple selections are not accepted.             *
    '* Note:  The subroutine does not assume that the range has a non numeric  *
    '*        variable name in the first row.                                  *
    '***************************************************************************

    'Test to see that there are data available
    If VariableList.ListCount < 2 Then
        MsgBox Prompt:="The procedure requires at least two " & _
               "columns of data in the worksheet.", _
               Buttons:=48, Title:="Input Error!"
        Me.Hide
        Exit Sub
    Else
        'Test for multiple selections
        For i = 0 To VariableList.ListCount - 1
            If VariableList.Selected(i) Then j = j + 1
            If j > 1 Then Exit For
        Next i
        If j > 1 Then
            MsgBox Prompt:="Please select a single response variable.", _
                   Buttons:=48, Title:="Input Error!"
            Exit Sub
        Else
            'Check that the Yvar listbox is empty
            If Yvar.ListIndex = -1 Then
                'Test to see if a selection has been made
                If j > 0 Then
                    Yvar.AddItem VariableList.List(VariableList.ListIndex)
                Else
                    MsgBox Prompt:="Please select a response variable.", _
                           Buttons:=48, Title:="Input Error!"
                End If
            Else
                MsgBox Prompt:="A response variable has already been chosen.", _
                       Buttons:=48, Title:="Input Error!"
            End If
        End If
    End If

End Sub

Private Sub Remove_Y_Click()

    '***************************************************************************
    '* The purpose of this subroutine is to remove a response variable name    *
    '* from the Yvar listbox.                                                  *
    '***************************************************************************

    'Test to see that there is a variable in the listbox
    If Yvar.ListIndex = -1 Then
        MsgBox Prompt:="There is no response variable to remove.", _
            Buttons:=48, Title:="Input Error!"
    Else
        Yvar.RemoveItem Yvar.ListIndex
    End If

End Sub
Private Sub ADD_X_Click()
Dim i As Integer, j As Integer, Present As Boolean

    '***************************************************************************
    '* The purpose of this subroutine is to copy predictor variables chosen    *
    '* from the VariableList listbox into the Xvars listbox.  Mutliple         *
    '* selections are accepted. The response variable cannot be selected.  If  *
    '* the selection is already present, it is not added to the listbox.       *
    '* Note:  The subroutine does not assume that the variable names are non   *
    '*        numeric.                                                         *
    '***************************************************************************

    If Yvar.ListIndex <> -1 Then
        'Test for whether predictor variables are present
        If VariableList.ListIndex <> -1 Then
            'Check if a selection has been made
            For i = 0 To VariableList.ListCount - 1
                If VariableList.Selected(i) Then
                    j = 1
                    Exit For
                End If
            Next i
            
            'If there has been a selection, load into Xvars listbox
            If j = 1 Then
                'Selection has been made
                For i = 0 To VariableList.ListCount - 1
                    'Do not add if selection is already present
                    For j = 0 To Xvars.ListCount - 1
                        If Xvars.List(j) = VariableList.List(i) Then
                            Present = 1
                            Exit For
                        Else
                            Present = 0
                        End If
                    Next j
                    'Cannot reselect the response variable
                    If VariableList.Selected(i) And _
                                VariableList.List(i) <> Yvar.List(0) And _
                                Not Present Then
                        Xvars.AddItem VariableList.List(i)
                    End If
                Next i
            Else
                'No selection has been made
                MsgBox Prompt:="Please select predictor variable(s).", _
                               Buttons:=48, Title:="Input Error!"
                
            End If
            
        Else
            'No predictors present
            MsgBox Prompt:="No predictors variables are present." _
                           , Buttons:=48, Title:="Input Error!"
        End If
    Else
        MsgBox Prompt:="Please select a response variable prior to " & _
                       "selecting predictor variables.", Buttons:=48, _
                       Title:="Input Error!"
    End If
    
End Sub

Private Sub Remove_X_Click()
Dim i As Integer, temp As Integer

    '***************************************************************************
    '* The purpose of this subroutine is to remove a response variable name    *
    '* from the Yvar listbox.                                                  *
    '***************************************************************************

    'Check to make sure there are variables to remove
    If Xvars.ListIndex = -1 Then
        MsgBox Prompt:="There are no predictor variables to remove.", _
                       Buttons:=48, Title:="Input Error!"
    Else
        'Check that a variable(s) have been chosen
        For i = 0 To Xvars.ListCount - 1
            If Xvars.Selected(i) Then temp = 1
        Next i
        
        If temp <> 1 Then
            MsgBox Prompt:="Please select variable(s) to remove.", _
                   Buttons:=48, Title:="Input Error!"
        Else
            'Removes the top entry or the entry selected
            For i = Xvars.ListCount - 1 To 0 Step -1
                If Xvars.Selected(i) Then Xvars.RemoveItem (i)
            Next i
        End If
    End If
   
End Sub

Private Sub Cancel_Btn_Click()
   Unload Me
End Sub

Private Sub Help_Btn_Click()
   Me.Hide
   MsgBox Prompt:="Procedure performs multiple linear regression.  " & vbCr & _
                  "The response variable must be chosen prior to " & vbCr & _
                  "the predictor(s).  No more than 50 predictor " & vbCr & _
                  "variables may be chosen at a time.", Buttons:=544, _
                  Title:="Help"
   Me.Show
   Exit Sub
End Sub
Sub OK_Btn_Click()
Dim Missing As Variant, temp As Variant, Varnames As Variant, ref As String
Dim ShtName As String, FinalCol As Double, Time As Double, Data As Variant
Dim cntsheets, newsheet As Worksheet, p As Double, wholestring As Variant
Dim partstring As Variant, Model As String, n As Variant, wks As Variant
Dim i As Double, j As Double, ret_err As Boolean, Intercept As Variant


'***************************************************************************
'* Subroutine executes a multiple linear regression analysis.              *
'* Developed by Chad A. Rankin 2007 chad_rankin@hotmail.com                *
'***************************************************************************

    On Error GoTo EndProc
    Application.Calculation = xlCalculationManual
    
    'Start the timer
    Time = Timer

'****************************
' Import data from listboxes
'****************************
    
    If Yvar.ListCount = 0 Or Xvars.ListCount = 0 Then
        Me.Hide
        MsgBox Prompt:="Please enter data to analyze.", _
                       Buttons:=48, Title:="Input Error!"
        VariableList.Clear  'Starting over-clear contents
        Me.Show
        Exit Sub
    End If
    
    If Xvars.ListCount > 50 Then
        Me.Hide
        MsgBox Prompt:="Procedure accepts no more than 50 predictors.", _
                       Buttons:=48, Title:="Input Error!"
        Me.Show
    End If
    
    'Find the last column with first row non empty
    FinalCol = Application.Cells(1, 255).End(xlToLeft).Column
    
    'Find the response variable and assign the reference
    For j = 1 To FinalCol
        If Cells(1, j) = Yvar.List(i) Then
            wholestring = Range(Cells(1, j), Cells(1, j)).AddressLocal
            partstring = Split(wholestring, "$")
            ref = "$" & partstring(1) & ":$" & partstring(1)
            Exit For
        End If
    Next j
    
    'Find the x variables and assemble a complete reference
    For i = 0 To Xvars.ListCount - 1
        For j = 1 To FinalCol
            If Cells(1, j) = Xvars.List(i) Then
                wholestring = Range(Cells(1, j), Cells(1, j)).AddressLocal
                partstring = Split(wholestring, "$")
                ref = ref & ", $" & partstring(1) & ":$" & partstring(1)
                Exit For
            End If
        Next j
    Next i
    
    Call Progress(0.1)    'update procedure's progress
    
    'Prepare Data: submit data matrix to be prepared
    Call Range_Prep(ref, temp, Varnames, ret_err)
    If ret_err = True Then GoTo EndProc             'error with range reference
    
    Call Progress(0.35)    'update procedure's progress
    
    'Remove any observations with missing values-response included
    'Variable "missing" is returned as an array containing the _
     observation numbers removed (if any).  The count of the array _
     is the number of rows removed.
     
    Call Remove_Missing(temp, Missing)
    n = UBound(temp, 1)
    p = UBound(temp, 2)
    
    Call Progress(0.55)    'update procedure's progress
    
    'Set value of intercept variable
    If ckb_INT.Value = True Then
        Intercept = 1#
    Else
        Intercept = 0#
    End If
    p = p - 1 + Intercept   'p = count of predictors including intercept

    'If intercept is specified, add to data matrix
    ReDim Y(1 To n, 1 To 1), Data(1 To n, 1 To p)
    For i = 1 To n
        Y(i, 1) = temp(i, 1)
        If Intercept = 1 Then Data(i, 1) = 1#
        For j = 1 + Intercept To p
            Data(i, j) = temp(i, j + 1 - Intercept)
        Next j
    Next i
    
    Call Progress(0.65)    'update procedure's progress
    
'****************************
' Regression Model Statement
'****************************
    
    Model = "= "
    Model = Model & """Y"""
    Model = Model & " & "
    Model = Model & """ = """
    
    'The information for the model statement is taken from the _
     worksheet and not hard coded.
    For i = 1 To p
        If Intercept = i Then
            temp = " & " & "round(B19,2)"
        Else
            temp = " & if(sign(B" & 18 + i & ")=-1, "" "","" + "")" & _
                   " & " & "Round(B" & 18 + i & ", 2)" & " & " & _
                   """ """ & " & " & "A" & 18 + i
        End If
        Model = Model & temp
    Next i
   
    Call Progress(0.75)   'update procedure's progress
         
         
'********************************************************************************
'******************************  OUTPUT  ****************************************
'********************************************************************************

    'Output in new worksheet
    'Check workbook for a worksheet named "Regression"
    For Each wks In Application.Worksheets
        If wks.Name = "Regression" Then wks.Delete
    Next
        
    'Place new worksheet after the last worksheet in the workbook
    cntsheets = Application.Sheets.Count
    Set newsheet = Application.Worksheets.Add(after:=Worksheets(cntsheets))
    newsheet.Name = "Regression"
    FinalCol = 0
        
    Call Progress(0.8)    'update procedure's progress
    
    'Get the sheet name-either new or existing
    ShtName = Application.ActiveSheet.Name
    
    With Application
        'Place the data in the worksheet along with variable names
        .Cells(1, 13 + p).Value = Varnames(1)
        .Range(Cells(2, 13 + p), Cells(n + 1, 13 + p)).Value = Y
        For i = 1 To p
            If Intercept = i Then
                .Cells(1, 13 + p + i).Value = "Intercept"
            Else
                .Cells(1, 13 + p + i).Value = Varnames(i + 1 - Intercept)
            End If
        Next i
        .Range(Cells(2, 14 + p), Cells(n + 1, 13 + p + p)) = Data
        
        'Insert the range formula for the X'Xinv
        .Cells(1, 8).Value = "X'X inverse"
        .Range(Cells(2, 8), Cells(1 + p, 7 + p)).FormulaArray = _
                    "=MINVERSE(MMULT(TRANSPOSE(RC[" & 6 + p & _
                    "]:R[" & n - 1 & "]C[" & 5 + p + p & "]),RC[" & 6 + p & _
                    "]:R[" & n - 1 & "]C[" & 5 + p + p & "]))"
        
        'Insert the fomulae for the variance-covariance matrix
        .Cells(2 + p, 8).Value = "Variance-covariance matrix"
        .Range(Cells(3 + p, 8), Cells(p * 2 + 2, 7 + p)).FormulaR1C1 = _
                                            "=R8C4*R[-" & 1 + p & "]C"
        
        'Build the correlation matrix using the 'Correl' function
        'Must apply the function to all combinations to get lower _
         triangular of correlation matrix--get other half by symmetry
        .Cells(3 + 2 * p, 8).Value = "Correlation matrix"
        For i = 1 To p - Intercept
            .Cells(3 + 2 * p + i, 7 + i).Value = 1#
            For j = i + 1 To p - Intercept
                .Cells(3 + 2 * p + j, 7 + i).FormulaR1C1 = _
                        "=Correl(R2C" & 13 + p + Intercept + i & _
                        ":R" & n + 1 & "C" & 13 + p + Intercept + i & "," & _
                        "R2C" & 13 + p + Intercept + j & ":R" & n + 1 & "C" & _
                        13 + p + Intercept + j & ")"
                .Cells(3 + 2 * p + i, 7 + j).FormulaR1C1 = _
                        "=R[" & j - i & "]C[-" & j - i & "]"
            Next j
        Next i
        
        'Calculate the inverse of the correlation matrix
        Cells(4 + 2 * p + p - Intercept, 8).Value = "Inverse Correlation Matrix"
        .Range(Cells(5 + 2 * p + p - Intercept, 8), _
            Cells(4 + 2 * p + 2 * (p - Intercept), 7 + p - Intercept)).FormulaArray = _
                    "=MINVERSE(R" & 4 + 2 * p & "C8:R" & 3 + 2 * p + p - Intercept & _
                    "C" & 7 + p - Intercept & ")"
        
        Call Progress(0.85)    'update procedure's progress
        
        'Output ANOVA table
        .Cells(1, 1).Value = "Regression Analysis of " & Varnames(1)
        .Cells(1, 1).Font.Bold = True
        .Cells(3, 1).Value = "Regression equation:"
        On Error Resume Next
        .Cells(3, 2).Value = Model
        On Error GoTo 0
        .Cells(5, 2).Value = "Sum of"
        .Cells(5, 3).Value = "Degrees of"
        .Cells(5, 4).Value = "Mean"
        .Cells(6, 1).Value = "Source of Variation"
        .Cells(6, 2).Value = "Squares"
        .Cells(6, 3).Value = "Freedom"
        .Cells(6, 4).Value = "Square"
        .Cells(6, 5).Value = "F"
        .Cells(6, 6).Value = "P-value"
        
        'Output fitted values
        .Cells(1, 9 + p).Value = "Fits"
        .Range(Cells(2, 9 + p), Cells(n + 1, 9 + p)).FormulaR1C1 = _
                "=MMULT(RC[5]:RC[" & 4 + p & "],R19C2:R" & 18 + p & "C2)"
        
        'Output residuals
        .Cells(1, 10 + p).Value = "Resids"
        .Range(Cells(2, 10 + p), Cells(n + 1, 10 + p)).FormulaR1C1 = _
                "=RC[3]-RC[-1]"
        
        'Output regression sum of squares
        .Cells(7, 1).Value = "Regression"
        .Cells(7, 2).FormulaR1C1 = "=R[2]C-R[1]C"
        .Cells(7, 3).Value = p - Intercept
        .Cells(7, 4).FormulaR1C1 = "=RC[-2]/RC[-1]"
        .Cells(7, 5).FormulaR1C1 = "=RC[-1]/R[1]C[-1]"
        
        'Output error sum of squares
        .Cells(8, 1).Value = "Error"
        .Cells(8, 2).FormulaR1C1 = _
            "=SUMSQ(R2C" & 10 + p & ":R" & n + 1 & "C" & 10 + p & ")"
        .Cells(8, 3).Value = n - p
        .Cells(8, 4).FormulaR1C1 = "=RC[-2]/RC[-1]"
        
        'Output total sum of squares
        .Cells(9, 1).Value = "Total"
        If Intercept = 1 Then
            .Cells(9, 2).FormulaR1C1 = _
                "=DEVSQ(R2C" & 13 + p & ":R" & n + 1 & "C" & 13 + p & ")"
        Else
            .Cells(9, 2).FormulaR1C1 = _
                "=SUMSQ(R2C" & 13 + p & ":R" & n + 1 & "C" & 13 + p & ")"
        End If
        
        'Output error degrees of freedom
        .Cells(9, 3).Value = n - Intercept
        
        'Output RMSE
        .Cells(11, 2).Value = "s"
        .Cells(11, 3).FormulaR1C1 = "=SQRT(R[-3]C[1])"
        .Cells(11, 3).NumberFormat = "0.0000"
        
        'Output Rsq only with intercept model
        If Intercept = 1 Then
            .Cells(12, 2).Value = "R-sq"
            .Cells(12, 3).FormulaR1C1 = "=R[-5]C[-1]/R[-3]C[-1]"
            .Cells(12, 3).NumberFormat = "0.00%"
            .Cells(13, 2).Value = "R-Sq(adj)"
            .Cells(13, 3).FormulaR1C1 = "=1-R8C4/(R9C2/R8C3)"
            .Cells(13, 3).NumberFormat = "0.00%"
        End If
        
        'Output table of coefficient estimates, etc.
        .Cells(16, 1).Value = "Parameter Estimates"
        .Cells(18, 1).Value = "Predictor"
        .Cells(18, 2).Value = "Coef Est"
        .Cells(18, 3).Value = "Std Error"
        .Cells(18, 4).Value = "t value"
        .Cells(18, 5).Value = "P-value"
        
        'General formatting
        'Draw lines on ANOVA table
        Range(.Cells(4, 1), Cells(4, 6)) _
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range(.Cells(6, 1), Cells(6, 6)) _
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range(.Cells(9, 1), Cells(9, 6)) _
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
        'Draw line for table of coefs, se, VIFs, t & p statistics
        Range(.Cells(18, 1), Cells(18, 5)) _
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        .Columns(1).ColumnWidth = 18
        .Columns(3).ColumnWidth = 11
        .Columns(4).ColumnWidth = 9.75
        .Range(Cells(7, 5), Cells(7, 6)).NumberFormat = "0.0000"
        
        'Output the coefficient estimates
        .Range(Cells(19, 2), Cells(18 + p, 2)).FormulaArray = _
            "=MMULT(R2C8:R" & 1 + p & "C" & 7 + p & _
            ",MMULT(TRANSPOSE(R2C" & 14 + p & ":R" & _
            n + 1 & "C" & 13 + 2 * p & ")," & _
            "R2C" & 13 + p & ":R" & n + 1 & "C" & 13 + p & "))"
        
        
        'Ouput SEs, t values, pvalues, and VIFs
        For i = 1 To p
            'Output variable names
            If i = 1 Then
                If Intercept = 1 Then
                    .Cells(i + 18, 1).Value = "Constant"
                Else
                    .Cells(i + 18, 1).Value = Varnames(i + 1)
                End If
            Else
                .Cells(i + 18, 1).Value = Varnames(i)
            End If
            .Cells(i + 18, 2).NumberFormat = "0.0000"
            
            'Output standard errors
            .Cells(i + 18, 3).FormulaR1C1 = _
                "=SQRT(R" & 2 + p + i & "C[" & 4 + i & "])"
            .Cells(i + 18, 3).NumberFormat = "0.0000"
            .Cells(i + 18, 4).NumberFormat = "0.0000"
            .Cells(i + 18, 5).NumberFormat = "0.0000"
            
            'Output VIFs
            If i > 1 And Intercept = 1 And p > 2 Then
                .Cells(18, 6) = "VIFs"
                .Cells(i + 18, 6).FormulaR1C1 = _
                        "=R" & 3 + 3 * p - Intercept + i & "C[" & i & "]"
                .Cells(i + 18, 6).NumberFormat = "0.0000"
                .Cells(18, 6).Borders(xlEdgeBottom) _
                                        .LineStyle = xlContinuous
            End If
        Next i
       
        Call Progress(0.9)    'update procedure's progress
        
        'Write note detailing the use of observations
        If IsEmpty(Missing) Then
            .Cells(i + 19, 1) = n & _
                " observations were used in the analysis."
        Else
            .Cells(i + 19, 1) = n & _
                " observations were used in the analysis."
            'Two statements to get the verb tense correct
            If UBound(Missing, 1) = 1 Then
                .Cells(i + 20, 1) = UBound(Missing, 1) & _
                    " observation was excluded due to missing values."
            Else
                .Cells(i + 20, 1) = UBound(Missing, 1) & _
                    " observations were excluded due to missing values."
            End If
        End If
        
'****************************
' Diagnostic Calculations
'****************************
        
        'Output the value of the determinant of the correlation matrix
        .Cells(11, 4) = "Determinant"
        .Range(Cells(11, 5), Cells(11, 5)).FormulaArray = _
                                "=MDETERM(R" & 4 + 2 * p & _
                                "C8:R" & 3 + 2 * p + p - Intercept & _
                                "C" & 7 + p - Intercept & ")"
        
        
        Call Progress(0.95)    'update procedure's progress
        
        'Durbin-Watson statistic
        .Cells(1, 11 + p).Value = "Durbin-Watson"
        .Range(Cells(3, 11 + p), Cells(n + 1, 11 + p)).FormulaR1C1 = _
                "=(RC[-1]-R[-1]C[-1])^2"
        .Cells(12, 4) = "DW"
        .Cells(12, 5).FormulaR1C1 = _
                    "=SUM(R3C" & 11 + p & ":R" & n + 1 & "C" & 11 + p & ")/R8C2"
        .Cells(12, 5).NumberFormat = "0.00"
    
        'Output processing time
        .Worksheets(ShtName).Cells(22 + p, 1) = _
            "Computational time: " & .Round(Timer - Time, 2) & " seconds."
        
        Call Progress(1)   'update procedure's progress
        
        'resume worksheet calculations
        .Calculation = xlCalculationAutomatic
        
        'Calculate probabilities after worksheet calculations _
         have been set to automatic
        't values
        .Range(Cells(19, 4), Cells(18 + p, 4)). _
                        FormulaR1C1 = "=RC[-2]/RC[-1]"
        'p values
        .Range(Cells(19, 5), Cells(18 + p, 5)). _
                        FormulaR1C1 = "=TDIST(abs(RC[-1]),R8C3,1)*2"
        'F value
        .Cells(7, 6).FormulaR1C1 = "=FDIST(RC[-1],RC[-3],R[1]C[-3])"

    End With
    
    GoTo endmacro

 
Unload Me

EndProc:
Application.Calculation = xlCalculationAutomatic    'resume worksheet calculations
MsgBox ("Procedure has encountered a fatal error and will terminate.  " & _
         "Error code: " & Err)
endmacro:
Unload Me
End Sub

Sub Progress(Pct)
'This sub updates the width of the bar moving across the _
 progress indicator frame and the % complete caption
   With Me
      .Progress_Frame.Caption = FormatPercent(Pct, 0)
      .ProgressBar.Width = Pct * .Progress_Frame.Width
      .Repaint
   End With
End Sub

