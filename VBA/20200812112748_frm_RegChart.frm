VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_RegChart 
   Caption         =   "Linear Regression Chart"
   ClientHeight    =   11676
   ClientLeft      =   156
   ClientTop       =   624
   ClientWidth     =   17136
   OleObjectBlob   =   "20200812112748_frm_RegChart.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_RegChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ChartBtn_Click()
    Call Chart_Preparation.LoadChart
End Sub

Private Sub UserForm_Initialize()

Application.ScreenUpdating = False
Application.Calculation = xlManual
Application.DisplayAlerts = False
Application.EnableEvents = False

On Error GoTo errhandler

Dim NrOfElements As Long
Dim nOfX As Long
Dim allData As Variant
Dim pVal As Variant
Dim iCol As Long
Dim StartCol As Long
Dim RegColumns() As Variant

    nOfX = frm_LinReg_Wks.Xvars.ListCount + 18

    allData = Worksheets("Regression").UsedRange
    
    For iCol = 1 To UBound(allData, 2)
        
        If allData(1, iCol) = "Intercept" Then
            StartCol = iCol - 1
            Exit For
        End If
    Next iCol
    
    pVal = Worksheets("Regression").Range("A18:F" & nOfX).Value
    pValLbo.List = pVal
    
    For iCol = StartCol To UBound(allData, 2)
        
        If Not allData(1, iCol) = "" Then
            If Not allData(1, iCol) = "Intercept" Then
            
                NrOfElements = NrOfElements + 1 ' Increase the variable that is used to increase the nr of elements in array arr.
                ReDim Preserve RegColumns(1 To NrOfElements) '
                RegColumns(UBound(RegColumns)) = allData(1, iCol)
            End If
        End If
    Next iCol
    
    
    Xcbo.List = RegColumns
    Ycbo.List = RegColumns
    
    GoTo endmacro

errhandler:

   Call bCentralErrorHandler( _
                  "frm_RegChart", _
                  "UserForm_Initialize", , True)
    
endmacro:
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.DisplayAlerts = True
Application.EnableEvents = True
End Sub

Private Sub Xcbo_Change()
Application.ScreenUpdating = False
Application.Calculation = xlManual
Application.DisplayAlerts = False
Application.EnableEvents = False

On Error GoTo errhandler

    Call Chart_Preparation.LoadChart
    
    GoTo endmacro

errhandler:

   Call bCentralErrorHandler( _
                  "frm_RegChart", _
                  "Xcbo_Change", , True)


endmacro:
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.DisplayAlerts = True
Application.EnableEvents = True
End Sub

Private Sub Ycbo_Change()
Application.ScreenUpdating = False
Application.Calculation = xlManual
Application.DisplayAlerts = False
Application.EnableEvents = False

On Error GoTo errhandler

    Call Chart_Preparation.LoadChart
    
    GoTo endmacro

errhandler:

   Call bCentralErrorHandler( _
                  "frm_RegChart", _
                  "Ycbo_Change", , True)


endmacro:
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.DisplayAlerts = True
Application.EnableEvents = True
End Sub

