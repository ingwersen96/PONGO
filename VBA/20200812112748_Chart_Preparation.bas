Attribute VB_Name = "Chart_Preparation"
Option Explicit

Sub LoadChart()

    Dim MyChart As Chart
    Dim ChartData As Variant
    Dim ChartName As String
    Dim ochartObj As ChartObject
    Dim oChart As Chart
    
    Dim lastRow As Long
    Dim allData As Variant
    Dim sht As Worksheet
    Dim CurrentSheet As Worksheet
    Dim cht As ChartObject
    
    If frm_RegChart.Xcbo.Value <> "" Then
        If frm_RegChart.Ycbo.Value = "" Then
        
'            Dim answer As Integer
'            answer = MsgBox("Please select Values for both axis!", vbCritical + vbOKOnly, "Value Error")
            Exit Sub
        Else
        
            Call GetArray(frm_RegChart.Xcbo.Value, 1)
            Call GetArray(frm_RegChart.Ycbo.Value, 2)
            
            allData = ThisWorkbook.Worksheets("Chart").UsedRange
            lastRow = UBound(allData, 1)

            ChartName = frm_RegChart.Xcbo.Value & " X " & frm_RegChart.Ycbo.Value
            
            Set cht = Worksheets("Chart").ChartObjects.Add( _
              Left:=0, _
              Width:=600, _
              Top:=0, _
              Height:=500)
              
            cht.Chart.ChartType = xlXYScatter
            'Give chart some data
            cht.Chart.Axes(xlCategory).HasTitle = True
            cht.Chart.Axes(xlCategory).AxisTitle.Caption = frm_RegChart.Xcbo.Value
            cht.Chart.Axes(xlValue).HasTitle = True
            cht.Chart.Axes(xlValue).AxisTitle.Caption = frm_RegChart.Ycbo.Value
            cht.Chart.HasTitle = True
            cht.Chart.ChartTitle.Text = ChartName
            cht.Chart.SetSourceData Source:=Worksheets("Chart").Range("A2:B" & lastRow)
            cht.Chart.SeriesCollection(1).Trendlines.Add(Type:=xlLinear, Forward:=0, Backward:=0, DisplayEquation:=True, DisplayRSquared:=True).Select
            
            Dim imageName As String
            imageName = Application.DefaultFilePath & Application.PathSeparator & "TempChart.gif"
            
            DeleteFile (imageName)
                    
            cht.Chart.Export Filename:=imageName
            
            Call DeleteChartsOnActiveSheet
            
            frm_RegChart.ChartImg.Picture = LoadPicture(imageName)
            
        End If
    End If
    
End Sub

Sub DeleteChartsOnActiveSheet()
  ActiveSheet.ChartObjects.Delete
End Sub

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      ' First remove readonly attribute, if set
      SetAttr FileToDelete, vbNormal
      ' Then delete the file
      Kill FileToDelete
   End If
End Sub

Sub GetArray(colName As String, ColValue As Long)
Application.ScreenUpdating = False
Application.Calculation = xlManual
Application.DisplayAlerts = False
Application.EnableEvents = False

On Error GoTo errhandler

Dim allData As Variant
Dim iCol As Long
Dim iRow As Long

    allData = Worksheets("Regression").UsedRange
    
    If ColValue = 1 Then
        Worksheets("Chart").Range("A:A").Clear
    Else
        Worksheets("Chart").Range("B:B").Clear
    End If
    
    For iCol = 1 To UBound(allData, 2)
        
        If allData(1, iCol) = colName Then
            For iRow = 1 To UBound(allData, 1)
                ThisWorkbook.Worksheets("Chart").Cells(iRow, ColValue) = allData(iRow, iCol)
            Next iRow
        End If
    Next iCol

    GoTo endmacro

errhandler:

   Call bCentralErrorHandler( _
                  "Chart_Preparation", _
                  "GetArray", , True)
    
endmacro:
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.DisplayAlerts = True
Application.EnableEvents = True
End Sub

Sub Trendline()
Dim sht As Worksheet
Dim CurrentSheet As Worksheet
Dim cht As ChartObject

Application.ScreenUpdating = False
Application.EnableEvents = False

Worksheets("Chart").Activate
For Each cht In ActiveSheet.ChartObjects
cht.Activate
ActiveChart.SeriesCollection(1).Trendlines.Add(Type:=xlLinear, Forward:=0, Backward:=0, DisplayEquation:=False, DisplayRSquared:=False).Select
cht.Delete
Next cht
End Sub
