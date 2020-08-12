Attribute VB_Name = "Column_Check"
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
'' Program:   Column_Check
'' Desc:      Verificação dos Valores das colunas. Elimina colunas que possuem valores
''            não numéricos.
'' Comments: (1)
'' Creators----------------------------------------------------------------------------
'' Programmer          GPN                Email
'' Erik I Santos     BR014359831    erik.ingwersen@br.ey.com
''======================================================================================


Function Column_Check(wsName As String) As Variant
'Essas 4 primeiras linhas de código desabilitam algumas
'funções nativas do Excel fazendo com que o código rode
'mais rápido
Application.ScreenUpdating = False
Application.Calculation = xlManual
Application.DisplayAlerts = False
Application.EnableEvents = False

On Error GoTo errhandler

Dim NumericCols() As Variant
Dim NrOfElements As Long
Dim allData As Variant
Dim iCol As Long
Dim iRow As Long
Dim bNonNumeric As Boolean

NrOfElements = 0
    
    allData = Worksheets(wsName).UsedRange
    For iCol = 1 To UBound(allData, 2)
        
        If Not allData(1, iCol) = "" Then
            
            bNonNumeric = False
            For iRow = 2 To UBound(allData, 1)
                
                If Not IsNumeric(allData(iRow, iCol)) = True Then
                    
                    bNonNumeric = True
                    
                    Exit For
                End If
            Next iRow
            
            If bNonNumeric = False Then
                NrOfElements = NrOfElements + 1 ' Increase the variable that is used to increase the nr of elements in array arr.
                ReDim Preserve NumericCols(1 To NrOfElements) '
                NumericCols(UBound(NumericCols)) = allData(1, iCol)
            End If
        End If
    Next iCol

    Column_Check = NumericCols
    
GoTo endmacro

errhandler:

   Call bCentralErrorHandler( _
                  "Column_Check", _
                  "Column_Check", , True)

endmacro:
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.DisplayAlerts = True
Application.EnableEvents = True
End Function
