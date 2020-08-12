Attribute VB_Name = "CentralErrorHandler"
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
'' Program:   Central Error Handler
'' Desc:      Código fonte da central de erros.
'' Comments: (1)
'' Creators----------------------------------------------------------------------------
'' Programmer          GPN                Email
'' Erik I Santos     BR014359831    erik.ingwersen@br.ey.com
''======================================================================================

Public Const gbDEBUG_MODE As Boolean = False
Public Const glHANDLED_ERROR As Long = 9999
Public Const glUSER_CANCEL As Long = 18
Public Const gsAPP_TITLE As String = "Error"

Private Const msSILENT_ERROR As String = "UserCancel"
Private Const msFILE_ERROR_LOG As String = "Error.log"


Public Function bCentralErrorHandler( _
            ByVal sModule As String, _
            ByVal sProc As String, _
            Optional ByVal sFile As String, _
            Optional ByVal bEntryPoint As Boolean, _
            Optional erl) As Boolean

    Static sErrMsg As String

    Dim iFile       As Integer
    Dim lErrNum     As Long
    Dim sFullSource As String
    Dim sPath       As String
    Dim sLogText    As String

    ' Pega o erro antes de ser deletado pelo sistema
    ' On Error Resume Next below.
    lErrNum = Err.Number
    ' Se for um cancelamento pelo usuário, configura silent error flag
    ' message. Isso vai fazer com que o erro seja ignorado
    If lErrNum = glUSER_CANCEL Then sErrMsg = msSILENT_ERROR
    ' Se o erro for originado ai, a mensagem de erro estático virá
    ' em branco. Nesse caso, guardamos
    ' a mensagem originária na variável estática.
    If Len(sErrMsg) = 0 Then sErrMsg = Err.Description

    ' Não podem haver erros na CentralErrorHandler
    On Error Resume Next

    ' Carrega o default filename se necessário
    If Len(sFile) = 0 Then sFile = ThisWorkbook.Name

    ' Busca o diretório da aplicação
    sPath = ThisWorkbook.Path
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"

    ' Constroi o fully-qualified error source name.
    sFullSource = "[" & sFile & "]" & sModule & "." & sProc

    ' Criação do error text para ser adicionado ao log.
    sLogText = "  " & sFullSource & ", Error " & _
                        CStr(lErrNum) & ": " & sErrMsg

    ' Abre o Log file, escreve as informações do erro e
    ' fecha o log file.
    iFile = FreeFile()
    Open sPath & msFILE_ERROR_LOG For Append As #iFile
    Print #iFile, Format$(Now(), "mm/dd/yy hh:mm:ss"); sLogText
    If bEntryPoint Then Print #iFile,
    Close #iFile

    ' Não mostra silent errors.
    If sErrMsg <> msSILENT_ERROR Then

        ' Mostra a mensagem de erro quando chegamos no entry point
        ' procedure ou imediatamente se estivermos em modo de debug.
        If bEntryPoint Or gbDEBUG_MODE Then
            Application.ScreenUpdating = True
            MsgBox sErrMsg, vbCritical, gsAPP_TITLE
            ' Limpa a variável de erro estático uma vez
            ' Chegamos no entry-point, preparar tudo para o próximo erro
            sErrMsg = vbNullString
        End If

        bCentralErrorHandler = gbDEBUG_MODE

    Else

        If bEntryPoint Then sErrMsg = vbNullString
        bCentralErrorHandler = False
    End If

End Function

