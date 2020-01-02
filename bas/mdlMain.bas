Attribute VB_Name = "mdlMain"
'' [ imageMSO ]
'' https://bert-toolkit.com/imagemso-list.html

'' [ git ]
'' https://githowto.com/pt-BR/create_a_project

Private Const ColumnIndex As Integer = 3
Private Const InicioDaPesquisa As Long = 3
Private Const ColunaTipoDeArquivo As String = "C"
Private Const ColunaStatus As String = "E"
Private Const ColunaTarefa As String = "D"
Private cell As Range
Private strBody As String
Private tmp As String

Option Explicit


Sub open_Repository(ByVal control As IRibbonControl) '' Criar repositorio

MsgBox "Em testes", vbInformation + vbOKOnly

'Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
'Dim t As String: t = Etiqueta("pathRepository") & ws.Range(Etiqueta("robot_07_Environment")).Value '& ActiveSheet.Name
'Dim strTemp As String: strTemp = Etiqueta("nameFolders")
'Dim item As Variant
'
'    '' BASE
'    CreateDir t
'
'    For Each item In Split(strTemp, "|")
'        CreateDir t & "\" & item
'    Next
'
'
'    Shell "explorer.exe " + t, vbMaximizedFocus

End Sub


Sub open_List_Tasks(ByVal control As IRibbonControl) '' listar tarefas
'' Global
strBody = ""

'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim eMail As New clsOutlook

'' Principal
Dim strFiltro As String: strFiltro = Etiqueta("eMail_Search")
Dim strSubject As String: strSubject = Etiqueta("eMail_Subject")
'Dim strTo As String: strTo = Etiqueta("eMail_To")
'Dim strCC As String: strCC = Etiqueta("eMail_CC")

'' Confirma��o de envio de e-mail
Dim sTitle As String:       sTitle = ws.Name
Dim sMessage As String:     sMessage = "Deseja criar uma tarefa com a posi��o atual ?"
Dim resposta As Variant

'' criar tmp_file apenas para apresentacao
Dim pathExit As String: pathExit = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\tmp_file" & ".txt"
If (Dir(pathExit) <> "") Then Kill pathExit

'' linhas e colunas
Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row
Dim linha As Long: linha = InicioDaPesquisa

    With eMail

        '' To
        .strSubject = ws.Range(strSubject).Value

        '' Subject
        For Each cell In ws.Range("$" & ColunaTarefa & "$" & linha & ":$" & ColunaTarefa & "$" & lRow)
            If InStr(ws.Range("$" & ColunaStatus & "$" & linha).Value, ws.Range(strFiltro).Value) <> 0 Then
                If (Len(cell.Value) > 0) Then strBody = strBody & cell.Value & vbNewLine
                '' START_TIME
                .strStart = IIf(ws.Range("G" & linha).Value <> "", ws.Range("G" & linha).Value, Now())
                '' REMIND_TIME
                .strRemind_Time = ws.Range("H" & linha).Value * 1 * 60
            End If
            linha = linha + 1
        Next cell

        '' Send
        .strBody = strBody

        '' Apresenta��o
        TextFile_Append pathExit, ws.Range(strSubject).Value & vbNewLine
        TextFile_Append pathExit, strBody
        Shell "notepad.exe " & pathExit, vbMaximizedFocus
        Kill pathExit

        resposta = MsgBox(sMessage, vbQuestion + vbYesNo, sTitle)
        If (resposta = vbYes) Then
            .reportAll
            MsgBox "Concluido!", vbInformation + vbOKOnly, sTitle
        End If

    End With

End Sub

Sub send_communication_current(ByVal control As IRibbonControl) '' Enviar e-mail com posi��o da tarefa atual

'' Global
strBody = ""

'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim eMail As New clsOutlook

'' Principal
Dim strFiltro As String: strFiltro = Etiqueta("eMail_Search")
Dim strSubject As String: strSubject = Etiqueta("eMail_Subject")
Dim strTo As String: strTo = Etiqueta("eMail_To")
Dim strCC As String: strCC = Etiqueta("eMail_CC")

'' Confirma��o de envio de e-mail
Dim sTitle As String:       sTitle = ws.Name
Dim sMessage As String:     sMessage = "Deseja criar uma tarefa com a posi��o atual ?"
Dim resposta As Variant

'' criar tmp_file apenas para apresentacao
Dim pathExit As String: pathExit = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\tmp_file" & ".txt"
If (Dir(pathExit) <> "") Then Kill pathExit

'' linhas e colunas
Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row
Dim linha As Long: linha = InicioDaPesquisa

    With eMail

        '' To
        .strTo = strTo
        .strCC = strCC
        .strSubject = ws.Range(strSubject).Value
        .strCategory = ws.Range("C1").Value

        '' Subject
        For Each cell In ws.Range("$" & ColunaTarefa & "$" & linha & ":$" & ColunaTarefa & "$" & lRow)
            If InStr(ws.Range("$" & ColunaStatus & "$" & linha).Value, ws.Range(strFiltro).Value) <> 0 Then
                If (Len(cell.Value) > 0) Then strBody = strBody & cell.Value & vbNewLine
            End If
            linha = linha + 1
        Next cell

        '' Send
        .strBody = strBody

        '' Apresenta��o
        TextFile_Append pathExit, ws.Range(strSubject).Value & vbNewLine
        TextFile_Append pathExit, strBody
        Shell "notepad.exe " & pathExit, vbMaximizedFocus
        Kill pathExit

        resposta = MsgBox(sMessage, vbQuestion + vbYesNo, sTitle)
        If (resposta = vbYes) Then
            .EnviarEmail
            MsgBox "Concluido!", vbInformation + vbOKOnly, sTitle
        End If

    End With

End Sub


Sub show_Help(ByVal control As IRibbonControl) '' Listar fun��es da aplica��o

MsgBox "Em testes", vbInformation + vbOKOnly

''    ShowVersion

End Sub







