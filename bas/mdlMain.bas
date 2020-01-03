Attribute VB_Name = "mdlMain"
'' [ imageMSO ]
'' https://bert-toolkit.com/imagemso-list.html

'' [ git ]
'' https://githowto.com/pt-BR/create_a_project

'' [ Excel named range - how to define and use names in Excel ]
'' https://www.ablebits.com/office-addins-blog/2017/07/11/excel-name-named-range-define-use/

Private Const ColumnIndex As Integer = 3
Private Const InicioDaPesquisa As Long = 3
Private Const ColunaStatus As String = "E"
Private Const ColunaTarefa As String = "F"
Private Const ColunaDiario As String = "D"
Private cell As Range
Private strBody As String


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
Dim strSearch As String: strSearch = Etiqueta("eMail_Search")
Dim strSubject As String: strSubject = Etiqueta("eMail_Subject")

'' Confirmação de envio de e-mail
Dim sTitle As String:       sTitle = ws.Name
Dim sMessage As String:     sMessage = "Deseja criar uma tarefa com a posição atual ?"
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
        .strCategory = ws.Range(strSearch).Value

        '' Body
        For Each cell In ws.Range("$" & ColunaTarefa & "$" & linha & ":$" & ColunaTarefa & "$" & lRow)
            If InStr(ws.Range("$" & ColunaStatus & "$" & linha).Value, ws.Range(strSearch).Value) <> 0 Then
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

        '' Apresentação
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

Sub send_communication_current(ByVal control As IRibbonControl) '' Enviar e-mail com posição da tarefa atual

'' Global
strBody = ""

'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim eMail As New clsOutlook

'' Principal
Dim strSearch As String: strSearch = Etiqueta("eMail_Search")
Dim strSubject As String: strSubject = Etiqueta("eMail_Subject")
Dim strTo As String: strTo = Etiqueta("eMail_To")
Dim strCC As String: strCC = Etiqueta("eMail_CC")

'' Confirmação de envio de e-mail
Dim sTitle As String:       sTitle = ws.Name
Dim sMessage As String:     sMessage = "Deseja criar uma tarefa com a posição atual ?"
Dim resposta As Variant

'' criar tmp_file apenas para apresentacao
Dim pathExit As String: pathExit = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\tmp_file" & ".txt"
If (Dir(pathExit) <> "") Then Kill pathExit

'' linhas e colunas
Dim lRow As Long: lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Row
Dim linha As Long: linha = InicioDaPesquisa

'' #dbDados - Usando base auxiliar de emails
Dim A As Variant, i As Integer, tmp As String: tmp = ""
A = Range(strTo)


    With eMail

        '' To
        '' #dbDados - Usando base auxiliar de emails
        For i = 1 To UBound(A)
            tmp = tmp & A(i, 1) & ";"
        Next i
                
        .strTo = Left(tmp, Len(tmp) - 1) ''ws.Range(strTo).Value ''strTo
        
        
        .strCC = strCC
        .strSubject = ws.Range(strSubject).Value
        .strCategory = ws.Range(strSearch).Value

        '' Subject
        For Each cell In ws.Range("$" & ColunaDiario & "$" & linha & ":$" & ColunaDiario & "$" & lRow)
            If InStr(ws.Range("$" & ColunaStatus & "$" & linha).Value, ws.Range(strSearch).Value) <> 0 Then
                If (Len(cell.Value) > 0) Then strBody = strBody & cell.Value & vbNewLine
            End If
            linha = linha + 1
        Next cell

        '' Send
        .strBody = strBody

        '' Apresentação
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


Sub show_Help(ByVal control As IRibbonControl) '' Listar funções da aplicação

MsgBox "Em testes", vbInformation + vbOKOnly

''    ShowVersion

End Sub







