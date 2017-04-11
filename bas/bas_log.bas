Attribute VB_Name = "bas_log"
Option Explicit

Public Sub psub_gerar_log_sql(ByVal pstr_query_sql As String, ByVal pstr_modulo As String, ByVal pstr_metodo As String, ByVal pstr_erro As String)
    On Error GoTo erro_psub_gerar_log_sql
    Dim lstr_caminho As String
    Dim lstr_linha_log As String
    If (p_modo_debug) Then
        lstr_caminho = p_banco.str_caminho_log
        If (Right$(lstr_caminho, 1) <> "\") Then
            lstr_caminho = lstr_caminho + "\"
        End If
        lstr_caminho = lstr_caminho & "sql.log"
        pstr_query_sql = UCase$(pstr_query_sql)
        lstr_linha_log = Format$(Now, "dd/mm/yyyy hh:mm:ss") & ";" & LCase$(pstr_modulo) & "." & LCase$(pstr_metodo) & ";" & CStr(pstr_query_sql) & ""
        If (Len(pstr_erro) > 0) Then
            lstr_linha_log = lstr_linha_log & ";erro: [" & pstr_erro & "];"
        Else
            lstr_linha_log = lstr_linha_log & ";ok;"
        End If
        Open lstr_caminho For Append As #1
            Print #1, lstr_linha_log
        Close #1
    End If
fim_psub_gerar_log_sql:
    Exit Sub
erro_psub_gerar_log_sql:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_log", "psub_gerar_log_sql"
    GoTo fim_psub_gerar_log_sql
End Sub

Public Sub psub_gerar_log_erro(ByVal plng_erro As Long, ByVal pstr_erro As String, ByVal pstr_modulo As String, ByVal pstr_metodo As String, Optional ByVal pbln_exibe_mensagem As Boolean = True)
    On Error GoTo erro_psub_gerar_log_erro
    Dim lstr_caminho As String
    Dim lstr_linha_log As String
    Dim lstr_mensagem_erro As String
    If (p_modo_debug) Then
        lstr_caminho = p_banco.str_caminho_log
        If (Right$(lstr_caminho, 1) <> "\") Then
            lstr_caminho = lstr_caminho + "\"
        End If
        lstr_caminho = lstr_caminho & "err.log"
        lstr_linha_log = Format$(Now, "dd/mm/yyyy hh:mm:ss") & ";" & LCase$(pstr_modulo) & "." & LCase$(pstr_metodo) & ";" & CStr(plng_erro) & ";" & CStr(pstr_erro) & ";"
        Open lstr_caminho For Append As #1
            Print #1, lstr_linha_log
        Close #1
    End If
    If (pbln_exibe_mensagem) Then
        'monta a mensagem de erro
        lstr_mensagem_erro = "Atenção!"
        lstr_mensagem_erro = lstr_mensagem_erro & vbCrLf
        lstr_mensagem_erro = lstr_mensagem_erro & "Ocorreu um erro durante a operação do sistema."
        lstr_mensagem_erro = lstr_mensagem_erro & vbCrLf & vbCrLf
        lstr_mensagem_erro = lstr_mensagem_erro & "Nº.: " & CStr(plng_erro)
        lstr_mensagem_erro = lstr_mensagem_erro & vbCrLf
        lstr_mensagem_erro = lstr_mensagem_erro & "Descrição: " & CStr(pstr_erro)
        'exibe a mensagem de erro
        MsgBox lstr_mensagem_erro, vbOKOnly + vbCritical, pcst_nome_aplicacao
    End If
fim_psub_gerar_log_erro:
    Exit Sub
erro_psub_gerar_log_erro:
    GoTo fim_psub_gerar_log_erro
End Sub

