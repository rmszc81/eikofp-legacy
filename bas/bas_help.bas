Attribute VB_Name = "bas_help"
Option Explicit

Private Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" _
    (ByVal hWndCaller As Long, _
     ByVal pszFile As String, _
     ByVal uCommand As Long, _
     dwData As Any) As Long

Private Const HH_DISPLAY_TOPIC As Long = 0
Private Const HH_HELP_CONTEXT As Long = &HF

Public Sub psub_exibir_ajuda(ByRef pobj_form As Form, ByVal pstr_topico As String, ByVal plng_contexto As Long)
    On Error GoTo erro_psub_exibir_ajuda
    
    'muda a pasta corrente
    ChDir App.Path
    
    'se for por tópico
    If (pstr_topico <> "") Then
        HtmlHelp pobj_form.hWnd, "EikoFP.chm", HH_DISPLAY_TOPIC, ByVal pstr_topico
        GoTo fim_psub_exibir_ajuda
    End If
    
    'se for por contexto
    If (plng_contexto <> 0) Then
        HtmlHelp pobj_form.hWnd, "EikoFP.chm", HH_HELP_CONTEXT, ByVal plng_contexto
    End If

fim_psub_exibir_ajuda:
    Exit Sub
erro_psub_exibir_ajuda:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_help", "psub_exibir_ajuda"
    GoTo fim_psub_exibir_ajuda
End Sub

