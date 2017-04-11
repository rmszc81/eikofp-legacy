Attribute VB_Name = "bas_campos"
Option Explicit

Public Const pcst_formato_numerico_parcela As String = "00"
Public Const pcst_formato_numerico_padrao As String = "000"
Public Const pcst_formato_numerico As String = "###,###,##0.00"
Public Const pcst_formato_data As String = "dd/mm/yyyy"
Public Const pcst_formato_hora As String = "hh:mm:ss"
Public Const pcst_formato_data_sql As String = "yyyy-mm-dd"
Public Const pcst_formato_hora_sql As String = "hh:mm:ss"
Public Const pcst_formato_mes_ano As String = "mm/yyyy"

Public Sub psub_tratar_campo(ByRef pobj_campo As Object)
    On Error GoTo erro_psub_tratar_campo
    If (TypeOf pobj_campo Is TextBox) Then
        'remove os caracteres inválidos
        pobj_campo.Text = Replace$(pobj_campo.Text, "'", "")
        pobj_campo.Text = Replace$(pobj_campo.Text, "&", "")
        pobj_campo.Text = Replace$(pobj_campo.Text, "@@", "")
        pobj_campo.Text = Replace$(pobj_campo.Text, Chr$(160), "") 'este caracter está presente nos arquivos do banco do brasil
        'remove o excesso de espaços
        pobj_campo.Text = pfct_remover_excesso_espacos(UCase$(pobj_campo.Text))
    End If
fim_psub_tratar_campo:
    Exit Sub
erro_psub_tratar_campo:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_campos", "psub_tratar_campo"
    GoTo fim_psub_tratar_campo
End Sub

Public Sub psub_campo_got_focus(ByRef pobj_campo As Object)
    On Error GoTo erro_psub_campo_got_focus
    Dim llng_contador As Long
    If (TypeOf pobj_campo Is DTPicker) Then
        pobj_campo.CalendarBackColor = vbInfoBackground
    ElseIf (TypeOf pobj_campo Is MSFlexGrid) Then
        If (pobj_campo.Row > 0) Then
            pobj_campo.Redraw = False
            pobj_campo.Col = 0
            For llng_contador = 0 To pobj_campo.Cols - 1
                pobj_campo.Col = llng_contador
                pobj_campo.CellBackColor = vbInfoBackground
            Next
            pobj_campo.Redraw = True
            pobj_campo.Col = 0
        End If
    Else
        pobj_campo.BackColor = vbInfoBackground
    End If
    If (TypeOf pobj_campo Is TextBox) Then
        If (Not pobj_campo.MultiLine) Then
            pobj_campo.SelStart = 0
            pobj_campo.SelLength = Len(pobj_campo.Text)
        End If
    End If
    pobj_campo.Refresh
fim_psub_campo_got_focus:
    Exit Sub
erro_psub_campo_got_focus:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_campos", "psub_campo_got_focus"
    GoTo fim_psub_campo_got_focus
End Sub

Public Sub psub_campo_lost_focus(ByRef pobj_campo As Object)
    On Error GoTo erro_psub_campo_lost_focus
    Dim llng_contador As Long
    If (TypeOf pobj_campo Is DTPicker) Then
        pobj_campo.CalendarBackColor = vbWindowBackground
    ElseIf (TypeOf pobj_campo Is MSFlexGrid) Then
        If (pobj_campo.Row > 0) Then
            pobj_campo.Redraw = False
            pobj_campo.Col = 0
            For llng_contador = 0 To pobj_campo.Cols - 1
                pobj_campo.Col = llng_contador
                pobj_campo.CellBackColor = vbWindowBackground
            Next
            pobj_campo.Redraw = True
            pobj_campo.Col = 0
        End If
    Else
        pobj_campo.BackColor = vbWindowBackground
    End If
    pobj_campo.Refresh
fim_psub_campo_lost_focus:
    Exit Sub
erro_psub_campo_lost_focus:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_campos", "psub_campo_lost_focus"
    GoTo fim_psub_campo_lost_focus
End Sub

Public Sub psub_campo_keypress(ByRef pint_key_ascii As Integer)
    On Error GoTo erro_psub_campo_keypress
    If (pint_key_ascii = vbKeyReturn) Then
        SendKeys "{TAB}", 0
        pint_key_ascii = 0
    ElseIf (pint_key_ascii = vbKeyEscape) Then
        SendKeys "+{TAB}", 0
        pint_key_ascii = 0
    End If
fim_psub_campo_keypress:
    Exit Sub
erro_psub_campo_keypress:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_campos", "psub_campo_keypress"
    GoTo fim_psub_campo_keypress
End Sub

Public Function pfct_validar_campo(ByRef pobj_campo As Object, ByVal pnum_tipo_campo As enm_tipo_campo) As Boolean
    On Error GoTo erro_pfct_validar_campo
    If (pobj_campo.Text <> "") Then
        If (pnum_tipo_campo = tc_inteiro) Or (pnum_tipo_campo = tc_monetario) Then
            If (IsNumeric(pobj_campo.Text)) Then
                If (pnum_tipo_campo = tc_monetario) Then
                    pobj_campo.Text = Format$(pobj_campo.Text, pcst_formato_numerico)
                End If
                If (pnum_tipo_campo = tc_inteiro) Then
                    pobj_campo.Text = Val(Format$(pobj_campo.Text, pcst_formato_numerico_padrao))
                End If
            Else
                MsgBox "Valor numérico inválido.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                pobj_campo.Text = ""
                pfct_validar_campo = False
                GoTo fim_pfct_validar_campo
            End If
        Else
            If ((InStr(1, pobj_campo.Text, "'") > 0) Or _
                (InStr(1, pobj_campo.Text, "&") > 0) Or _
                (InStr(1, pobj_campo.Text, "@@") > 0)) Then
                MsgBox "Atenção! Os caracteres ( ' & @@ ) não são permitidos.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                pobj_campo.Text = ""
                pfct_validar_campo = False
                GoTo fim_pfct_validar_campo
            End If
        End If
    End If
    pfct_validar_campo = True
fim_pfct_validar_campo:
    Exit Function
erro_pfct_validar_campo:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_campos", "pfct_validar_campo"
    GoTo fim_pfct_validar_campo
End Function
