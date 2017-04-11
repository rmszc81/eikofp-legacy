VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_baixar_contas_receber 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa de Conta a Receber"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   5370
   Begin VB.ComboBox cbo_forma_pagamento 
      Height          =   315
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2775
      Width           =   3255
   End
   Begin VB.ComboBox cbo_tipo_receita 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1980
      Width           =   3255
   End
   Begin MSComCtl2.DTPicker dtp_vencimento 
      Height          =   315
      Left            =   3420
      TabIndex        =   3
      Top             =   420
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   39608
   End
   Begin VB.TextBox txt_descricao 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   5115
   End
   Begin VB.TextBox txt_documento 
      Height          =   315
      Left            =   3480
      TabIndex        =   13
      Top             =   2790
      Width           =   1755
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar (F3)"
      Height          =   375
      Left            =   4020
      TabIndex        =   18
      Top             =   4980
      Width           =   1215
   End
   Begin VB.CommandButton cmd_baixar 
      Caption         =   "&Baixar (F2)"
      Height          =   375
      Left            =   2700
      TabIndex        =   17
      Top             =   4980
      Width           =   1215
   End
   Begin VB.TextBox txt_observacoes 
      Height          =   1275
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   3540
      Width           =   5115
   End
   Begin VB.TextBox txt_valor 
      Height          =   315
      Left            =   3480
      TabIndex        =   9
      Top             =   1980
      Width           =   1755
   End
   Begin VB.ComboBox cbo_lancar_conta 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   420
      Width           =   3195
   End
   Begin VB.Label lblCtrlEnter 
      AutoSize        =   -1  'True
      Caption         =   "CTRL + Enter para nova linha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   165
      Left            =   3360
      TabIndex        =   15
      Top             =   3300
      Width           =   1875
   End
   Begin VB.Label lbl_forma_pagamento 
      AutoSize        =   -1  'True
      Caption         =   "&Forma de pagamento:"
      Height          =   195
      Left            =   135
      TabIndex        =   11
      Top             =   2475
      Width           =   1590
   End
   Begin VB.Label lbl_tipo_receita 
      AutoSize        =   -1  'True
      Caption         =   "&Tipo de receita:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1230
   End
   Begin VB.Label lbl_vencimento 
      AutoSize        =   -1  'True
      Caption         =   "&Vencimento:"
      Height          =   195
      Left            =   3420
      TabIndex        =   1
      Top             =   120
      Width           =   885
   End
   Begin VB.Label lbl_descricao 
      AutoSize        =   -1  'True
      Caption         =   "&Descrição:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   750
   End
   Begin VB.Label lbl_documento 
      AutoSize        =   -1  'True
      Caption         =   "&Documento:"
      Height          =   195
      Left            =   3510
      TabIndex        =   10
      Top             =   2460
      Width           =   870
   End
   Begin VB.Label lbl_observacoes 
      AutoSize        =   -1  'True
      Caption         =   "&Observações:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Label lbl_valor 
      AutoSize        =   -1  'True
      Caption         =   "&Valor:"
      Height          =   195
      Left            =   3480
      TabIndex        =   7
      Top             =   1680
      Width           =   420
   End
   Begin VB.Label lbl_lancar_conta 
      AutoSize        =   -1  'True
      Caption         =   "&Lançar na conta:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frm_baixar_contas_receber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng_codigo_conta_receber As Long

'parcela atual e total parcelas
Private mlng_parcela As Long
Private mlng_total_parcelas As Long
Private mbln_baixa_automatica As Boolean
Private mlng_conta_baixa_automatica As Long

Private Function lfct_verificar_campos() As Boolean
    On Error GoTo Erro_lfct_verificar_campos
    Me.ValidateControls
    'valida a data e se pode fazer lançamentos retroativos
    If ((dtp_vencimento.Value < Date) And (Not p_usuario.bln_lancamentos_retroativos)) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [data do vencimento] não pode ser menor que a data de hoje.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        dtp_vencimento.Value = Date
        dtp_vencimento.SetFocus
        GoTo Fim_lfct_verificar_campos
    End If
    If (txt_valor.Text = "") Then
        MsgBox "Atenção!" & vbCrLf & "O campo [valor] é de preenchimento obrigatório.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_valor.SetFocus
        GoTo Fim_lfct_verificar_campos
    End If
    If (txt_descricao.Text = "") Then
        MsgBox "Atenção!" & vbCrLf & "O campo [descrição] é de preenchimento obrigatório.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_descricao.SetFocus
        GoTo Fim_lfct_verificar_campos
    End If
    'valida o combo lançar na conta
    If (cbo_lancar_conta.ItemData(cbo_lancar_conta.ListIndex) = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [lançar na conta].", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_lancar_conta.SetFocus
        GoTo Fim_lfct_verificar_campos
    End If
    'valida o combo tipo de receita
    If (cbo_tipo_receita.ItemData(cbo_tipo_receita.ListIndex) = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [tipo de receita].", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_tipo_receita.SetFocus
        GoTo Fim_lfct_verificar_campos
    End If
    'valida o combo forma de pagamento
    If (cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [forma de pagamento].", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_forma_pagamento.SetFocus
        GoTo Fim_lfct_verificar_campos
    End If
    lfct_verificar_campos = True
Fim_lfct_verificar_campos:
    Exit Function
Erro_lfct_verificar_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "lfct_verificar_campos"
    GoTo Fim_lfct_verificar_campos
End Function

Private Sub lsub_preencher_campos(ByVal plng_codigo As Long)
    On Error GoTo erro_lsub_preencher_campos
    Dim lobj_conta_receber As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    'atribui o código da conta a receber na variável modular
    mlng_codigo_conta_receber = plng_codigo
    'monta o comando sql
    lstr_sql = "select * from [tb_contas_receber] where [int_codigo] = " & pfct_tratar_numero_sql(plng_codigo)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_conta_receber, lstr_sql, "frm_baixar_contas_receber", "lsub_preencher_campos")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_preencher_campos
    End If
    llng_registros = lobj_conta_receber.Count
    If (llng_registros > 0) Then
        'preenche os valores nas variáveis
        mlng_parcela = CLng(lobj_conta_receber(1)("int_parcela"))
        mlng_total_parcelas = CLng(lobj_conta_receber(1)("int_total_parcelas"))
        mbln_baixa_automatica = IIf(lobj_conta_receber(1)("chr_baixa_automatica") = "S", True, False)
        mlng_conta_baixa_automatica = CLng(lobj_conta_receber(1)("int_conta_baixa_automatica"))
        'preenche os valores nos campos
        txt_valor.Text = Format$(lobj_conta_receber(1)("num_valor"), pcst_formato_numerico)
        txt_descricao.Text = lobj_conta_receber(1)("str_descricao")
        dtp_vencimento.Value = CDate(lobj_conta_receber(1)("dt_vencimento"))
        txt_documento.Text = lobj_conta_receber(1)("str_documento")
        txt_observacoes.Text = lobj_conta_receber(1)("str_observacoes")
        'se for uma conta de baixa automática
        If (mlng_conta_baixa_automatica) Then
            'seleciona o item no combo de contas
            For llng_contador = 0 To cbo_lancar_conta.ListCount - 1
                If (cbo_lancar_conta.ItemData(llng_contador) = lobj_conta_receber(1)("int_conta_baixa_automatica")) Then
                    cbo_lancar_conta.ListIndex = llng_contador
                    Exit For
                End If
            Next
        End If
        'seleciona o item no combo de receitas
        For llng_contador = 0 To cbo_tipo_receita.ListCount - 1
            If (cbo_tipo_receita.ItemData(llng_contador) = lobj_conta_receber(1)("int_receita")) Then
                cbo_tipo_receita.ListIndex = llng_contador
                Exit For
            End If
        Next
        'seleciona o item no combo formas de pagamento
        For llng_contador = 0 To cbo_forma_pagamento.ListCount - 1
            If (cbo_forma_pagamento.ItemData(llng_contador) = lobj_conta_receber(1)("int_forma_pagamento")) Then
                cbo_forma_pagamento.ListIndex = llng_contador
                Exit For
            End If
        Next
    End If
fim_lsub_preencher_campos:
    'destrói os objetos
    Set lobj_conta_receber = Nothing
    Exit Sub
erro_lsub_preencher_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "lsub_preencher_campos"
    GoTo fim_lsub_preencher_campos
End Sub

Private Sub cbo_forma_pagamento_DropDown()
    On Error GoTo erro_cbo_forma_pagamento_DropDown
    psub_campo_got_focus cbo_forma_pagamento
fim_cbo_forma_pagamento_DropDown:
    Exit Sub
erro_cbo_forma_pagamento_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "cbo_forma_pagamento_DropDown"
    GoTo fim_cbo_forma_pagamento_DropDown
End Sub

Private Sub cbo_forma_pagamento_GotFocus()
    On Error GoTo erro_cbo_forma_pagamento_GotFocus
    psub_campo_got_focus cbo_forma_pagamento
fim_cbo_forma_pagamento_GotFocus:
    Exit Sub
erro_cbo_forma_pagamento_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "cbo_forma_pagamento_GotFocus"
    GoTo fim_cbo_forma_pagamento_GotFocus
End Sub

Private Sub cbo_forma_pagamento_LostFocus()
    On Error GoTo erro_cbo_forma_pagamento_LostFocus
    psub_campo_lost_focus cbo_forma_pagamento
fim_cbo_forma_pagamento_LostFocus:
    Exit Sub
erro_cbo_forma_pagamento_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "cbo_forma_pagamento_LostFocus"
    GoTo fim_cbo_forma_pagamento_LostFocus
End Sub

Private Sub cbo_lancar_conta_GotFocus()
    On Error GoTo erro_cbo_lancar_conta_gotFocus
    psub_campo_got_focus cbo_lancar_conta
fim_cbo_lancar_conta_gotFocus:
    Exit Sub
erro_cbo_lancar_conta_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "cbo_lancar_conta_GotFocus"
    GoTo fim_cbo_lancar_conta_gotFocus
End Sub

Private Sub cbo_lancar_conta_LostFocus()
    On Error GoTo erro_cbo_lancar_conta_LostFocus
    psub_campo_lost_focus cbo_lancar_conta
fim_cbo_lancar_conta_LostFocus:
    Exit Sub
erro_cbo_lancar_conta_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "cbo_lancar_conta_LostFocus"
    GoTo fim_cbo_lancar_conta_LostFocus
End Sub

Private Sub cbo_lancar_conta_DropDown()
    On Error GoTo erro_cbo_lancar_conta_DropDown
    psub_campo_got_focus cbo_lancar_conta
fim_cbo_lancar_conta_DropDown:
    Exit Sub
erro_cbo_lancar_conta_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "cbo_lancar_conta_DropDown"
    GoTo fim_cbo_lancar_conta_DropDown
End Sub

Private Sub cbo_tipo_receita_DropDown()
    On Error GoTo erro_cbo_tipo_receita_DropDown
    psub_campo_got_focus cbo_tipo_receita
fim_cbo_tipo_receita_DropDown:
    Exit Sub
erro_cbo_tipo_receita_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "cbo_tipo_receita_DropDown"
    GoTo fim_cbo_tipo_receita_DropDown
End Sub

Private Sub cbo_tipo_receita_GotFocus()
    On Error GoTo erro_cbo_tipo_receita_gotFocus
    psub_campo_got_focus cbo_tipo_receita
fim_cbo_tipo_receita_gotFocus:
    Exit Sub
erro_cbo_tipo_receita_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "cbo_tipo_receita_GotFocus"
    GoTo fim_cbo_tipo_receita_gotFocus
End Sub

Private Sub cbo_tipo_receita_LostFocus()
    On Error GoTo erro_cbo_tipo_receita_LostFocus
    psub_campo_lost_focus cbo_tipo_receita
fim_cbo_tipo_receita_LostFocus:
    Exit Sub
erro_cbo_tipo_receita_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "cbo_tipo_receita_LostFocus"
    GoTo fim_cbo_tipo_receita_LostFocus
End Sub

Private Sub cmd_baixar_Click()
    On Error GoTo erro_cmd_baixar_Click
    Dim lint_resposta As Integer
    Dim lobj_movimentacao As Object
    Dim lobj_conta As Object
    Dim lobj_baixar_conta_receber As Object
    Dim lobj_remover_conta_receber As Object
    Dim lobj_atualizar_saldo_conta As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    'variáveis para armazenar os dados dos campos
    Dim llng_conta As Long
    Dim llng_despesa As Long
    Dim llng_receita As Long
    Dim llng_forma_pagamento As Long
    Dim lstr_data_pagamento As String
    Dim lstr_data_vencimento As String
    Dim lstr_descricao As String
    Dim lstr_documento As String
    Dim lstr_codigo_barras As String
    Dim lstr_observacoes As String
    Dim llng_parcela As Long
    Dim llng_total_parcelas As Long
    Dim ldbl_valor_movimentacao As Double
    Dim ldbl_saldo_atual As Double
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_baixar.Enabled) Then
        Exit Sub
    End If
    'verifica os campos
    If (lfct_verificar_campos) Then
        If (mbln_baixa_automatica) Then
            lint_resposta = MsgBox("Esta conta a receber está programada para ser baixada automaticamente em sua data de vencimento." & vbCrLf & "Deseja baixar essa conta mesmo assim?", vbYesNo + vbQuestion + vbDefaultButton2, pcst_nome_aplicacao)
        Else
            lint_resposta = MsgBox("Confirma a baixa?", vbYesNo + vbQuestion + vbDefaultButton2, pcst_nome_aplicacao)
        End If
        If (lint_resposta = vbYes) Then
            'atribui os valores dos campos às variáveis
            llng_conta = CLng(cbo_lancar_conta.ItemData(cbo_lancar_conta.ListIndex))
            llng_despesa = 0
            llng_receita = CLng(cbo_tipo_receita.ItemData(cbo_tipo_receita.ListIndex))
            'parcelas
            llng_parcela = mlng_parcela
            llng_total_parcelas = mlng_total_parcelas
            '
            llng_forma_pagamento = CLng(cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex))
            lstr_data_pagamento = Format$(Date, pcst_formato_data_sql)
            lstr_data_vencimento = Format$(dtp_vencimento.Value, pcst_formato_data_sql)
            lstr_descricao = pfct_tratar_texto_sql(txt_descricao.Text)
            ldbl_valor_movimentacao = CDbl(txt_valor.Text)
            lstr_documento = pfct_tratar_texto_sql(txt_documento.Text)
            lstr_codigo_barras = ""
            lstr_observacoes = pfct_tratar_texto_sql(txt_observacoes.Text)
            ' --- verificamos se não estamos duplicando a conta a receber - início --- '
            If ((p_usuario.bln_lancamentos_duplicados) And (lstr_documento <> Empty)) Then 'só verificamos se houver também um número de documento
                'monta o comando sql
                lstr_sql = "select * from [tb_movimentacao] where ([dt_pagamento] = '" & lstr_data_pagamento & "' or [dt_vencimento] = '" & lstr_data_vencimento + "') and [chr_tipo] = 'E' and [num_valor] = " & pfct_tratar_numero_sql(ldbl_valor_movimentacao) & " and [str_documento] = '" & lstr_documento & "'"
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_movimentacao, lstr_sql, "frm_baixar_contas_receber", "cmd_baixar_Click")) Then
                    MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                    GoTo fim_cmd_baixar_Click
                End If
                llng_registros = lobj_movimentacao.Count
                If (llng_registros > 0) Then
                    'exibe mensagem ao usuário e desvia a execução para o bloco fim
                    MsgBox "Este lançamento não pode ser baixado pois já existe um registro equivalente na movimentação.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                    GoTo fim_cmd_baixar_Click
                End If
            End If
            ' --- verificamos se não estamos duplicando a conta a receber - fim --- '
            'monta o comando sql
            lstr_sql = "select * from [tb_contas] where [int_codigo] = " & pfct_tratar_numero_sql(llng_conta)
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_conta, lstr_sql, "frm_baixar_contas_receber", "cmd_baixar_Click")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_cmd_baixar_Click
            End If
            llng_registros = lobj_conta.Count
            If (llng_registros > 0) Then
                ldbl_saldo_atual = CDbl(lobj_conta(1)("num_saldo"))
                '-- atualiza o saldo da conta --'
                'monta o comando sql
                lstr_sql = ""
                lstr_sql = lstr_sql & " update "
                lstr_sql = lstr_sql & " [tb_contas] "
                lstr_sql = lstr_sql & " set "
                lstr_sql = lstr_sql & " [num_saldo] = " & pfct_tratar_numero_sql((ldbl_saldo_atual + ldbl_valor_movimentacao))
                lstr_sql = lstr_sql & " where "
                lstr_sql = lstr_sql & " [int_codigo] = " & pfct_tratar_numero_sql(llng_conta)
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_atualizar_saldo_conta, lstr_sql, "frm_baixar_contas_receber", "cmd_baixar_Click")) Then
                    MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                    GoTo fim_cmd_baixar_Click
                End If
                '-- insere uma nova movimentação --'
                'monta o comando sql
                lstr_sql = ""
                lstr_sql = lstr_sql & " insert into [tb_movimentacao] "
                lstr_sql = lstr_sql & " ( "
                lstr_sql = lstr_sql & " [int_conta], "
                lstr_sql = lstr_sql & " [int_receita], "
                lstr_sql = lstr_sql & " [int_despesa], "
                lstr_sql = lstr_sql & " [int_forma_pagamento], "
                lstr_sql = lstr_sql & " [chr_tipo], "
                lstr_sql = lstr_sql & " [dt_vencimento], "
                lstr_sql = lstr_sql & " [dt_pagamento], "
                lstr_sql = lstr_sql & " [num_valor], "
                lstr_sql = lstr_sql & " [int_parcela], "
                lstr_sql = lstr_sql & " [int_total_parcelas], "
                lstr_sql = lstr_sql & " [str_descricao], "
                lstr_sql = lstr_sql & " [str_documento], "
                lstr_sql = lstr_sql & " [str_codigo_barras], "
                lstr_sql = lstr_sql & " [str_observacoes] "
                lstr_sql = lstr_sql & " ) values ( "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_conta) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_receita) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_despesa) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_forma_pagamento) & ", "
                lstr_sql = lstr_sql & " 'E', "
                lstr_sql = lstr_sql & " '" & lstr_data_vencimento & "', "
                lstr_sql = lstr_sql & " '" & lstr_data_pagamento & "', "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(ldbl_valor_movimentacao) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_parcela) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_total_parcelas) & ", "
                lstr_sql = lstr_sql & " '" & lstr_descricao & "', "
                lstr_sql = lstr_sql & " '" & lstr_documento & "', "
                lstr_sql = lstr_sql & " '" & lstr_codigo_barras & "',"
                lstr_sql = lstr_sql & " '" & lstr_observacoes & "' "
                lstr_sql = lstr_sql & " ) "
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_baixar_conta_receber, lstr_sql, "frm_baixar_contas_receber", "cmd_baixar_Click")) Then
                    MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                    GoTo fim_cmd_baixar_Click
                End If
                '-- exclui a conta a receber da tabela --'
                'monta o comando sql
                lstr_sql = "delete from [tb_contas_receber] where [int_codigo] = " & pfct_tratar_numero_sql(mlng_codigo_conta_receber)
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_remover_conta_receber, lstr_sql, "frm_baixar_contas_receber", "cmd_baixar_Click")) Then
                    MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                    GoTo fim_cmd_baixar_Click
                End If
                'exibe mensagem ao usuário, descarrega o form e desvia a execução para o bloco fim
                MsgBox "Baixa de conta realizada com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                Unload Me
                GoTo fim_cmd_baixar_Click
            End If
        End If
    End If
fim_cmd_baixar_Click:
    'destrói os objetos
    Set lobj_movimentacao = Nothing
    Set lobj_conta = Nothing
    Set lobj_baixar_conta_receber = Nothing
    Set lobj_remover_conta_receber = Nothing
    Set lobj_atualizar_saldo_conta = Nothing
    Exit Sub
erro_cmd_baixar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "cmd_baixar_Click"
    GoTo fim_cmd_baixar_Click
End Sub

Private Sub cmd_cancelar_Click()
    On Error GoTo erro_cmd_cancelar_Click
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_cancelar.Enabled) Then
        Exit Sub
    End If
    Unload Me
fim_cmd_cancelar_Click:
    Exit Sub
erro_cmd_cancelar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "cmd_cancelar_Click"
    GoTo fim_cmd_cancelar_Click
End Sub

Private Sub dtp_vencimento_GotFocus()
    On Error GoTo erro_dtp_vencimento_GotFocus
    psub_campo_got_focus dtp_vencimento
fim_dtp_vencimento_GotFocus:
    Exit Sub
erro_dtp_vencimento_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "dtp_vencimento_GotFocus"
    GoTo fim_dtp_vencimento_GotFocus
End Sub

Private Sub dtp_vencimento_LostFocus()
    On Error GoTo erro_dtp_vencimento_LostFocus
    psub_campo_lost_focus dtp_vencimento
fim_dtp_vencimento_LostFocus:
    Exit Sub
erro_dtp_vencimento_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "dtp_vencimento_LostFocus"
    GoTo fim_dtp_vencimento_LostFocus
End Sub

Private Sub dtp_vencimento_DropDown()
    On Error GoTo erro_dtp_vencimento_DropDown
    psub_campo_got_focus dtp_vencimento
fim_dtp_vencimento_DropDown:
    Exit Sub
erro_dtp_vencimento_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "dtp_vencimento_DropDown"
    GoTo fim_dtp_vencimento_DropDown
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/financeiro_contas_receber_baixar.htm", 0
        Case vbKeyF2
            cmd_baixar_Click
        Case vbKeyF3
            cmd_cancelar_Click
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_Initialize()
    On Error GoTo Erro_Form_Initialize
    InitCommonControls
Fim_Form_Initialize:
    Exit Sub
Erro_Form_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "Form_Initialize"
    GoTo Fim_Form_Initialize
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
    psub_preencher_contas cbo_lancar_conta
    psub_preencher_formas_pagamento cbo_forma_pagamento
    psub_preencher_receitas cbo_tipo_receita, False
    lsub_preencher_campos frm_cadastro_contas_receber.msf_grade.RowData(frm_cadastro_contas_receber.msf_grade.Row)
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "Form_Load"
    GoTo fim_Form_Load
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo erro_Form_Unload
    With frm_cadastro_contas_receber
        .Enabled = True
        .lsub_preencher_combos
        .lsub_ajustar_grade .msf_grade
    End With
fim_Form_Unload:
    Exit Sub
erro_Form_Unload:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "Form_Unload"
    GoTo fim_Form_Unload
End Sub

Private Sub txt_descricao_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_descricao_validate
    psub_tratar_campo txt_descricao
fim_txt_descricao_validate:
    Exit Sub
erro_txt_descricao_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "txt_descricao_validate"
    GoTo fim_txt_descricao_validate
End Sub

Private Sub txt_descricao_GotFocus()
    On Error GoTo erro_txt_descricao_GotFocus
    psub_campo_got_focus txt_descricao
fim_txt_descricao_GotFocus:
    Exit Sub
erro_txt_descricao_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "txt_descricao_GotFocus"
    GoTo fim_txt_descricao_GotFocus
End Sub

Private Sub txt_descricao_LostFocus()
    On Error GoTo erro_txt_descricao_LostFocus
    psub_campo_lost_focus txt_descricao
fim_txt_descricao_LostFocus:
    Exit Sub
erro_txt_descricao_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "txt_descricao_LostFocus"
    GoTo fim_txt_descricao_LostFocus
End Sub

Private Sub txt_documento_GotFocus()
    On Error GoTo erro_txt_documento_GotFocus
    psub_campo_got_focus txt_documento
fim_txt_documento_GotFocus:
    Exit Sub
erro_txt_documento_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "txt_documento_GotFocus"
    GoTo fim_txt_documento_GotFocus
End Sub

Private Sub txt_documento_LostFocus()
    On Error GoTo erro_txt_documento_LostFocus
    psub_campo_lost_focus txt_documento
fim_txt_documento_LostFocus:
    Exit Sub
erro_txt_documento_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "txt_documento_LostFocus"
    GoTo fim_txt_documento_LostFocus
End Sub

Private Sub txt_documento_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_documento_validate
    psub_tratar_campo txt_documento
fim_txt_documento_validate:
    Exit Sub
erro_txt_documento_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "txt_documento_validate"
    GoTo fim_txt_documento_validate
End Sub

Private Sub txt_observacoes_GotFocus()
    On Error GoTo erro_txt_observacoes_gotFocus
    psub_campo_got_focus txt_observacoes
fim_txt_observacoes_gotFocus:
    Exit Sub
erro_txt_observacoes_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "txt_observacoes_GotFocus"
    GoTo fim_txt_observacoes_gotFocus
End Sub

Private Sub txt_observacoes_LostFocus()
    On Error GoTo erro_txt_observacoes_LostFocus
    psub_campo_lost_focus txt_observacoes
fim_txt_observacoes_LostFocus:
    Exit Sub
erro_txt_observacoes_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "txt_observacoes_LostFocus"
    GoTo fim_txt_observacoes_LostFocus
End Sub

Private Sub txt_observacoes_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_observacoes_validate
    psub_tratar_campo txt_observacoes
fim_txt_observacoes_validate:
    Exit Sub
erro_txt_observacoes_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "txt_observacoes_validate"
    GoTo fim_txt_observacoes_validate
End Sub

Private Sub txt_valor_GotFocus()
    On Error GoTo erro_txt_valor_GotFocus
    psub_campo_got_focus txt_valor
fim_txt_valor_GotFocus:
    Exit Sub
erro_txt_valor_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "txt_valor_GotFocus"
    GoTo fim_txt_valor_GotFocus
End Sub

Private Sub txt_valor_LostFocus()
    On Error GoTo erro_txt_valor_LostFocus
    psub_campo_lost_focus txt_valor
fim_txt_valor_LostFocus:
    Exit Sub
erro_txt_valor_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "txt_valor_LostFocus"
    GoTo fim_txt_valor_LostFocus
End Sub

Private Sub txt_valor_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_valor_validate
    psub_tratar_campo txt_valor
    Cancel = Not pfct_validar_campo(txt_valor, tc_monetario)
fim_txt_valor_validate:
    Exit Sub
erro_txt_valor_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_baixar_contas_receber", "txt_valor_validate"
    GoTo fim_txt_valor_validate
End Sub

