VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_movimentacao_geral_detalhes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalhes da Movimentação"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9315
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
   ScaleHeight     =   4635
   ScaleWidth      =   9315
   Begin VB.TextBox txt_parcela 
      Height          =   315
      Left            =   1980
      TabIndex        =   19
      Top             =   1980
      Width           =   1005
   End
   Begin VB.ComboBox cbo_forma_pagamento 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1200
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker dtp_pagamento 
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Top             =   420
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   39591
   End
   Begin MSComCtl2.DTPicker dtp_vencimento 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   420
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   39591
   End
   Begin VB.ComboBox cbo_receita_despesa 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmd_gravar 
      Caption         =   "&Gravar (F7)"
      Height          =   375
      Left            =   6540
      TabIndex        =   25
      Top             =   4140
      Width           =   1275
   End
   Begin VB.TextBox txt_codigo_barras 
      Height          =   315
      Left            =   4920
      TabIndex        =   21
      Top             =   1980
      Width           =   4275
   End
   Begin VB.TextBox txt_tipo 
      Height          =   315
      Left            =   6900
      TabIndex        =   7
      Top             =   420
      Width           =   2295
   End
   Begin VB.TextBox txt_conta 
      Height          =   315
      Left            =   4920
      TabIndex        =   6
      Top             =   420
      Width           =   1875
   End
   Begin VB.TextBox txt_observacoes 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   2760
      Width           =   9075
   End
   Begin VB.TextBox txt_documento 
      Height          =   315
      Left            =   3060
      TabIndex        =   20
      Top             =   1980
      Width           =   1755
   End
   Begin VB.TextBox txt_descricao 
      Height          =   315
      Left            =   4920
      TabIndex        =   13
      Top             =   1200
      Width           =   4275
   End
   Begin VB.TextBox txt_valor 
      Height          =   315
      Left            =   120
      TabIndex        =   18
      Top             =   1980
      Width           =   1755
   End
   Begin VB.CommandButton cmd_fechar 
      Caption         =   "&Fechar (F8)"
      Height          =   375
      Left            =   7920
      TabIndex        =   26
      Top             =   4140
      Width           =   1275
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
      Left            =   7320
      TabIndex        =   23
      Top             =   2520
      Width           =   1875
   End
   Begin VB.Label lbl_parcela 
      AutoSize        =   -1  'True
      Caption         =   "&Parcela:"
      Height          =   195
      Left            =   1980
      TabIndex        =   15
      Top             =   1680
      Width           =   585
   End
   Begin VB.Label lbl_codigo_barras 
      AutoSize        =   -1  'True
      Caption         =   "&Código de barras:"
      Height          =   195
      Left            =   4920
      TabIndex        =   17
      Top             =   1680
      Width           =   1290
   End
   Begin VB.Label lbl_receita_despesa 
      AutoSize        =   -1  'True
      Caption         =   "&Receita/Despesa:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   900
      Width           =   1275
   End
   Begin VB.Label lbl_forma_pagamento 
      AutoSize        =   -1  'True
      Caption         =   "&Forma de pagamento:"
      Height          =   195
      Left            =   2520
      TabIndex        =   9
      Top             =   900
      Width           =   1590
   End
   Begin VB.Label lbl_tipo 
      AutoSize        =   -1  'True
      Caption         =   "&Tipo:"
      Height          =   195
      Left            =   6900
      TabIndex        =   3
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lbl_conta 
      AutoSize        =   -1  'True
      Caption         =   "&Conta:"
      Height          =   195
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lbl_pagamento 
      AutoSize        =   -1  'True
      Caption         =   "&Pagamento:"
      Height          =   195
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   870
   End
   Begin VB.Label lbl_vencimento 
      AutoSize        =   -1  'True
      Caption         =   "&Vencimento:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   885
   End
   Begin VB.Label lbl_observacoes 
      AutoSize        =   -1  'True
      Caption         =   "&Observações:"
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   2460
      Width           =   1005
   End
   Begin VB.Label lbl_documento 
      AutoSize        =   -1  'True
      Caption         =   "&Documento:"
      Height          =   195
      Left            =   3060
      TabIndex        =   16
      Top             =   1680
      Width           =   870
   End
   Begin VB.Label lbl_descricao 
      AutoSize        =   -1  'True
      Caption         =   "&Descrição:"
      Height          =   195
      Left            =   4920
      TabIndex        =   10
      Top             =   900
      Width           =   750
   End
   Begin VB.Label lbl_valor 
      AutoSize        =   -1  'True
      Caption         =   "&Valor:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   420
   End
End
Attribute VB_Name = "frm_movimentacao_geral_detalhes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'guarda o form anterior
Public mobj_form_anterior As Object

'código do registro
Public mlng_codigo As Long

'tipo da movimentação
Private mstr_tipo As String * 1

Private Sub lsub_carregar_dados(ByVal plng_codigo As Long)
    On Error GoTo erro_lsub_carregar_dados
    'declaração de variáveis
    Dim lobj_movimentacao As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_codigo], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_conta], "
    lstr_sql = lstr_sql & " [tb_contas].[str_descricao] as [str_descricao_conta], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_despesa], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_receita], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_forma_pagamento], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_parcela], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_total_parcelas], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[chr_tipo], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[dt_pagamento], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[dt_vencimento], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[num_valor], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[str_descricao], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[str_documento], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[str_codigo_barras], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[str_observacoes] "
    lstr_sql = lstr_sql & " from "
    lstr_sql = lstr_sql & " [tb_movimentacao] "
    lstr_sql = lstr_sql & " inner join "
    lstr_sql = lstr_sql & " [tb_contas] on [tb_contas].[int_codigo] = [tb_movimentacao].[int_conta] "
    lstr_sql = lstr_sql & " left outer join "
    lstr_sql = lstr_sql & " [tb_formas_pagamento] on [tb_formas_pagamento].[int_codigo] = [tb_movimentacao].[int_forma_pagamento] "
    lstr_sql = lstr_sql & " left outer join "
    lstr_sql = lstr_sql & " [tb_despesas] on [tb_despesas].[int_codigo] = [tb_movimentacao].[int_despesa] "
    lstr_sql = lstr_sql & " left outer join "
    lstr_sql = lstr_sql & " [tb_receitas] on [tb_receitas].[int_codigo] = [tb_movimentacao].[int_receita] "
    lstr_sql = lstr_sql & " where "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_codigo] = " & pfct_tratar_numero_sql(plng_codigo) & " "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_movimentacao, lstr_sql, "frm_movimentacao_geral_detalhes", "lsub_carregar_dados")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_carregar_dados
    End If
    llng_registros = lobj_movimentacao.Count
    If (llng_registros > 0) Then
        '-- preenche os campos da tela --'
        
        'data de vencimento
        dtp_vencimento.Value = Format$(lobj_movimentacao(1)("dt_vencimento"), pcst_formato_data)
        
        'data de pagamento
        dtp_pagamento.Value = Format$(lobj_movimentacao(1)("dt_pagamento"), pcst_formato_data)
        
        'nome da conta
        txt_conta.Text = lobj_movimentacao(1)("str_descricao_conta")
        
        'variável modular tipo
        mstr_tipo = lobj_movimentacao(1)("chr_tipo")
        
        'receitas/despesas
        If (mstr_tipo = "S") Then
            'atribui texto ao componente
            txt_tipo.Text = "- SAÍDA"
            'preenche o combo com as despesas
            psub_preencher_despesas cbo_receita_despesa, False
        Else
            'atribui texto ao componente
            txt_tipo.Text = "- ENTRADA"
            'preenche o combo com as receitas
            psub_preencher_receitas cbo_receita_despesa, False
        End If
        
        'seleciona o item no combo de receitas/despesas
        For llng_contador = 0 To cbo_receita_despesa.ListCount - 1
            If (mstr_tipo = "S") Then
                If (cbo_receita_despesa.ItemData(llng_contador) = lobj_movimentacao(1)("int_despesa")) Then
                    cbo_receita_despesa.ListIndex = llng_contador
                    Exit For
                End If
            Else
                If (cbo_receita_despesa.ItemData(llng_contador) = lobj_movimentacao(1)("int_receita")) Then
                    cbo_receita_despesa.ListIndex = llng_contador
                    Exit For
                End If
            End If
        Next
        
        'seleciona o item no combo de formas de pagamento
        For llng_contador = 0 To cbo_forma_pagamento.ListCount - 1
            If (cbo_forma_pagamento.ItemData(llng_contador) = lobj_movimentacao(1)("int_forma_pagamento")) Then
                cbo_forma_pagamento.ListIndex = llng_contador
                Exit For
            End If
        Next
        
        'descrição
        txt_descricao.Text = lobj_movimentacao(1)("str_descricao")
        
        'valor
        txt_valor.Text = Format$(lobj_movimentacao(1)("num_valor"), pcst_formato_numerico)
        
        'código de barras
        txt_codigo_barras.Text = lobj_movimentacao(1)("str_codigo_barras")
        
        'parcela
        txt_parcela.Text = Format$(lobj_movimentacao(1)("int_parcela"), pcst_formato_numerico_parcela) & "/" & _
                           Format$(lobj_movimentacao(1)("int_total_parcelas"), pcst_formato_numerico_parcela)
        
        'documento
        txt_documento.Text = lobj_movimentacao(1)("str_documento")
        
        'observações
        txt_observacoes.Text = lobj_movimentacao(1)("str_observacoes")
        
    End If
fim_lsub_carregar_dados:
    'destrói os objetos
    Set lobj_movimentacao = Nothing
    Exit Sub
erro_lsub_carregar_dados:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "lsub_carregar_dados"
    GoTo fim_lsub_carregar_dados
    Resume 0
End Sub

Private Sub lsub_habilitar_campos(ByVal pbln_habilitar As Boolean)
    On Error GoTo erro_lsub_habilitar_campos
    
    'campos sempre desabilitados
    txt_conta.Enabled = False
    txt_tipo.Enabled = False
    txt_valor.Enabled = False
    txt_parcela.Enabled = False
    
    'campos que podem ser habilitados/desabilitados
    dtp_vencimento.Enabled = pbln_habilitar
    dtp_pagamento.Enabled = pbln_habilitar
    cbo_receita_despesa.Enabled = pbln_habilitar
    cbo_forma_pagamento.Enabled = pbln_habilitar
    txt_descricao.Enabled = pbln_habilitar
    txt_codigo_barras.Enabled = pbln_habilitar
    txt_documento.Enabled = pbln_habilitar
    txt_observacoes.Enabled = pbln_habilitar
    
fim_lsub_habilitar_campos:
    Exit Sub
erro_lsub_habilitar_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "lsub_habilitar_campos"
    GoTo fim_lsub_habilitar_campos
End Sub

Private Function lfct_salvar_registro(ByVal plng_codigo As Long) As Boolean
    On Error GoTo erro_lfct_salvar_registro
    Dim lobj_movimentacao As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    'variáveis para dados
    Dim llng_receita_despesa As Long
    Dim llng_forma_pagamento As Long
    Dim lstr_data_pagamento As String
    Dim lstr_data_vencimento As String
    Dim lstr_descricao As String
    Dim ldbl_valor_movimentacao As Double
    Dim lstr_documento As String
    Dim lstr_codigo_barras As String
    Dim lstr_observacoes As String
    'preenche as variáveis
    llng_receita_despesa = CLng(cbo_receita_despesa.ItemData(cbo_receita_despesa.ListIndex))
    llng_forma_pagamento = CLng(cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex))
    lstr_data_pagamento = Format$(dtp_pagamento.Value, pcst_formato_data_sql)
    lstr_data_vencimento = Format$(dtp_vencimento.Value, pcst_formato_data_sql)
    lstr_descricao = pfct_tratar_texto_sql(txt_descricao.Text)
    ldbl_valor_movimentacao = CDbl(txt_valor.Text)
    lstr_documento = pfct_tratar_texto_sql(txt_documento.Text)
    lstr_codigo_barras = pfct_tratar_texto_sql(txt_codigo_barras.Text)
    lstr_observacoes = pfct_tratar_texto_sql(txt_observacoes.Text)
    ' --- verificamos se não estamos duplicando a conta a receber - início --- '
    If ((p_usuario.bln_lancamentos_duplicados) And (lstr_documento <> Empty)) Then 'só verificamos se houver também um número de documento
        'monta o comando sql
        lstr_sql = "select * from [tb_movimentacao] where ([dt_pagamento] = '" & lstr_data_pagamento & "' or [dt_vencimento] = '" & lstr_data_vencimento & "') and [chr_tipo] = '" & pfct_tratar_texto_sql(mstr_tipo) & "' and [num_valor] = " & pfct_tratar_numero_sql(ldbl_valor_movimentacao) & " and [str_documento] = '" & lstr_documento & "'"
        'executa o comando sql e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_movimentacao, lstr_sql, "frm_movimentacao_geral_detalhes", "lfct_salvar_registro")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_lfct_salvar_registro
        End If
        llng_registros = lobj_movimentacao.Count
        If (llng_registros > 0) Then
            If (lobj_movimentacao(1)("int_codigo") <> plng_codigo) Then
                'exibe mensagem ao usuário e desvia a execução para o bloco fim
                MsgBox "Este lançamento não pode ser atualizado pois já existe um registro equivalente na movimentação.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                GoTo fim_lfct_salvar_registro
            End If
        End If
    End If
    ' --- verificamos se não estamos duplicando a conta a receber - fim ---
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " update [tb_movimentacao] set "
    'receita
    If (mstr_tipo = "E") Then
        lstr_sql = lstr_sql & " [int_receita] = " & pfct_tratar_numero_sql(llng_receita_despesa) & ", "
    End If
    'despesa
    If (mstr_tipo = "S") Then
        lstr_sql = lstr_sql & " [int_despesa] = " & pfct_tratar_numero_sql(llng_receita_despesa) & ", "
    End If
    lstr_sql = lstr_sql & " [int_forma_pagamento] = " & pfct_tratar_numero_sql(llng_forma_pagamento) & ", "
    lstr_sql = lstr_sql & " [dt_vencimento] = '" & lstr_data_vencimento & "', "
    lstr_sql = lstr_sql & " [dt_pagamento] = '" & lstr_data_pagamento & "', "
    lstr_sql = lstr_sql & " [str_descricao] = '" & lstr_descricao & "', "
    lstr_sql = lstr_sql & " [str_documento] = '" & lstr_documento & "', "
    lstr_sql = lstr_sql & " [str_codigo_barras] = '" & lstr_codigo_barras & "', "
    lstr_sql = lstr_sql & " [str_observacoes] = '" & lstr_observacoes & "' "
    lstr_sql = lstr_sql & " where "
    lstr_sql = lstr_sql & " [int_codigo] = " & pfct_tratar_numero_sql(plng_codigo) & " "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_movimentacao, lstr_sql, "frm_movimentacao_geral_detalhes", "lfct_salvar_registro")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lfct_salvar_registro
    End If
    'devolve valor
    lfct_salvar_registro = True
fim_lfct_salvar_registro:
    Set lobj_movimentacao = Nothing
    Exit Function
erro_lfct_salvar_registro:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "lfct_salvar_registro"
    GoTo fim_lfct_salvar_registro
End Function

Private Function lfct_validar_campos() As Boolean
    On Error GoTo erro_lfct_validar_campos
    'data de vencimento
    If ((dtp_vencimento.Value < Date) And (Not p_usuario.bln_lancamentos_retroativos)) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [data do vencimento] não pode ser menor que a data de hoje.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        dtp_vencimento.Value = Date
        dtp_vencimento.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    'data de pagamento
    If ((dtp_pagamento.Value < Date) And (Not p_usuario.bln_lancamentos_retroativos)) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [data do pagamento] não pode ser menor que a data de hoje.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        dtp_pagamento.Value = Date
        dtp_pagamento.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    'receita/despesa
    If (cbo_receita_despesa.ItemData(cbo_receita_despesa.ListIndex) = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [receita/despesa].", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_receita_despesa.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    'forma de pagamento
    If (cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [forma de pagamento].", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_forma_pagamento.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    'descrição
    If (txt_descricao.Text = "") Then
        MsgBox "Atenção!" & vbCrLf & "O campo [descrição] é de preenchimento obrigatório.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_descricao.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    'devolve valor
    lfct_validar_campos = True
fim_lfct_validar_campos:
    Exit Function
erro_lfct_validar_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "lfct_validar_campos"
    GoTo fim_lfct_validar_campos
End Function

Private Sub cbo_receita_despesa_DropDown()
    On Error GoTo erro_cbo_receita_despesa_DropDown
    psub_campo_got_focus cbo_receita_despesa
fim_cbo_receita_despesa_DropDown:
    Exit Sub
erro_cbo_receita_despesa_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "cbo_receita_despesa_DropDown"
    GoTo fim_cbo_receita_despesa_DropDown
End Sub

Private Sub cbo_receita_despesa_GotFocus()
    On Error GoTo erro_cbo_receita_despesa_GotFocus
    psub_campo_got_focus cbo_receita_despesa
fim_cbo_receita_despesa_GotFocus:
    Exit Sub
erro_cbo_receita_despesa_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "cbo_receita_despesa_GotFocus"
    GoTo fim_cbo_receita_despesa_GotFocus
End Sub

Private Sub cbo_receita_despesa_LostFocus()
    On Error GoTo erro_cbo_receita_despesa_LostFocus
    psub_campo_lost_focus cbo_receita_despesa
fim_cbo_receita_despesa_LostFocus:
    Exit Sub
erro_cbo_receita_despesa_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "cbo_receita_despesa_LostFocus"
    GoTo fim_cbo_receita_despesa_LostFocus
End Sub

Private Sub cbo_forma_pagamento_DropDown()
    On Error GoTo erro_cbo_forma_pagamento_DropDown
    psub_campo_got_focus cbo_forma_pagamento
fim_cbo_forma_pagamento_DropDown:
    Exit Sub
erro_cbo_forma_pagamento_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "cbo_forma_pagamento_DropDown"
    GoTo fim_cbo_forma_pagamento_DropDown
End Sub

Private Sub cbo_forma_pagamento_GotFocus()
    On Error GoTo erro_cbo_forma_pagamento_GotFocus
    psub_campo_got_focus cbo_forma_pagamento
fim_cbo_forma_pagamento_GotFocus:
    Exit Sub
erro_cbo_forma_pagamento_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "cbo_forma_pagamento_GotFocus"
    GoTo fim_cbo_forma_pagamento_GotFocus
End Sub

Private Sub cbo_forma_pagamento_LostFocus()
    On Error GoTo erro_cbo_forma_pagamento_LostFocus
    psub_campo_lost_focus cbo_forma_pagamento
fim_cbo_forma_pagamento_LostFocus:
    Exit Sub
erro_cbo_forma_pagamento_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "cbo_forma_pagamento_LostFocus"
    GoTo fim_cbo_forma_pagamento_LostFocus
End Sub

Private Sub cmd_fechar_Click()
    On Error GoTo erro_cmd_fechar_Click
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_fechar.Enabled) Then
        Exit Sub
    End If
    Unload Me
fim_cmd_fechar_Click:
    Exit Sub
erro_cmd_fechar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "cmd_fechar_Click"
    GoTo fim_cmd_fechar_Click
End Sub

Private Sub cmd_gravar_Click()
    On Error GoTo erro_cmd_gravar_Click
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_gravar.Enabled) Then
        Exit Sub
    Else
        If (lfct_validar_campos()) Then
            If (lfct_salvar_registro(mlng_codigo)) Then
                'exibe mensagem ao usuário, descarrega o form e desvia a execução para o bloco fim
                MsgBox "Movimentação alterada com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                'descarrega o form
                Unload Me
                'dispara a atualização do form anterior
                mobj_form_anterior.cmd_filtrar_Click
            End If
        End If
    End If
fim_cmd_gravar_Click:
    Exit Sub
erro_cmd_gravar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "cmd_gravar_Click"
    GoTo fim_cmd_gravar_Click
End Sub

Private Sub dtp_vencimento_DropDown()
    On Error GoTo erro_dtp_vencimento_DropDown
    psub_campo_got_focus dtp_vencimento
fim_dtp_vencimento_DropDown:
    Exit Sub
erro_dtp_vencimento_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "dtp_vencimento_DropDown"
    GoTo fim_dtp_vencimento_DropDown
End Sub

Private Sub dtp_vencimento_GotFocus()
    On Error GoTo erro_dtp_vencimento_GotFocus
    psub_campo_got_focus dtp_vencimento
fim_dtp_vencimento_GotFocus:
    Exit Sub
erro_dtp_vencimento_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "dtp_vencimento_GotFocus"
    GoTo fim_dtp_vencimento_GotFocus
End Sub

Private Sub dtp_vencimento_LostFocus()
    On Error GoTo erro_dtp_vencimento_LostFocus
    psub_campo_lost_focus dtp_vencimento
fim_dtp_vencimento_LostFocus:
    Exit Sub
erro_dtp_vencimento_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "dtp_vencimento_LostFocus"
    GoTo fim_dtp_vencimento_LostFocus
End Sub

Private Sub dtp_pagamento_DropDown()
    On Error GoTo erro_dtp_pagamento_DropDown
    psub_campo_got_focus dtp_pagamento
fim_dtp_pagamento_DropDown:
    Exit Sub
erro_dtp_pagamento_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "dtp_pagamento_DropDown"
    GoTo fim_dtp_pagamento_DropDown
End Sub

Private Sub dtp_pagamento_GotFocus()
    On Error GoTo erro_dtp_pagamento_GotFocus
    psub_campo_got_focus dtp_pagamento
fim_dtp_pagamento_GotFocus:
    Exit Sub
erro_dtp_pagamento_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "dtp_pagamento_GotFocus"
    GoTo fim_dtp_pagamento_GotFocus
End Sub

Private Sub dtp_pagamento_LostFocus()
    On Error GoTo erro_dtp_pagamento_LostFocus
    psub_campo_lost_focus dtp_pagamento
fim_dtp_pagamento_LostFocus:
    Exit Sub
erro_dtp_pagamento_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "dtp_pagamento_LostFocus"
    GoTo fim_dtp_pagamento_LostFocus
End Sub

Private Sub Form_Activate()
    On Error GoTo Erro_Form_Activate
    'preenche o combo de formas de pagamento
    psub_preencher_formas_pagamento cbo_forma_pagamento
    'carrega os dados
    lsub_carregar_dados mlng_codigo
Fim_Form_Activate:
    Exit Sub
Erro_Form_Activate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "Form_Activate"
    GoTo Fim_Form_Activate
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            If (mobj_form_anterior.Name = "frm_movimentacao_geral") Then
                psub_exibir_ajuda Me, "html/movimentacao_geral_detalhes.htm", 0
            ElseIf (mobj_form_anterior.Name = "frm_movimentacao_por_receitas_despesas") Then
                psub_exibir_ajuda Me, "html/movimentacao_por_receitas_despesas_detalhes.htm", 0
            End If
        Case vbKeyF7
            cmd_gravar_Click
        Case vbKeyF8
            cmd_fechar_Click
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
    
    'ajusta as configurações do form
    Me.Left = mobj_form_anterior.Left + 250
    Me.Top = mobj_form_anterior.Top + 250
    
    'desabilita o form anterior
    mobj_form_anterior.Enabled = False
    
    'configurações
    If (p_usuario.bln_alteracoes_detalhes) Then
        'habilita/desabilita campos
        lsub_habilitar_campos (True)
        'habilita o botão
        cmd_gravar.Enabled = True
        cmd_gravar.Visible = True
    Else
        'habilita/desabilita campos
        lsub_habilitar_campos (False)
        'desabilita o botão
        cmd_gravar.Enabled = False
        cmd_gravar.Visible = False
    End If
    
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "Form_Load"
    GoTo fim_Form_Load
End Sub

Private Sub Form_Terminate()
    On Error GoTo erro_Form_Terminate
    
    'destrói os objetos
    Set mobj_form_anterior = Nothing
    
fim_Form_Terminate:
    Exit Sub
erro_Form_Terminate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "Form_Terminate"
    GoTo fim_Form_Terminate
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo erro_Form_Unload
    
    'reabilita o form anterior
    mobj_form_anterior.Enabled = True
    'destrói o próprio form
    Set frm_movimentacao_geral_detalhes = Nothing

fim_Form_Unload:
    Exit Sub
erro_Form_Unload:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "Form_Unload"
    GoTo fim_Form_Unload
End Sub

Private Sub txt_codigo_barras_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_codigo_barras_validate
    psub_tratar_campo txt_codigo_barras
fim_txt_codigo_barras_validate:
    Exit Sub
erro_txt_codigo_barras_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "txt_codigo_barras_validate"
    GoTo fim_txt_codigo_barras_validate
End Sub

Private Sub txt_conta_GotFocus()
    On Error GoTo erro_txt_conta_GotFocus
    psub_campo_got_focus txt_conta
fim_txt_conta_GotFocus:
    Exit Sub
erro_txt_conta_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "txt_conta_GotFocus"
    GoTo fim_txt_conta_GotFocus
End Sub

Private Sub txt_conta_LostFocus()
    On Error GoTo erro_txt_conta_LostFocus
    psub_campo_lost_focus txt_conta
fim_txt_conta_LostFocus:
    Exit Sub
erro_txt_conta_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "txt_conta_LostFocus"
    GoTo fim_txt_conta_LostFocus
End Sub

Private Sub txt_descricao_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_descricao_validate
    psub_tratar_campo txt_descricao
fim_txt_descricao_validate:
    Exit Sub
erro_txt_descricao_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "txt_descricao_validate"
    GoTo fim_txt_descricao_validate
End Sub

Private Sub txt_documento_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_documento_validate
    psub_tratar_campo txt_documento
fim_txt_documento_validate:
    Exit Sub
erro_txt_documento_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "txt_documento_validate"
    GoTo fim_txt_documento_validate
End Sub

Private Sub txt_observacoes_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_observacoes_validate
    psub_tratar_campo txt_observacoes
fim_txt_observacoes_validate:
    Exit Sub
erro_txt_observacoes_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_detalhes", "txt_observacoes_validate"
    GoTo fim_txt_observacoes_validate
End Sub

Private Sub txt_tipo_GotFocus()
    On Error GoTo erro_txt_tipo_GotFocus
    psub_campo_got_focus txt_tipo
fim_txt_tipo_GotFocus:
    Exit Sub
erro_txt_tipo_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_tipo_detalhes", "txt_tipo_GotFocus"
    GoTo fim_txt_tipo_GotFocus
End Sub

Private Sub txt_tipo_LostFocus()
    On Error GoTo erro_txt_tipo_LostFocus
    psub_campo_lost_focus txt_tipo
fim_txt_tipo_LostFocus:
    Exit Sub
erro_txt_tipo_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_tipo_detalhes", "txt_tipo_LostFocus"
    GoTo fim_txt_tipo_LostFocus
End Sub

Private Sub txt_descricao_GotFocus()
    On Error GoTo erro_txt_descricao_GotFocus
    psub_campo_got_focus txt_descricao
fim_txt_descricao_GotFocus:
    Exit Sub
erro_txt_descricao_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_descricao_detalhes", "txt_descricao_GotFocus"
    GoTo fim_txt_descricao_GotFocus
End Sub

Private Sub txt_descricao_LostFocus()
    On Error GoTo erro_txt_descricao_LostFocus
    psub_campo_lost_focus txt_descricao
fim_txt_descricao_LostFocus:
    Exit Sub
erro_txt_descricao_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_descricao_detalhes", "txt_descricao_LostFocus"
    GoTo fim_txt_descricao_LostFocus
End Sub

Private Sub txt_valor_GotFocus()
    On Error GoTo erro_txt_valor_GotFocus
    psub_campo_got_focus txt_valor
fim_txt_valor_GotFocus:
    Exit Sub
erro_txt_valor_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_valor_detalhes", "txt_valor_GotFocus"
    GoTo fim_txt_valor_GotFocus
End Sub

Private Sub txt_valor_LostFocus()
    On Error GoTo erro_txt_valor_LostFocus
    psub_campo_lost_focus txt_valor
fim_txt_valor_LostFocus:
    Exit Sub
erro_txt_valor_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_valor_detalhes", "txt_valor_LostFocus"
    GoTo fim_txt_valor_LostFocus
End Sub

Private Sub txt_codigo_barras_GotFocus()
    On Error GoTo erro_txt_codigo_barras_GotFocus
    psub_campo_got_focus txt_codigo_barras
fim_txt_codigo_barras_GotFocus:
    Exit Sub
erro_txt_codigo_barras_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_codigo_barras_detalhes", "txt_codigo_barras_GotFocus"
    GoTo fim_txt_codigo_barras_GotFocus
End Sub

Private Sub txt_codigo_barras_LostFocus()
    On Error GoTo erro_txt_codigo_barras_LostFocus
    psub_campo_lost_focus txt_codigo_barras
fim_txt_codigo_barras_LostFocus:
    Exit Sub
erro_txt_codigo_barras_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_codigo_barras_detalhes", "txt_codigo_barras_LostFocus"
    GoTo fim_txt_codigo_barras_LostFocus
End Sub

Private Sub txt_parcela_GotFocus()
    On Error GoTo erro_txt_parcela_GotFocus
    psub_campo_got_focus txt_parcela
fim_txt_parcela_GotFocus:
    Exit Sub
erro_txt_parcela_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_parcela_detalhes", "txt_parcela_GotFocus"
    GoTo fim_txt_parcela_GotFocus
End Sub

Private Sub txt_parcela_LostFocus()
    On Error GoTo erro_txt_parcela_LostFocus
    psub_campo_lost_focus txt_parcela
fim_txt_parcela_LostFocus:
    Exit Sub
erro_txt_parcela_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_parcela_detalhes", "txt_parcela_LostFocus"
    GoTo fim_txt_parcela_LostFocus
End Sub


Private Sub txt_documento_GotFocus()
    On Error GoTo erro_txt_documento_GotFocus
    psub_campo_got_focus txt_documento
fim_txt_documento_GotFocus:
    Exit Sub
erro_txt_documento_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_documento_detalhes", "txt_documento_GotFocus"
    GoTo fim_txt_documento_GotFocus
End Sub

Private Sub txt_documento_LostFocus()
    On Error GoTo erro_txt_documento_LostFocus
    psub_campo_lost_focus txt_documento
fim_txt_documento_LostFocus:
    Exit Sub
erro_txt_documento_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_documento_detalhes", "txt_documento_LostFocus"
    GoTo fim_txt_documento_LostFocus
End Sub

Private Sub txt_observacoes_GotFocus()
    On Error GoTo erro_txt_observacoes_gotFocus
    psub_campo_got_focus txt_observacoes
fim_txt_observacoes_gotFocus:
    Exit Sub
erro_txt_observacoes_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_observacoes_detalhes", "txt_observacoes_GotFocus"
    GoTo fim_txt_observacoes_gotFocus
End Sub

Private Sub txt_observacoes_LostFocus()
    On Error GoTo erro_txt_observacoes_LostFocus
    psub_campo_lost_focus txt_observacoes
fim_txt_observacoes_LostFocus:
    Exit Sub
erro_txt_observacoes_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral_observacoes_detalhes", "txt_observacoes_LostFocus"
    GoTo fim_txt_observacoes_LostFocus
End Sub
