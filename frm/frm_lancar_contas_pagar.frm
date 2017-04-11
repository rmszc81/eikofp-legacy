VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_lancar_contas_pagar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançar contas a pagar"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
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
   ScaleHeight     =   7575
   ScaleWidth      =   5205
   Begin VB.ComboBox cbo_conta_baixa_automatica 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   420
      Width           =   2850
   End
   Begin VB.CheckBox chk_baixa_automatica 
      Caption         =   "&Baixar conta automaticamente"
      Height          =   255
      Left            =   2300
      TabIndex        =   1
      Top             =   120
      Width           =   2865
   End
   Begin VB.Frame fme_baixa_conta 
      Caption         =   " Baixa de conta imediata "
      Enabled         =   0   'False
      Height          =   885
      Left            =   90
      TabIndex        =   30
      Top             =   7560
      Width           =   5040
      Begin VB.ComboBox cbo_conta_baixa_imediata 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   375
         Width           =   2850
      End
      Begin VB.Label lbl_conta 
         AutoSize        =   -1  'True
         Caption         =   "&Conta:"
         Height          =   195
         Left            =   675
         TabIndex        =   31
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CheckBox chk_baixar_conta 
      Caption         =   "&Baixar conta imediatamente"
      Height          =   315
      Left            =   135
      TabIndex        =   27
      Top             =   7155
      Width           =   2325
   End
   Begin VB.ComboBox cbo_forma_pagamento 
      Height          =   315
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   4260
      Width           =   2895
   End
   Begin VB.TextBox txt_codigo_barras 
      Height          =   315
      Left            =   120
      TabIndex        =   23
      Top             =   5040
      Width           =   4995
   End
   Begin VB.TextBox txt_documento 
      Height          =   315
      Left            =   3150
      TabIndex        =   21
      Top             =   4275
      Width           =   1950
   End
   Begin VB.CheckBox chk_multiplos 
      Caption         =   " &Lançar múltiplos vencimentos"
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Top             =   815
      Width           =   2475
   End
   Begin VB.Frame fme_multiplos_vencimentos 
      Enabled         =   0   'False
      Height          =   1515
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   4995
      Begin VB.CheckBox chk_antecipar_pagamento 
         Caption         =   "&Antecipar pagamento caso data calculada não seja dia útil"
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   1080
         Width           =   4515
      End
      Begin VB.TextBox txt_quantidade 
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   660
         Width           =   915
      End
      Begin VB.ComboBox cbo_tempo 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   660
         Width           =   1335
      End
      Begin VB.TextBox txt_tempo 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   660
         Width           =   795
      End
      Begin VB.Label lbl_quantidade 
         AutoSize        =   -1  'True
         Caption         =   "&Quantidade:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lbl_cada 
         AutoSize        =   -1  'True
         Caption         =   "L&ançar a cada:"
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.TextBox txt_observacoes 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   5790
      Width           =   4995
   End
   Begin VB.TextBox txt_descricao 
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Top             =   3510
      Width           =   4995
   End
   Begin VB.TextBox txt_valor 
      Height          =   315
      Left            =   3120
      TabIndex        =   15
      Top             =   2775
      Width           =   1995
   End
   Begin VB.ComboBox cbo_tipo_despesa 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2775
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker dtp_vencimento 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Format          =   156958721
      CurrentDate     =   39595
   End
   Begin VB.CommandButton cmd_lancar 
      Caption         =   "&Lançar (F2)"
      Height          =   375
      Left            =   2520
      TabIndex        =   28
      Top             =   7110
      Width           =   1275
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar (F3)"
      Height          =   375
      Left            =   3840
      TabIndex        =   29
      Top             =   7110
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
      Left            =   3240
      TabIndex        =   25
      Top             =   5550
      Width           =   1875
   End
   Begin VB.Label lbl_forma_pagamento 
      AutoSize        =   -1  'True
      Caption         =   "&Forma de Pagamento:"
      Height          =   195
      Left            =   135
      TabIndex        =   18
      Top             =   3960
      Width           =   1590
   End
   Begin VB.Label lbl_codigo_barras 
      AutoSize        =   -1  'True
      Caption         =   "&Código de barras:"
      Height          =   195
      Left            =   135
      TabIndex        =   22
      Top             =   4725
      Width           =   1290
   End
   Begin VB.Label lbl_documento 
      AutoSize        =   -1  'True
      Caption         =   "&Documento:"
      Height          =   195
      Left            =   3150
      TabIndex        =   19
      Top             =   3990
      Width           =   870
   End
   Begin VB.Label lbl_observacoes 
      AutoSize        =   -1  'True
      Caption         =   "&Observações:"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   5490
      Width           =   1005
   End
   Begin VB.Label lbl_descricao 
      AutoSize        =   -1  'True
      Caption         =   "&Descrição:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   3210
      Width           =   750
   End
   Begin VB.Label lbl_valor 
      AutoSize        =   -1  'True
      Caption         =   "&Valor:"
      Height          =   195
      Left            =   3120
      TabIndex        =   13
      Top             =   2475
      Width           =   420
   End
   Begin VB.Label lbl_tipo_despesa 
      AutoSize        =   -1  'True
      Caption         =   "&Tipo de despesa:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   2475
      Width           =   1230
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
End
Attribute VB_Name = "frm_lancar_contas_pagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function lfct_verificar_campos() As Boolean
    On Error GoTo Erro_lfct_verificar_campos
    'força a validação de todos os campos
    Me.ValidateControls
    'verifica se é uma baixa automática
    If (chk_baixa_automatica.Value = vbChecked) Then
        'valida o combo contas
        If (cbo_conta_baixa_automatica.ItemData(cbo_conta_baixa_automatica.ListIndex) = 0) Then
            MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [conta] para a baixa automática.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            cbo_conta_baixa_automatica.SetFocus
            GoTo Fim_lfct_verificar_campos
        End If
    End If
    'valida a data e se pode fazer lançamentos retroativos
    If ((dtp_vencimento.Value < Date) And (Not p_usuario.bln_lancamentos_retroativos)) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [data do vencimento] não pode ser menor que a data de hoje.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        dtp_vencimento.Value = Date
        dtp_vencimento.SetFocus
        GoTo Fim_lfct_verificar_campos
    End If
    'se for múltiplos vencimentos
    If (chk_multiplos.Value = vbChecked) Then
        'valida a quantidade
        If (txt_quantidade.Text = "") Then
            MsgBox "Atenção!" & vbCrLf & "Campo [quantidade] não pode estar em branco.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            txt_quantidade.SetFocus
            GoTo Fim_lfct_verificar_campos
        End If
        'valida o tempo
        If (txt_tempo.Text = "") Then
            MsgBox "Atenção!" & vbCrLf & "Campo [cada] não pode estar em branco.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            txt_tempo.SetFocus
            GoTo Fim_lfct_verificar_campos
        End If
        'valida o combo tempo
        If (cbo_tempo.ItemData(cbo_tempo.ListIndex) = 0) Then
            MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [tempo].", vbOKOnly + vbInformation, pcst_nome_aplicacao
            cbo_tempo.SetFocus
            GoTo Fim_lfct_verificar_campos
        End If
    End If
    'valida o combo despesa
    If (cbo_tipo_despesa.ItemData(cbo_tipo_despesa.ListIndex) = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [tipo de despesa].", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_tipo_despesa.SetFocus
        GoTo Fim_lfct_verificar_campos
    End If
    'valida o campo valor
    If (txt_valor.Text = "") Then
        MsgBox "Atenção!" & vbCrLf & "Campo [valor] não pode estar em branco.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_valor.SetFocus
        GoTo Fim_lfct_verificar_campos
    End If
    'valida o campo descrição
    If (txt_descricao.Text = "") Then
        MsgBox "Atenção!" & vbCrLf & "Campo [descrição] não pode estar em branco.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_descricao.SetFocus
        GoTo Fim_lfct_verificar_campos
    End If
    'valida o combo forma pagamento
    If (cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [forma de pagamento].", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_forma_pagamento.SetFocus
        GoTo Fim_lfct_verificar_campos
    End If
    'verifica se é uma baixa imediata
    If (chk_baixar_conta.Value = vbChecked) Then
        'valida o combo contas
        If (cbo_conta_baixa_imediata.ItemData(cbo_conta_baixa_imediata.ListIndex) = 0) Then
            MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [conta] para a baixa imediata.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            cbo_conta_baixa_imediata.SetFocus
            GoTo Fim_lfct_verificar_campos
        End If
    End If
    'retorna true
    lfct_verificar_campos = True
Fim_lfct_verificar_campos:
    Exit Function
Erro_lfct_verificar_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "lfct_verificar_campos"
    GoTo Fim_lfct_verificar_campos
End Function

Private Sub lsub_ajustar_data()
    On Error GoTo erro_lsub_ajustar_data
    dtp_vencimento.Value = Date
fim_lsub_ajustar_data:
    Exit Sub
erro_lsub_ajustar_data:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "lsub_ajustar_data"
    GoTo fim_lsub_ajustar_data
End Sub

Private Function lfct_lancar_conta_pagar() As Boolean
    On Error GoTo erro_lfct_lancar_conta_pagar
    Dim lobj_lancar_conta_pagar As Object
    Dim lstr_sql As String
    Dim llng_contador As Long
    Dim llng_quantidade As Long
    'variáveis para armazenar dados dos campos
    Dim ldt_data_vencimento As Date
    Dim llng_cada As Long
    Dim llng_tempo As Long
    Dim llng_conta_baixa_automatica As Long
    Dim llng_despesa As Long
    Dim llng_forma_pagamento As Long
    Dim ldbl_valor As Double
    Dim llng_parcela As Long
    Dim llng_total_parcelas As Long
    Dim lstr_baixa_automatica As String
    Dim lstr_descricao As String
    Dim lstr_documento As String
    Dim lstr_codigo_barras As String
    Dim lstr_observacoes As String
    Dim lstr_chave As String
    'atribui os valores dos campos às variáveis
    ldt_data_vencimento = dtp_vencimento.Value
    llng_conta_baixa_automatica = CLng(cbo_conta_baixa_automatica.ItemData(cbo_conta_baixa_automatica.ListIndex))
    llng_despesa = CLng(cbo_tipo_despesa.ItemData(cbo_tipo_despesa.ListIndex))
    llng_forma_pagamento = CLng(cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex))
    ldbl_valor = CDbl(txt_valor.Text)
    'lançamento único
    llng_parcela = 1
    llng_total_parcelas = 1
    '
    lstr_baixa_automatica = IIf(chk_baixa_automatica.Value = vbChecked, "S", "N")
    lstr_descricao = pfct_tratar_texto_sql(txt_descricao.Text)
    lstr_documento = pfct_tratar_texto_sql(txt_documento.Text)
    lstr_codigo_barras = pfct_tratar_texto_sql(txt_codigo_barras.Text)
    lstr_observacoes = pfct_tratar_texto_sql(txt_observacoes.Text)
    'múltiplos vencimentos
    If (chk_multiplos.Value = vbChecked) Then
        'se for múltiplos lançamentos, gera a chave de ligação dos registros
        lstr_chave = pfct_gerar_chave(15)
        'tempo (valor)
        llng_cada = CLng(txt_tempo.Text)
        'tempo (dias/meses/anos)
        llng_tempo = CLng(cbo_tempo.ItemData(cbo_tempo.ListIndex))
        'quantidade de lançamentos
        llng_quantidade = CLng(txt_quantidade.Text)
        'laço para múltipla inserção
        For llng_contador = 1 To llng_quantidade
            'parcelas, atual e total
            llng_parcela = llng_contador
            llng_total_parcelas = llng_quantidade
            'monta o comando sql
            lstr_sql = ""
            lstr_sql = lstr_sql & " insert into [tb_contas_pagar] "
            lstr_sql = lstr_sql & " ( "
            lstr_sql = lstr_sql & " [chr_baixa_automatica], "
            lstr_sql = lstr_sql & " [int_conta_baixa_automatica], "
            lstr_sql = lstr_sql & " [int_despesa], "
            lstr_sql = lstr_sql & " [int_forma_pagamento], "
            lstr_sql = lstr_sql & " [dt_vencimento], "
            lstr_sql = lstr_sql & " [num_valor], "
            lstr_sql = lstr_sql & " [int_parcela], "
            lstr_sql = lstr_sql & " [int_total_parcelas], "
            lstr_sql = lstr_sql & " [str_descricao], "
            lstr_sql = lstr_sql & " [str_documento], "
            lstr_sql = lstr_sql & " [str_chave], "
            lstr_sql = lstr_sql & " [str_codigo_barras], "
            lstr_sql = lstr_sql & " [str_observacoes] "
            lstr_sql = lstr_sql & " ) values ( "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lstr_baixa_automatica) & "', "
            lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_conta_baixa_automatica) & ", "
            lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_despesa) & ", "
            lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_forma_pagamento) & ", "
            lstr_sql = lstr_sql & " '" & Format$(ldt_data_vencimento, pcst_formato_data_sql) & "', "
            lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(ldbl_valor) & ", "
            lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_parcela) & ", " 'parcela atual
            lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_total_parcelas) & ", " 'total parcelas
            lstr_sql = lstr_sql & " '" & lstr_descricao & "', "
            lstr_sql = lstr_sql & " '" & lstr_documento & "', "
            lstr_sql = lstr_sql & " '" & lstr_chave & "',"
            lstr_sql = lstr_sql & " '" & lstr_codigo_barras & "',"
            lstr_sql = lstr_sql & " '" & lstr_observacoes & "' "
            lstr_sql = lstr_sql & " ) "
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_lancar_conta_pagar, lstr_sql, "frm_lancar_contas_pagar", "lfct_lancar_conta_pagar")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_lfct_lancar_conta_pagar
            End If
            'calcula a data de acordo com seleção de período
            Select Case llng_tempo
                Case 30 'meses
                    ldt_data_vencimento = DateAdd("m", llng_cada, ldt_data_vencimento)
                Case 365 'anos
                    ldt_data_vencimento = DateAdd("yyyy", llng_cada, ldt_data_vencimento)
            End Select
            'antecipar pagamento
            If (chk_antecipar_pagamento.Value = vbChecked) Then
                Do While (Weekday(ldt_data_vencimento) = vbSunday) Or (Weekday(ldt_data_vencimento) = vbSaturday)
                    ldt_data_vencimento = ldt_data_vencimento - 1
                Loop
            End If
        Next
    Else
        'limpa a variável
        lstr_chave = ""
        'monta o comando sql
        lstr_sql = ""
        lstr_sql = lstr_sql & " insert into [tb_contas_pagar] "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " [chr_baixa_automatica], "
        lstr_sql = lstr_sql & " [int_conta_baixa_automatica], "
        lstr_sql = lstr_sql & " [int_despesa], "
        lstr_sql = lstr_sql & " [int_forma_pagamento], "
        lstr_sql = lstr_sql & " [dt_vencimento], "
        lstr_sql = lstr_sql & " [num_valor], "
        lstr_sql = lstr_sql & " [int_parcela], "
        lstr_sql = lstr_sql & " [int_total_parcelas], "
        lstr_sql = lstr_sql & " [str_descricao], "
        lstr_sql = lstr_sql & " [str_documento], "
        lstr_sql = lstr_sql & " [str_chave], "
        lstr_sql = lstr_sql & " [str_codigo_barras], "
        lstr_sql = lstr_sql & " [str_observacoes] "
        lstr_sql = lstr_sql & " ) values ( "
        lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lstr_baixa_automatica) & "', "
        lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_conta_baixa_automatica) & ", "
        lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_despesa) & ", "
        lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_forma_pagamento) & ", "
        lstr_sql = lstr_sql & " '" & Format$(ldt_data_vencimento, pcst_formato_data_sql) & "', "
        lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(ldbl_valor) & ", "
        lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_parcela) & ", "
        lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_total_parcelas) & ", "
        lstr_sql = lstr_sql & " '" & lstr_descricao & "', "
        lstr_sql = lstr_sql & " '" & lstr_documento & "', "
        lstr_sql = lstr_sql & " '" & lstr_chave & "',"
        lstr_sql = lstr_sql & " '" & lstr_codigo_barras & "',"
        lstr_sql = lstr_sql & " '" & lstr_observacoes & "' "
        lstr_sql = lstr_sql & " ) "
        'executa o comando sql e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_lancar_conta_pagar, lstr_sql, "frm_lancar_contas_pagar", "lfct_lancar_conta_pagar")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_lfct_lancar_conta_pagar
        End If
    End If
    lfct_lancar_conta_pagar = True
fim_lfct_lancar_conta_pagar:
    'destrói os objetos
    Set lobj_lancar_conta_pagar = Nothing
    Exit Function
erro_lfct_lancar_conta_pagar:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "lfct_lancar_conta_pagar"
    GoTo fim_lfct_lancar_conta_pagar
End Function

Private Function lfct_lancar_conta_movimentacao() As Boolean
    On Error GoTo erro_lfct_lancar_conta_movimentacao
    Dim lobj_movimentacao As Object
    Dim lobj_conta As Object
    Dim lobj_baixar_conta_pagar As Object
    Dim lobj_atualizar_saldo_conta As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
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
    Dim ldbl_valor_movimentacao As Double
    Dim llng_parcela As Long
    Dim llng_total_parcelas As Long
    Dim ldbl_saldo_atual As Double
    Dim ldbl_limite_negativo As Double
    'atribui os valores dos campos às variáveis
    llng_conta = CLng(cbo_conta_baixa_imediata.ItemData(cbo_conta_baixa_imediata.ListIndex))
    llng_despesa = CLng(cbo_tipo_despesa.ItemData(cbo_tipo_despesa.ListIndex))
    llng_receita = 0
    llng_forma_pagamento = CLng(cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex))
    lstr_data_vencimento = Format$(dtp_vencimento.Value, pcst_formato_data_sql)
    lstr_data_pagamento = IIf(p_usuario.bln_data_vencimento_baixa_imediata, lstr_data_vencimento, Format$(Date, pcst_formato_data_sql))
    'lançamento único
    llng_parcela = 1
    llng_total_parcelas = 1
    '
    lstr_descricao = pfct_tratar_texto_sql(txt_descricao.Text)
    ldbl_valor_movimentacao = CDbl(txt_valor.Text)
    lstr_documento = pfct_tratar_texto_sql(txt_documento.Text)
    lstr_codigo_barras = pfct_tratar_texto_sql(txt_codigo_barras.Text)
    lstr_observacoes = pfct_tratar_texto_sql(txt_observacoes.Text)
    ' --- verificamos se não estamos duplicando a conta a pagar - início --- '
    If ((p_usuario.bln_lancamentos_duplicados) And (lstr_documento <> Empty)) Then 'só verificamos se houver também um número de documento
        'monta o comando sql
        lstr_sql = "select * from [tb_movimentacao] where ([dt_pagamento] = '" & lstr_data_pagamento & "' or [dt_vencimento] = '" & lstr_data_vencimento + "') and [chr_tipo] = 'S' and [num_valor] = " & pfct_tratar_numero_sql(ldbl_valor_movimentacao) & " and [str_documento] = '" & lstr_documento & "'"
        'executa o comando sql e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_movimentacao, lstr_sql, "frm_lancar_contas_pagar", "lfct_lancar_conta_movimentacao")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_lfct_lancar_conta_movimentacao
        End If
        llng_registros = lobj_movimentacao.Count
        If (llng_registros > 0) Then
            'exibe mensagem ao usuário e desvia a execução para o bloco fim
            MsgBox "Este lançamento não pode ser baixado pois já existe um registro equivalente na movimentação.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            GoTo fim_lfct_lancar_conta_movimentacao
        End If
    End If
    ' --- verificamos se não estamos duplicando a conta a pagar - fim --- '
    'monta o comando sql
    lstr_sql = "select * from [tb_contas] where [int_codigo] = " & pfct_tratar_numero_sql(llng_conta)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_conta, lstr_sql, "frm_lancar_contas_pagar", "lfct_lancar_conta_movimentacao")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lfct_lancar_conta_movimentacao
    End If
    llng_registros = lobj_conta.Count
    If (llng_registros > 0) Then
        ldbl_saldo_atual = CDbl(lobj_conta(1)("num_saldo"))
        ldbl_limite_negativo = CDbl(lobj_conta(1)("num_limite_negativo"))
        If ((ldbl_saldo_atual - ldbl_valor_movimentacao) >= (ldbl_limite_negativo * -1)) Then
            '-- atualiza o saldo da conta --'
            'monta o comando sql
            lstr_sql = ""
            lstr_sql = lstr_sql & " update "
            lstr_sql = lstr_sql & " [tb_contas] "
            lstr_sql = lstr_sql & " set "
            lstr_sql = lstr_sql & " [num_saldo] = " & pfct_tratar_numero_sql((ldbl_saldo_atual - ldbl_valor_movimentacao))
            lstr_sql = lstr_sql & " where "
            lstr_sql = lstr_sql & " [int_codigo] = " & pfct_tratar_numero_sql(llng_conta)
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_atualizar_saldo_conta, lstr_sql, "frm_lancar_contas_pagar", "lfct_lancar_conta_movimentacao")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_lfct_lancar_conta_movimentacao
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
            lstr_sql = lstr_sql & " 'S', "
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
            If (Not pfct_executar_comando_sql(lobj_baixar_conta_pagar, lstr_sql, "frm_lancar_contas_pagar", "lfct_lancar_conta_movimentacao")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_lfct_lancar_conta_movimentacao
            End If
        Else
            MsgBox "Conta com saldo insuficiente." & vbCrLf & "Verique o limite negativo e tente novamente.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            GoTo fim_lfct_lancar_conta_movimentacao
        End If
    End If
    'retorna true
    lfct_lancar_conta_movimentacao = True
fim_lfct_lancar_conta_movimentacao:
    'destrói os objetos
    Set lobj_movimentacao = Nothing
    Set lobj_conta = Nothing
    Set lobj_baixar_conta_pagar = Nothing
    Set lobj_atualizar_saldo_conta = Nothing
    Exit Function
erro_lfct_lancar_conta_movimentacao:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "lfct_lancar_conta_movimentacao"
    GoTo fim_lfct_lancar_conta_movimentacao
End Function

Private Sub cbo_conta_baixa_automatica_DropDown()
    On Error GoTo erro_cbo_conta_baixa_automatica_DropDown
    psub_campo_got_focus cbo_conta_baixa_automatica
fim_cbo_conta_baixa_automatica_DropDown:
    Exit Sub
erro_cbo_conta_baixa_automatica_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_conta_baixa_automatica_DropDown"
    GoTo fim_cbo_conta_baixa_automatica_DropDown
End Sub

Private Sub cbo_conta_baixa_automatica_GotFocus()
    On Error GoTo erro_cbo_conta_baixa_automatica_gotFocus
    psub_campo_got_focus cbo_conta_baixa_automatica
fim_cbo_conta_baixa_automatica_gotFocus:
    Exit Sub
erro_cbo_conta_baixa_automatica_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_conta_baixa_automatica_GotFocus"
    GoTo fim_cbo_conta_baixa_automatica_gotFocus
End Sub

Private Sub cbo_conta_baixa_automatica_LostFocus()
    On Error GoTo erro_cbo_conta_baixa_automatica_LostFocus
    psub_campo_lost_focus cbo_conta_baixa_automatica
fim_cbo_conta_baixa_automatica_LostFocus:
    Exit Sub
erro_cbo_conta_baixa_automatica_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_conta_baixa_automatica_LostFocus"
    GoTo fim_cbo_conta_baixa_automatica_LostFocus
End Sub

Private Sub cbo_conta_baixa_imediata_DropDown()
    On Error GoTo erro_cbo_conta_baixa_imediata_DropDown
    psub_campo_got_focus cbo_conta_baixa_imediata
fim_cbo_conta_baixa_imediata_DropDown:
    Exit Sub
erro_cbo_conta_baixa_imediata_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_conta_baixa_imediata_DropDown"
    GoTo fim_cbo_conta_baixa_imediata_DropDown
End Sub

Private Sub cbo_conta_baixa_imediata_GotFocus()
    On Error GoTo erro_cbo_conta_baixa_imediata_GotFocus
    psub_campo_got_focus cbo_conta_baixa_imediata
fim_cbo_conta_baixa_imediata_GotFocus:
    Exit Sub
erro_cbo_conta_baixa_imediata_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_conta_baixa_imediata_GotFocus"
    GoTo fim_cbo_conta_baixa_imediata_GotFocus
End Sub

Private Sub cbo_conta_baixa_imediata_LostFocus()
    On Error GoTo erro_cbo_conta_baixa_imediata_LostFocus
    psub_campo_lost_focus cbo_conta_baixa_imediata
fim_cbo_conta_baixa_imediata_LostFocus:
    Exit Sub
erro_cbo_conta_baixa_imediata_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_conta_baixa_imediata_LostFocus"
    GoTo fim_cbo_conta_baixa_imediata_LostFocus
End Sub

Private Sub cbo_forma_pagamento_DropDown()
    On Error GoTo erro_cbo_forma_pagamento_DropDown
    psub_campo_got_focus cbo_forma_pagamento
fim_cbo_forma_pagamento_DropDown:
    Exit Sub
erro_cbo_forma_pagamento_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_forma_pagamento_DropDown"
    GoTo fim_cbo_forma_pagamento_DropDown
End Sub

Private Sub cbo_forma_pagamento_GotFocus()
    On Error GoTo erro_cbo_forma_pagamento_GotFocus
    psub_campo_got_focus cbo_forma_pagamento
fim_cbo_forma_pagamento_GotFocus:
    Exit Sub
erro_cbo_forma_pagamento_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_forma_pagamento_GotFocus"
    GoTo fim_cbo_forma_pagamento_GotFocus
End Sub

Private Sub cbo_forma_pagamento_LostFocus()
    On Error GoTo erro_cbo_forma_pagamento_LostFocus
    psub_campo_lost_focus cbo_forma_pagamento
fim_cbo_forma_pagamento_LostFocus:
    Exit Sub
erro_cbo_forma_pagamento_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_forma_pagamento_LostFocus"
    GoTo fim_cbo_forma_pagamento_LostFocus
End Sub

Private Sub cbo_tempo_DropDown()
    On Error GoTo erro_cbo_tempo_DropDown
    psub_campo_got_focus cbo_tempo
fim_cbo_tempo_DropDown:
    Exit Sub
erro_cbo_tempo_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_tempo_DropDown"
    GoTo fim_cbo_tempo_DropDown
End Sub

Private Sub cbo_tempo_GotFocus()
    On Error GoTo erro_cbo_tempo_gotFocus
    psub_campo_got_focus cbo_tempo
fim_cbo_tempo_gotFocus:
    Exit Sub
erro_cbo_tempo_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_tempo_GotFocus"
    GoTo fim_cbo_tempo_gotFocus
End Sub

Private Sub cbo_tempo_LostFocus()
    On Error GoTo erro_cbo_tempo_LostFocus
    psub_campo_lost_focus cbo_tempo
fim_cbo_tempo_LostFocus:
    Exit Sub
erro_cbo_tempo_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_tempo_LostFocus"
    GoTo fim_cbo_tempo_LostFocus
End Sub

Private Sub cbo_tipo_despesa_DropDown()
    On Error GoTo erro_cbo_tipo_despesa_DropDown
    psub_campo_got_focus cbo_tipo_despesa
fim_cbo_tipo_despesa_DropDown:
    Exit Sub
erro_cbo_tipo_despesa_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_tipo_despesa_DropDown"
    GoTo fim_cbo_tipo_despesa_DropDown
End Sub

Private Sub cbo_tipo_despesa_GotFocus()
    On Error GoTo erro_cbo_tipo_despesa_gotFocus
    psub_campo_got_focus cbo_tipo_despesa
fim_cbo_tipo_despesa_gotFocus:
    Exit Sub
erro_cbo_tipo_despesa_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_tipo_despesa_GotFocus"
    GoTo fim_cbo_tipo_despesa_gotFocus
End Sub

Private Sub cbo_tipo_despesa_LostFocus()
    On Error GoTo erro_cbo_tipo_despesa_LostFocus
    psub_campo_lost_focus cbo_tipo_despesa
fim_cbo_tipo_despesa_LostFocus:
    Exit Sub
erro_cbo_tipo_despesa_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cbo_tipo_despesa_LostFocus"
    GoTo fim_cbo_tipo_despesa_LostFocus
End Sub

Private Sub chk_baixa_automatica_Click()
    On Error GoTo Erro_chk_baixa_automatica_Click
    If (chk_baixa_automatica.Value = vbChecked) Then
        cbo_conta_baixa_automatica.Enabled = True
        chk_baixar_conta.Value = vbUnchecked
        chk_baixar_conta.Enabled = False
        If (cbo_conta_baixa_imediata.ListCount > 0) Then
            cbo_conta_baixa_imediata.ListIndex = 0
        End If
    Else
        If (cbo_conta_baixa_automatica.ListCount > 0) Then
            cbo_conta_baixa_automatica.ListIndex = 0
        End If
        cbo_conta_baixa_automatica.Enabled = False
        chk_baixar_conta.Enabled = True
    End If
Fim_chk_baixa_automatica_Click:
    Exit Sub
Erro_chk_baixa_automatica_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "chk_baixa_automatica_Click"
    GoTo Fim_chk_baixa_automatica_Click
End Sub

Private Sub chk_baixar_conta_Click()
    On Error GoTo erro_chk_baixar_conta_Click
    If (chk_baixar_conta.Value = vbChecked) Then
        If (p_usuario.bln_data_vencimento_baixa_imediata) Then
            lbl_vencimento.Caption = "&Vencimento/Pagamento:"
        End If
        frm_lancar_contas_pagar.Height = 9000
        fme_baixa_conta.Enabled = True
        chk_baixa_automatica.Value = vbUnchecked
        chk_baixa_automatica.Enabled = False
        chk_multiplos.Enabled = False
        psub_preencher_contas cbo_conta_baixa_imediata
    Else
        If (p_usuario.bln_data_vencimento_baixa_imediata) Then
            lbl_vencimento.Caption = "&Vencimento:"
        End If
        frm_lancar_contas_pagar.Height = 8055
        fme_baixa_conta.Enabled = False
        chk_baixa_automatica.Enabled = True
        chk_multiplos.Enabled = True
        cbo_conta_baixa_imediata.Clear
    End If
fim_chk_baixar_conta_Click:
    Exit Sub
erro_chk_baixar_conta_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "chk_baixar_conta_Click"
    GoTo fim_chk_baixar_conta_Click
End Sub

Private Sub chk_multiplos_Click()
    On Error GoTo erro_chk_multiplos_Click
    If (chk_multiplos.Value = vbChecked) Then
        fme_multiplos_vencimentos.Enabled = True
        'desabilita baixa imediata
        chk_baixar_conta.Value = vbUnchecked
        chk_baixar_conta.Enabled = False
        chk_baixar_conta_Click
        txt_quantidade.SetFocus
    Else
        'limpa os campos
        txt_quantidade.Text = ""
        txt_tempo.Text = ""
        cbo_tempo.ListIndex = 0
        chk_antecipar_pagamento.Value = vbUnchecked
        '
        fme_multiplos_vencimentos.Enabled = False
        chk_baixar_conta.Enabled = True
        dtp_vencimento.SetFocus
    End If
fim_chk_multiplos_Click:
    Exit Sub
erro_chk_multiplos_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "chk_multiplos_Click"
    GoTo fim_chk_multiplos_Click
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cmd_cancelar_Click"
    GoTo fim_cmd_cancelar_Click
End Sub

Private Sub cmd_lancar_Click()
    On Error GoTo erro_cmd_lancar_Click
    Dim lint_resposta As Integer
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_lancar.Enabled) Then
        Exit Sub
    End If
    If (lfct_verificar_campos) Then
        lint_resposta = MsgBox("Confirma o lançamento?", vbYesNo + vbQuestion + vbDefaultButton2, pcst_nome_aplicacao)
        If (lint_resposta = vbYes) Then
            If (frm_lancar_contas_pagar.chk_baixar_conta.Value = vbChecked) Then
                If (lfct_lancar_conta_movimentacao) Then
                    MsgBox "Atenção!" & vbCrLf & "Lançamento de contas a pagar e baixa imediata realizado com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                    Unload Me
                Else
                    GoTo fim_cmd_lancar_Click
                End If
            Else
                If (lfct_lancar_conta_pagar) Then
                    MsgBox "Atenção!" & vbCrLf & "Lançamento de contas a pagar realizado com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                    Unload Me
                Else
                    GoTo fim_cmd_lancar_Click
                End If
            End If
        End If
    End If
fim_cmd_lancar_Click:
    Exit Sub
erro_cmd_lancar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "cmd_lancar_Click"
    GoTo fim_cmd_lancar_Click
End Sub

Private Sub dtp_vencimento_DropDown()
    On Error GoTo erro_dtp_vencimento_DropDown
    psub_campo_got_focus dtp_vencimento
fim_dtp_vencimento_DropDown:
    Exit Sub
erro_dtp_vencimento_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "dtp_vencimento_DropDown"
    GoTo fim_dtp_vencimento_DropDown
End Sub

Private Sub dtp_vencimento_GotFocus()
    On Error GoTo erro_dtp_vencimento_GotFocus
    psub_campo_got_focus dtp_vencimento
fim_dtp_vencimento_GotFocus:
    Exit Sub
erro_dtp_vencimento_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "dtp_vencimento_GotFocus"
    GoTo fim_dtp_vencimento_GotFocus
End Sub

Private Sub dtp_vencimento_LostFocus()
    On Error GoTo erro_dtp_vencimento_LostFocus
    psub_campo_lost_focus dtp_vencimento
fim_dtp_vencimento_LostFocus:
    Exit Sub
erro_dtp_vencimento_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "dtp_vencimento_LostFocus"
    GoTo fim_dtp_vencimento_LostFocus
End Sub

Private Sub Form_Initialize()
    On Error GoTo Erro_Form_Initialize
    InitCommonControls
Fim_Form_Initialize:
    Exit Sub
Erro_Form_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "Form_Initialize"
    GoTo Fim_Form_Initialize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/financeiro_contas_pagar_lancar.htm", 0
        Case vbKeyF2
            cmd_lancar_Click
        Case vbKeyF3
            cmd_cancelar_Click
        Case Else
            GoTo Fim_Form_KeyUp
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
    lsub_ajustar_data
    psub_preencher_contas cbo_conta_baixa_automatica
    psub_preencher_tempo cbo_tempo
    psub_preencher_despesas cbo_tipo_despesa, False
    psub_preencher_formas_pagamento cbo_forma_pagamento
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "Form_Load"
    GoTo fim_Form_Load
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo erro_Form_Unload
    With frm_cadastro_contas_pagar
        .Enabled = True
        .lsub_preencher_combos
        .lsub_ajustar_grade .msf_grade
    End With
fim_Form_Unload:
    Exit Sub
erro_Form_Unload:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "Form_Unload"
    GoTo fim_Form_Unload
End Sub

Private Sub txt_codigo_barras_GotFocus()
    On Error GoTo erro_txt_codigo_barras_GotFocus
    psub_campo_got_focus txt_codigo_barras
fim_txt_codigo_barras_GotFocus:
    Exit Sub
erro_txt_codigo_barras_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_codigo_barras_GotFocus"
    GoTo fim_txt_codigo_barras_GotFocus
End Sub

Private Sub txt_codigo_barras_LostFocus()
    On Error GoTo erro_txt_codigo_barras_LostFocus
    psub_campo_lost_focus txt_codigo_barras
fim_txt_codigo_barras_LostFocus:
    Exit Sub
erro_txt_codigo_barras_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_codigo_barras_LostFocus"
    GoTo fim_txt_codigo_barras_LostFocus
End Sub

Private Sub txt_codigo_barras_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_codigo_barras_validate
    psub_tratar_campo txt_codigo_barras
fim_txt_codigo_barras_validate:
    Exit Sub
erro_txt_codigo_barras_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_codigo_barras_validate"
    GoTo fim_txt_codigo_barras_validate
End Sub

Private Sub txt_descricao_GotFocus()
    On Error GoTo erro_txt_descricao_GotFocus
    psub_campo_got_focus txt_descricao
fim_txt_descricao_GotFocus:
    Exit Sub
erro_txt_descricao_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_descricao_GotFocus"
    GoTo fim_txt_descricao_GotFocus
End Sub

Private Sub txt_descricao_LostFocus()
    On Error GoTo erro_txt_descricao_LostFocus
    psub_campo_lost_focus txt_descricao
fim_txt_descricao_LostFocus:
    Exit Sub
erro_txt_descricao_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_descricao_LostFocus"
    GoTo fim_txt_descricao_LostFocus
End Sub

Private Sub txt_descricao_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_descricao_validate
    psub_tratar_campo txt_descricao
fim_txt_descricao_validate:
    Exit Sub
erro_txt_descricao_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_descricao_validate"
    GoTo fim_txt_descricao_validate
End Sub

Private Sub txt_documento_GotFocus()
    On Error GoTo erro_txt_documento_GotFocus
    psub_campo_got_focus txt_documento
fim_txt_documento_GotFocus:
    Exit Sub
erro_txt_documento_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_documento_GotFocus"
    GoTo fim_txt_documento_GotFocus
End Sub

Private Sub txt_documento_LostFocus()
    On Error GoTo erro_txt_documento_LostFocus
    psub_campo_lost_focus txt_documento
fim_txt_documento_LostFocus:
    Exit Sub
erro_txt_documento_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_documento_LostFocus"
    GoTo fim_txt_documento_LostFocus
End Sub

Private Sub txt_documento_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_documento_validate
    psub_tratar_campo txt_documento
fim_txt_documento_validate:
    Exit Sub
erro_txt_documento_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_documento_validate"
    GoTo fim_txt_documento_validate
End Sub

Private Sub txt_observacoes_GotFocus()
    On Error GoTo erro_txt_observacoes_gotFocus
    psub_campo_got_focus txt_observacoes
fim_txt_observacoes_gotFocus:
    Exit Sub
erro_txt_observacoes_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_observacoes_GotFocus"
    GoTo fim_txt_observacoes_gotFocus
End Sub

Private Sub txt_observacoes_LostFocus()
    On Error GoTo erro_txt_observacoes_LostFocus
    psub_campo_lost_focus txt_observacoes
fim_txt_observacoes_LostFocus:
    Exit Sub
erro_txt_observacoes_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_observacoes_LostFocus"
    GoTo fim_txt_observacoes_LostFocus
End Sub

Private Sub txt_observacoes_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_observacoes_validate
    psub_tratar_campo txt_observacoes
fim_txt_observacoes_validate:
    Exit Sub
erro_txt_observacoes_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_observacoes_validate"
    GoTo fim_txt_observacoes_validate
End Sub

Private Sub txt_quantidade_GotFocus()
    On Error GoTo erro_txt_quantidade_GotFocus
    psub_campo_got_focus txt_quantidade
fim_txt_quantidade_GotFocus:
    Exit Sub
erro_txt_quantidade_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_quantidade_GotFocus"
    GoTo fim_txt_quantidade_GotFocus
End Sub

Private Sub txt_quantidade_LostFocus()
    On Error GoTo erro_txt_quantidade_LostFocus
    psub_campo_lost_focus txt_quantidade
fim_txt_quantidade_LostFocus:
    Exit Sub
erro_txt_quantidade_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_quantidade_LostFocus"
    GoTo fim_txt_quantidade_LostFocus
End Sub

Private Sub txt_quantidade_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_quantidade_Validate
    psub_tratar_campo txt_quantidade
    Cancel = Not pfct_validar_campo(txt_quantidade, tc_inteiro)
fim_txt_quantidade_Validate:
    Exit Sub
erro_txt_quantidade_Validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_quantidade_validate"
    GoTo fim_txt_quantidade_Validate
End Sub

Private Sub txt_tempo_GotFocus()
    On Error GoTo erro_txt_tempo_gotFocus
    psub_campo_got_focus txt_tempo
fim_txt_tempo_gotFocus:
    Exit Sub
erro_txt_tempo_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_tempo_GotFocus"
    GoTo fim_txt_tempo_gotFocus
End Sub

Private Sub txt_tempo_LostFocus()
    On Error GoTo erro_txt_tempo_LostFocus
    psub_campo_lost_focus txt_tempo
fim_txt_tempo_LostFocus:
    Exit Sub
erro_txt_tempo_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_tempo_LostFocus"
    GoTo fim_txt_tempo_LostFocus
End Sub

Private Sub txt_tempo_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_tempo_validate
    psub_tratar_campo txt_tempo
    Cancel = Not pfct_validar_campo(txt_tempo, tc_inteiro)
fim_txt_tempo_validate:
    Exit Sub
erro_txt_tempo_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_tempo_validate"
    GoTo fim_txt_tempo_validate
End Sub

Private Sub txt_valor_GotFocus()
    On Error GoTo erro_txt_valor_GotFocus
    psub_campo_got_focus txt_valor
fim_txt_valor_GotFocus:
    Exit Sub
erro_txt_valor_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_valor_GotFocus"
    GoTo fim_txt_valor_GotFocus
End Sub

Private Sub txt_valor_LostFocus()
    On Error GoTo erro_txt_valor_LostFocus
    psub_campo_lost_focus txt_valor
fim_txt_valor_LostFocus:
    Exit Sub
erro_txt_valor_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_valor_LostFocus"
    GoTo fim_txt_valor_LostFocus
End Sub

Private Sub txt_valor_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_valor_validate
    psub_tratar_campo txt_valor
    Cancel = Not pfct_validar_campo(txt_valor, tc_monetario)
fim_txt_valor_validate:
    Exit Sub
erro_txt_valor_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_lancar_contas_pagar", "txt_valor_validate"
    GoTo fim_txt_valor_validate
End Sub
