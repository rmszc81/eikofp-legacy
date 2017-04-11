VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_movimentacao_geral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimentação Geral"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13110
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
   ScaleHeight     =   7350
   ScaleWidth      =   13110
   Begin VB.CommandButton cmd_iniciar 
      Caption         =   "&Iniciar (F10)"
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmd_detalhes 
      Caption         =   "&Detalhes (F2)"
      Height          =   375
      Left            =   1365
      TabIndex        =   1
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmd_filtrar 
      Caption         =   "&Filtrar (F7)"
      Height          =   375
      Left            =   4005
      TabIndex        =   3
      Top             =   60
      Width           =   1275
   End
   Begin VB.Frame fme_filtros 
      Caption         =   " Filtros "
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   12855
      Begin VB.CheckBox chk_considerar_contas_inativas 
         Caption         =   "&Considerar valores no período de contas inativas"
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   18
         Top             =   1020
         Width           =   4515
      End
      Begin VB.ComboBox cbo_tipo 
         Height          =   315
         Left            =   2580
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   600
         Width           =   1875
      End
      Begin VB.ComboBox cbo_contas 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cbo_ordenar_por 
         Height          =   315
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cbo_ordem 
         Height          =   315
         Left            =   9720
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   1995
      End
      Begin MSComCtl2.DTPicker dtp_de 
         Height          =   315
         Left            =   4560
         TabIndex        =   14
         Top             =   600
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59834369
         CurrentDate     =   39591
      End
      Begin MSComCtl2.DTPicker dtp_ate 
         Height          =   315
         Left            =   5940
         TabIndex        =   15
         Top             =   600
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59834369
         CurrentDate     =   39591
      End
      Begin VB.Label lbl_tipo 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo:"
         Height          =   195
         Left            =   2580
         TabIndex        =   7
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lbl_selecionar_conta 
         AutoSize        =   -1  'True
         Caption         =   "&Selecione a conta:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label lbl_ate 
         AutoSize        =   -1  'True
         Caption         =   "até:"
         Height          =   195
         Left            =   5940
         TabIndex        =   9
         Top             =   300
         Width           =   300
      End
      Begin VB.Label lbl_periodo 
         AutoSize        =   -1  'True
         Caption         =   "Exibir de:"
         Height          =   195
         Left            =   4560
         TabIndex        =   8
         Top             =   300
         Width           =   675
      End
      Begin VB.Label lbl_ordenar_por 
         AutoSize        =   -1  'True
         Caption         =   "&Ordenar por:"
         Height          =   195
         Left            =   7320
         TabIndex        =   10
         Top             =   300
         Width           =   945
      End
      Begin VB.Label lbl_ordem 
         AutoSize        =   -1  'True
         Caption         =   "&Em ordem:"
         Height          =   195
         Left            =   9720
         TabIndex        =   11
         Top             =   300
         Width           =   765
      End
   End
   Begin VB.CommandButton cmd_cancelar_movimentacao 
      Caption         =   "&Cancelar (F3)"
      Height          =   375
      Left            =   2685
      TabIndex        =   2
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmd_fechar 
      Caption         =   "&Fechar (F8)"
      Height          =   375
      Left            =   5325
      TabIndex        =   4
      Top             =   60
      Width           =   1275
   End
   Begin MSComctlLib.StatusBar stb_status 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   20
      Top             =   7065
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23072
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid msf_grade 
      Height          =   4890
      Left            =   60
      TabIndex        =   19
      Top             =   2100
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   8625
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Menu mnu_msf_grade 
      Caption         =   "&Grade"
      Visible         =   0   'False
      Begin VB.Menu mnu_msf_grade_copiar 
         Caption         =   "&Copiar conteúdo"
      End
      Begin VB.Menu mnu_msf_grade_exportar 
         Caption         =   "&Exportar para arquivo..."
      End
   End
End
Attribute VB_Name = "frm_movimentacao_geral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mcst_qtde_colunas_grade As Byte = 11
Private Const mcst_todas_contas As Long = 9999

Private Enum enm_movimentacao_geral
    col_vencimento = 0
    col_pagamento = 1
    col_codconta = 2
    col_conta = 3
    col_tipo = 4
    col_receita_despesa = 5
    col_valor = 6
    col_parcela = 7
    col_descricao = 8
    col_forma_pagamento = 9
    col_documento = 10
End Enum

Private Enum enm_ordenar_por
    col_selecione = 0
    col_lancamento = 1
    col_vencimento = 2
    col_pagamento = 3
    col_conta = 4
    col_tipo = 5
    col_receita_despesa = 6
    col_valor = 7
    col_descricao = 8
    col_forma_pagamento = 9
    col_documento = 10
End Enum

Private Enum enm_tipo_movimentacao
    op_selecione_tipo = 0
    op_entrada = 1
    op_saida = 2
    op_ambas = 3
End Enum

Private Enum enm_status_entradas
    pnl_entradas = 1
End Enum

Private Enum enm_status_saidas
    pnl_saidas = 1
End Enum

Private Enum enm_status
    pnl_entradas = 1
    pnl_saidas = 2
    pnl_total = 3
End Enum

Private mobj_frm_detalhes As Object

Private Sub lsub_preencher_combos()
    On Error GoTo erro_lsub_preencher_combos
    With cbo_tipo
        .Clear
        .AddItem "- Selecione o tipo -", enm_tipo_movimentacao.op_selecione_tipo
        .AddItem "- Entrada", enm_tipo_movimentacao.op_entrada
        .AddItem "- Saída", enm_tipo_movimentacao.op_saida
        .AddItem "- Ambas", enm_tipo_movimentacao.op_ambas
        .ListIndex = 0
    End With
    With cbo_ordenar_por
        .Clear
        .AddItem "- Selecione o campo -", enm_ordenar_por.col_selecione
        .AddItem "- Lançamento", enm_ordenar_por.col_lancamento
        .AddItem "- Vencimento", enm_ordenar_por.col_vencimento
        .AddItem "- Pagamento", enm_ordenar_por.col_pagamento
        .AddItem "- Conta", enm_ordenar_por.col_conta
        .AddItem "- Tipo", enm_ordenar_por.col_tipo
        .AddItem "- Receita/Despesa", enm_ordenar_por.col_receita_despesa
        .AddItem "- Valor", enm_ordenar_por.col_valor
        .AddItem "- Descrição", enm_ordenar_por.col_descricao
        .AddItem "- Forma de Pagamento", enm_ordenar_por.col_forma_pagamento
        .AddItem "- Documento", enm_ordenar_por.col_documento
        .ListIndex = enm_ordenar_por.col_selecione
    End With
    psub_preencher_ordem cbo_ordem
fim_lsub_preencher_combos:
    Exit Sub
erro_lsub_preencher_combos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "lsub_preencher_combos"
    GoTo fim_lsub_preencher_combos
    Resume 0
End Sub

Private Sub lsub_ajustar_status(ByRef pstb_status As StatusBar)
    On Error GoTo erro_lsub_ajustar_status
    pstb_status.Panels.Clear
    pstb_status.Panels.Add
    pstb_status.Panels.Item(1).AutoSize = sbrSpring
    pstb_status.Panels.Item(1).Text = ""
fim_lsub_ajustar_status:
    Exit Sub
erro_lsub_ajustar_status:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "lsub_ajustar_status"
    GoTo fim_lsub_ajustar_status
End Sub

Private Sub lsub_ajustar_grade(ByRef pgrd_grade As MSFlexGrid)
    On Error GoTo erro_lsub_ajustar_grade
    Dim llng_contador       As Long
    Dim llng_total_linhas   As Long
    
    'atribui valor
    llng_total_linhas = pgrd_grade.Rows
    
    'limpa a propriedade rowdata de todas as linhas
    pgrd_grade.Redraw = False
    For llng_contador = 1 To (llng_total_linhas - 1)
        pgrd_grade.Row = llng_contador
        pgrd_grade.RowData(llng_contador) = 0
    Next
    pgrd_grade.Redraw = True
    
    'reajusta a grade
    With pgrd_grade
        .Clear
        .Cols = mcst_qtde_colunas_grade
        .Rows = 2
        .ColWidth(enm_movimentacao_geral.col_vencimento) = 1050
        .ColWidth(enm_movimentacao_geral.col_pagamento) = 1050
        .ColWidth(enm_movimentacao_geral.col_codconta) = 0
        .ColWidth(enm_movimentacao_geral.col_conta) = 2150
        .ColWidth(enm_movimentacao_geral.col_tipo) = 850
        .ColWidth(enm_movimentacao_geral.col_receita_despesa) = 2130
        .ColWidth(enm_movimentacao_geral.col_valor) = 1020
        .ColWidth(enm_movimentacao_geral.col_parcela) = 1020
        .ColWidth(enm_movimentacao_geral.col_descricao) = 3350
        .ColWidth(enm_movimentacao_geral.col_forma_pagamento) = 1850
        .ColWidth(enm_movimentacao_geral.col_documento) = 1800
        .TextMatrix(0, enm_movimentacao_geral.col_vencimento) = " Vencimento"
        .TextMatrix(0, enm_movimentacao_geral.col_pagamento) = " Pagamento"
        .TextMatrix(0, enm_movimentacao_geral.col_codconta) = " Código da Conta"
        .TextMatrix(0, enm_movimentacao_geral.col_conta) = " Conta"
        .TextMatrix(0, enm_movimentacao_geral.col_tipo) = " Tipo"
        .TextMatrix(0, enm_movimentacao_geral.col_receita_despesa) = " Receita/Despesa"
        .TextMatrix(0, enm_movimentacao_geral.col_valor) = " Valor"
        .TextMatrix(0, enm_movimentacao_geral.col_parcela) = " Parcela"
        .TextMatrix(0, enm_movimentacao_geral.col_descricao) = " Descrição"
        .TextMatrix(0, enm_movimentacao_geral.col_forma_pagamento) = " Forma de Pagamento"
        .TextMatrix(0, enm_movimentacao_geral.col_documento) = " Documento"
    End With
fim_lsub_ajustar_grade:
    Exit Sub
erro_lsub_ajustar_grade:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "lsub_ajustar_grade"
    GoTo fim_lsub_ajustar_grade
End Sub

Private Sub lsub_preencher_grade(ByVal pstr_conta As String, _
                                 ByVal pstr_tipo As String, _
                                 ByVal pdt_data_de As Date, _
                                 ByVal pdt_data_ate As Date, _
                                 ByVal pstr_ordenar_por As String, _
                                 ByVal pstr_ordem As String, _
                                 ByVal pbln_contas_inativas As Boolean)
    On Error GoTo erro_lsub_preencher_grade
    'declaração de variáveis
    Dim lobj_movimentacao As Object
    Dim lstr_entradas As String
    Dim lstr_saidas As String
    Dim lstr_total As String
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    Dim llng_quantidade_entrada As Long
    Dim llng_quantidade_saida As Long
    Dim llng_quantidade_total As Long
    Dim ldbl_valor_entrada As Double
    Dim ldbl_valor_saida As Double
    Dim ldbl_valor_total As Double
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_codigo], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_conta], "
    lstr_sql = lstr_sql & " [tb_contas].[str_descricao] as [str_descricao_conta], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_despesa], "
    lstr_sql = lstr_sql & " [tb_despesas].[str_descricao] as [str_descricao_despesa], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_receita], "
    lstr_sql = lstr_sql & " [tb_receitas].[str_descricao] as [str_descricao_receita], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_forma_pagamento], "
    lstr_sql = lstr_sql & " [tb_formas_pagamento].[str_descricao] as [str_descricao_forma_pagamento], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[chr_tipo], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[dt_pagamento], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[dt_vencimento], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[num_valor], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_parcela], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[int_total_parcelas], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[str_descricao], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[str_documento], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[str_codigo_barras], "
    lstr_sql = lstr_sql & " [tb_movimentacao].[str_observacoes] "
    lstr_sql = lstr_sql & " from "
    lstr_sql = lstr_sql & " [tb_movimentacao]"
    lstr_sql = lstr_sql & " inner join [tb_contas] on [tb_contas].[int_codigo] = [tb_movimentacao].[int_conta] "
    lstr_sql = lstr_sql & " left outer join [tb_formas_pagamento] on [tb_formas_pagamento].[int_codigo] = [tb_movimentacao].[int_forma_pagamento] "
    lstr_sql = lstr_sql & " left outer join [tb_despesas] on [tb_despesas].[int_codigo] = [tb_movimentacao].[int_despesa] "
    lstr_sql = lstr_sql & " left outer join [tb_receitas] on [tb_receitas].[int_codigo] = [tb_movimentacao].[int_receita] "
    lstr_sql = lstr_sql & " where "
    lstr_sql = lstr_sql & " 1 = 1 "
    If (pstr_conta <> "9999") Then
        lstr_sql = lstr_sql & " and [tb_movimentacao].[int_conta] = " & pstr_conta & " "
    Else
        lstr_sql = lstr_sql & " and [tb_movimentacao].[int_conta] in "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " select "
        lstr_sql = lstr_sql & " [int_codigo] "
        lstr_sql = lstr_sql & " from "
        lstr_sql = lstr_sql & " [tb_contas] "
        lstr_sql = lstr_sql & " where "
        lstr_sql = lstr_sql & " 1 = 1 "
        If (Not pbln_contas_inativas) Then
            lstr_sql = lstr_sql & " and [chr_ativo] = 'S' "
        End If
        lstr_sql = lstr_sql & " ) "
    End If
    If (pstr_tipo <> "") Then
        lstr_sql = lstr_sql & " and [tb_movimentacao].[chr_tipo] = '" & pstr_tipo & "' "
    End If
    lstr_sql = lstr_sql & " and [tb_movimentacao].[dt_pagamento] between '" & Format$(pdt_data_de, pcst_formato_data_sql) & "' "
    lstr_sql = lstr_sql & " and '" & Format$(pdt_data_ate, pcst_formato_data_sql) & "' "
    lstr_sql = lstr_sql & " order by " & pstr_ordenar_por & " " & pstr_ordem & " "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_movimentacao, lstr_sql, "frm_movimentacao_geral", "lsub_preencher_grade ")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_preencher_grade
    End If
    llng_registros = lobj_movimentacao.Count
    'se não houver registros
    If (llng_registros = 0) Then
        'limpa a barra de status
        lsub_ajustar_status stb_status
        'preenche a barra de status
        stb_status.Panels.Add
        stb_status.Panels.Item(enm_status.pnl_entradas).Text = " Não há movimentação com os critérios de filtros selecionados. "
        stb_status.Panels.Item(enm_status.pnl_entradas).AutoSize = sbrSpring
        'exibe mensagem
        MsgBox "Atenção!" & vbCrLf & "Não há movimentação com os critérios de filtros selecionados.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        'desvia ao final do método
        GoTo fim_lsub_preencher_grade
    ElseIf (llng_registros > 0) Then 'se houver registros
        'zera as variáveis de quantidade
        llng_quantidade_entrada = 0
        llng_quantidade_saida = 0
        llng_quantidade_total = llng_registros
        'zera as variáveis de valores
        ldbl_valor_entrada = 0
        ldbl_valor_saida = 0
        ldbl_valor_total = 0
        'limpa as variáveis de texto
        lstr_entradas = ""
        lstr_saidas = ""
        lstr_total = ""
        'preenche a grade
        msf_grade.Redraw = False
        For llng_contador = 1 To llng_registros
            msf_grade.Row = llng_contador
            msf_grade.Col = enm_movimentacao_geral.col_vencimento
            msf_grade.RowData(llng_contador) = lobj_movimentacao(llng_contador)("int_codigo")
            msf_grade.TextMatrix(llng_contador, enm_movimentacao_geral.col_vencimento) = " " & Format$(lobj_movimentacao(llng_contador)("dt_vencimento"), pcst_formato_data)
            msf_grade.TextMatrix(llng_contador, enm_movimentacao_geral.col_pagamento) = " " & Format$(lobj_movimentacao(llng_contador)("dt_pagamento"), pcst_formato_data)
            'formata coluna pagamento
            msf_grade.Col = enm_movimentacao_geral.col_pagamento
            msf_grade.Row = llng_contador
            If (lobj_movimentacao(llng_contador)("dt_pagamento") < lobj_movimentacao(llng_contador)("dt_vencimento")) Then
                msf_grade.CellForeColor = vbBlue
            ElseIf (lobj_movimentacao(llng_contador)("dt_pagamento") > lobj_movimentacao(llng_contador)("dt_vencimento")) Then
                msf_grade.CellForeColor = vbRed
            Else
                msf_grade.CellForeColor = vbWindowText
            End If
            'coluna código da conta
            msf_grade.TextMatrix(llng_contador, enm_movimentacao_geral.col_codconta) = lobj_movimentacao(llng_contador)("int_conta")
            'coluna conta
            msf_grade.TextMatrix(llng_contador, enm_movimentacao_geral.col_conta) = lobj_movimentacao(llng_contador)("str_descricao_conta")
            msf_grade.TextMatrix(llng_contador, enm_movimentacao_geral.col_tipo) = " " & IIf(lobj_movimentacao(llng_contador)("chr_tipo") = "S", "Saída", "Entrada")
            If (lobj_movimentacao(llng_contador)("chr_tipo") = "S") Then
                llng_quantidade_saida = llng_quantidade_saida + 1
                ldbl_valor_saida = ldbl_valor_saida + CDbl(lobj_movimentacao(llng_contador)("num_valor"))
                msf_grade.TextMatrix(llng_contador, enm_movimentacao_geral.col_receita_despesa) = " " & lobj_movimentacao(llng_contador)("str_descricao_despesa")
                msf_grade.Col = enm_movimentacao_geral.col_valor
                msf_grade.Row = llng_contador
                msf_grade.CellForeColor = vbRed
            Else
                llng_quantidade_entrada = llng_quantidade_entrada + 1
                ldbl_valor_entrada = ldbl_valor_entrada + CDbl(lobj_movimentacao(llng_contador)("num_valor"))
                msf_grade.TextMatrix(llng_contador, enm_movimentacao_geral.col_receita_despesa) = " " & lobj_movimentacao(llng_contador)("str_descricao_receita")
                msf_grade.Col = enm_movimentacao_geral.col_valor
                msf_grade.Row = llng_contador
                msf_grade.CellForeColor = vbBlue
            End If
            msf_grade.TextMatrix(llng_contador, enm_movimentacao_geral.col_valor) = " " & Format$(lobj_movimentacao(llng_contador)("num_valor"), pcst_formato_numerico)
            'parcelas
             msf_grade.TextMatrix(llng_contador, enm_movimentacao_geral.col_parcela) = " " & _
                Format$(lobj_movimentacao(llng_contador)("int_parcela"), pcst_formato_numerico_parcela) & "/" & _
                Format$(lobj_movimentacao(llng_contador)("int_total_parcelas"), pcst_formato_numerico_parcela)
            '
            msf_grade.ColAlignment(enm_movimentacao_geral.col_valor) = flexAlignRightCenter
            msf_grade.TextMatrix(llng_contador, enm_movimentacao_geral.col_descricao) = " " & lobj_movimentacao(llng_contador)("str_descricao")
            msf_grade.TextMatrix(llng_contador, enm_movimentacao_geral.col_forma_pagamento) = " " & lobj_movimentacao(llng_contador)("str_descricao_forma_pagamento")
            msf_grade.TextMatrix(llng_contador, enm_movimentacao_geral.col_documento) = " " & lobj_movimentacao(llng_contador)("str_documento")
            'incrementa uma linha
            If (llng_contador < llng_registros) Then
                msf_grade.Rows = msf_grade.Rows + 1
            End If
        Next
        msf_grade.Redraw = True
        msf_grade.Col = enm_movimentacao_geral.col_vencimento
        msf_grade.Row = 1
        'ajusta a barra de status
        stb_status.Panels.Clear
        'calcula o valor total no período
        ldbl_valor_total = (ldbl_valor_entrada - ldbl_valor_saida)
        'entradas
        If (llng_quantidade_entrada > 0) Then
            lstr_entradas = "Entradas: [" & Format$(llng_quantidade_entrada, "0000") & "] ->" & " " & pfct_retorna_simbolo_moeda() & " " & Format$(ldbl_valor_entrada, pcst_formato_numerico)
        Else
            lstr_entradas = "Não há entradas no período selecionado."
        End If
        'saidas
        If (llng_quantidade_saida > 0) Then
            lstr_saidas = "Saídas: [" & Format$(llng_quantidade_saida, "0000") & "] ->" & " " & pfct_retorna_simbolo_moeda() & " " & Format$(ldbl_valor_saida, pcst_formato_numerico)
        Else
            lstr_saidas = "Não há saídas no período selecionado."
        End If
        'total
        lstr_total = "Total no período: [" & Format$(llng_quantidade_total, "0000") & "] ->" & " " & pfct_retorna_simbolo_moeda() & " " & Format$(ldbl_valor_total, pcst_formato_numerico)
        If (pstr_tipo = "E") Then
            stb_status.Panels.Add
            stb_status.Panels.Item(enm_status_entradas.pnl_entradas).AutoSize = sbrSpring
            stb_status.Panels.Item(enm_status_entradas.pnl_entradas).Text = lstr_entradas
        End If
        If (pstr_tipo = "S") Then
            stb_status.Panels.Add
            stb_status.Panels.Item(enm_status_saidas.pnl_saidas).AutoSize = sbrSpring
            stb_status.Panels.Item(enm_status_saidas.pnl_saidas).Text = lstr_saidas
        End If
        If ((pstr_tipo <> "E") And (pstr_tipo <> "S")) Then
            'entrada
            stb_status.Panels.Add
            stb_status.Panels.Item(enm_status.pnl_entradas).AutoSize = sbrSpring
            stb_status.Panels.Item(enm_status.pnl_entradas).Text = lstr_entradas
            'saida
            stb_status.Panels.Add
            stb_status.Panels.Item(enm_status.pnl_saidas).AutoSize = sbrSpring
            stb_status.Panels.Item(enm_status.pnl_saidas).Text = lstr_saidas
            'total
            stb_status.Panels.Add
            stb_status.Panels.Item(enm_status.pnl_total).AutoSize = sbrSpring
            stb_status.Panels.Item(enm_status.pnl_total).Text = lstr_total
        End If
    End If
fim_lsub_preencher_grade:
    Set lobj_movimentacao = Nothing
    Exit Sub
erro_lsub_preencher_grade:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "lsub_preencher_grade"
    GoTo fim_lsub_preencher_grade
End Sub

Private Sub cbo_contas_Click()
    On Error GoTo erro_cbo_contas_Click
    If (cbo_contas.ItemData(cbo_contas.ListIndex) = mcst_todas_contas) Then
        chk_considerar_contas_inativas.Enabled = True
    Else
        chk_considerar_contas_inativas.Value = vbUnchecked
        chk_considerar_contas_inativas.Enabled = False
    End If
fim_cbo_contas_Click:
    Exit Sub
erro_cbo_contas_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cbo_contas_Click"
    GoTo fim_cbo_contas_Click
End Sub

Private Sub cbo_contas_DropDown()
    On Error GoTo erro_cbo_contas_DropDown
    psub_campo_got_focus cbo_contas
fim_cbo_contas_DropDown:
    Exit Sub
erro_cbo_contas_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cbo_contas_DropDown"
    GoTo fim_cbo_contas_DropDown
End Sub

Private Sub cbo_contas_GotFocus()
    On Error GoTo erro_cbo_contas_GotFocus
    psub_campo_got_focus cbo_contas
fim_cbo_contas_GotFocus:
    Exit Sub
erro_cbo_contas_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cbo_contas_GotFocus"
    GoTo fim_cbo_contas_GotFocus
End Sub

Private Sub cbo_contas_LostFocus()
    On Error GoTo erro_cbo_contas_LostFocus
    psub_campo_lost_focus cbo_contas
fim_cbo_contas_LostFocus:
    Exit Sub
erro_cbo_contas_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cbo_contas_LostFocus"
    GoTo fim_cbo_contas_LostFocus
End Sub

Private Sub cbo_ordem_DropDown()
    On Error GoTo erro_cbo_ordem_DropDown
    psub_campo_got_focus cbo_ordem
fim_cbo_ordem_DropDown:
    Exit Sub
erro_cbo_ordem_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cbo_ordem_DropDown"
    GoTo fim_cbo_ordem_DropDown
End Sub

Private Sub cbo_ordem_GotFocus()
    On Error GoTo erro_cbo_ordem_GotFocus
    psub_campo_got_focus cbo_ordem
fim_cbo_ordem_GotFocus:
    Exit Sub
erro_cbo_ordem_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cbo_ordem_GotFocus"
    GoTo fim_cbo_ordem_GotFocus
End Sub

Private Sub cbo_ordem_LostFocus()
    On Error GoTo erro_cbo_ordem_LostFocus
    psub_campo_lost_focus cbo_ordem
fim_cbo_ordem_LostFocus:
    Exit Sub
erro_cbo_ordem_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cbo_ordem_LostFocus"
    GoTo fim_cbo_ordem_LostFocus
End Sub

Private Sub cbo_ordenar_por_DropDown()
    On Error GoTo erro_cbo_ordenar_por_DropDown
    psub_campo_got_focus cbo_ordenar_por
fim_cbo_ordenar_por_DropDown:
    Exit Sub
erro_cbo_ordenar_por_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cbo_ordenar_por_DropDown"
    GoTo fim_cbo_ordenar_por_DropDown
End Sub

Private Sub cbo_ordenar_por_GotFocus()
    On Error GoTo erro_cbo_ordenar_por_GotFocus
    psub_campo_got_focus cbo_ordenar_por
fim_cbo_ordenar_por_GotFocus:
    Exit Sub
erro_cbo_ordenar_por_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cbo_ordenar_por_GotFocus"
    GoTo fim_cbo_ordenar_por_GotFocus
End Sub

Private Sub cbo_ordenar_por_LostFocus()
    On Error GoTo erro_cbo_ordenar_por_LostFocus
    psub_campo_lost_focus cbo_ordenar_por
fim_cbo_ordenar_por_LostFocus:
    Exit Sub
erro_cbo_ordenar_por_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cbo_ordenar_por_LostFocus"
    GoTo fim_cbo_ordenar_por_LostFocus
End Sub

Private Sub cbo_tipo_DropDown()
    On Error GoTo erro_cbo_tipo_DropDown
    psub_campo_got_focus cbo_tipo
fim_cbo_tipo_DropDown:
    Exit Sub
erro_cbo_tipo_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cbo_tipo_DropDown"
    GoTo fim_cbo_tipo_DropDown
End Sub

Private Sub cbo_tipo_GotFocus()
    On Error GoTo erro_cbo_tipo_GotFocus
    psub_campo_got_focus cbo_tipo
fim_cbo_tipo_GotFocus:
    Exit Sub
erro_cbo_tipo_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cbo_tipo_GotFocus"
    GoTo fim_cbo_tipo_GotFocus
End Sub

Private Sub cbo_tipo_LostFocus()
    On Error GoTo erro_cbo_tipo_LostFocus
    psub_campo_lost_focus cbo_tipo
fim_cbo_tipo_LostFocus:
    Exit Sub
erro_cbo_tipo_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cbo_tipo_LostFocus"
    GoTo fim_cbo_tipo_LostFocus
End Sub

Private Sub cmd_cancelar_movimentacao_Click()
    On Error GoTo erro_cmd_cancelar_movimentacao_Click
    Dim lobj_movimentacao As Object
    Dim lobj_conta As Object
    Dim lstr_sql As String
    Dim lstr_tipo_movimentacao As String * 1
    Dim llng_registros As Long
    Dim llng_codigo_conta As Long
    Dim llng_codigo_item As Long
    Dim ldbl_saldo_conta As Double
    Dim ldbl_limite_negativo As Double
    Dim ldbl_valor_movimentacao As Double
    Dim lint_resposta As Integer
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_cancelar_movimentacao.Enabled) Then
        Exit Sub
    End If
    'atribui às variáveis os valores selecionados
    llng_codigo_conta = msf_grade.TextMatrix(msf_grade.Row, enm_movimentacao_geral.col_codconta)
    llng_codigo_item = msf_grade.RowData(msf_grade.Row)
    'se não houver item selecionado na grade
    If (llng_codigo_item = 0) Then
        MsgBox "Selecione um item na grade.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        GoTo fim_cmd_cancelar_movimentacao_Click
    End If
    'exibe mensagem ao usuário
    lint_resposta = MsgBox("Cancelar esta movimentação atualizará o saldo da conta ao qual está vinculada." & vbCrLf & "Deseja continuar?", vbYesNo + vbQuestion + vbDefaultButton2, pcst_nome_aplicacao)
    If (lint_resposta = vbYes) Then
        'monta o comando sql
        lstr_sql = " select * from [tb_movimentacao] where [int_codigo] = " & pfct_tratar_numero_sql(llng_codigo_item)
        'executa o comando sql e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_movimentacao, lstr_sql, "frm_movimentacao_geral", "cmd_cancelar_movimentacao_Click")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_cmd_cancelar_movimentacao_Click
        End If
        llng_registros = lobj_movimentacao.Count
        If (llng_registros > 0) Then
            lstr_tipo_movimentacao = lobj_movimentacao(1)("chr_tipo")
            ldbl_valor_movimentacao = CDbl(lobj_movimentacao(1)("num_valor"))
            If (lstr_tipo_movimentacao = "E") Then
                '-- se for entrada, verifica se é possível atualizar o saldo da conta sem estourar o limite negativo --'
                'monta o comando sql
                lstr_sql = "select * from [tb_contas] where [int_codigo] = " & pfct_tratar_numero_sql(llng_codigo_conta)
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_conta, lstr_sql, "frm_movimentacao_geral", "cmd_cancelar_movimentacao_Click")) Then
                    MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                    GoTo fim_cmd_cancelar_movimentacao_Click
                End If
                If (llng_registros > 0) Then
                    ldbl_saldo_conta = CDbl(lobj_conta(1)("num_saldo"))
                    ldbl_limite_negativo = CDbl(lobj_conta(1)("num_limite_negativo"))
                    If ((ldbl_saldo_conta - ldbl_valor_movimentacao) >= (ldbl_limite_negativo * -1)) Then
                        '-- se houver saldo suficiente na conta atualiza o saldo --'
                        'monta o comando sql
                        lstr_sql = ""
                        lstr_sql = lstr_sql & " update "
                        lstr_sql = lstr_sql & " [tb_contas] "
                        lstr_sql = lstr_sql & " set "
                        lstr_sql = lstr_sql & " [num_saldo] = " & pfct_tratar_numero_sql((ldbl_saldo_conta - ldbl_valor_movimentacao))
                        lstr_sql = lstr_sql & " where "
                        lstr_sql = lstr_sql & " [int_codigo] = " & pfct_tratar_numero_sql(llng_codigo_conta)
                        'executa o comando sql e devolve o objeto
                        If (Not pfct_executar_comando_sql(lobj_conta, lstr_sql, "frm_movimentacao_geral", "cmd_cancelar_movimentacao_Click")) Then
                            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                            GoTo fim_cmd_cancelar_movimentacao_Click
                        End If
                        '-- em seguida, cancela a movimentação selecionada --'
                        'monta o comando sql
                        lstr_sql = "delete from [tb_movimentacao] where [int_codigo] = " & pfct_tratar_numero_sql(llng_codigo_item)
                        'executa o comando sql e devolve o objeto
                        If (Not pfct_executar_comando_sql(lobj_movimentacao, lstr_sql, "frm_movimentacao_geral", "cmd_cancelar_movimentacao_Click")) Then
                            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                            GoTo fim_cmd_cancelar_movimentacao_Click
                        End If
                        'exibe mensagem ao usuário, atualiza a grade e desvia para o final do método
                        MsgBox "Movimentação cancelada com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                        cmd_filtrar_Click
                        GoTo fim_cmd_cancelar_movimentacao_Click
                    Else
                        'exibe mensagem ao usuário e desvia ao final do método
                        MsgBox "Conta com saldo insuficiente." & vbCrLf & "Verique o limite negativo e tente novamente.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                        GoTo fim_cmd_cancelar_movimentacao_Click
                    End If
                End If
            Else
                '-- se for saída, atualiza o saldo da conta imediatamente --'
                'monta o comando sql
                lstr_sql = "select * from [tb_contas] where [int_codigo] = " & pfct_tratar_numero_sql(llng_codigo_conta)
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_conta, lstr_sql, "frm_movimentacao_geral", "cmd_cancelar_movimentacao_Click")) Then
                    MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                    GoTo fim_cmd_cancelar_movimentacao_Click
                End If
                 If (llng_registros > 0) Then
                    ldbl_saldo_conta = CDbl(lobj_conta(1)("num_saldo"))
                    'monta o comando sql
                    lstr_sql = ""
                    lstr_sql = lstr_sql & " update "
                    lstr_sql = lstr_sql & " [tb_contas] "
                    lstr_sql = lstr_sql & " set "
                    lstr_sql = lstr_sql & " [num_saldo] = " & pfct_tratar_numero_sql((ldbl_saldo_conta + ldbl_valor_movimentacao))
                    lstr_sql = lstr_sql & " where "
                    lstr_sql = lstr_sql & " [int_codigo] = " & pfct_tratar_numero_sql(llng_codigo_conta)
                    'executa o comando sql e devolve o objeto
                    If (Not pfct_executar_comando_sql(lobj_conta, lstr_sql, "frm_movimentacao_geral", "cmd_cancelar_movimentacao_Click")) Then
                        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                        GoTo fim_cmd_cancelar_movimentacao_Click
                    End If
                    '-- em seguida, cancela a movimentação selecionada --'
                    'monta o comando sql
                    lstr_sql = "delete from [tb_movimentacao] where [int_codigo] = " & pfct_tratar_numero_sql(llng_codigo_item)
                    'executa o comando sql e devolve o objeto
                    If (Not pfct_executar_comando_sql(lobj_movimentacao, lstr_sql, "frm_movimentacao_geral", "cmd_cancelar_movimentacao_Click")) Then
                        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                        GoTo fim_cmd_cancelar_movimentacao_Click
                    End If
                    'exibe mensagem ao usuário, atualiza a grade e desvia para o final do método
                    MsgBox "Movimentação cancelada com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                    cmd_filtrar_Click
                    GoTo fim_cmd_cancelar_movimentacao_Click
                 End If
            End If
        End If
    End If
fim_cmd_cancelar_movimentacao_Click:
    'destrói os objetos
    Set lobj_conta = Nothing
    Set lobj_movimentacao = Nothing
    Exit Sub
erro_cmd_cancelar_movimentacao_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cmd_cancelar_movimentacao_Click"
    GoTo fim_cmd_cancelar_movimentacao_Click
End Sub

Private Sub cmd_detalhes_Click()
    On Error GoTo erro_cmd_detalhes_Click
    
    Dim llng_codigo_item    As Long
    
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_detalhes.Enabled) Then
        Exit Sub
    End If
    
    llng_codigo_item = msf_grade.RowData(msf_grade.Row)
    If (llng_codigo_item = 0) Then
        MsgBox "Selecione um item na grade.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        GoTo fim_cmd_detalhes_Click
    Else
    
        'força a destruição do objeto
        Set mobj_frm_detalhes = Nothing
        'cria a nova instância do objeto
        Set mobj_frm_detalhes = New frm_movimentacao_geral_detalhes
        
        'ajusta as propriedades
        Set mobj_frm_detalhes.mobj_form_anterior = Me
        mobj_frm_detalhes.mlng_codigo = llng_codigo_item
        mobj_frm_detalhes.Show
    
    End If

fim_cmd_detalhes_Click:
    Exit Sub
erro_cmd_detalhes_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cmd_detalhes_Click"
    GoTo fim_cmd_detalhes_Click
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cmd_fechar_Click"
    GoTo fim_cmd_fechar_Click
End Sub

Public Sub cmd_filtrar_Click()
    On Error GoTo erro_cmd_filtrar_Click
    'declaração de variáveis
    Dim lstr_tipo As String
    Dim lstr_ordenar_por As String
    Dim lstr_ordem As String
    Dim lstr_conta As String
    Dim lbln_contas_inativas As Boolean
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_filtrar.Enabled) Then
        Exit Sub
    End If
    'valida o combo contas
    If (cbo_contas.ListIndex = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [conta] é obrigatório.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_contas.SetFocus
        GoTo fim_cmd_filtrar_Click
    End If
    'valida o combo tipo
    If (cbo_tipo.ListIndex = enm_tipo_movimentacao.op_selecione_tipo) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [tipo] é obrigatório.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_tipo.SetFocus
        GoTo fim_cmd_filtrar_Click
    End If
    'valida as datas
    If (dtp_de.Value > dtp_ate.Value) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [data inicial] deve ser menor que data final.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        psub_ajustar_combos_data dtp_de, dtp_ate
        dtp_de.SetFocus
        GoTo fim_cmd_filtrar_Click
    End If
    'valida o combo ordenar por
    If (cbo_ordenar_por.ListIndex = enm_ordenar_por.col_selecione) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [ordenar por] é obrigatório.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_ordenar_por.SetFocus
        GoTo fim_cmd_filtrar_Click
    End If
    'valida o combo ordem
    If (cbo_ordem.ListIndex = enm_em_ordem.op_selecione) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [ordem] é obrigatório.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_ordem.SetFocus
        GoTo fim_cmd_filtrar_Click
    End If
    'valida novamente o combo ordenar por
    If (cbo_ordenar_por.ListIndex = enm_ordenar_por.col_receita_despesa) Then
        If (cbo_tipo.ListIndex = enm_tipo_movimentacao.op_ambas) Then
            MsgBox "Atenção!" & vbCrLf & "É obrigatória a seleção do tipo [entrada] ou [saída] ao classificar por [receita/despesa].", vbOKOnly + vbInformation, pcst_nome_aplicacao
            cbo_tipo.SetFocus
            GoTo fim_cmd_filtrar_Click
        End If
    End If
    Select Case cbo_tipo.ListIndex
        Case enm_tipo_movimentacao.op_entrada
            lstr_tipo = "E"
        Case enm_tipo_movimentacao.op_saida
            lstr_tipo = "S"
    End Select
    Select Case cbo_ordenar_por.ListIndex
        Case enm_ordenar_por.col_lancamento
            lstr_ordenar_por = "[tb_movimentacao].[int_codigo]"
        Case enm_ordenar_por.col_vencimento
            lstr_ordenar_por = "[tb_movimentacao].[dt_vencimento]"
        Case enm_ordenar_por.col_pagamento
            lstr_ordenar_por = "[tb_movimentacao].[dt_pagamento]"
        Case enm_ordenar_por.col_conta
            lstr_ordenar_por = "[str_descricao_conta]"
        Case enm_ordenar_por.col_tipo
            lstr_ordenar_por = "[tb_movimentacao].[chr_tipo]"
        Case enm_ordenar_por.col_receita_despesa
            If (lstr_tipo = "E") Then
                lstr_ordenar_por = "[str_descricao_receita]"
            ElseIf (lstr_tipo = "S") Then
                lstr_ordenar_por = "[str_descricao_despesa]"
            End If
        Case enm_ordenar_por.col_valor
            lstr_ordenar_por = "[tb_movimentacao].[num_valor]"
        Case enm_ordenar_por.col_descricao
            lstr_ordenar_por = "[tb_movimentacao].[str_descricao]"
        Case enm_ordenar_por.col_forma_pagamento
            lstr_ordenar_por = "[str_descricao_forma_pagamento]"
        Case enm_ordenar_por.col_documento
            lstr_ordenar_por = "[tb_movimentacao].[str_documento]"
    End Select
    Select Case cbo_ordem.ListIndex
        Case enm_em_ordem.op_crescente
            lstr_ordem = "asc"
        Case enm_em_ordem.op_decrescente
            lstr_ordem = "desc"
    End Select
    'incluir contas inativas?
    lbln_contas_inativas = IIf(chk_considerar_contas_inativas.Value = vbChecked, True, False)
    'conta
    lstr_conta = CStr(cbo_contas.ItemData(cbo_contas.ListIndex))
    'configura a grade
    lsub_ajustar_grade msf_grade
    'executa a consultae preenche a grade
    lsub_preencher_grade lstr_conta, lstr_tipo, dtp_de.Value, dtp_ate.Value, lstr_ordenar_por, lstr_ordem, lbln_contas_inativas
fim_cmd_filtrar_Click:
    Exit Sub
erro_cmd_filtrar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cmd_filtrar_Click"
    GoTo fim_cmd_filtrar_Click
End Sub

Private Sub cmd_iniciar_Click()
    On Error GoTo erro_cmd_iniciar_Click
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_iniciar.Enabled) Then
        Exit Sub
    End If
    Form_Load
fim_cmd_iniciar_Click:
    Exit Sub
erro_cmd_iniciar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "cmd_iniciar_Click"
    GoTo fim_cmd_iniciar_Click
End Sub

Private Sub dtp_ate_DropDown()
    On Error GoTo erro_dtp_ate_DropDown
    psub_campo_got_focus dtp_ate
fim_dtp_ate_DropDown:
    Exit Sub
erro_dtp_ate_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "dtp_ate_DropDown"
    GoTo fim_dtp_ate_DropDown
End Sub

Private Sub dtp_ate_GotFocus()
    On Error GoTo erro_dtp_ate_GotFocus
    psub_campo_got_focus dtp_ate
fim_dtp_ate_GotFocus:
    Exit Sub
erro_dtp_ate_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "dtp_ate_GotFocus"
    GoTo fim_dtp_ate_GotFocus
End Sub

Private Sub dtp_ate_LostFocus()
    On Error GoTo erro_dtp_ate_LostFocus
    psub_campo_lost_focus dtp_ate
fim_dtp_ate_LostFocus:
    Exit Sub
erro_dtp_ate_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "dtp_ate_LostFocus"
    GoTo fim_dtp_ate_LostFocus
End Sub

Private Sub dtp_de_DropDown()
    On Error GoTo erro_dtp_de_DropDown
    psub_campo_got_focus dtp_de
fim_dtp_de_DropDown:
    Exit Sub
erro_dtp_de_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "dtp_de_DropDown"
    GoTo fim_dtp_de_DropDown
End Sub

Private Sub dtp_de_GotFocus()
    On Error GoTo erro_dtp_de_GotFocus
    psub_campo_got_focus dtp_de
fim_dtp_de_GotFocus:
    Exit Sub
erro_dtp_de_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "dtp_de_GotFocus"
    GoTo fim_dtp_de_GotFocus
End Sub

Private Sub dtp_de_LostFocus()
    On Error GoTo erro_dtp_de_LostFocus
    psub_campo_lost_focus dtp_de
fim_dtp_de_LostFocus:
    Exit Sub
erro_dtp_de_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "dtp_de_LostFocus"
    GoTo fim_dtp_de_LostFocus
End Sub

Private Sub Form_Initialize()
    On Error GoTo Erro_Form_Initialize
    InitCommonControls
Fim_Form_Initialize:
    Exit Sub
Erro_Form_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "Form_Initialize"
    GoTo Fim_Form_Initialize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/movimentacao_geral.htm", 0
        Case vbKeyF2
            cmd_detalhes_Click
        Case vbKeyF3
            cmd_cancelar_movimentacao_Click
        Case vbKeyF7
            cmd_filtrar_Click
        Case vbKeyF8
            cmd_fechar_Click
        Case vbKeyF10
            cmd_iniciar_Click
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
    psub_ajustar_combos_data dtp_de, dtp_ate
    lsub_preencher_combos
    psub_preencher_contas cbo_contas
    'customização combo contas'
    If (cbo_contas.ListCount > 1) Then
        cbo_contas.AddItem "- Todas as Contas -"
        cbo_contas.ItemData(cbo_contas.NewIndex) = mcst_todas_contas
    End If
    lsub_ajustar_grade msf_grade
    lsub_ajustar_status stb_status
    'desmarca o checkbox
    chk_considerar_contas_inativas.Value = vbUnchecked
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "Form_Load"
    GoTo fim_Form_Load
End Sub

Private Sub Form_Terminate()
    On Error GoTo erro_Form_Terminate
    
    'destrói objetos
    Set mobj_frm_detalhes = Nothing

fim_Form_Terminate:
    Exit Sub
erro_Form_Terminate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "Form_Terminate"
    GoTo fim_Form_Terminate
End Sub

Private Sub msf_grade_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_msf_grade_MouseUp
    If (Button = 2) Then 'botão direito do mouse
        PopupMenu mnu_msf_grade 'exibimos o popup
    End If
fim_msf_grade_MouseUp:
    Exit Sub
erro_msf_grade_MouseUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "msf_grade_MouseUp"
    GoTo fim_msf_grade_MouseUp
End Sub

Private Sub mnu_msf_grade_copiar_Click()
    On Error GoTo erro_mnu_msf_grade_copiar_Click
    pfct_copiar_conteudo_grade msf_grade
fim_mnu_msf_grade_copiar_Click:
    Exit Sub
erro_mnu_msf_grade_copiar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "mnu_msf_grade_copiar_Click"
    GoTo fim_mnu_msf_grade_copiar_Click
End Sub

Private Sub mnu_msf_grade_exportar_Click()
    On Error GoTo erro_mnu_msf_grade_exportar_Click
    pfct_exportar_conteudo_grade msf_grade, "movimentacao_geral"
fim_mnu_msf_grade_exportar_Click:
    Exit Sub
erro_mnu_msf_grade_exportar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_geral", "mnu_msf_grade_exportar_Click"
    GoTo fim_mnu_msf_grade_exportar_Click
End Sub

