VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_movimentacao_por_receitas_despesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimentação por Receitas/Despesas"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
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
   ScaleHeight     =   6690
   ScaleWidth      =   9015
   Begin VB.CommandButton cmd_detalhes 
      Caption         =   "&Detalhes (F2)"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmd_iniciar 
      Caption         =   "&Iniciar (F10)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1275
   End
   Begin VB.Frame fme_filtros 
      Caption         =   " Filtros "
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   540
      Width           =   8775
      Begin VB.ComboBox cbo_receitas_despesas 
         Height          =   315
         ItemData        =   "frm_movimentacao_por_receitas_despesas.frx":0000
         Left            =   2400
         List            =   "frm_movimentacao_por_receitas_despesas.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   6255
      End
      Begin VB.ComboBox cbo_ordem 
         Height          =   315
         ItemData        =   "frm_movimentacao_por_receitas_despesas.frx":0004
         Left            =   6600
         List            =   "frm_movimentacao_por_receitas_despesas.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox cbo_ordenar_por 
         Height          =   315
         ItemData        =   "frm_movimentacao_por_receitas_despesas.frx":0008
         Left            =   4560
         List            =   "frm_movimentacao_por_receitas_despesas.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox cbo_tipo 
         Height          =   315
         ItemData        =   "frm_movimentacao_por_receitas_despesas.frx":000C
         Left            =   240
         List            =   "frm_movimentacao_por_receitas_despesas.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtp_de 
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   39591
      End
      Begin MSComCtl2.DTPicker dtp_ate 
         Height          =   315
         Left            =   2400
         TabIndex        =   14
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   39591
      End
      Begin VB.Label lbl_receitas_despesas 
         AutoSize        =   -1  'True
         Caption         =   "&Receita/Despesa:"
         Height          =   195
         Left            =   2400
         TabIndex        =   6
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label lbl_ordem 
         AutoSize        =   -1  'True
         Caption         =   "&Em ordem:"
         Height          =   195
         Left            =   6600
         TabIndex        =   12
         Top             =   1020
         Width           =   765
      End
      Begin VB.Label lbl_ordenar_por 
         AutoSize        =   -1  'True
         Caption         =   "&Ordenar por:"
         Height          =   195
         Left            =   4560
         TabIndex        =   11
         Top             =   1020
         Width           =   945
      End
      Begin VB.Label lbl_selecionar_tipo 
         AutoSize        =   -1  'True
         Caption         =   "&Selecione o tipo:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label lbl_periodo 
         AutoSize        =   -1  'True
         Caption         =   "&Período:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1020
         Width           =   600
      End
      Begin VB.Label lbl_ate 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
         Height          =   195
         Left            =   2400
         TabIndex        =   10
         Top             =   1020
         Width           =   315
      End
   End
   Begin VB.CommandButton cmd_filtrar 
      Caption         =   "&Filtrar (F7)"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmd_fechar 
      Caption         =   "&Fechar (F8)"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   60
      Width           =   1275
   End
   Begin MSComctlLib.StatusBar stb_status 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   18
      Top             =   6405
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15849
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid msf_grade 
      Height          =   3870
      Left            =   120
      TabIndex        =   17
      Top             =   2460
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6826
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
Attribute VB_Name = "frm_movimentacao_por_receitas_despesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enm_movimentacao
    col_vencimento = 0
    col_pagamento = 1
    col_descricao = 2
    col_parcela = 3
    col_valor = 4
    col_diferenca = 5
End Enum

Private Enum enm_status
    pnl_totais = 1
    pnl_valor_medio = 2
End Enum

Private mdbl_valor_total As Double
Private mlng_quantidade As Long

Private mobj_frm_detalhes As Object

'variáveis privadas
Private mobj_frm_anterior As Object
Private mstr_tipo_movimentacao As String * 1
Private mint_codigo_registro As Integer
Private mdt_data_de As Date
Private mdt_data_ate As Date

'propriedades
Public Property Let dt_data_ate(ByVal vData As Date)
    mdt_data_ate = vData
End Property

Public Property Get dt_data_ate() As Date
    dt_data_ate = mdt_data_ate
End Property

Public Property Let dt_data_de(ByVal vData As Date)
    mdt_data_de = vData
End Property

Public Property Get dt_data_de() As Date
    dt_data_de = mdt_data_de
End Property

Public Property Let int_codigo_registro(ByVal vData As Integer)
    mint_codigo_registro = vData
End Property

Public Property Get int_codigo_registro() As Integer
    int_codigo_registro = mint_codigo_registro
End Property

Public Property Let str_tipo_movimentacao(ByVal vData As String)
    mstr_tipo_movimentacao = vData
End Property

Public Property Get str_tipo_movimentacao() As String
    str_tipo_movimentacao = mstr_tipo_movimentacao
End Property

Public Property Set obj_frm_anterior(ByVal vData As Object)
    Set mobj_frm_anterior = vData
End Property

Public Property Get obj_frm_anterior() As Object
    Set obj_frm_anterior = mobj_frm_anterior
End Property


Private Sub lsub_preencher_combos()
    On Error GoTo erro_lsub_preencher_combos
    With cbo_tipo
        .Clear
        .AddItem "- Selecione -", 0
        .AddItem "- Receitas", 1
        .AddItem "- Despesas", 2
        .ListIndex = 0
    End With
    With cbo_ordenar_por
        .Clear
        .AddItem "- Selecione o campo -", 0
        .AddItem "- Lançamento", 1
        .AddItem "- Vencimento", 2
        .AddItem "- Pagamento", 3
        .AddItem "- Descrição", 4
        .AddItem "- Valor", 5
        .ListIndex = 0
    End With
    psub_preencher_ordem cbo_ordem
fim_lsub_preencher_combos:
    Exit Sub
erro_lsub_preencher_combos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "lsub_preencher_combos"
    GoTo fim_lsub_preencher_combos
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
        .Cols = 6
        .Rows = 2
        .ColWidth(enm_movimentacao.col_vencimento) = 1110
        .ColWidth(enm_movimentacao.col_pagamento) = 1110
        .ColWidth(enm_movimentacao.col_descricao) = 3350
        .ColWidth(enm_movimentacao.col_parcela) = 800
        .ColWidth(enm_movimentacao.col_valor) = 1020
        .ColWidth(enm_movimentacao.col_diferenca) = 1020
        .TextMatrix(0, enm_movimentacao.col_vencimento) = " Vencimento"
        .TextMatrix(0, enm_movimentacao.col_pagamento) = " Pagamento"
        .TextMatrix(0, enm_movimentacao.col_descricao) = " Descrição"
        .TextMatrix(0, enm_movimentacao.col_parcela) = " Parcela"
        .TextMatrix(0, enm_movimentacao.col_valor) = " Valor"
        .TextMatrix(0, enm_movimentacao.col_diferenca) = " Diferença"
    End With
fim_lsub_ajustar_grade:
    Exit Sub
erro_lsub_ajustar_grade:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "lsub_ajustar_grade"
    GoTo fim_lsub_ajustar_grade
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "lsub_ajustar_status"
    GoTo fim_lsub_ajustar_status
End Sub

Private Sub lsub_preencher_barra_status(ByVal plng_quantidade As Long, ByVal pdbl_valor As Double)
    On Error GoTo erro_lsub_preencher_barra_status
    With stb_status.Panels
        'limpa a status bar
        .Clear
        If (plng_quantidade > 0) Then
            .Add 'totais
            .Item(enm_status.pnl_totais).AutoSize = sbrSpring
            .Item(enm_status.pnl_totais).Text = "Total no período: [" & Format$(plng_quantidade, "0000") & "] ->" & " " & pfct_retorna_simbolo_moeda() & " " & Format$(pdbl_valor, pcst_formato_numerico)
            .Add 'valor médio
            .Item(enm_status.pnl_valor_medio).AutoSize = sbrSpring
            .Item(enm_status.pnl_valor_medio).Text = "Valor médio: " & pfct_retorna_simbolo_moeda() & " " & Format$(pdbl_valor / plng_quantidade, pcst_formato_numerico)
        Else
            .Add 'não há movimentações
            .Item(enm_status.pnl_totais).AutoSize = sbrSpring
            .Item(enm_status.pnl_totais).Text = "Não há movimentaçao no período selecionado."
        End If
    End With
fim_lsub_preencher_barra_status:
    Exit Sub
erro_lsub_preencher_barra_status:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "lsub_preencher_barra_status"
    GoTo fim_lsub_preencher_barra_status
End Sub

Private Sub lsub_preencher_grade(ByVal pstr_tipo As String, _
                                 ByVal pstr_receita_despesa As String, _
                                 ByVal pstr_data_de As String, _
                                 ByVal pstr_data_ate As String, _
                                 ByVal pstr_ordenar_por As String, _
                                 ByVal pstr_ordem As String)
    On Error GoTo erro_lsub_preencher_grade
    'declaração de variáveis
    Dim lobj_movimentacao As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    Dim ldbl_valor_anterior As Double
    Dim ldbl_valor_atual As Double
    Dim ldbl_valor_diferenca As Double
    Dim lstr_sinal As String
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select * from [tb_movimentacao] "
    lstr_sql = lstr_sql & " where 1=1 "
    lstr_sql = lstr_sql & "     and [chr_tipo] = '" & UCase$(pstr_tipo) & "' "
    If (pstr_tipo = "E") Then
        lstr_sql = lstr_sql & "     and [int_receita] = " & pstr_receita_despesa & " "
    ElseIf (pstr_tipo = "S") Then
        lstr_sql = lstr_sql & "     and [int_despesa] = " & pstr_receita_despesa & " "
    End If
    lstr_sql = lstr_sql & "     and [dt_pagamento] between '" & pstr_data_de & "' and '" & pstr_data_ate & "' "
    lstr_sql = lstr_sql & " order by " & pstr_ordenar_por & " " & pstr_ordem
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_movimentacao, lstr_sql, "frm_movimentacao_por_receitas_despesas", "lsub_preencher_grade")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_preencher_grade
    End If
    llng_registros = lobj_movimentacao.Count
    'se houver registros na tabela
    If (llng_registros > 0) Then
        'zera as variáveis
        mdbl_valor_total = 0
        ldbl_valor_atual = 0
        ldbl_valor_anterior = 0
        ldbl_valor_diferenca = 0
        'atualiza a quantidade
        mlng_quantidade = llng_registros
        'desabilita a atualização da grade
        msf_grade.Redraw = False
        For llng_contador = 1 To llng_registros
            'dados da linha
            msf_grade.Row = llng_contador
            msf_grade.Col = enm_movimentacao.col_vencimento
            msf_grade.RowData(llng_contador) = lobj_movimentacao(llng_contador)("int_codigo")
            'continua
            msf_grade.TextMatrix(llng_contador, enm_movimentacao.col_vencimento) = " " & Format$(lobj_movimentacao(llng_contador)("dt_vencimento"), pcst_formato_data)
            msf_grade.TextMatrix(llng_contador, enm_movimentacao.col_pagamento) = " " & Format$(lobj_movimentacao(llng_contador)("dt_pagamento"), pcst_formato_data)
            msf_grade.TextMatrix(llng_contador, enm_movimentacao.col_descricao) = " " & lobj_movimentacao(llng_contador)("str_descricao")
            'ini parcelas
             msf_grade.TextMatrix(llng_contador, enm_movimentacao.col_parcela) = " " & _
                Format$(lobj_movimentacao(llng_contador)("int_parcela"), pcst_formato_numerico_parcela) & "/" & _
                Format$(lobj_movimentacao(llng_contador)("int_total_parcelas"), pcst_formato_numerico_parcela)
            'fim parcelas
            msf_grade.ColAlignment(enm_movimentacao.col_valor) = flexAlignRightCenter
            msf_grade.TextMatrix(llng_contador, enm_movimentacao.col_valor) = " " & Format$(lobj_movimentacao(llng_contador)("num_valor"), pcst_formato_numerico)
            
            '--- coluna diferença ---'
            
            'atribui valor atual
            ldbl_valor_atual = lobj_movimentacao(llng_contador)("num_valor")
            
            'se o contador for maior que 1
            If (llng_contador > 1) Then
                'calcula a diferença
                ldbl_valor_diferenca = ldbl_valor_atual - ldbl_valor_anterior
            End If
            
            'atribui valor anterior
            ldbl_valor_anterior = lobj_movimentacao(llng_contador)("num_valor")
            
            'se a diferença for diferente de 0
            If (ldbl_valor_diferenca <> 0) Then
                If (Abs(ldbl_valor_diferenca) = ldbl_valor_diferenca) Then
                    lstr_sinal = "+"
                Else
                    lstr_sinal = "-"
                End If
            Else
                lstr_sinal = ""
            End If
                        
            'alinha a coluna à direita
            msf_grade.ColAlignment(enm_movimentacao.col_diferenca) = flexAlignRightCenter
            
            'atribui texto à célula
            msf_grade.TextMatrix(llng_contador, enm_movimentacao.col_diferenca) = " " & lstr_sinal & Format$(Abs(ldbl_valor_diferenca), pcst_formato_numerico)
            
            '--- coluna diferença ---'
                        
            'se a data de pagamento for menor que a data de vencimento, conta foi paga com atraso
            If (lobj_movimentacao(llng_contador)("dt_pagamento") > lobj_movimentacao(llng_contador)("dt_vencimento")) Then
                'cor da fonte da linha em vermelho
                psub_ajustar_cor_linha_grade msf_grade, llng_contador, vbRed
            End If
            'se a data de pagamento for maior que a data de vencimento, conta foi paga adiantada
            If (lobj_movimentacao(llng_contador)("dt_pagamento") < lobj_movimentacao(llng_contador)("dt_vencimento")) Then
                'cor da fonte da linha em azul
                psub_ajustar_cor_linha_grade msf_grade, llng_contador, vbBlue
            End If
            'se a data de pagamento for igual a data de vencimento, conta foi paga na data
            If (lobj_movimentacao(llng_contador)("dt_pagamento") = lobj_movimentacao(llng_contador)("dt_vencimento")) Then
                'cor da fonte da linha em preto
                psub_ajustar_cor_linha_grade msf_grade, llng_contador, vbWindowText
            End If
            'se ainda houver registros
            If (llng_contador < llng_registros) Then
                'adiciona mais uma linha
                msf_grade.Rows = msf_grade.Rows + 1
            End If
            'alimenta a variável com o valor total
            mdbl_valor_total = mdbl_valor_total + CDbl(lobj_movimentacao(llng_contador)("num_valor"))
        Next
        msf_grade.Col = enm_movimentacao.col_pagamento
        msf_grade.Row = 1
        'reabilita a atualização da grade
        msf_grade.Redraw = True
    Else
        'zera o total
        mdbl_valor_total = 0
        'zera a quantidade
        mlng_quantidade = 0
        'exibe mensagem
        MsgBox "Atenção!" & vbCrLf & "Não há movimentaçao no período selecionado.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        'desvia ao final do método
        GoTo fim_lsub_preencher_grade
    End If
fim_lsub_preencher_grade:
    'destrói os objetos
    Set lobj_movimentacao = Nothing
    Exit Sub
erro_lsub_preencher_grade:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "lsub_preencher_grade"
    GoTo fim_lsub_preencher_grade
End Sub

Private Sub cbo_tipo_Click()
    On Error GoTo erro_cbo_tipo_Click
    Select Case cbo_tipo.ListIndex
        Case 0 'selecione
            lbl_receitas_despesas.Caption = "&Receita/Despesa:"
            lbl_receitas_despesas.Enabled = False
            cbo_receitas_despesas.Enabled = False
            cbo_receitas_despesas.Clear
        Case 1 'receitas
            lbl_receitas_despesas.Caption = "&Receitas:"
            lbl_receitas_despesas.Enabled = True
            cbo_receitas_despesas.Enabled = True
            psub_preencher_receitas cbo_receitas_despesas, False
        Case 2 'despesas
            lbl_receitas_despesas.Caption = "&Despesas:"
            lbl_receitas_despesas.Enabled = True
            cbo_receitas_despesas.Enabled = True
            psub_preencher_despesas cbo_receitas_despesas, False
    End Select
fim_cbo_tipo_Click:
    Exit Sub
erro_cbo_tipo_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cbo_tipo_Click"
    GoTo fim_cbo_tipo_Click
End Sub

Private Sub cbo_tipo_DropDown()
    On Error GoTo erro_cbo_tipo_DropDown
    psub_campo_got_focus cbo_tipo
fim_cbo_tipo_DropDown:
    Exit Sub
erro_cbo_tipo_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cbo_tipo_DropDown"
    GoTo fim_cbo_tipo_DropDown
End Sub

Private Sub cbo_tipo_GotFocus()
    On Error GoTo erro_cbo_tipo_GotFocus
    psub_campo_got_focus cbo_tipo
fim_cbo_tipo_GotFocus:
    Exit Sub
erro_cbo_tipo_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cbo_tipo_GotFocus"
    GoTo fim_cbo_tipo_GotFocus
End Sub

Private Sub cbo_tipo_LostFocus()
    On Error GoTo erro_cbo_tipo_LostFocus
    psub_campo_lost_focus cbo_tipo
fim_cbo_tipo_LostFocus:
    Exit Sub
erro_cbo_tipo_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cbo_tipo_LostFocus"
    GoTo fim_cbo_tipo_LostFocus
End Sub

Private Sub cbo_receitas_despesas_DropDown()
    On Error GoTo erro_cbo_receitas_despesas_DropDown
    psub_campo_got_focus cbo_receitas_despesas
fim_cbo_receitas_despesas_DropDown:
    Exit Sub
erro_cbo_receitas_despesas_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cbo_receitas_despesas_DropDown"
    GoTo fim_cbo_receitas_despesas_DropDown
End Sub

Private Sub cbo_receitas_despesas_GotFocus()
    On Error GoTo erro_cbo_receitas_despesas_GotFocus
    psub_campo_got_focus cbo_receitas_despesas
fim_cbo_receitas_despesas_GotFocus:
    Exit Sub
erro_cbo_receitas_despesas_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cbo_receitas_despesas_GotFocus"
    GoTo fim_cbo_receitas_despesas_GotFocus
End Sub

Private Sub cbo_receitas_despesas_LostFocus()
    On Error GoTo erro_cbo_receitas_despesas_LostFocus
    psub_campo_lost_focus cbo_receitas_despesas
fim_cbo_receitas_despesas_LostFocus:
    Exit Sub
erro_cbo_receitas_despesas_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cbo_receitas_despesas_LostFocus"
    GoTo fim_cbo_receitas_despesas_LostFocus
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cmd_detalhes_Click"
    GoTo fim_cmd_detalhes_Click
End Sub

Private Sub dtp_de_DropDown()
    On Error GoTo erro_dtp_de_DropDown
    psub_campo_got_focus dtp_de
fim_dtp_de_DropDown:
    Exit Sub
erro_dtp_de_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "dtp_de_DropDown"
    GoTo fim_dtp_de_DropDown
End Sub

Private Sub dtp_de_GotFocus()
    On Error GoTo erro_dtp_de_GotFocus
    psub_campo_got_focus dtp_de
fim_dtp_de_GotFocus:
    Exit Sub
erro_dtp_de_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "dtp_de_GotFocus"
    GoTo fim_dtp_de_GotFocus
End Sub

Private Sub dtp_de_LostFocus()
    On Error GoTo erro_dtp_de_LostFocus
    psub_campo_lost_focus dtp_de
fim_dtp_de_LostFocus:
    Exit Sub
erro_dtp_de_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "dtp_de_LostFocus"
    GoTo fim_dtp_de_LostFocus
End Sub

Private Sub dtp_ate_DropDown()
    On Error GoTo erro_dtp_ate_DropDown
    psub_campo_got_focus dtp_ate
fim_dtp_ate_DropDown:
    Exit Sub
erro_dtp_ate_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "dtp_ate_DropDown"
    GoTo fim_dtp_ate_DropDown
End Sub

Private Sub dtp_ate_GotFocus()
    On Error GoTo erro_dtp_ate_GotFocus
    psub_campo_got_focus dtp_ate
fim_dtp_ate_GotFocus:
    Exit Sub
erro_dtp_ate_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "dtp_ate_GotFocus"
    GoTo fim_dtp_ate_GotFocus
End Sub

Private Sub dtp_ate_LostFocus()
    On Error GoTo erro_dtp_ate_LostFocus
    psub_campo_lost_focus dtp_ate
fim_dtp_ate_LostFocus:
    Exit Sub
erro_dtp_ate_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "dtp_ate_LostFocus"
    GoTo fim_dtp_ate_LostFocus
End Sub

Private Sub cbo_ordenar_por_DropDown()
    On Error GoTo erro_cbo_ordenar_por_DropDown
    psub_campo_got_focus cbo_ordenar_por
fim_cbo_ordenar_por_DropDown:
    Exit Sub
erro_cbo_ordenar_por_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cbo_ordenar_por_DropDown"
    GoTo fim_cbo_ordenar_por_DropDown
End Sub

Private Sub cbo_ordenar_por_GotFocus()
    On Error GoTo erro_cbo_ordenar_por_GotFocus
    psub_campo_got_focus cbo_ordenar_por
fim_cbo_ordenar_por_GotFocus:
    Exit Sub
erro_cbo_ordenar_por_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cbo_ordenar_por_GotFocus"
    GoTo fim_cbo_ordenar_por_GotFocus
End Sub

Private Sub cbo_ordenar_por_LostFocus()
    On Error GoTo erro_cbo_ordenar_por_LostFocus
    psub_campo_lost_focus cbo_ordenar_por
fim_cbo_ordenar_por_LostFocus:
    Exit Sub
erro_cbo_ordenar_por_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cbo_ordenar_por_LostFocus"
    GoTo fim_cbo_ordenar_por_LostFocus
End Sub

Private Sub cbo_ordem_DropDown()
    On Error GoTo erro_cbo_ordem_DropDown
    psub_campo_got_focus cbo_ordem
fim_cbo_ordem_DropDown:
    Exit Sub
erro_cbo_ordem_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cbo_ordem_DropDown"
    GoTo fim_cbo_ordem_DropDown
End Sub

Private Sub cbo_ordem_GotFocus()
    On Error GoTo erro_cbo_ordem_GotFocus
    psub_campo_got_focus cbo_ordem
fim_cbo_ordem_GotFocus:
    Exit Sub
erro_cbo_ordem_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cbo_ordem_GotFocus"
    GoTo fim_cbo_ordem_GotFocus
End Sub

Private Sub cbo_ordem_LostFocus()
    On Error GoTo erro_cbo_ordem_LostFocus
    psub_campo_lost_focus cbo_ordem
fim_cbo_ordem_LostFocus:
    Exit Sub
erro_cbo_ordem_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cbo_ordem_LostFocus"
    GoTo fim_cbo_ordem_LostFocus
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cmd_fechar_Click"
    GoTo fim_cmd_fechar_Click
End Sub

Public Sub cmd_filtrar_Click()
    On Error GoTo erro_cmd_filtrar_Click
    'declaração de variáveis
    Dim lstr_tipo As String
    Dim lstr_desc_tipo As String
    Dim lstr_receita_despesa As String
    Dim lstr_data_de As String
    Dim lstr_data_ate As String
    Dim lstr_ordenar_por As String
    Dim lstr_ordem As String
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_filtrar.Enabled) Then
        Exit Sub
    End If
    'valida os campos da tela
    If (cbo_tipo.ListIndex = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [tipo] é obrigatório.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_tipo.SetFocus
        GoTo fim_cmd_filtrar_Click
    Else
        If (cbo_tipo.ListIndex = 1) Then
            lstr_desc_tipo = "receita"
        ElseIf (cbo_tipo.ListIndex = 2) Then
            lstr_desc_tipo = "despesa"
        End If
        If (cbo_receitas_despesas.ListIndex = 0) Then
            MsgBox "Atenção!" & vbCrLf & "Campo [" & lstr_desc_tipo & "] é obrigatório.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            cbo_receitas_despesas.SetFocus
            GoTo fim_cmd_filtrar_Click
        End If
    End If
    If (dtp_de.Value > dtp_ate.Value) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [data inicial] deve ser menor que data final.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        psub_ajustar_combos_data dtp_de, dtp_ate
        dtp_de.SetFocus
        GoTo fim_cmd_filtrar_Click
    End If
    If (cbo_ordenar_por.ListIndex = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [ordenar por] é obrigatório.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_ordenar_por.SetFocus
        GoTo fim_cmd_filtrar_Click
    End If
    If (cbo_ordem.ListIndex = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [ordem] é obrigatório.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_ordem.SetFocus
        GoTo fim_cmd_filtrar_Click
    End If
    'atribui os valores
    If (cbo_tipo.ListIndex = 1) Then
        lstr_tipo = "E"
    ElseIf (cbo_tipo.ListIndex = 2) Then
        lstr_tipo = "S"
    End If
    lstr_receita_despesa = CStr(cbo_receitas_despesas.ItemData(cbo_receitas_despesas.ListIndex))
    lstr_data_de = Format$(dtp_de, pcst_formato_data_sql)
    lstr_data_ate = Format$(dtp_ate, pcst_formato_data_sql)
    'ordem de campo da grade
    Select Case cbo_ordenar_por.ListIndex
        Case 1
            lstr_ordenar_por = "[int_codigo]"
        Case 2
            lstr_ordenar_por = "[dt_vencimento]"
        Case 3
            lstr_ordenar_por = "[dt_pagamento]"
        Case 4
            lstr_ordenar_por = "[str_descricao]"
        Case 5
            lstr_ordenar_por = "[num_valor]"
    End Select
    'ordem da grade
    Select Case cbo_ordem.ListIndex
        Case 1
            lstr_ordem = "asc"
        Case 2
            lstr_ordem = "desc"
    End Select
    'reajusta as grades
    lsub_ajustar_grade msf_grade
    'faz a chamada aos métodos de consulta ao banco e preenchimento de grade
    lsub_preencher_grade lstr_tipo, lstr_receita_despesa, lstr_data_de, lstr_data_ate, lstr_ordenar_por, lstr_ordem
    'ajusta a barra de status
    lsub_preencher_barra_status mlng_quantidade, mdbl_valor_total
fim_cmd_filtrar_Click:
    Exit Sub
erro_cmd_filtrar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cmd_filtrar_Click"
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "cmd_iniciar_Click"
    GoTo fim_cmd_iniciar_Click
End Sub

Private Sub Form_Initialize()
    On Error GoTo Erro_Form_Initialize
    InitCommonControls
Fim_Form_Initialize:
    Exit Sub
Erro_Form_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "Form_Initialize"
    GoTo Fim_Form_Initialize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/movimentacao_por_receitas_despesas.htm", 0
        Case vbKeyF2
            cmd_detalhes_Click
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
    'declara variáveis
    Dim llng_contador As Long
    
    'preenche os combos
    lsub_preencher_combos
    
    'ajusta a grade
    lsub_ajustar_grade msf_grade
    
    'ajusta a barra de status
    lsub_ajustar_status stb_status
    
    'zera as variáveis
    mlng_quantidade = 0
    mdbl_valor_total = 0
    
    'verifica como foi chamado o form
    If (Not mobj_frm_anterior Is Nothing) Then
    
        'ajusta as configurações do form
        Me.Left = mobj_frm_anterior.Left + 250
        Me.Top = mobj_frm_anterior.Top + 250
        
        'desabilita o form anterior
        mobj_frm_anterior.Enabled = False
    
        'se o tipo de movimentação for entrada
        'posiciona no item receitas, senão, despesas
        If (mstr_tipo_movimentacao = "E") Then
            cbo_tipo.ListIndex = 1 'receitas
        Else
            cbo_tipo.ListIndex = 2 'despesas
        End If
        
        'combo ordenar por 'lançamento
        cbo_ordenar_por.ListIndex = 1
        
        'combo ordem 'crescente
        cbo_ordem.ListIndex = 1
        
        'combo de receitas/despesas 'posiciona no item desejado
        For llng_contador = 0 To cbo_receitas_despesas.ListCount - 1
            If (cbo_receitas_despesas.ItemData(llng_contador) = mint_codigo_registro) Then
                cbo_receitas_despesas.ListIndex = llng_contador
                Exit For
            End If
        Next
        
        'ajusta as datas
        dtp_de.Value = mdt_data_de
        dtp_ate.Value = mdt_data_ate
        
        'dispara o evento click do botão filtrar
        cmd_filtrar_Click
        
    Else
        'ajusta a data dos combos
        psub_ajustar_combos_data dtp_de, dtp_ate
    End If

fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "Form_Load"
    GoTo fim_Form_Load
End Sub

Private Sub Form_Terminate()
    On Error GoTo erro_Form_Terminate

    'destrói objetos
    Set mobj_frm_detalhes = Nothing
    Set mobj_frm_anterior = Nothing
    
fim_Form_Terminate:
    Exit Sub
erro_Form_Terminate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "Form_Terminate"
    GoTo fim_Form_Terminate
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo erro_Form_Unload
    
    'se houver instância do objeto
    If (Not mobj_frm_anterior Is Nothing) Then
        'reabilita o form anterior
        mobj_frm_anterior.Enabled = True
    End If
    
    'destrói o próprio form
    Set frm_movimentacao_por_receitas_despesas = Nothing

fim_Form_Unload:
    Exit Sub
erro_Form_Unload:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "Form_Unload"
    GoTo fim_Form_Unload
End Sub

Private Sub msf_grade_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_msf_grade_MouseUp
    If (Button = 2) Then 'botão direito do mouse
        PopupMenu mnu_msf_grade 'exibimos o popup
    End If
fim_msf_grade_MouseUp:
    Exit Sub
erro_msf_grade_MouseUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "msf_grade_MouseUp"
    GoTo fim_msf_grade_MouseUp
End Sub

Private Sub mnu_msf_grade_copiar_Click()
    On Error GoTo erro_mnu_msf_grade_copiar_Click
    pfct_copiar_conteudo_grade msf_grade
fim_mnu_msf_grade_copiar_Click:
    Exit Sub
erro_mnu_msf_grade_copiar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "mnu_msf_grade_copiar_Click"
    GoTo fim_mnu_msf_grade_copiar_Click
End Sub

Private Sub mnu_msf_grade_exportar_Click()
    On Error GoTo erro_mnu_msf_grade_exportar_Click
    pfct_exportar_conteudo_grade msf_grade, "movimentacao_por_receitas_despesas"
fim_mnu_msf_grade_exportar_Click:
    Exit Sub
erro_mnu_msf_grade_exportar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_por_receitas_despesas", "mnu_msf_grade_exportar_Click"
    GoTo fim_mnu_msf_grade_exportar_Click
End Sub
