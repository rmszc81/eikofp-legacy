VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_movimentacao_receitas_despesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimentação Receitas x Despesas"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
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
   ScaleHeight     =   6660
   ScaleWidth      =   13110
   Begin VB.CommandButton cmd_detalhes_despesa 
      Caption         =   "&Detalhes da Despesa (F3)"
      Height          =   375
      Left            =   10920
      TabIndex        =   18
      Top             =   2100
      Width           =   2115
   End
   Begin VB.CommandButton cmd_detalhes_receita 
      Caption         =   "&Detalhes da Receita (F2)"
      Height          =   375
      Left            =   4380
      TabIndex        =   16
      Top             =   2100
      Width           =   2115
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
      Height          =   1455
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   12975
      Begin VB.CheckBox chk_considerar_contas_inativas 
         Caption         =   "&Considerar valores no período de contas inativas"
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   1020
         Width           =   4515
      End
      Begin VB.ComboBox cbo_ordem 
         Height          =   315
         ItemData        =   "frm_movimentacao_receitas_despesas.frx":0000
         Left            =   9180
         List            =   "frm_movimentacao_receitas_despesas.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   600
         Width           =   1995
      End
      Begin VB.ComboBox cbo_ordenar_por 
         Height          =   315
         ItemData        =   "frm_movimentacao_receitas_despesas.frx":0004
         Left            =   6780
         List            =   "frm_movimentacao_receitas_despesas.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cbo_contas 
         Height          =   315
         ItemData        =   "frm_movimentacao_receitas_despesas.frx":0008
         Left            =   180
         List            =   "frm_movimentacao_receitas_despesas.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker dtp_de 
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   600
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   39591
      End
      Begin MSComCtl2.DTPicker dtp_ate 
         Height          =   315
         Left            =   4620
         TabIndex        =   11
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   39591
      End
      Begin VB.Label lbl_ordem 
         AutoSize        =   -1  'True
         Caption         =   "&Em ordem:"
         Height          =   195
         Left            =   9180
         TabIndex        =   7
         Top             =   300
         Width           =   765
      End
      Begin VB.Label lbl_ordenar_por 
         AutoSize        =   -1  'True
         Caption         =   "&Ordenar por:"
         Height          =   195
         Left            =   6780
         TabIndex        =   6
         Top             =   300
         Width           =   945
      End
      Begin VB.Label lbl_selecionar_conta 
         AutoSize        =   -1  'True
         Caption         =   "&Selecione a conta:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label lbl_periodo 
         AutoSize        =   -1  'True
         Caption         =   "&Período:"
         Height          =   195
         Left            =   2160
         TabIndex        =   5
         Top             =   300
         Width           =   600
      End
      Begin VB.Label lbl_ate 
         AutoSize        =   -1  'True
         Caption         =   "até:"
         Height          =   195
         Left            =   4260
         TabIndex        =   10
         Top             =   660
         Width           =   300
      End
   End
   Begin VB.CommandButton cmd_filtrar 
      Caption         =   "&Filtrar (F7)"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmd_fechar 
      Caption         =   "&Fechar (F8)"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   60
      Width           =   1275
   End
   Begin MSComctlLib.StatusBar stb_status 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   21
      Top             =   6375
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
   Begin MSFlexGridLib.MSFlexGrid msf_grade_receitas 
      Height          =   3810
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6720
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf_grade_despesas 
      Height          =   3810
      Left            =   6660
      TabIndex        =   20
      Top             =   2520
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6720
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lbl_despesas 
      AutoSize        =   -1  'True
      Caption         =   "&Despesas:"
      Height          =   195
      Left            =   6660
      TabIndex        =   17
      Top             =   2220
      Width           =   750
   End
   Begin VB.Label lbl_receitas 
      AutoSize        =   -1  'True
      Caption         =   "&Receitas:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   2220
      Width           =   675
   End
   Begin VB.Menu mnu_msf_grade_receitas 
      Caption         =   "&Receitas"
      Visible         =   0   'False
      Begin VB.Menu mnu_msf_grade_receitas_copiar 
         Caption         =   "&Copiar conteúdo"
      End
      Begin VB.Menu mnu_msf_grade_receitas_exportar 
         Caption         =   "&Exportar para arquivo..."
      End
   End
   Begin VB.Menu mnu_msf_grade_despesas 
      Caption         =   "&Despesas"
      Visible         =   0   'False
      Begin VB.Menu mnu_msf_grade_despesas_copiar 
         Caption         =   "&Copiar conteúdo"
      End
      Begin VB.Menu mnu_msf_grade_despesas_exportar 
         Caption         =   "&Exportar para arquivo..."
      End
   End
End
Attribute VB_Name = "frm_movimentacao_receitas_despesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enm_receitas
    col_receita = 0
    col_quantidade = 1
    col_valor_medio = 2
    col_valor_total = 3
End Enum

Private Enum enm_despesas
    col_despesa = 0
    col_quantidade = 1
    col_valor_medio = 2
    col_valor_total = 3
End Enum

Private Enum enm_status
    pnl_receitas = 1
    pnl_despesas = 2
    pnl_total = 3
End Enum

Private Const mcst_todas_contas As Long = 9999

Private mdbl_valor_receitas As Double
Private mdbl_valor_despesas As Double

Private Sub lsub_preencher_combos()
    On Error GoTo erro_lsub_preencher_combos
    With cbo_ordenar_por
        .Clear
        .AddItem "- Selecione o campo -", 0
        .AddItem "- Receita/Despesa", 1
        .AddItem "- Quantidade", 2
        .AddItem "- Valor Médio", 3
        .AddItem "- Valor Total", 4
        .ListIndex = 0
    End With
    psub_preencher_ordem cbo_ordem
fim_lsub_preencher_combos:
    Exit Sub
erro_lsub_preencher_combos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "lsub_preencher_combos"
    GoTo fim_lsub_preencher_combos
End Sub

Private Sub lsub_ajustar_grade_receitas(ByRef pgrd_grade As MSFlexGrid)
    On Error GoTo erro_lsub_ajustar_grade_receitas
    Dim llng_contador As Long
    With pgrd_grade
        'limpa a propriedade rowdata de todas as linhas
        .Redraw = False
        For llng_contador = 1 To (.Rows - 1)
            .Row = llng_contador
            .RowData(llng_contador) = 0
        Next
        .Redraw = True
        'reajusta as configurações
        .Clear
        .Cols = 4
        .Rows = 2
        .ColWidth(enm_receitas.col_receita) = 2800
        .ColWidth(enm_receitas.col_quantidade) = 800
        .ColWidth(enm_receitas.col_valor_medio) = 1200
        .ColWidth(enm_receitas.col_valor_total) = 1200
        .TextMatrix(0, enm_receitas.col_receita) = " Receita"
        .TextMatrix(0, enm_receitas.col_quantidade) = " Quantid. "
        .TextMatrix(0, enm_receitas.col_valor_medio) = " Valor Médio"
        .TextMatrix(0, enm_receitas.col_valor_total) = " Valor Total"
        .ColAlignment(enm_receitas.col_quantidade) = flexAlignCenterCenter
        .ColAlignment(enm_receitas.col_valor_medio) = flexAlignRightCenter
        .ColAlignment(enm_receitas.col_valor_total) = flexAlignRightCenter
    End With
fim_lsub_ajustar_grade_receitas:
    Exit Sub
erro_lsub_ajustar_grade_receitas:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "lsub_ajustar_grade_receitas"
    GoTo fim_lsub_ajustar_grade_receitas
End Sub

Private Sub lsub_ajustar_grade_despesas(ByRef pgrd_grade As MSFlexGrid)
    On Error GoTo erro_lsub_ajustar_grade_despesas
    Dim llng_contador As Long
    With pgrd_grade
        'limpa a propriedade rowdata de todas as linhas
        .Redraw = False
        For llng_contador = 1 To (.Rows - 1)
            .Row = llng_contador
            .RowData(llng_contador) = 0
        Next
        .Redraw = True
        'reajusta as configurações
        .Clear
        .Cols = 4
        .Rows = 2
        .ColWidth(enm_despesas.col_despesa) = 2800
        .ColWidth(enm_despesas.col_quantidade) = 800
        .ColWidth(enm_despesas.col_valor_medio) = 1200
        .ColWidth(enm_despesas.col_valor_total) = 1200
        .TextMatrix(0, enm_despesas.col_despesa) = " Despesa"
        .TextMatrix(0, enm_despesas.col_quantidade) = " Quantid. "
        .TextMatrix(0, enm_despesas.col_valor_medio) = " Valor Médio"
        .TextMatrix(0, enm_despesas.col_valor_total) = " Valor Total"
        .ColAlignment(enm_despesas.col_quantidade) = flexAlignCenterCenter
        .ColAlignment(enm_despesas.col_valor_medio) = flexAlignRightCenter
        .ColAlignment(enm_despesas.col_valor_total) = flexAlignRightCenter
    End With
fim_lsub_ajustar_grade_despesas:
    Exit Sub
erro_lsub_ajustar_grade_despesas:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "lsub_ajustar_grade_despesas"
    GoTo fim_lsub_ajustar_grade_despesas
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "lsub_ajustar_status"
    GoTo fim_lsub_ajustar_status
End Sub

Private Sub lsub_preencher_barra_status(ByVal pdbl_valor_receitas As Double, ByVal pdbl_valor_despesas As Double)
    On Error GoTo erro_lsub_preencher_barra_status
    With stb_status.Panels
        'limpa a status bar
        .Clear
        'receitas
        .Add
        .Item(enm_status.pnl_receitas).AutoSize = sbrSpring
        .Item(enm_status.pnl_receitas).Text = " Receitas (" & pfct_retorna_simbolo_moeda() & "): " & Format$(pdbl_valor_receitas, pcst_formato_numerico)
        'despesas
        .Add
        .Item(enm_status.pnl_despesas).AutoSize = sbrSpring
        .Item(enm_status.pnl_despesas).Text = " Despesas (" & pfct_retorna_simbolo_moeda() & "): " & Format$(pdbl_valor_despesas, pcst_formato_numerico)
        'total
        .Add
        .Item(enm_status.pnl_total).AutoSize = sbrSpring
        .Item(enm_status.pnl_total).Text = " Total (" & pfct_retorna_simbolo_moeda() & "): " & Format$((pdbl_valor_receitas - pdbl_valor_despesas), pcst_formato_numerico)
    End With
fim_lsub_preencher_barra_status:
    Exit Sub
erro_lsub_preencher_barra_status:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "lsub_preencher_barra_status"
    GoTo fim_lsub_preencher_barra_status
End Sub

Private Sub lsub_preencher_grade_receitas(ByVal pstr_conta As String, _
                                          ByVal pstr_data_de As String, _
                                          ByVal pstr_data_ate As String, _
                                          ByVal pstr_ordenar_por As String, _
                                          ByVal pstr_ordem As String, _
                                          ByVal pbln_contas_inativas As Boolean)
    On Error GoTo erro_lsub_preencher_grade_receitas
    'declaração de variáveis
    Dim lobj_receitas As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select "
    lstr_sql = lstr_sql & " [tb_receitas].[int_codigo] as [int_codigo], "
    lstr_sql = lstr_sql & " ifnull([tb_receitas].[str_descricao], '') as [descricao], "
    lstr_sql = lstr_sql & " count(*) as [quantidade], "
    lstr_sql = lstr_sql & " (total([tb_movimentacao].[num_valor]) / count(*)) as [valor_medio], "
    lstr_sql = lstr_sql & " total([tb_movimentacao].[num_valor]) as [valor_total] "
    lstr_sql = lstr_sql & " from "
    lstr_sql = lstr_sql & " [tb_movimentacao] "
    lstr_sql = lstr_sql & " inner join [tb_receitas] on [tb_receitas].[int_codigo] = [tb_movimentacao].[int_receita] "
    lstr_sql = lstr_sql & " where "
    lstr_sql = lstr_sql & " [tb_movimentacao].[chr_tipo] = 'E' "
    lstr_sql = lstr_sql & " and [tb_movimentacao].[dt_pagamento] between '" & pstr_data_de & "' and '" & pstr_data_ate & "'"
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
    lstr_sql = lstr_sql & " group by [descricao] "
    lstr_sql = lstr_sql & " order by " & pstr_ordenar_por & " " & pstr_ordem
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_receitas, lstr_sql, "frm_movimentacao_receitas_despesas", "lsub_preencher_grade_receitas ")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_preencher_grade_receitas
    End If
    llng_registros = lobj_receitas.Count
    'se houver registros na tabela
    If (llng_registros > 0) Then
        'zera a variável com o total das receitas
        mdbl_valor_receitas = 0
        'desabilita a atualização da grade
        msf_grade_receitas.Redraw = False
        For llng_contador = 1 To llng_registros
            msf_grade_receitas.RowData(llng_contador) = lobj_receitas(llng_contador)("int_codigo")
            msf_grade_receitas.TextMatrix(llng_contador, enm_receitas.col_receita) = " " & lobj_receitas(llng_contador)("descricao")
            msf_grade_receitas.TextMatrix(llng_contador, enm_receitas.col_quantidade) = " " & Format$(lobj_receitas(llng_contador)("quantidade"), pcst_formato_numerico_padrao)
            msf_grade_receitas.TextMatrix(llng_contador, enm_receitas.col_valor_medio) = " " & Format$(lobj_receitas(llng_contador)("valor_medio"), pcst_formato_numerico)
            msf_grade_receitas.TextMatrix(llng_contador, enm_receitas.col_valor_total) = " " & Format$(lobj_receitas(llng_contador)("valor_total"), pcst_formato_numerico)
            'alimenta a variável com o valor total das receitas
            mdbl_valor_receitas = mdbl_valor_receitas + CDbl(lobj_receitas(llng_contador)("valor_total"))
            'incrementa uma linha
            If (llng_contador < llng_registros) Then
                msf_grade_receitas.Rows = msf_grade_receitas.Rows + 1
            End If
        Next
        msf_grade_receitas.Col = enm_receitas.col_receita
        msf_grade_receitas.Row = 1
        'reabilita a atualização da grade
        msf_grade_receitas.Redraw = True
    End If
fim_lsub_preencher_grade_receitas:
    'destrói os objetos
    Set lobj_receitas = Nothing
    Exit Sub
erro_lsub_preencher_grade_receitas:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "lsub_preencher_grade_receitas"
    GoTo fim_lsub_preencher_grade_receitas
End Sub

Private Sub lsub_preencher_grade_despesas(ByVal pstr_conta As String, _
                                          ByVal pstr_data_de As String, _
                                          ByVal pstr_data_ate As String, _
                                          ByVal pstr_ordenar_por As String, _
                                          ByVal pstr_ordem As String, _
                                          ByVal pbln_contas_inativas As Boolean)
    On Error GoTo erro_lsub_preencher_grade_despesas
    'declaração de variáveis
    Dim lobj_despesas As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select "
    lstr_sql = lstr_sql & " [tb_despesas].[int_codigo] as [int_codigo], "
    lstr_sql = lstr_sql & " ifnull([tb_despesas].[str_descricao], '') as [descricao], "
    lstr_sql = lstr_sql & " count(*) as [quantidade], "
    lstr_sql = lstr_sql & " (total([tb_movimentacao].[num_valor]) / count(*)) as [valor_medio], "
    lstr_sql = lstr_sql & " total([tb_movimentacao].[num_valor]) as [valor_total] "
    lstr_sql = lstr_sql & " from "
    lstr_sql = lstr_sql & " [tb_movimentacao] "
    lstr_sql = lstr_sql & " inner join [tb_despesas] on [tb_despesas].[int_codigo] = [tb_movimentacao].[int_despesa] "
    lstr_sql = lstr_sql & " where "
    lstr_sql = lstr_sql & " [tb_movimentacao].[chr_tipo] = 'S' "
    lstr_sql = lstr_sql & " and [tb_movimentacao].[dt_pagamento] between '" & pstr_data_de & "' and '" & pstr_data_ate & "'"
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
    lstr_sql = lstr_sql & " group by [descricao] "
    lstr_sql = lstr_sql & " order by " & pstr_ordenar_por & " " & pstr_ordem
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_despesas, lstr_sql, "frm_movimentacao_receitas_despesas", "lsub_preencher_grade_despesas ")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_preencher_grade_despesas
    End If
    llng_registros = lobj_despesas.Count
    'se houver registros na tabela
    If (llng_registros > 0) Then
        'zera a variável com o total das despesas
        mdbl_valor_despesas = 0
        'desabilita a atualização da grade
        msf_grade_despesas.Redraw = False
        For llng_contador = 1 To llng_registros
            msf_grade_despesas.RowData(llng_contador) = lobj_despesas(llng_contador)("int_codigo")
            msf_grade_despesas.TextMatrix(llng_contador, enm_despesas.col_despesa) = " " & lobj_despesas(llng_contador)("descricao")
            msf_grade_despesas.TextMatrix(llng_contador, enm_despesas.col_quantidade) = " " & Format$(lobj_despesas(llng_contador)("quantidade"), pcst_formato_numerico_padrao)
            msf_grade_despesas.TextMatrix(llng_contador, enm_despesas.col_valor_medio) = " " & Format$(lobj_despesas(llng_contador)("valor_medio"), pcst_formato_numerico)
            msf_grade_despesas.TextMatrix(llng_contador, enm_despesas.col_valor_total) = " " & Format$(lobj_despesas(llng_contador)("valor_total"), pcst_formato_numerico)
            'alimenta a variável com o valor total das despesas
            mdbl_valor_despesas = mdbl_valor_despesas + CDbl(lobj_despesas(llng_contador)("valor_total"))
            'incrementa uma linha
            If (llng_contador < llng_registros) Then
                msf_grade_despesas.Rows = msf_grade_despesas.Rows + 1
            End If
        Next
        msf_grade_despesas.Col = enm_despesas.col_despesa
        msf_grade_despesas.Row = 1
        'reabilita a atualização da grade
        msf_grade_despesas.Redraw = True
    End If
fim_lsub_preencher_grade_despesas:
    'destrói os objetos
    Set lobj_despesas = Nothing
    Exit Sub
erro_lsub_preencher_grade_despesas:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "lsub_preencher_grade_despesas"
    GoTo fim_lsub_preencher_grade_despesas
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "cbo_contas_Click"
    GoTo fim_cbo_contas_Click
End Sub

Private Sub cbo_contas_DropDown()
    On Error GoTo erro_cbo_contas_DropDown
    psub_campo_got_focus cbo_contas
fim_cbo_contas_DropDown:
    Exit Sub
erro_cbo_contas_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "cbo_contas_DropDown"
    GoTo fim_cbo_contas_DropDown
End Sub

Private Sub cbo_contas_GotFocus()
    On Error GoTo erro_cbo_contas_GotFocus
    psub_campo_got_focus cbo_contas
fim_cbo_contas_GotFocus:
    Exit Sub
erro_cbo_contas_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "cbo_contas_GotFocus"
    GoTo fim_cbo_contas_GotFocus
End Sub

Private Sub cbo_contas_LostFocus()
    On Error GoTo erro_cbo_contas_LostFocus
    psub_campo_lost_focus cbo_contas
fim_cbo_contas_LostFocus:
    Exit Sub
erro_cbo_contas_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "cbo_contas_LostFocus"
    GoTo fim_cbo_contas_LostFocus
End Sub

Private Sub cmd_detalhes_receita_Click()
    On Error GoTo erro_cmd_detalhes_receita_Click
    'declara variáveis
    Dim lobj_frm_detalhes As Form
    Dim llng_codigo_item    As Long

    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_detalhes_receita.Enabled) Then
        Exit Sub
    End If
    
    llng_codigo_item = msf_grade_receitas.RowData(msf_grade_receitas.Row)
    If (llng_codigo_item = 0) Then
        MsgBox "Selecione um item na grade de receitas.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        GoTo fim_cmd_detalhes_receita_Click
    Else
    
        'força a destruição do objeto
        Set lobj_frm_detalhes = Nothing
        'cria uma nova instância do formulário
        Set lobj_frm_detalhes = New frm_movimentacao_por_receitas_despesas
        
        'atribui as propriedades
        With lobj_frm_detalhes
            Set .obj_frm_anterior = Me
            .str_tipo_movimentacao = "E" 'entrada
            .int_codigo_registro = llng_codigo_item
            .dt_data_de = dtp_de.Value
            .dt_data_ate = dtp_ate.Value
            'exibe o form
            .Show
        End With
    
    End If

fim_cmd_detalhes_receita_Click:
    Exit Sub
erro_cmd_detalhes_receita_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "cmd_detalhes_receita_Click"
    GoTo fim_cmd_detalhes_receita_Click
End Sub

Private Sub cmd_detalhes_despesa_Click()
    On Error GoTo erro_cmd_detalhes_despesa_Click
    'declara variáveis
    Dim lobj_frm_detalhes As Form
    Dim llng_codigo_item    As Long

    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_detalhes_despesa.Enabled) Then
        Exit Sub
    End If
    
    llng_codigo_item = msf_grade_despesas.RowData(msf_grade_despesas.Row)
    If (llng_codigo_item = 0) Then
        MsgBox "Selecione um item na grade de despesas.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        GoTo fim_cmd_detalhes_despesa_Click
    Else
    
        'força a destruição do objeto
        Set lobj_frm_detalhes = Nothing
        'cria uma nova instância do formulário
        Set lobj_frm_detalhes = New frm_movimentacao_por_receitas_despesas
        
        'atribui as propriedades
        With lobj_frm_detalhes
            Set .obj_frm_anterior = Me
            .str_tipo_movimentacao = "S" 'saída
            .int_codigo_registro = llng_codigo_item
            .dt_data_de = dtp_de.Value
            .dt_data_ate = dtp_ate.Value
            'exibe o form
            .Show
        End With
    
    End If

fim_cmd_detalhes_despesa_Click:
    Exit Sub
erro_cmd_detalhes_despesa_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_despesas_despesas", "cmd_detalhes_despesa_Click"
    GoTo fim_cmd_detalhes_despesa_Click
End Sub

Private Sub dtp_de_DropDown()
    On Error GoTo erro_dtp_de_DropDown
    psub_campo_got_focus dtp_de
fim_dtp_de_DropDown:
    Exit Sub
erro_dtp_de_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "dtp_de_DropDown"
    GoTo fim_dtp_de_DropDown
End Sub

Private Sub dtp_de_GotFocus()
    On Error GoTo erro_dtp_de_GotFocus
    psub_campo_got_focus dtp_de
fim_dtp_de_GotFocus:
    Exit Sub
erro_dtp_de_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "dtp_de_GotFocus"
    GoTo fim_dtp_de_GotFocus
End Sub

Private Sub dtp_de_LostFocus()
    On Error GoTo erro_dtp_de_LostFocus
    psub_campo_lost_focus dtp_de
fim_dtp_de_LostFocus:
    Exit Sub
erro_dtp_de_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "dtp_de_LostFocus"
    GoTo fim_dtp_de_LostFocus
End Sub

Private Sub dtp_ate_DropDown()
    On Error GoTo erro_dtp_ate_DropDown
    psub_campo_got_focus dtp_ate
fim_dtp_ate_DropDown:
    Exit Sub
erro_dtp_ate_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "dtp_ate_DropDown"
    GoTo fim_dtp_ate_DropDown
End Sub

Private Sub dtp_ate_GotFocus()
    On Error GoTo erro_dtp_ate_GotFocus
    psub_campo_got_focus dtp_ate
fim_dtp_ate_GotFocus:
    Exit Sub
erro_dtp_ate_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "dtp_ate_GotFocus"
    GoTo fim_dtp_ate_GotFocus
End Sub

Private Sub dtp_ate_LostFocus()
    On Error GoTo erro_dtp_ate_LostFocus
    psub_campo_lost_focus dtp_ate
fim_dtp_ate_LostFocus:
    Exit Sub
erro_dtp_ate_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "dtp_ate_LostFocus"
    GoTo fim_dtp_ate_LostFocus
End Sub

Private Sub cbo_ordenar_por_DropDown()
    On Error GoTo erro_cbo_ordenar_por_DropDown
    psub_campo_got_focus cbo_ordenar_por
fim_cbo_ordenar_por_DropDown:
    Exit Sub
erro_cbo_ordenar_por_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "cbo_ordenar_por_DropDown"
    GoTo fim_cbo_ordenar_por_DropDown
End Sub

Private Sub cbo_ordenar_por_GotFocus()
    On Error GoTo erro_cbo_ordenar_por_GotFocus
    psub_campo_got_focus cbo_ordenar_por
fim_cbo_ordenar_por_GotFocus:
    Exit Sub
erro_cbo_ordenar_por_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "cbo_ordenar_por_GotFocus"
    GoTo fim_cbo_ordenar_por_GotFocus
End Sub

Private Sub cbo_ordenar_por_LostFocus()
    On Error GoTo erro_cbo_ordenar_por_LostFocus
    psub_campo_lost_focus cbo_ordenar_por
fim_cbo_ordenar_por_LostFocus:
    Exit Sub
erro_cbo_ordenar_por_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "cbo_ordenar_por_LostFocus"
    GoTo fim_cbo_ordenar_por_LostFocus
End Sub

Private Sub cbo_ordem_DropDown()
    On Error GoTo erro_cbo_ordem_DropDown
    psub_campo_got_focus cbo_ordem
fim_cbo_ordem_DropDown:
    Exit Sub
erro_cbo_ordem_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "cbo_ordem_DropDown"
    GoTo fim_cbo_ordem_DropDown
End Sub

Private Sub cbo_ordem_GotFocus()
    On Error GoTo erro_cbo_ordem_GotFocus
    psub_campo_got_focus cbo_ordem
fim_cbo_ordem_GotFocus:
    Exit Sub
erro_cbo_ordem_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "cbo_ordem_GotFocus"
    GoTo fim_cbo_ordem_GotFocus
End Sub

Private Sub cbo_ordem_LostFocus()
    On Error GoTo erro_cbo_ordem_LostFocus
    psub_campo_lost_focus cbo_ordem
fim_cbo_ordem_LostFocus:
    Exit Sub
erro_cbo_ordem_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "cbo_ordem_LostFocus"
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "cmd_fechar_Click"
    GoTo fim_cmd_fechar_Click
End Sub

Private Sub cmd_filtrar_Click()
    On Error GoTo erro_cmd_filtrar_Click
    'declaração de variáveis
    Dim lstr_conta As String
    Dim lstr_data_de As String
    Dim lstr_data_ate As String
    Dim lstr_ordenar_por As String
    Dim lstr_ordem As String
    Dim lbln_contas_inativas As Boolean
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_filtrar.Enabled) Then
        Exit Sub
    End If
    'valida os campos da tela
    If (cbo_contas.ListIndex = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [conta] é obrigatório.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_contas.SetFocus
        GoTo fim_cmd_filtrar_Click
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
    lstr_conta = CStr(cbo_contas.ItemData(cbo_contas.ListIndex))
    lstr_data_de = Format$(dtp_de, pcst_formato_data_sql)
    lstr_data_ate = Format$(dtp_ate, pcst_formato_data_sql)
    'ordem de campo das grades
    Select Case cbo_ordenar_por.ListIndex
        Case 1
            lstr_ordenar_por = "[descricao]"
        Case 2
            lstr_ordenar_por = "[quantidade]"
        Case 3
            lstr_ordenar_por = "[valor_medio]"
        Case 4
            lstr_ordenar_por = "[valor_total]"
    End Select
    'ordem das grades
    Select Case cbo_ordem.ListIndex
        Case 1
            lstr_ordem = "asc"
        Case 2
            lstr_ordem = "desc"
    End Select
    'incluir contas inativas?
    lbln_contas_inativas = IIf(chk_considerar_contas_inativas.Value = vbChecked, True, False)
    'reajusta as grades
    lsub_ajustar_grade_despesas msf_grade_despesas
    lsub_ajustar_grade_receitas msf_grade_receitas
    'faz a chamada aos métodos de consulta ao banco e preenchimento de grade
    lsub_preencher_grade_receitas lstr_conta, lstr_data_de, lstr_data_ate, lstr_ordenar_por, lstr_ordem, lbln_contas_inativas
    lsub_preencher_grade_despesas lstr_conta, lstr_data_de, lstr_data_ate, lstr_ordenar_por, lstr_ordem, lbln_contas_inativas
    'ajusta a barra de status
    lsub_preencher_barra_status mdbl_valor_receitas, mdbl_valor_despesas
fim_cmd_filtrar_Click:
    Exit Sub
erro_cmd_filtrar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "cmd_filtrar_Click"
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "cmd_iniciar_Click"
    GoTo fim_cmd_iniciar_Click
End Sub

Private Sub Form_Initialize()
    On Error GoTo Erro_Form_Initialize
    InitCommonControls
Fim_Form_Initialize:
    Exit Sub
Erro_Form_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "Form_Initialize"
    GoTo Fim_Form_Initialize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/movimentacao_receitas_x_despesas.htm", 0
        Case vbKeyF2
            cmd_detalhes_receita_Click
        Case vbKeyF3
            cmd_detalhes_despesa_Click
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "Form_KeyUp"
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
    lsub_ajustar_grade_receitas msf_grade_receitas
    lsub_ajustar_grade_despesas msf_grade_despesas
    lsub_ajustar_status stb_status
    'zera as variáveis
    mdbl_valor_despesas = 0
    mdbl_valor_receitas = 0
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "Form_Load"
    GoTo fim_Form_Load
End Sub

Private Sub msf_grade_receitas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_msf_grade_receitas_MouseUp
    If (Button = 2) Then 'botão direito do mouse
        PopupMenu mnu_msf_grade_receitas 'exibimos o popup
    End If
fim_msf_grade_receitas_MouseUp:
    Exit Sub
erro_msf_grade_receitas_MouseUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "msf_grade_receitas_MouseUp"
    GoTo fim_msf_grade_receitas_MouseUp
End Sub

Private Sub mnu_msf_grade_receitas_copiar_Click()
    On Error GoTo erro_mnu_msf_grade_receitas_copiar_Click
    pfct_copiar_conteudo_grade msf_grade_receitas
fim_mnu_msf_grade_receitas_copiar_Click:
    Exit Sub
erro_mnu_msf_grade_receitas_copiar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "mnu_msf_grade_receitas_copiar_Click"
    GoTo fim_mnu_msf_grade_receitas_copiar_Click
End Sub

Private Sub mnu_msf_grade_receitas_exportar_Click()
    On Error GoTo erro_mnu_msf_grade_receitas_exportar_Click
    pfct_exportar_conteudo_grade msf_grade_receitas, "movimentacao_receitas_x_despesas"
fim_mnu_msf_grade_receitas_exportar_Click:
    Exit Sub
erro_mnu_msf_grade_receitas_exportar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "mnu_msf_grade_receitas_exportar_Click"
    GoTo fim_mnu_msf_grade_receitas_exportar_Click
End Sub

Private Sub msf_grade_despesas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_msf_grade_despesas_MouseUp
    If (Button = 2) Then 'botão direito do mouse
        PopupMenu mnu_msf_grade_despesas 'exibimos o popup
    End If
fim_msf_grade_despesas_MouseUp:
    Exit Sub
erro_msf_grade_despesas_MouseUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "msf_grade_despesas_MouseUp"
    GoTo fim_msf_grade_despesas_MouseUp
End Sub

Private Sub mnu_msf_grade_despesas_copiar_Click()
    On Error GoTo erro_mnu_msf_grade_despesas_copiar_Click
    pfct_copiar_conteudo_grade msf_grade_despesas
fim_mnu_msf_grade_despesas_copiar_Click:
    Exit Sub
erro_mnu_msf_grade_despesas_copiar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "mnu_msf_grade_despesas_copiar_Click"
    GoTo fim_mnu_msf_grade_despesas_copiar_Click
End Sub

Private Sub mnu_msf_grade_despesas_exportar_Click()
    On Error GoTo erro_mnu_msf_grade_despesas_exportar_Click
    pfct_exportar_conteudo_grade msf_grade_despesas, "movimentacao_receitas_x_despesas"
fim_mnu_msf_grade_despesas_exportar_Click:
    Exit Sub
erro_mnu_msf_grade_despesas_exportar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_movimentacao_receitas_despesas", "mnu_msf_grade_despesas_exportar_Click"
    GoTo fim_mnu_msf_grade_despesas_exportar_Click
End Sub
