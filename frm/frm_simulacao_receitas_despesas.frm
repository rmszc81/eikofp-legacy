VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_simulacao_receitas_despesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulação Receitas x Despesas"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13935
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
   ScaleHeight     =   7320
   ScaleWidth      =   13935
   Begin VB.Frame fme_base_calculo 
      Caption         =   " Base de cálculo "
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   4695
      Begin VB.ComboBox cbo_tempo 
         Height          =   315
         Left            =   3180
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox txt_quantidade 
         Height          =   315
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   5
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lbl_considerar_quantidade 
         AutoSize        =   -1  'True
         Caption         =   "&Calcular com base nos últimos:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   405
         Width           =   2190
      End
   End
   Begin VB.Frame fme_ordenacao 
      Caption         =   " Ordenação dos resultados "
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   4695
      Begin VB.ComboBox cbo_ordenar_por 
         Height          =   315
         ItemData        =   "frm_simulacao_receitas_despesas.frx":0000
         Left            =   120
         List            =   "frm_simulacao_receitas_despesas.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   675
         Width           =   2235
      End
      Begin VB.ComboBox cbo_ordem 
         Height          =   315
         ItemData        =   "frm_simulacao_receitas_despesas.frx":0004
         Left            =   2460
         List            =   "frm_simulacao_receitas_despesas.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   675
         Width           =   2115
      End
      Begin VB.Label lbl_ordenar_por 
         AutoSize        =   -1  'True
         Caption         =   "&Ordenar por:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lbl_ordem 
         AutoSize        =   -1  'True
         Caption         =   "&Em ordem:"
         Height          =   195
         Left            =   2460
         TabIndex        =   14
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Frame fme_parametros 
      Caption         =   " Parâmetros "
      Height          =   1995
      Left            =   4920
      TabIndex        =   7
      Top             =   660
      Width           =   8895
      Begin VB.ListBox lst_receitas 
         Height          =   1185
         ItemData        =   "frm_simulacao_receitas_despesas.frx":0008
         Left            =   120
         List            =   "frm_simulacao_receitas_despesas.frx":000A
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   720
         Width           =   4275
      End
      Begin VB.ListBox lst_despesas 
         Height          =   1185
         ItemData        =   "frm_simulacao_receitas_despesas.frx":000C
         Left            =   4500
         List            =   "frm_simulacao_receitas_despesas.frx":000E
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   720
         Width           =   4275
      End
      Begin VB.Label lbl_receitas 
         AutoSize        =   -1  'True
         Caption         =   "&Receitas:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   420
         Width           =   675
      End
      Begin VB.Label lbl_despesas 
         AutoSize        =   -1  'True
         Caption         =   "&Despesas:"
         Height          =   195
         Left            =   4500
         TabIndex        =   9
         Top             =   420
         Width           =   750
      End
   End
   Begin VB.CommandButton cmd_fechar 
      Caption         =   "&Fechar (F8)"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton cmd_filtrar 
      Caption         =   "&Filtrar (F7)"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton cmd_iniciar 
      Caption         =   "&Iniciar (F10)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1275
   End
   Begin MSComctlLib.StatusBar stb_status 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   21
      Top             =   7035
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   24527
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid msf_grade_receitas 
      Height          =   3810
      Left            =   120
      TabIndex        =   19
      Top             =   3120
      Width           =   6795
      _ExtentX        =   11986
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
      Left            =   7020
      TabIndex        =   20
      Top             =   3120
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   6720
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lbl_grade_receitas 
      AutoSize        =   -1  'True
      Caption         =   "&Receitas:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2820
      Width           =   675
   End
   Begin VB.Label lbl_grade_despesas 
      AutoSize        =   -1  'True
      Caption         =   "&Despesas:"
      Height          =   195
      Left            =   7020
      TabIndex        =   18
      Top             =   2820
      Width           =   750
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
Attribute VB_Name = "frm_simulacao_receitas_despesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enm_receitas
    col_receita = 0
    col_valor_simulado = 1
    col_valor_recebido = 2
    col_valor_diferenca = 3
End Enum

Private Enum enm_despesas
    col_despesa = 0
    col_valor_simulado = 1
    col_valor_pago = 2
    col_valor_diferenca = 3
End Enum

Private Enum enm_status
    pnl_receitas = 1
    pnl_despesas = 2
    pnl_total = 3
End Enum

'totais receita
Private mdbl_receita_total_simulado As Double
Private mdbl_receita_total_pago As Double
Private mdbl_receita_total_diferenca As Double

'totais despesa
Private mdbl_despesa_total_simulado As Double
Private mdbl_despesa_total_pago As Double
Private mdbl_despesa_total_diferenca As Double


Private Sub lsub_preencher_grade_despesas(ByVal pstr_in As String, _
                                          ByVal pdt_data_inicial As Date, _
                                          ByVal pdt_data_final As Date, _
                                          ByVal pint_dividir_por As Integer, _
                                          ByVal pstr_ordenar_por As String, _
                                          ByVal pstr_ordem As String)
    On Error GoTo erro_lsub_preencher_grade_despesas
    Dim lobj_despesas As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select "
    lstr_sql = lstr_sql & "     [t].[codigo] as [int_codigo], "
    lstr_sql = lstr_sql & "     [t].[descricao] as [str_descricao], "
    lstr_sql = lstr_sql & "     [t].[valor simulado] as [num_simulado], "
    lstr_sql = lstr_sql & "     [t].[valor pago] as [num_pago], "
    lstr_sql = lstr_sql & "     ([t].[valor pago] - [t].[valor simulado]) as [num_diferenca] "
    lstr_sql = lstr_sql & " from "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & "     select "
    lstr_sql = lstr_sql & "         [d].[int_codigo] as [codigo], "
    lstr_sql = lstr_sql & "         [d].[str_descricao] as [descricao], "
    lstr_sql = lstr_sql & "         ifnull(( "
    lstr_sql = lstr_sql & "                     select "
    lstr_sql = lstr_sql & "                         (total([m].[num_valor]) / " & CStr(Abs(pint_dividir_por)) & ") " 'count(*)
    lstr_sql = lstr_sql & "                     from "
    lstr_sql = lstr_sql & "                         [tb_movimentacao] [m] "
    lstr_sql = lstr_sql & "                     where "
    lstr_sql = lstr_sql & "                             [m].[int_despesa] = [d].[int_codigo] "
    lstr_sql = lstr_sql & "                         and [m].[dt_pagamento] between '" & pfct_tratar_data_sql(pfct_retorna_primeiro_dia_mes(pdt_data_inicial)) & "' and '" & pfct_tratar_data_sql(pfct_retorna_ultimo_dia_mes(pdt_data_final)) & "' "
    lstr_sql = lstr_sql & "                 ),0) as [valor simulado], "
    lstr_sql = lstr_sql & "         ifnull(( "
    lstr_sql = lstr_sql & "                     select "
    lstr_sql = lstr_sql & "                         (total([m].[num_valor]) / 1) " 'count(*)
    lstr_sql = lstr_sql & "                     from "
    lstr_sql = lstr_sql & "                         [tb_movimentacao] [m] "
    lstr_sql = lstr_sql & "                     where "
    lstr_sql = lstr_sql & "                             [m].[int_despesa] = [d].[int_codigo] "
    lstr_sql = lstr_sql & "                         and [m].[dt_pagamento] between '" & pfct_tratar_data_sql(pfct_retorna_primeiro_dia_mes(Now)) & "' and '" & pfct_tratar_data_sql(pfct_retorna_ultimo_dia_mes(Now)) & "' "
    lstr_sql = lstr_sql & "                 ),0) as [valor pago] "
    lstr_sql = lstr_sql & "     from "
    lstr_sql = lstr_sql & "         [tb_despesas] [d] "
    lstr_sql = lstr_sql & " ) t "
    lstr_sql = lstr_sql & " where "
    lstr_sql = lstr_sql & "    [t].[codigo] IN (" & pstr_in & ")"
    lstr_sql = lstr_sql & " order by " & pstr_ordenar_por & " " & pstr_ordem & " "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_despesas, lstr_sql, "frm_simulacao_receitas_despesas", "lsub_preencher_grade_despesas ")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_preencher_grade_despesas
    End If
    llng_registros = lobj_despesas.Count
    'se houver registros na tabela
    If (llng_registros > 0) Then
        'zera os totais da receita
        mdbl_despesa_total_simulado = 0
        mdbl_despesa_total_pago = 0
        mdbl_despesa_total_diferenca = 0
        'desabilita a atualização da grade
        msf_grade_despesas.Redraw = False
        For llng_contador = 1 To llng_registros
            msf_grade_despesas.RowData(llng_contador) = lobj_despesas(llng_contador)("int_codigo")
            msf_grade_despesas.TextMatrix(llng_contador, enm_despesas.col_despesa) = " " & lobj_despesas(llng_contador)("str_descricao")
            msf_grade_despesas.TextMatrix(llng_contador, enm_despesas.col_valor_simulado) = " " & Format$(lobj_despesas(llng_contador)("num_simulado"), pcst_formato_numerico)
            msf_grade_despesas.TextMatrix(llng_contador, enm_despesas.col_valor_pago) = " " & Format$(lobj_despesas(llng_contador)("num_pago"), pcst_formato_numerico)
            msf_grade_despesas.TextMatrix(llng_contador, enm_despesas.col_valor_diferenca) = " " & Format$(lobj_despesas(llng_contador)("num_diferenca"), pcst_formato_numerico)
            'soma os totais das despesas para exibir na barra de status
            mdbl_despesa_total_simulado = mdbl_despesa_total_simulado + CDbl(lobj_despesas(llng_contador)("num_simulado"))
            mdbl_despesa_total_pago = mdbl_despesa_total_pago + CDbl(lobj_despesas(llng_contador)("num_pago"))
            mdbl_despesa_total_diferenca = mdbl_despesa_total_diferenca + CDbl(lobj_despesas(llng_contador)("num_diferenca"))
            'incrementa uma linha
            If (llng_contador < llng_registros) Then
                msf_grade_despesas.Rows = msf_grade_despesas.Rows + 1
            End If
            'insere linha com os totais
            If (llng_contador = llng_registros) Then
                'incrementa mais uma linha
                msf_grade_despesas.Rows = msf_grade_despesas.Rows + 1
                'incrementa mais um no contador
                llng_contador = llng_contador + 1
                'atribui os valores totais
                msf_grade_despesas.RowData(llng_contador) = 99999
                msf_grade_despesas.TextMatrix(llng_contador, enm_despesas.col_despesa) = " -- TOTAL -- "
                msf_grade_despesas.TextMatrix(llng_contador, enm_despesas.col_valor_simulado) = " " & Format$(mdbl_despesa_total_simulado, pcst_formato_numerico)
                msf_grade_despesas.TextMatrix(llng_contador, enm_despesas.col_valor_pago) = " " & Format$(mdbl_despesa_total_pago, pcst_formato_numerico)
                msf_grade_despesas.TextMatrix(llng_contador, enm_despesas.col_valor_diferenca) = " " & Format$(mdbl_despesa_total_diferenca, pcst_formato_numerico)
            End If
        Next
        msf_grade_despesas.Col = enm_despesas.col_despesa
        msf_grade_despesas.Row = 1
        'reabilita a atualização da grade
        msf_grade_despesas.Redraw = True
        'ajusta barra de status
        lsub_preencher_barra_status mdbl_receita_total_simulado, mdbl_despesa_total_simulado
    End If
fim_lsub_preencher_grade_despesas:
    'destrói objetos
    Set lobj_despesas = Nothing
    'sai do método
    Exit Sub
erro_lsub_preencher_grade_despesas:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "lsub_preencher_grade_despesas"
    GoTo fim_lsub_preencher_grade_despesas
End Sub

Private Sub lsub_preencher_grade_receitas(ByVal pstr_in As String, _
                                          ByVal pdt_data_inicial As Date, _
                                          ByVal pdt_data_final As Date, _
                                          ByVal pint_dividir_por As Integer, _
                                          ByVal pstr_ordenar_por As String, _
                                          ByVal pstr_ordem As String)
    On Error GoTo erro_lsub_preencher_grade_receitas
    Dim lobj_receitas As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select "
    lstr_sql = lstr_sql & "     [t].[codigo] as [int_codigo], "
    lstr_sql = lstr_sql & "     [t].[descricao] as [str_descricao], "
    lstr_sql = lstr_sql & "     [t].[valor simulado] as [num_simulado], "
    lstr_sql = lstr_sql & "     [t].[valor pago] as [num_pago], "
    lstr_sql = lstr_sql & "     ([t].[valor pago] - [t].[valor simulado]) as [num_diferenca] "
    lstr_sql = lstr_sql & " from "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & "     select "
    lstr_sql = lstr_sql & "         [r].[int_codigo] as [codigo], "
    lstr_sql = lstr_sql & "         [r].[str_descricao] as [descricao], "
    lstr_sql = lstr_sql & "         ifnull(( "
    lstr_sql = lstr_sql & "                     select "
    lstr_sql = lstr_sql & "                         (total([m].[num_valor]) / " & CStr(pint_dividir_por) & ") " 'count(*)
    lstr_sql = lstr_sql & "                     from "
    lstr_sql = lstr_sql & "                         [tb_movimentacao] [m] "
    lstr_sql = lstr_sql & "                     where "
    lstr_sql = lstr_sql & "                             [m].[int_receita] = [r].[int_codigo] "
    lstr_sql = lstr_sql & "                         and [m].[dt_pagamento] between '" & pfct_tratar_data_sql(pfct_retorna_primeiro_dia_mes(pdt_data_inicial)) & "' and '" & pfct_tratar_data_sql(pfct_retorna_ultimo_dia_mes(pdt_data_final)) & "' "
    lstr_sql = lstr_sql & "                 ),0) as [valor simulado], "
    lstr_sql = lstr_sql & "         ifnull(( "
    lstr_sql = lstr_sql & "                     select "
    lstr_sql = lstr_sql & "                         (total([m].[num_valor]) / 1) " 'count(*)
    lstr_sql = lstr_sql & "                     from "
    lstr_sql = lstr_sql & "                         [tb_movimentacao] [m] "
    lstr_sql = lstr_sql & "                     where "
    lstr_sql = lstr_sql & "                             [m].[int_receita] = [r].[int_codigo] "
    lstr_sql = lstr_sql & "                         and [m].[dt_pagamento] between '" & pfct_tratar_data_sql(pfct_retorna_primeiro_dia_mes(Now)) & "' and '" & pfct_tratar_data_sql(pfct_retorna_ultimo_dia_mes(Now)) & "' "
    lstr_sql = lstr_sql & "                 ),0) as [valor pago] "
    lstr_sql = lstr_sql & "     from "
    lstr_sql = lstr_sql & "         [tb_receitas] [r] "
    lstr_sql = lstr_sql & " ) t "
    lstr_sql = lstr_sql & " where "
    lstr_sql = lstr_sql & "    [t].[codigo] IN (" & pstr_in & ")"
    lstr_sql = lstr_sql & " order by " & pstr_ordenar_por & " " & pstr_ordem & " "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_receitas, lstr_sql, "frm_simulacao_receitas_receitas", "lsub_preencher_grade_receitas ")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_preencher_grade_receitas
    End If
    llng_registros = lobj_receitas.Count
    'se houver registros na tabela
    If (llng_registros > 0) Then
        'zera os totais da receita
        mdbl_receita_total_simulado = 0
        mdbl_receita_total_pago = 0
        mdbl_receita_total_diferenca = 0
        'desabilita a atualização da grade
        msf_grade_receitas.Redraw = False
        For llng_contador = 1 To llng_registros
            msf_grade_receitas.RowData(llng_contador) = lobj_receitas(llng_contador)("int_codigo")
            msf_grade_receitas.TextMatrix(llng_contador, enm_receitas.col_receita) = " " & lobj_receitas(llng_contador)("str_descricao")
            msf_grade_receitas.TextMatrix(llng_contador, enm_receitas.col_valor_simulado) = " " & Format$(lobj_receitas(llng_contador)("num_simulado"), pcst_formato_numerico)
            msf_grade_receitas.TextMatrix(llng_contador, enm_receitas.col_valor_recebido) = " " & Format$(lobj_receitas(llng_contador)("num_pago"), pcst_formato_numerico)
            msf_grade_receitas.TextMatrix(llng_contador, enm_receitas.col_valor_diferenca) = " " & Format$(lobj_receitas(llng_contador)("num_diferenca"), pcst_formato_numerico)
            'soma os totais das receitas para exibir na barra de status
            mdbl_receita_total_simulado = mdbl_receita_total_simulado + CDbl(lobj_receitas(llng_contador)("num_simulado"))
            mdbl_receita_total_pago = mdbl_receita_total_pago + CDbl(lobj_receitas(llng_contador)("num_pago"))
            mdbl_receita_total_diferenca = mdbl_receita_total_diferenca + CDbl(lobj_receitas(llng_contador)("num_diferenca"))
            'incrementa uma linha
            If (llng_contador < llng_registros) Then
                msf_grade_receitas.Rows = msf_grade_receitas.Rows + 1
            End If
            'insere linha com os totais
            If (llng_contador = llng_registros) Then
                'incrementa mais uma linha
                msf_grade_receitas.Rows = msf_grade_receitas.Rows + 1
                'incrementa mais um no contador
                llng_contador = llng_contador + 1
                'atribui os valores totais
                msf_grade_receitas.RowData(llng_contador) = 99999
                msf_grade_receitas.TextMatrix(llng_contador, enm_receitas.col_receita) = " -- TOTAL -- "
                msf_grade_receitas.TextMatrix(llng_contador, enm_receitas.col_valor_simulado) = " " & Format$(mdbl_receita_total_simulado, pcst_formato_numerico)
                msf_grade_receitas.TextMatrix(llng_contador, enm_receitas.col_valor_recebido) = " " & Format$(mdbl_receita_total_pago, pcst_formato_numerico)
                msf_grade_receitas.TextMatrix(llng_contador, enm_receitas.col_valor_diferenca) = " " & Format$(mdbl_receita_total_diferenca, pcst_formato_numerico)
            End If
        Next
        msf_grade_receitas.Col = enm_receitas.col_receita
        msf_grade_receitas.Row = 1
        'reabilita a atualização da grade
        msf_grade_receitas.Redraw = True
        'ajusta barra de status
        lsub_preencher_barra_status mdbl_receita_total_simulado, mdbl_despesa_total_simulado
    End If
fim_lsub_preencher_grade_receitas:
    'destrói objetos
    Set lobj_receitas = Nothing
    Exit Sub
erro_lsub_preencher_grade_receitas:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "lsub_preencher_grade_receitas"
    GoTo fim_lsub_preencher_grade_receitas
End Sub

Private Function lfct_validar_campos() As Boolean
    On Error GoTo erro_lfct_validar_campos
    'quantidade
    If (txt_quantidade.Text = "") Then
        MsgBox "Atenção!" & vbCrLf & "Campo [quantidade] não pode estar em branco.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_quantidade.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    If (Val(txt_quantidade.Text) = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [quantidade] deve ser maior que 0 [zero].", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_quantidade.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    'tempo
    If (cbo_tempo.ItemData(cbo_tempo.ListIndex) = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [tempo].", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_tempo.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    'receitas
    If (pfct_retorna_in(lst_receitas) = "") Then
        MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [receitas].", vbOKOnly + vbInformation, pcst_nome_aplicacao
        lst_receitas.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    'despesas
    If (pfct_retorna_in(lst_despesas) = "") Then
        MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [despesas].", vbOKOnly + vbInformation, pcst_nome_aplicacao
        lst_despesas.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    'valida o combo ordenar por
    If (cbo_ordenar_por.ListIndex = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [ordenar por] é obrigatório.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_ordenar_por.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    'valida o combo ordem
    If (cbo_ordem.ListIndex = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Campo [ordem] é obrigatório.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_ordem.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    'devolve true
    lfct_validar_campos = True
fim_lfct_validar_campos:
    Exit Function
erro_lfct_validar_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "lfct_validar_campos"
    GoTo fim_lfct_validar_campos
End Function

Private Sub lsub_ajustar_status(ByRef pstb_status As StatusBar)
    On Error GoTo erro_lsub_ajustar_status
    pstb_status.Panels.Clear
    pstb_status.Panels.Add
    pstb_status.Panels.Item(1).AutoSize = sbrSpring
    pstb_status.Panels.Item(1).Text = ""
fim_lsub_ajustar_status:
    Exit Sub
erro_lsub_ajustar_status:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "lsub_ajustar_status"
    GoTo fim_lsub_ajustar_status
End Sub

Private Sub lsub_preencher_barra_status(ByVal pdbl_valor_simulado_receitas As Double, ByVal pdbl_valor_simulado_despesas As Double)
    On Error GoTo erro_lsub_preencher_barra_status
    With stb_status.Panels
        'limpa a status bar
        .Clear
        'receitas
        .Add
        .Item(enm_status.pnl_receitas).AutoSize = sbrSpring
        .Item(enm_status.pnl_receitas).Text = " Simulado Receitas (" & pfct_retorna_simbolo_moeda() & "): " & Format$(pdbl_valor_simulado_receitas, pcst_formato_numerico)
        'despesas
        .Add
        .Item(enm_status.pnl_despesas).AutoSize = sbrSpring
        .Item(enm_status.pnl_despesas).Text = " Simulado Despesas (" & pfct_retorna_simbolo_moeda() & "): " & Format$(pdbl_valor_simulado_despesas, pcst_formato_numerico)
        'total
        .Add
        .Item(enm_status.pnl_total).AutoSize = sbrSpring
        .Item(enm_status.pnl_total).Text = " Total (" & pfct_retorna_simbolo_moeda() & "): " & Format$((pdbl_valor_simulado_receitas - pdbl_valor_simulado_despesas), pcst_formato_numerico)
    End With
fim_lsub_preencher_barra_status:
    Exit Sub
erro_lsub_preencher_barra_status:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "lsub_preencher_barra_status"
    GoTo fim_lsub_preencher_barra_status
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
        .ColWidth(enm_receitas.col_valor_simulado) = 1200
        .ColWidth(enm_receitas.col_valor_recebido) = 1200
        .ColWidth(enm_receitas.col_valor_diferenca) = 1200
        .TextMatrix(0, enm_receitas.col_receita) = " Receita"
        .TextMatrix(0, enm_receitas.col_valor_simulado) = " Simulado "
        .TextMatrix(0, enm_receitas.col_valor_recebido) = " Rec. " & Format$(Now, "mm/yyyy")
        .TextMatrix(0, enm_receitas.col_valor_diferenca) = " Diferença"
        .ColAlignment(enm_receitas.col_valor_simulado) = flexAlignRightCenter
        .ColAlignment(enm_receitas.col_valor_recebido) = flexAlignRightCenter
        .ColAlignment(enm_receitas.col_valor_diferenca) = flexAlignRightCenter
    End With
fim_lsub_ajustar_grade_receitas:
    Exit Sub
erro_lsub_ajustar_grade_receitas:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "lsub_ajustar_grade_receitas"
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
        .ColWidth(enm_despesas.col_valor_simulado) = 1200
        .ColWidth(enm_despesas.col_valor_pago) = 1200
        .ColWidth(enm_despesas.col_valor_diferenca) = 1200
        .TextMatrix(0, enm_despesas.col_despesa) = " Despesa"
        .TextMatrix(0, enm_despesas.col_valor_simulado) = " Simulado "
        .TextMatrix(0, enm_despesas.col_valor_pago) = " Pago " & Format$(Now, "mm/yyyy")
        .TextMatrix(0, enm_despesas.col_valor_diferenca) = " Diferença"
        .ColAlignment(enm_despesas.col_valor_simulado) = flexAlignRightCenter
        .ColAlignment(enm_despesas.col_valor_pago) = flexAlignRightCenter
        .ColAlignment(enm_despesas.col_valor_diferenca) = flexAlignRightCenter
    End With
fim_lsub_ajustar_grade_despesas:
    Exit Sub
erro_lsub_ajustar_grade_despesas:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "lsub_ajustar_grade_despesas"
    GoTo fim_lsub_ajustar_grade_despesas
End Sub

Private Sub lsub_preencher_combos()
    On Error GoTo erro_lsub_preencher_combos
    With cbo_ordenar_por
        .Clear
        .AddItem "- Selecione o campo -", 0
        .AddItem "- Receita/Despesa", 1
        .AddItem "- Valor Simulado", 2
        .AddItem "- Valor Pago/Recebido", 3
        .AddItem "- Diferença", 4
        .ListIndex = 0
    End With
    psub_preencher_ordem cbo_ordem
fim_lsub_preencher_combos:
    Exit Sub
erro_lsub_preencher_combos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "lsub_preencher_combos"
    GoTo fim_lsub_preencher_combos
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "cmd_fechar_Click"
    GoTo fim_cmd_fechar_Click
End Sub

Private Sub cmd_filtrar_Click()
    On Error GoTo erro_cmd_filtrar_Click
    Dim lint_tempo As Integer
    Dim lstr_ordenar_por As String
    Dim lstr_em_ordem As String
    If (lfct_validar_campos) Then
        'verifica qual opção foi selecionada no combo tempo
        Select Case cbo_tempo.ItemData(cbo_tempo.ListIndex)
            Case 30
                lint_tempo = ((Val(txt_quantidade.Text) + 1) * -1)
            Case 365
                lint_tempo = (((Val(txt_quantidade.Text)) * 12 + 1) * -1)
        End Select
        'verifica qual opção foi selecionada no combo ordenar por
        Select Case cbo_ordenar_por.ListIndex
            Case 1
                lstr_ordenar_por = "[str_descricao]"
            Case 2
                lstr_ordenar_por = "[num_simulado]"
            Case 3
                lstr_ordenar_por = "[num_pago]"
            Case 4
                lstr_ordenar_por = "[num_diferenca]"
        End Select
        'verifica qual opção foi selecionada no combo em ordem
        Select Case cbo_ordem.ListIndex
            Case 1
                lstr_em_ordem = "asc"
            Case 2
                lstr_em_ordem = "desc"
        End Select
        'reajusta as grades
        lsub_ajustar_grade_receitas msf_grade_receitas
        lsub_ajustar_grade_despesas msf_grade_despesas
        'ajusta a barra de status
        lsub_ajustar_status stb_status
        'busca os dados
        lsub_preencher_grade_receitas (pfct_retorna_in(lst_receitas)), DateAdd("m", lint_tempo, Now), DateAdd("m", -1, Now), IIf(Abs(lint_tempo) = 0, 1, Abs(lint_tempo) - 1), lstr_ordenar_por, lstr_em_ordem
        lsub_preencher_grade_despesas (pfct_retorna_in(lst_despesas)), DateAdd("m", lint_tempo, Now), DateAdd("m", -1, Now), IIf(Abs(lint_tempo) = 0, 1, Abs(lint_tempo) - 1), lstr_ordenar_por, lstr_em_ordem
    End If
fim_cmd_filtrar_Click:
    Exit Sub
erro_cmd_filtrar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "cmd_filtrar_Click"
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
    Form_Activate
fim_cmd_iniciar_Click:
    Exit Sub
erro_cmd_iniciar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "cmd_iniciar_Click"
    GoTo fim_cmd_iniciar_Click
End Sub

Private Sub Form_Activate()
    On Error GoTo Erro_Form_Activate
    'posiciona o foco no campo
    txt_quantidade.SetFocus
Fim_Form_Activate:
    Exit Sub
Erro_Form_Activate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "Form_Activate"
    GoTo Fim_Form_Activate
End Sub

Private Sub Form_Initialize()
    On Error GoTo Erro_Form_Initialize
    InitCommonControls
Fim_Form_Initialize:
    Exit Sub
Erro_Form_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "Form_Initialize"
    GoTo Fim_Form_Initialize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/simulacao_receitas_x_despesas.htm", 0
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
    'preenche o combo tempo
    psub_preencher_tempo cbo_tempo
    'preenche os combos ordenar por e ordem
    lsub_preencher_combos
    'preenche as receitas
    psub_preencher_receitas lst_receitas, True
    'preenche as despesas
    psub_preencher_despesas lst_despesas, True
    'ajusta grade receitas
    lsub_ajustar_grade_receitas msf_grade_receitas
    'ajusta grade despesas
    lsub_ajustar_grade_despesas msf_grade_despesas
    'limpa o campo quantidade
    txt_quantidade.Text = ""
    'ajusta a barra de status
    lsub_ajustar_status stb_status
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "Form_Load"
    GoTo fim_Form_Load
End Sub

Private Sub txt_quantidade_GotFocus()
    On Error GoTo erro_txt_quantidade_GotFocus
    psub_campo_got_focus txt_quantidade
fim_txt_quantidade_GotFocus:
    Exit Sub
erro_txt_quantidade_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "txt_quantidade_GotFocus"
    GoTo fim_txt_quantidade_GotFocus
End Sub

Private Sub txt_quantidade_LostFocus()
    On Error GoTo erro_txt_quantidade_LostFocus
    psub_campo_lost_focus txt_quantidade
fim_txt_quantidade_LostFocus:
    Exit Sub
erro_txt_quantidade_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "txt_quantidade_LostFocus"
    GoTo fim_txt_quantidade_LostFocus
End Sub

Private Sub txt_quantidade_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_quantidade_Validate
    psub_tratar_campo txt_quantidade
    Cancel = Not pfct_validar_campo(txt_quantidade, tc_inteiro)
fim_txt_quantidade_Validate:
    Exit Sub
erro_txt_quantidade_Validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "txt_quantidade_Validate"
    GoTo fim_txt_quantidade_Validate
End Sub

Private Sub cbo_ordem_DropDown()
    On Error GoTo erro_cbo_ordem_DropDown
    psub_campo_got_focus cbo_ordem
fim_cbo_ordem_DropDown:
    Exit Sub
erro_cbo_ordem_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "cbo_ordem_DropDown"
    GoTo fim_cbo_ordem_DropDown
End Sub

Private Sub cbo_ordem_GotFocus()
    On Error GoTo erro_cbo_ordem_GotFocus
    psub_campo_got_focus cbo_ordem
fim_cbo_ordem_GotFocus:
    Exit Sub
erro_cbo_ordem_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "cbo_ordem_GotFocus"
    GoTo fim_cbo_ordem_GotFocus
End Sub

Private Sub cbo_ordem_LostFocus()
    On Error GoTo erro_cbo_ordem_LostFocus
    psub_campo_lost_focus cbo_ordem
fim_cbo_ordem_LostFocus:
    Exit Sub
erro_cbo_ordem_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "cbo_ordem_LostFocus"
    GoTo fim_cbo_ordem_LostFocus
End Sub

Private Sub cbo_ordenar_por_DropDown()
    On Error GoTo erro_cbo_ordenar_por_DropDown
    psub_campo_got_focus cbo_ordenar_por
fim_cbo_ordenar_por_DropDown:
    Exit Sub
erro_cbo_ordenar_por_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "cbo_ordenar_por_DropDown"
    GoTo fim_cbo_ordenar_por_DropDown
End Sub

Private Sub cbo_ordenar_por_GotFocus()
    On Error GoTo erro_cbo_ordenar_por_GotFocus
    psub_campo_got_focus cbo_ordenar_por
fim_cbo_ordenar_por_GotFocus:
    Exit Sub
erro_cbo_ordenar_por_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "cbo_ordenar_por_GotFocus"
    GoTo fim_cbo_ordenar_por_GotFocus
End Sub

Private Sub cbo_ordenar_por_LostFocus()
    On Error GoTo erro_cbo_ordenar_por_LostFocus
    psub_campo_lost_focus cbo_ordenar_por
fim_cbo_ordenar_por_LostFocus:
    Exit Sub
erro_cbo_ordenar_por_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "cbo_ordenar_por_LostFocus"
    GoTo fim_cbo_ordenar_por_LostFocus
End Sub

Private Sub lst_receitas_GotFocus()
    On Error GoTo erro_lst_receitas_GotFocus
    psub_campo_got_focus lst_receitas
fim_lst_receitas_GotFocus:
    Exit Sub
erro_lst_receitas_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "lst_receitas_GotFocus"
    GoTo fim_lst_receitas_GotFocus
End Sub

Private Sub lst_receitas_LostFocus()
    On Error GoTo erro_lst_receitas_LostFocus
    psub_campo_lost_focus lst_receitas
fim_lst_receitas_LostFocus:
    Exit Sub
erro_lst_receitas_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "lst_receitas_LostFocus"
    GoTo fim_lst_receitas_LostFocus
End Sub

Private Sub lst_despesas_GotFocus()
    On Error GoTo erro_lst_despesas_GotFocus
    psub_campo_got_focus lst_despesas
fim_lst_despesas_GotFocus:
    Exit Sub
erro_lst_despesas_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "lst_despesas_GotFocus"
    GoTo fim_lst_despesas_GotFocus
End Sub

Private Sub lst_despesas_LostFocus()
    On Error GoTo erro_lst_despesas_LostFocus
    psub_campo_lost_focus lst_despesas
fim_lst_despesas_LostFocus:
    Exit Sub
erro_lst_despesas_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "lst_despesas_LostFocus"
    GoTo fim_lst_despesas_LostFocus
End Sub

Private Sub msf_grade_receitas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_msf_grade_receitas_MouseUp
    If (Button = 2) Then 'botão direito do mouse
        PopupMenu mnu_msf_grade_receitas 'exibimos o popup
    End If
fim_msf_grade_receitas_MouseUp:
    Exit Sub
erro_msf_grade_receitas_MouseUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "msf_grade_receitas_MouseUp"
    GoTo fim_msf_grade_receitas_MouseUp
End Sub

Private Sub mnu_msf_grade_receitas_copiar_Click()
    On Error GoTo erro_mnu_msf_grade_receitas_copiar_Click
    pfct_copiar_conteudo_grade msf_grade_receitas
fim_mnu_msf_grade_receitas_copiar_Click:
    Exit Sub
erro_mnu_msf_grade_receitas_copiar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "mnu_msf_grade_receitas_copiar_Click"
    GoTo fim_mnu_msf_grade_receitas_copiar_Click
End Sub

Private Sub mnu_msf_grade_receitas_exportar_Click()
    On Error GoTo erro_mnu_msf_grade_receitas_exportar_Click
    pfct_exportar_conteudo_grade msf_grade_receitas, "simulacao_receitas_x_despesas"
fim_mnu_msf_grade_receitas_exportar_Click:
    Exit Sub
erro_mnu_msf_grade_receitas_exportar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "mnu_msf_grade_receitas_exportar_Click"
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "msf_grade_despesas_MouseUp"
    GoTo fim_msf_grade_despesas_MouseUp
End Sub

Private Sub mnu_msf_grade_despesas_copiar_Click()
    On Error GoTo erro_mnu_msf_grade_despesas_copiar_Click
    pfct_copiar_conteudo_grade msf_grade_despesas
fim_mnu_msf_grade_despesas_copiar_Click:
    Exit Sub
erro_mnu_msf_grade_despesas_copiar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "mnu_msf_grade_despesas_copiar_Click"
    GoTo fim_mnu_msf_grade_despesas_copiar_Click
End Sub

Private Sub mnu_msf_grade_despesas_exportar_Click()
    On Error GoTo erro_mnu_msf_grade_despesas_exportar_Click
    pfct_exportar_conteudo_grade msf_grade_despesas, "simulacao_receitas_x_despesas"
fim_mnu_msf_grade_despesas_exportar_Click:
    Exit Sub
erro_mnu_msf_grade_despesas_exportar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_simulacao_receitas_despesas", "mnu_msf_grade_despesas_exportar_Click"
    GoTo fim_mnu_msf_grade_despesas_exportar_Click
End Sub
