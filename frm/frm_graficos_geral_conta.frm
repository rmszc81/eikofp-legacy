VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frm_graficos_geral_conta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gráficos por Conta"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11415
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
   ScaleHeight     =   1140
   ScaleWidth      =   11415
   Begin VB.PictureBox pic_grafico 
      BackColor       =   &H80000005&
      Height          =   5955
      Left            =   120
      ScaleHeight     =   5895
      ScaleWidth      =   11115
      TabIndex        =   12
      Top             =   1260
      Width           =   11175
      Begin MSChart20Lib.MSChart msc_grafico 
         Height          =   5895
         Left            =   0
         OleObjectBlob   =   "frm_graficos_geral_conta.frx":0000
         TabIndex        =   13
         Top             =   0
         Width           =   11115
      End
   End
   Begin VB.CommandButton cmd_fechar 
      Caption         =   "&Fechar (F8)"
      Height          =   375
      Left            =   10020
      TabIndex        =   9
      Top             =   390
      Width           =   1275
   End
   Begin VB.CommandButton cmd_gerar 
      Caption         =   "&Gerar (F7)"
      Height          =   375
      Left            =   8700
      TabIndex        =   8
      Top             =   390
      Width           =   1275
   End
   Begin VB.ComboBox cbo_tipo_grafico 
      Height          =   315
      ItemData        =   "frm_graficos_geral_conta.frx":39FB
      Left            =   6720
      List            =   "frm_graficos_geral_conta.frx":39FD
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   420
      Width           =   1875
   End
   Begin VB.ComboBox cbo_exibir 
      Height          =   315
      ItemData        =   "frm_graficos_geral_conta.frx":39FF
      Left            =   4740
      List            =   "frm_graficos_geral_conta.frx":3A01
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   420
      Width           =   1875
   End
   Begin MSComCtl2.DTPicker dtp_de 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   420
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   556
      _Version        =   393216
      Format          =   242221057
      CurrentDate     =   39591
   End
   Begin MSComCtl2.DTPicker dtp_ate 
      Height          =   315
      Left            =   2580
      TabIndex        =   5
      Top             =   420
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Format          =   242221057
      CurrentDate     =   39591
   End
   Begin MSComctlLib.StatusBar stb_status 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   855
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20082
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl_grafico 
      AutoSize        =   -1  'True
      Caption         =   "&Gráfico:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   570
   End
   Begin VB.Label lbl_tipo_grafico 
      AutoSize        =   -1  'True
      Caption         =   "&Tipo de gráfico:"
      Height          =   195
      Left            =   6720
      TabIndex        =   2
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label lbl_exibir 
      AutoSize        =   -1  'True
      Caption         =   "&Exibir:"
      Height          =   195
      Left            =   4740
      TabIndex        =   1
      Top             =   120
      Width           =   450
   End
   Begin VB.Label lbl_ate 
      AutoSize        =   -1  'True
      Caption         =   "até:"
      Height          =   195
      Left            =   2220
      TabIndex        =   4
      Top             =   480
      Width           =   300
   End
   Begin VB.Label lbl_periodo 
      AutoSize        =   -1  'True
      Caption         =   "&Período:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frm_graficos_geral_conta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enm_exibir
    val_selecione = 0
    val_entradas = 1
    val_saidas = 2
    val_ambas = 3
End Enum

Private Enum enm_tipo_grafico
    val_selecione = 0
    val_2d_barras = 1
    val_2d_pizza = 2
End Enum

Private Enum enm_status
    pnl_mensagem = 1
End Enum

Private Sub lsub_preencher_combos()
    On Error GoTo erro_lsub_preencher_combos
    With cbo_exibir
        .Clear
        .AddItem "- Selecione o tipo -", enm_exibir.val_selecione
        .AddItem "- Entradas", enm_exibir.val_entradas
        .AddItem "- Saídas", enm_exibir.val_saidas
        .AddItem "- Ambas", enm_exibir.val_ambas
        .ListIndex = enm_exibir.val_selecione
    End With
    With cbo_tipo_grafico
        .Clear
        .AddItem "- Selecione o tipo -", enm_tipo_grafico.val_selecione
        .AddItem "- Barras", enm_tipo_grafico.val_2d_barras
        .AddItem "- Pizza", enm_tipo_grafico.val_2d_pizza
        .ListIndex = enm_tipo_grafico.val_selecione
    End With
fim_lsub_preencher_combos:
    Exit Sub
erro_lsub_preencher_combos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_graficos_geral_conta", "lsub_preencher_combos"
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_graficos_geral_conta", "cmd_fechar_Click"
    GoTo fim_cmd_fechar_Click
End Sub

Private Sub cmd_gerar_Click()
    On Error GoTo erro_cmd_gerar_Click
    'declaração de variáveis
    Dim ldt_data_de As Date
    Dim ldt_data_ate As Date
    Dim lenm_exibir As enm_exibir
    Dim lenm_tipo_grafico As enm_tipo_grafico
    '
    Dim lstr_sql As String
    Dim llng_contador As Long
    Dim llng_registros As Long
    Dim lobj_contas As Object
    Dim lobj_valores As Object
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_gerar.Enabled) Then
        Exit Sub
    End If
    'atribui valores às variáveis
    ldt_data_de = dtp_de.Value
    ldt_data_ate = dtp_ate.Value
    'valida as datas
    If (ldt_data_de > ldt_data_ate) Then
        MsgBox "Atenção!" & vbCrLf & "Data inicial deve ser menor que data final.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        psub_ajustar_combos_data dtp_de, dtp_ate
        dtp_de.SetFocus
        GoTo fim_cmd_gerar_Click
    End If
    'valida o combo tipo movimentação
    If (cbo_exibir.ListIndex = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Selecione o tipo de movimentação.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_exibir.SetFocus
        GoTo fim_cmd_gerar_Click
    End If
    'valida o combo tipo gráfico
    If (cbo_tipo_grafico.ListIndex = 0) Then
        MsgBox "Atenção!" & vbCrLf & "Selecione o tipo de gráfico para exibição.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_tipo_grafico.SetFocus
        GoTo fim_cmd_gerar_Click
    End If
    'atribui valores às variáveis
    lenm_exibir = cbo_exibir.ListIndex
    lenm_tipo_grafico = cbo_tipo_grafico.ListIndex
    'ajusta a barra de status
    With stb_status
        .Panels.Clear
        .Panels.Add
        .Panels.Item(enm_status.pnl_mensagem).Text = "Aguarde. Processando dados."
        .Panels.Item(enm_status.pnl_mensagem).AutoSize = sbrSpring
    End With
    'monta o comando sql
    lstr_sql = "select * from [tb_contas] where [chr_ativo] = 'S'"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_contas, lstr_sql, "frm_graficos_geral_conta", "cmd_gerar_Click")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_cmd_gerar_Click
    End If
    llng_registros = lobj_contas.Count
    'se não houver contas ativas no banco
    If (llng_registros = 0) Then
        'limpa a barra de status
        With stb_status
            .Panels.Clear
            .Panels.Add
            .Panels.Item(enm_status.pnl_mensagem).Text = ""
            .Panels.Item(enm_status.pnl_mensagem).AutoSize = sbrSpring
        End With
        'exibe mensagem e desvia para o fim do método
        MsgBox "Atenção!" & vbCrLf & "Não há contas ativas cadastradas.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        GoTo fim_cmd_gerar_Click
    Else 'caso haja contas ativas
        'limpa a variável
        lstr_sql = ""
        'percorre o objeto de contas
        For llng_contador = 1 To llng_registros
            'concatena a string para gerar a query de consulta
            lstr_sql = lstr_sql & " select "
            lstr_sql = lstr_sql & " ('" & lobj_contas(llng_contador)("str_descricao") & "') as [str_descricao_conta], "
            If (lenm_exibir = enm_exibir.val_entradas) Then
                lstr_sql = lstr_sql & "(select sum(ifnull([num_valor], 0)) from [tb_movimentacao] where [chr_tipo] = 'E' and [int_conta] = " & lobj_contas(llng_contador)("int_codigo") & " and [dt_pagamento] between '" & Format$(ldt_data_de, pcst_formato_data_sql) & "' and '" & Format$(ldt_data_ate, pcst_formato_data_sql) & "') as [num_valor_entrada] "
            End If
            If (lenm_exibir = enm_exibir.val_saidas) Then
                lstr_sql = lstr_sql & "(select sum(ifnull([num_valor], 0)) from [tb_movimentacao] where [chr_tipo] = 'S' and [int_conta] = " & lobj_contas(llng_contador)("int_codigo") & " and [dt_pagamento] between '" & Format$(ldt_data_de, pcst_formato_data_sql) & "' and '" & Format$(ldt_data_ate, pcst_formato_data_sql) & "') as [num_valor_saida] "
            End If
            If (lenm_exibir = enm_exibir.val_ambas) Then
                lstr_sql = lstr_sql & "(select sum(ifnull([num_valor], 0)) from [tb_movimentacao] where [chr_tipo] = 'E' and [int_conta] = " & lobj_contas(llng_contador)("int_codigo") & " and [dt_pagamento] between '" & Format$(ldt_data_de, pcst_formato_data_sql) & "' and '" & Format$(ldt_data_ate, pcst_formato_data_sql) & "') as [num_valor_entrada], "
                lstr_sql = lstr_sql & "(select sum(ifnull([num_valor], 0)) from [tb_movimentacao] where [chr_tipo] = 'S' and [int_conta] = " & lobj_contas(llng_contador)("int_codigo") & " and [dt_pagamento] between '" & Format$(ldt_data_de, pcst_formato_data_sql) & "' and '" & Format$(ldt_data_ate, pcst_formato_data_sql) & "') as [num_valor_saida] "
            End If
            If (llng_contador < llng_registros) Then
                lstr_sql = lstr_sql & " union all "
            End If
        Next
        lstr_sql = lstr_sql & " "
    End If
    'caso tenha montado a query SQL corretamente
    If (Trim$(lstr_sql) <> "") Then
        'executa o comando sql e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_valores, lstr_sql, "frm_graficos_geral_conta", "cmd_gerar_Click")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_cmd_gerar_Click
        End If
        'quantidade de registros
        llng_registros = lobj_valores.Count
        'ajusta as configurações do gráfico
        With msc_grafico
            'ajusta o título
            .Title = "Gráfico de Movimentação de Contas por "
            If (lenm_exibir = enm_exibir.val_entradas) Then
                .Title = .Title & "Entradas"
                .RowCount = 1
                .Row = 1
                .RowLabel = "Entradas"
            End If
            If (lenm_exibir = enm_exibir.val_saidas) Then
                .Title = .Title & "Saídas"
                .RowCount = 1
                .Row = 1
                .RowLabel = "Saídas"
            End If
            If (lenm_exibir = enm_exibir.val_ambas) Then
                .Title = .Title & "Entradas e Saídas"
                .RowCount = 2
                .Row = 1
                .RowLabel = "Entradas"
                .Row = 2
                .RowLabel = "Saídas"
            End If
            'ajusta o rodapé
            .FootnoteText = "Período selecionado: de " & Format$(ldt_data_de, pcst_formato_data) & " até " & Format$(ldt_data_ate, pcst_formato_data)
            'tipo de gráfico
            If (lenm_tipo_grafico = enm_tipo_grafico.val_2d_barras) Then
                .chartType = VtChChartType2dBar
            End If
            If (lenm_tipo_grafico = enm_tipo_grafico.val_2d_pizza) Then
                .chartType = VtChChartType2dPie
            End If
            'percorre o objeto de valores
            For llng_contador = 1 To llng_registros
                'ajusta de acordo com os dados retornados
                .ColumnCount = llng_contador
                If (lenm_exibir = enm_exibir.val_entradas) Then
                    .Row = 1
                    .Column = llng_contador
                    .Data = Format$(lobj_valores(llng_contador)("num_valor_entrada"), pcst_formato_numerico)
                    .ColumnLabel = lobj_valores(llng_contador)("str_descricao_conta")
                End If
                If (lenm_exibir = enm_exibir.val_saidas) Then
                    .Row = 1
                    .Column = llng_contador
                    .Data = Format$(lobj_valores(llng_contador)("num_valor_saida"), pcst_formato_numerico)
                    .ColumnLabel = lobj_valores(llng_contador)("str_descricao_conta")
                End If
                If (lenm_exibir = enm_exibir.val_ambas) Then
                    .Column = llng_contador
                    .Row = 1
                    .Data = Format$(lobj_valores(llng_contador)("num_valor_entrada"), pcst_formato_numerico)
                    .ColumnLabel = lobj_valores(llng_contador)("str_descricao_conta")
                    .Row = 2
                    .Data = Format$(lobj_valores(llng_contador)("num_valor_saida"), pcst_formato_numerico)
                    .ColumnLabel = lobj_valores(llng_contador)("str_descricao_conta")
                End If
            Next
        End With
        'ajusta altura da janela
        Me.Height = 8025
    End If
    'limpa a barra de status
    With stb_status
        .Panels.Clear
        .Panels.Add
        .Panels.Item(enm_status.pnl_mensagem).Text = ""
        .Panels.Item(enm_status.pnl_mensagem).AutoSize = sbrSpring
    End With
fim_cmd_gerar_Click:
    'destrói os objetos
    Set lobj_contas = Nothing
    Set lobj_valores = Nothing
    Exit Sub
erro_cmd_gerar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_graficos_geral_conta", "cmd_gerar_Click"
    GoTo fim_cmd_gerar_Click
End Sub

Private Sub Form_Initialize()
    On Error GoTo erro_Form_Initialize
    InitCommonControls
fim_Form_Initialize:
    Exit Sub
erro_Form_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_graficos_geral_conta", "Form_Initialize"
    GoTo fim_Form_Initialize
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/graficos_por_conta.htm", 0
        Case vbKeyF7
            cmd_gerar_Click
        Case vbKeyF8
            cmd_fechar_Click
    End Select
fim_Form_KeyUp:
    Exit Sub
erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_graficos_geral_conta", "Form_KeyUp"
    GoTo fim_Form_KeyUp
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
    psub_ajustar_combos_data dtp_de, dtp_ate
    lsub_preencher_combos
    'ajusta a barra de status
    With stb_status
        .Panels.Clear
        .Panels.Add
        .Panels.Item(enm_status.pnl_mensagem).Text = "Informe os filtros para geração do gráfico."
        .Panels.Item(enm_status.pnl_mensagem).AutoSize = sbrSpring
    End With
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_graficos_geral_conta", "Form_Load"
    GoTo fim_Form_Load
End Sub
