VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_financeiro_agenda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Agenda Financeira"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13095
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
   ScaleHeight     =   9675
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid msf_grade_contas_receber 
      Height          =   1860
      Left            =   90
      TabIndex        =   3
      Top             =   345
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   3281
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf_grade_contas_pagar 
      Height          =   1860
      Left            =   90
      TabIndex        =   7
      Top             =   2580
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   3281
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf_grade_contas 
      Height          =   2160
      Left            =   8940
      TabIndex        =   11
      Top             =   7440
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   3810
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf_grade_ultimas_baixas 
      Height          =   2160
      Left            =   90
      TabIndex        =   10
      Top             =   4830
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   3810
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf_grade_receitas_despesas_fixas 
      Height          =   2160
      Left            =   120
      TabIndex        =   13
      Top             =   7440
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   3810
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lbl_receitas_despesas_fixas 
      AutoSize        =   -1  'True
      Caption         =   "Receitas/Despesas fixas:"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   7140
      Width           =   1815
   End
   Begin VB.Label lbl_ultimas_baixas 
      AutoSize        =   -1  'True
      Caption         =   "Úl&timas 50 baixas (receitas e despesas):"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label lbl_total_pagar_valor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11325
      TabIndex        =   6
      Top             =   2295
      Width           =   1590
   End
   Begin VB.Label lbl_total_pagar 
      AutoSize        =   -1  'True
      Caption         =   "&Total a pagar:"
      Height          =   195
      Left            =   10245
      TabIndex        =   5
      Top             =   2295
      Width           =   1020
   End
   Begin VB.Label lbl_total_receber_valor 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11325
      TabIndex        =   2
      Top             =   75
      Width           =   1590
   End
   Begin VB.Label lbl_total_receber 
      AutoSize        =   -1  'True
      Caption         =   "&Total a receber:"
      Height          =   195
      Left            =   10110
      TabIndex        =   1
      Top             =   75
      Width           =   1155
   End
   Begin VB.Label lbl_contas 
      AutoSize        =   -1  'True
      Caption         =   "&Contas:"
      Height          =   195
      Left            =   8940
      TabIndex        =   9
      Top             =   7140
      Width           =   570
   End
   Begin VB.Label lbl_contas_pagar 
      AutoSize        =   -1  'True
      Caption         =   "C&ontas a pagar:"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   2295
      Width           =   1170
   End
   Begin VB.Label lbl_contas_receber 
      AutoSize        =   -1  'True
      Caption         =   "&Contas a receber:"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   1305
   End
   Begin VB.Menu mnu_msf_grade_contas_receber 
      Caption         =   "Contas a &Receber"
      Visible         =   0   'False
      Begin VB.Menu mnu_msf_grade_contas_receber_copiar 
         Caption         =   "&Copiar conteúdo"
      End
      Begin VB.Menu mnu_msf_grade_contas_receber_exportar 
         Caption         =   "&Exportar para arquivo..."
      End
      Begin VB.Menu mnu_msf_grade_contas_receber_separador 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_msf_grade_contas_receber_cadastro 
         Caption         =   "C&ontas a receber..."
      End
   End
   Begin VB.Menu mnu_msf_grade_contas_pagar 
      Caption         =   "Contas a &Pagar"
      Visible         =   0   'False
      Begin VB.Menu mnu_msf_grade_contas_pagar_copiar 
         Caption         =   "&Copiar conteúdo"
      End
      Begin VB.Menu mnu_msf_grade_contas_pagar_exportar 
         Caption         =   "&Exportar para arquivo..."
      End
      Begin VB.Menu mnu_msf_grade_contas_pagar_separador 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_msf_grade_contas_pagar_cadastro 
         Caption         =   "C&ontas a pagar..."
      End
   End
   Begin VB.Menu mnu_msf_grade_ultimas_baixas 
      Caption         =   "Últimas &Baixas"
      Visible         =   0   'False
      Begin VB.Menu mnu_msf_grade_ultimas_baixas_copiar 
         Caption         =   "&Copiar conteúdo"
      End
      Begin VB.Menu mnu_msf_grade_ultimas_baixas_exportar 
         Caption         =   "&Exportar para arquivo..."
      End
      Begin VB.Menu mnu_msf_grade_ultimas_baixas_separador 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_msf_grade_ultimas_baixas_cadastro 
         Caption         =   "M&ovimentação..."
      End
   End
   Begin VB.Menu mnu_msf_grade_receitas_despesas_fixas 
      Caption         =   "&Receitas/Despesas Fixas"
      Visible         =   0   'False
      Begin VB.Menu mnu_msf_grade_receitas_despesas_fixas_copiar 
         Caption         =   "&Copiar conteúdo"
      End
      Begin VB.Menu mnu_msf_grade_receitas_despesas_fixas_exportar 
         Caption         =   "&Exportar para arquivo..."
      End
      Begin VB.Menu mnu_msf_grade_receitas_despesas_fixas_separador 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_msf_grade_receitas_despesas_fixas_cadastro 
         Caption         =   "M&ovimentação..."
      End
   End
   Begin VB.Menu mnu_msf_grade_contas 
      Caption         =   "&Contas"
      Visible         =   0   'False
      Begin VB.Menu mnu_msf_grade_contas_copiar 
         Caption         =   "&Copiar conteúdo"
      End
      Begin VB.Menu mnu_msf_grade_contas_exportar 
         Caption         =   "&Exportar para arquivo..."
      End
      Begin VB.Menu mnu_msf_grade_contas_separador 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_msf_grade_contas_cadastro 
         Caption         =   "C&ontas..."
      End
   End
End
Attribute VB_Name = "frm_financeiro_agenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enm_colunas_receber
    col_vencimento = 0
    col_baixa_automatica = 1
    col_conta_baixa_automatica = 2
    col_receita = 3
    col_valor = 4
    col_descricao = 5
End Enum

Private Enum enm_colunas_pagar
    col_vencimento = 0
    col_baixa_automatica = 1
    col_conta_baixa_automatica = 2
    col_despesa = 3
    col_valor = 4
    col_descricao = 5
End Enum

Private Enum enm_ultimas_baixas
    col_pagamento = 0
    col_conta = 1
    col_receita_despesa = 2
    col_descricao = 3
    col_valor = 4
End Enum

Private Enum enm_receitas_despesas_fixas
    col_tipo = 0
    col_descricao = 1
    col_situacao = 2
    col_data = 3
    col_valor = 4
End Enum

Private Enum enm_saldo_atual
    col_conta = 0
    col_valor = 1
End Enum

Public Sub Form_Activate()
    On Error GoTo Erro_Form_Activate
    lsub_preencher_grade_contas_receber
    lsub_preencher_grade_contas_pagar
    lsub_preencher_grade_saldos_atual
    lsub_preencher_grade_ultimas_baixas
    lsub_preencher_grade_receitas_despesas_fixas
Fim_Form_Activate:
    Exit Sub
Erro_Form_Activate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_agenda_financeira", "form_activate"
    GoTo Fim_Form_Activate
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            If (Me.Visible) Then
                psub_exibir_ajuda Me, "html/financeiro_agenda_financeira.htm", 0
            Else
                psub_exibir_ajuda Me, "html/menu_principal.htm", 0
            End If
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_agenda_financeira", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
     Form_Activate
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "Form_Load"
    GoTo fim_Form_Load
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo erro_Form_QueryUnload
    'esconde o form
    Me.Hide
    'se o usuário fechou o formulário
    If (UnloadMode = 0) Then
        'impede o form de ser destruído
        Cancel = True
    End If
fim_Form_QueryUnload:
    Exit Sub
erro_Form_QueryUnload:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "Form_QueryUnload"
    GoTo fim_Form_QueryUnload
End Sub

Private Sub lsub_preencher_grade_contas_pagar()
    On Error GoTo erro_lsub_preencher_grade_contas_pagar
    Dim lobj_contas_pagar As Object
    Dim lstr_sql As String
    Dim llng_contador As Long
    Dim llng_registros As Long
    Dim lcur_valor_total_pagar As Currency
    'visibilidade dos componentes
    lbl_total_pagar_valor.Visible = True
    lbl_total_pagar.Visible = True
    'monta a grade de contas a pagar
    With msf_grade_contas_pagar
        .Clear
        .Cols = 6
        .Rows = 2
        .ColWidth(enm_colunas_pagar.col_vencimento) = 1000
        .ColWidth(enm_colunas_pagar.col_baixa_automatica) = 1400
        .ColWidth(enm_colunas_pagar.col_conta_baixa_automatica) = 2150
        .ColWidth(enm_colunas_pagar.col_despesa) = 2790
        .ColWidth(enm_colunas_pagar.col_valor) = 1200
        .ColWidth(enm_colunas_pagar.col_descricao) = 3945
        .TextMatrix(0, enm_colunas_pagar.col_vencimento) = " Vencimento"
        .TextMatrix(0, enm_colunas_pagar.col_baixa_automatica) = " Baixa automática"
        .TextMatrix(0, enm_colunas_pagar.col_conta_baixa_automatica) = " Conta"
        .TextMatrix(0, enm_colunas_pagar.col_despesa) = " Despesa"
        .TextMatrix(0, enm_colunas_pagar.col_valor) = " Valor"
        .TextMatrix(0, enm_colunas_pagar.col_descricao) = " Descrição"
        .ColAlignment(enm_colunas_pagar.col_valor) = flexAlignRightCenter
    End With
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select "
    lstr_sql = lstr_sql & " [tb_despesas].[str_descricao] as [str_descricao_despesa], "
    lstr_sql = lstr_sql & " ifnull([tb_contas].[str_descricao],'--') as [str_descricao_conta], "
    lstr_sql = lstr_sql & " [tb_contas_pagar].* "
    lstr_sql = lstr_sql & " from "
    lstr_sql = lstr_sql & " [tb_contas_pagar] "
    lstr_sql = lstr_sql & " inner join "
    lstr_sql = lstr_sql & " [tb_despesas] on [tb_contas_pagar].[int_despesa] = [tb_despesas].[int_codigo] "
    lstr_sql = lstr_sql & " left join "
    lstr_sql = lstr_sql & " [tb_contas] on [tb_contas_pagar].[int_conta_baixa_automatica] = [tb_contas].[int_codigo] "
    lstr_sql = lstr_sql & " where "
    lstr_sql = lstr_sql & " [dt_vencimento] between '" & Format$((Date - (pfct_retorna_periodo_data())), pcst_formato_data_sql) & "' "
    lstr_sql = lstr_sql & " and '" & Format$((Date + (pfct_retorna_periodo_data())), pcst_formato_data_sql) & "' "
    lstr_sql = lstr_sql & " order by [dt_vencimento] asc "
    If (Not pfct_executar_comando_sql(lobj_contas_pagar, lstr_sql, "frm_agenda_financeira", "lsub_preencher_grade_contas_pagar")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_preencher_grade_contas_pagar
    End If
    'caso existam registros
    llng_registros = lobj_contas_pagar.Count
    If (llng_registros > 0) Then
        msf_grade_contas_pagar.Redraw = False
        For llng_contador = 1 To llng_registros
            msf_grade_contas_pagar.Row = llng_contador
            msf_grade_contas_pagar.Col = enm_colunas_pagar.col_vencimento
            msf_grade_contas_pagar.RowData(llng_contador) = lobj_contas_pagar(llng_contador)("int_codigo")
            msf_grade_contas_pagar.TextMatrix(llng_contador, enm_colunas_pagar.col_vencimento) = " " & Format$(lobj_contas_pagar(llng_contador)("dt_vencimento"), pcst_formato_data)
            msf_grade_contas_pagar.TextMatrix(llng_contador, enm_colunas_pagar.col_baixa_automatica) = " " & IIf(lobj_contas_pagar(llng_contador)("chr_baixa_automatica") = "S", "Sim", "Não")
            msf_grade_contas_pagar.TextMatrix(llng_contador, enm_colunas_pagar.col_conta_baixa_automatica) = " " & lobj_contas_pagar(llng_contador)("str_descricao_conta")
            msf_grade_contas_pagar.TextMatrix(llng_contador, enm_colunas_pagar.col_despesa) = " " & lobj_contas_pagar(llng_contador)("str_descricao_despesa")
            msf_grade_contas_pagar.TextMatrix(llng_contador, enm_colunas_pagar.col_valor) = " " & Format$(lobj_contas_pagar(llng_contador)("num_valor"), pcst_formato_numerico)
            msf_grade_contas_pagar.TextMatrix(llng_contador, enm_colunas_pagar.col_valor) = " " & Format$(lobj_contas_pagar(llng_contador)("num_valor"), pcst_formato_numerico)
            msf_grade_contas_pagar.TextMatrix(llng_contador, enm_colunas_pagar.col_descricao) = " " & lobj_contas_pagar(llng_contador)("str_descricao")
            lcur_valor_total_pagar = lcur_valor_total_pagar + (CCur(lobj_contas_pagar(llng_contador)("num_valor")))
            'se a data de vencimento for menor que a data de hoje, conta está atrasada
            If (CDate(lobj_contas_pagar(llng_contador)("dt_vencimento")) < Date) Then
                'cor da fonte da linha em vermelho
                psub_ajustar_cor_linha_grade msf_grade_contas_pagar, llng_contador, vbRed
            End If
            'se a data de vencimento for maior que a data de hoje, conta ainda vai vencer
            If (CDate(lobj_contas_pagar(llng_contador)("dt_vencimento")) > Date) Then
                'cor da fonte da linha em azul
                psub_ajustar_cor_linha_grade msf_grade_contas_pagar, llng_contador, vbBlue
            End If
            'se a data de vencimento for igual a data de hoje, conta vence no dia
            If (CDate(lobj_contas_pagar(llng_contador)("dt_vencimento")) = Date) Then
                'cor da fonte da linha em preto
                psub_ajustar_cor_linha_grade msf_grade_contas_pagar, llng_contador, vbWindowText
            End If
            'se ainda houver registros
            If (llng_contador < llng_registros) Then
                'adiciona mais uma linha
                msf_grade_contas_pagar.Rows = msf_grade_contas_pagar.Rows + 1
            End If
        Next
        msf_grade_contas_pagar.Redraw = True
        lbl_total_pagar_valor.Caption = Format$(llng_registros, "0000") & " - " & Format$(lcur_valor_total_pagar, pcst_formato_numerico)
        msf_grade_contas_pagar.Row = 1
    Else
        lbl_total_pagar_valor.Visible = False
        lbl_total_pagar.Visible = False
        With msf_grade_contas_pagar
            .Clear
            .Cols = 1
            .Rows = 2
            .ColWidth(enm_colunas_pagar.col_vencimento) = .Width - 100
            .TextMatrix(0, enm_colunas_pagar.col_vencimento) = " Mensagem"
            .TextMatrix(1, enm_colunas_pagar.col_vencimento) = " Não há contas a pagar no período de " & Format$((Date - (pfct_retorna_periodo_data())), pcst_formato_data) & " a " & Format$((Date + (pfct_retorna_periodo_data())), pcst_formato_data)
        End With
        GoTo fim_lsub_preencher_grade_contas_pagar
    End If
fim_lsub_preencher_grade_contas_pagar:
    'destrói os objetos
    Set lobj_contas_pagar = Nothing
    Exit Sub
erro_lsub_preencher_grade_contas_pagar:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_agenda_financeira", "lsub_preencher_grade_contas_pagar"
    GoTo fim_lsub_preencher_grade_contas_pagar
    Resume 0
End Sub

Private Sub lsub_preencher_grade_contas_receber()
    On Error GoTo erro_lsub_preencher_grade_contas_receber
    Dim lobj_contas_receber As Object
    Dim lstr_sql As String
    Dim llng_contador As Long
    Dim llng_registros As Long
    Dim lcur_valor_total_receber As Currency
    'visibilidade dos componentes
    lbl_total_receber_valor.Visible = True
    lbl_total_receber.Visible = True
    'monta a grade de contas a receber
    With msf_grade_contas_receber
        .Clear
        .Cols = 6
        .Rows = 2
        .ColWidth(enm_colunas_receber.col_vencimento) = 1000
        .ColWidth(enm_colunas_receber.col_baixa_automatica) = 1400
        .ColWidth(enm_colunas_receber.col_conta_baixa_automatica) = 2150
        .ColWidth(enm_colunas_receber.col_receita) = 2790
        .ColWidth(enm_colunas_receber.col_valor) = 1200
        .ColWidth(enm_colunas_receber.col_descricao) = 3945
        .TextMatrix(0, enm_colunas_receber.col_vencimento) = " Vencimento"
        .TextMatrix(0, enm_colunas_receber.col_baixa_automatica) = " Baixa automática"
        .TextMatrix(0, enm_colunas_receber.col_conta_baixa_automatica) = " Conta"
        .TextMatrix(0, enm_colunas_receber.col_receita) = " Receita"
        .TextMatrix(0, enm_colunas_receber.col_valor) = " Valor"
        .TextMatrix(0, enm_colunas_receber.col_descricao) = " Descrição"
        .ColAlignment(enm_colunas_receber.col_valor) = flexAlignRightCenter
    End With
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select "
    lstr_sql = lstr_sql & " [tb_receitas].[str_descricao] as [str_descricao_receita], "
    lstr_sql = lstr_sql & " ifnull([tb_contas].[str_descricao],'--') as [str_descricao_conta], "
    lstr_sql = lstr_sql & " [tb_contas_receber].* "
    lstr_sql = lstr_sql & " from "
    lstr_sql = lstr_sql & " [tb_contas_receber] "
    lstr_sql = lstr_sql & " inner join "
    lstr_sql = lstr_sql & " [tb_receitas] on [tb_contas_receber].[int_receita] = [tb_receitas].[int_codigo] "
    lstr_sql = lstr_sql & " left join "
    lstr_sql = lstr_sql & " [tb_contas] on [tb_contas_receber].[int_conta_baixa_automatica] = [tb_contas].[int_codigo] "
    lstr_sql = lstr_sql & " where "
    lstr_sql = lstr_sql & " [dt_vencimento] between '" & Format$((Date - (pfct_retorna_periodo_data())), pcst_formato_data_sql) & "' "
    lstr_sql = lstr_sql & " and '" & Format$((Date + (pfct_retorna_periodo_data())), pcst_formato_data_sql) & "' "
    lstr_sql = lstr_sql & " order by [dt_vencimento] asc "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_contas_receber, lstr_sql, "frm_agenda_financeira", "lsub_preencher_grade_contas_receber")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_preencher_grade_contas_receber
    End If
    'caso existam registros
    llng_registros = lobj_contas_receber.Count
    If (llng_registros > 0) Then
        msf_grade_contas_receber.Redraw = False
        For llng_contador = 1 To llng_registros
            msf_grade_contas_receber.Row = llng_contador
            msf_grade_contas_receber.Col = enm_colunas_receber.col_vencimento
            msf_grade_contas_receber.RowData(llng_contador) = lobj_contas_receber(llng_contador)("int_codigo")
            msf_grade_contas_receber.TextMatrix(llng_contador, enm_colunas_receber.col_vencimento) = " " & Format$(lobj_contas_receber(llng_contador)("dt_vencimento"), pcst_formato_data)
            msf_grade_contas_receber.TextMatrix(llng_contador, enm_colunas_receber.col_baixa_automatica) = " " & IIf(lobj_contas_receber(llng_contador)("chr_baixa_automatica") = "S", "Sim", "Não")
            msf_grade_contas_receber.TextMatrix(llng_contador, enm_colunas_receber.col_conta_baixa_automatica) = " " & lobj_contas_receber(llng_contador)("str_descricao_conta")
            msf_grade_contas_receber.TextMatrix(llng_contador, enm_colunas_receber.col_receita) = " " & lobj_contas_receber(llng_contador)("str_descricao_receita")
            msf_grade_contas_receber.TextMatrix(llng_contador, enm_colunas_receber.col_valor) = " " & Format$(lobj_contas_receber(llng_contador)("num_valor"), pcst_formato_numerico)
            msf_grade_contas_receber.TextMatrix(llng_contador, enm_colunas_receber.col_descricao) = " " & lobj_contas_receber(llng_contador)("str_descricao")
            lcur_valor_total_receber = lcur_valor_total_receber + (CCur(lobj_contas_receber(llng_contador)("num_valor")))
            'se a data de vencimento for menor que a data de hoje, conta está atrasada
            If (CDate(lobj_contas_receber(llng_contador)("dt_vencimento")) < Date) Then
                'cor da fonte da linha em vermelho
                psub_ajustar_cor_linha_grade msf_grade_contas_receber, llng_contador, vbRed
            End If
            'se a data de vencimento for maior que a data de hoje, conta ainda vai vencer
            If (CDate(lobj_contas_receber(llng_contador)("dt_vencimento")) > Date) Then
                'cor da fonte da linha em azul
                psub_ajustar_cor_linha_grade msf_grade_contas_receber, llng_contador, vbBlue
            End If
            'se a data de vencimento for igual a data de hoje, conta vence no dia
            If (CDate(lobj_contas_receber(llng_contador)("dt_vencimento")) = Date) Then
                'cor da fonte da linha em preto
                psub_ajustar_cor_linha_grade msf_grade_contas_receber, llng_contador, vbWindowText
            End If
            'se ainda houver registros
            If (llng_contador < llng_registros) Then
                'adiciona mais uma linha
                msf_grade_contas_receber.Rows = msf_grade_contas_receber.Rows + 1
            End If
        Next
        msf_grade_contas_receber.Redraw = True
        lbl_total_receber_valor.Caption = Format$(llng_registros, "0000") & " - " & Format$(lcur_valor_total_receber, pcst_formato_numerico)
        msf_grade_contas_receber.Row = 1
    Else
        lbl_total_receber_valor.Visible = False
        lbl_total_receber.Visible = False
        With msf_grade_contas_receber
            .Clear
            .Cols = 1
            .Rows = 2
            .ColWidth(enm_colunas_receber.col_vencimento) = .Width - 100
            .TextMatrix(0, enm_colunas_receber.col_vencimento) = " Mensagem"
            .TextMatrix(1, enm_colunas_receber.col_vencimento) = " Não há contas a receber no período de " & Format$((Date - (pfct_retorna_periodo_data())), pcst_formato_data) & " a " & Format$((Date + (pfct_retorna_periodo_data())), pcst_formato_data)
        End With
        GoTo fim_lsub_preencher_grade_contas_receber
    End If
fim_lsub_preencher_grade_contas_receber:
    'destrói os objetos
    Set lobj_contas_receber = Nothing
    Exit Sub
erro_lsub_preencher_grade_contas_receber:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_agenda_financeira", "lsub_preencher_grade_contas_receber"
    GoTo fim_lsub_preencher_grade_contas_receber
    Resume 0
End Sub

Private Sub lsub_preencher_grade_receitas_despesas_fixas()
    On Error GoTo erro_lsub_preencher_grade_receitas_despesas_fixas
    Dim lobj_receitas_despesas_fixas As Object
    Dim lstr_sql As String
    Dim llng_contador As Long
    Dim llng_registros As Long
    'ajustamos o texto do rótulo
    lbl_receitas_despesas_fixas.Caption = Replace$("Receitas/Despesas fixas (mês mm/aaaa):", "mm/aaaa", Format$(Now, pcst_formato_mes_ano))
    'monta a grade de últimas baixas
    With msf_grade_receitas_despesas_fixas
        .Clear
        .Cols = 5
        .Rows = 2
        .ColWidth(enm_receitas_despesas_fixas.col_tipo) = 1000
        .ColWidth(enm_receitas_despesas_fixas.col_descricao) = 3945
        .ColWidth(enm_receitas_despesas_fixas.col_situacao) = 1000
        .ColWidth(enm_receitas_despesas_fixas.col_data) = 1200
        .ColWidth(enm_receitas_despesas_fixas.col_valor) = 1200
        .TextMatrix(0, enm_receitas_despesas_fixas.col_tipo) = " Tipo"
        .TextMatrix(0, enm_receitas_despesas_fixas.col_descricao) = " Receita/Despesa"
        .TextMatrix(0, enm_receitas_despesas_fixas.col_situacao) = " Situação"
        .TextMatrix(0, enm_receitas_despesas_fixas.col_data) = " Data"
        .TextMatrix(0, enm_receitas_despesas_fixas.col_valor) = " Valor"
        .ColAlignment(enm_receitas_despesas_fixas.col_valor) = flexAlignRightCenter
    End With
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & " case "
    lstr_sql = lstr_sql & "     when [receitas_despesas_fixas].[chr_tipo] = 'E' then 'Entrada' "
    lstr_sql = lstr_sql & "     when [receitas_despesas_fixas].[chr_tipo] = 'S' then 'Saída' "
    lstr_sql = lstr_sql & " end "
    lstr_sql = lstr_sql & " ) as [str_tipo], "
    lstr_sql = lstr_sql & " [receitas_despesas_fixas].[chr_tipo], "
    lstr_sql = lstr_sql & " [receitas_despesas_fixas].[str_descricao], "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & " case "
    lstr_sql = lstr_sql & "     when ([tb_contas_pagar].dt_vencimento is not null or [tb_contas_receber].dt_vencimento is not null) then 'AGENDADO' "
    lstr_sql = lstr_sql & "     when [tb_movimentacao].[dt_pagamento] is null then 'PENDENTE' "
    lstr_sql = lstr_sql & "     when [receitas_despesas_fixas].[chr_tipo] = 'E' and [tb_movimentacao].[dt_pagamento] is not null then 'RECEBIDO' "
    lstr_sql = lstr_sql & "     when [receitas_despesas_fixas].[chr_tipo] = 'S' and [tb_movimentacao].[dt_pagamento] is not null then 'PAGO' "
    lstr_sql = lstr_sql & " end "
    lstr_sql = lstr_sql & " ) as [str_situacao], "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & " case "
    lstr_sql = lstr_sql & "     when [tb_contas_pagar].[int_despesa] is not null then [tb_contas_pagar].[dt_vencimento] "
    lstr_sql = lstr_sql & "     when [tb_contas_receber].[int_receita] is not null then [tb_contas_receber].[dt_vencimento] "
    lstr_sql = lstr_sql & "     else [tb_movimentacao].[dt_pagamento] "
    lstr_sql = lstr_sql & " end "
    lstr_sql = lstr_sql & " ) [dt_pagamento], "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & " case "
    lstr_sql = lstr_sql & "     when [tb_contas_pagar].[int_despesa] is not null then [tb_contas_pagar].[num_valor] "
    lstr_sql = lstr_sql & "     when [tb_contas_receber].[int_receita] is not null then [tb_contas_receber].[num_valor] "
    lstr_sql = lstr_sql & "     else [tb_movimentacao].[num_valor] "
    lstr_sql = lstr_sql & " end "
    lstr_sql = lstr_sql & " ) [num_valor] "
    lstr_sql = lstr_sql & " from "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & " select "
    lstr_sql = lstr_sql & "     [int_codigo], "
    lstr_sql = lstr_sql & "     'E' as [chr_tipo], "
    lstr_sql = lstr_sql & "     [str_descricao] "
    lstr_sql = lstr_sql & " from "
    lstr_sql = lstr_sql & "     [tb_receitas] "
    lstr_sql = lstr_sql & " where 1 = 1 "
    lstr_sql = lstr_sql & "     and [chr_ativo] = 'S' "
    lstr_sql = lstr_sql & "     and [chr_fixa] = 'S' "
    lstr_sql = lstr_sql & " union "
    lstr_sql = lstr_sql & " select "
    lstr_sql = lstr_sql & "     [int_codigo], "
    lstr_sql = lstr_sql & "     'S' as [chr_tipo], "
    lstr_sql = lstr_sql & "     [str_descricao] "
    lstr_sql = lstr_sql & " from "
    lstr_sql = lstr_sql & "     [tb_despesas] "
    lstr_sql = lstr_sql & " where 1 = 1 "
    lstr_sql = lstr_sql & "     and [chr_ativo] = 'S' "
    lstr_sql = lstr_sql & "     and [chr_fixa] = 'S' "
    lstr_sql = lstr_sql & " ) [receitas_despesas_fixas] "
    lstr_sql = lstr_sql & " left join "
    lstr_sql = lstr_sql & " [tb_movimentacao] on "
    lstr_sql = lstr_sql & "         [tb_movimentacao].[chr_tipo] = [receitas_despesas_fixas].[chr_tipo] "
    lstr_sql = lstr_sql & "     and "
    lstr_sql = lstr_sql & "     ( "
    lstr_sql = lstr_sql & "         case "
    lstr_sql = lstr_sql & "             when [receitas_despesas_fixas].[chr_tipo] = 'E' then [tb_movimentacao].[int_receita] = [receitas_despesas_fixas].[int_codigo] "
    lstr_sql = lstr_sql & "             when [receitas_despesas_fixas].[chr_tipo] = 'S' then [tb_movimentacao].[int_despesa] = [receitas_despesas_fixas].[int_codigo] "
    lstr_sql = lstr_sql & "         end "
    lstr_sql = lstr_sql & "     ) "
    lstr_sql = lstr_sql & "     and [tb_movimentacao].[dt_pagamento] between '" & Format$(pfct_retorna_primeiro_dia_mes(Now), pcst_formato_data_sql) & "' and '" & Format$(pfct_retorna_ultimo_dia_mes(Now), pcst_formato_data_sql) & "' "
    lstr_sql = lstr_sql & " left join "
    lstr_sql = lstr_sql & " [tb_contas_pagar] on "
    lstr_sql = lstr_sql & "         [tb_contas_pagar].[int_despesa] = [receitas_despesas_fixas].[int_codigo] "
    lstr_sql = lstr_sql & "     and [tb_contas_pagar].chr_baixa_automatica = 'S' "
    lstr_sql = lstr_sql & "     and [tb_contas_pagar].dt_vencimento between '" & Format$(pfct_retorna_primeiro_dia_mes(Now), pcst_formato_data_sql) & "' and '" & Format$(pfct_retorna_ultimo_dia_mes(Now), pcst_formato_data_sql) & "' "
    lstr_sql = lstr_sql & "     and [receitas_despesas_fixas].[chr_tipo] = 'S' "
    lstr_sql = lstr_sql & " left join "
    lstr_sql = lstr_sql & "     [tb_contas_receber] on "
    lstr_sql = lstr_sql & "             [tb_contas_receber].[int_receita] = [receitas_despesas_fixas].[int_codigo] "
    lstr_sql = lstr_sql & "         and [tb_contas_receber].chr_baixa_automatica = 'S' "
    lstr_sql = lstr_sql & "         and [tb_contas_receber].dt_vencimento between '" & Format$(pfct_retorna_primeiro_dia_mes(Now), pcst_formato_data_sql) & "' and '" & Format$(pfct_retorna_ultimo_dia_mes(Now), pcst_formato_data_sql) & "' "
    lstr_sql = lstr_sql & "         and [receitas_despesas_fixas].[chr_tipo] = 'E' "
    lstr_sql = lstr_sql & " order by "
    lstr_sql = lstr_sql & " [receitas_despesas_fixas].[chr_tipo] asc, "
    lstr_sql = lstr_sql & " [receitas_despesas_fixas].[str_descricao] asc "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_receitas_despesas_fixas, lstr_sql, "frm_agenda_financeira", "lsub_preencher_grade_receitas_despesas_fixas")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_preencher_grade_receitas_despesas_fixas
    End If
    'caso existam registros
    llng_registros = lobj_receitas_despesas_fixas.Count
    If (llng_registros > 0) Then
        msf_grade_receitas_despesas_fixas.Redraw = False
        For llng_contador = 1 To llng_registros
            msf_grade_receitas_despesas_fixas.Row = llng_contador
            'tipo
            msf_grade_receitas_despesas_fixas.TextMatrix(llng_contador, enm_receitas_despesas_fixas.col_tipo) = " " & lobj_receitas_despesas_fixas(llng_contador)("str_tipo")
            'receita/despesa
            msf_grade_receitas_despesas_fixas.TextMatrix(llng_contador, enm_receitas_despesas_fixas.col_descricao) = " " & lobj_receitas_despesas_fixas(llng_contador)("str_descricao")
            'situação
            msf_grade_receitas_despesas_fixas.TextMatrix(llng_contador, enm_receitas_despesas_fixas.col_situacao) = " " & lobj_receitas_despesas_fixas(llng_contador)("str_situacao")
            'data
            If (Not IsNull(lobj_receitas_despesas_fixas(llng_contador)("dt_pagamento"))) Then
                msf_grade_receitas_despesas_fixas.TextMatrix(llng_contador, enm_receitas_despesas_fixas.col_data) = " " & Format$(lobj_receitas_despesas_fixas(llng_contador)("dt_pagamento"), pcst_formato_data)
            End If
            'valor
            If (Not IsNull(lobj_receitas_despesas_fixas(llng_contador)("num_valor"))) Then
                msf_grade_receitas_despesas_fixas.TextMatrix(llng_contador, enm_receitas_despesas_fixas.col_valor) = " " & IIf(lobj_receitas_despesas_fixas(llng_contador)("chr_tipo") = "E", "+", "-") & Format$(lobj_receitas_despesas_fixas(llng_contador)("num_valor"), pcst_formato_numerico)
            End If
            'se está pendente
            If (UCase$(lobj_receitas_despesas_fixas(llng_contador)("str_situacao")) = "PENDENTE") Then
                'cor da fonte da linha em vermelho
                psub_ajustar_cor_linha_grade msf_grade_receitas_despesas_fixas, llng_contador, vbRed
            'se está agendado
            ElseIf (UCase$(lobj_receitas_despesas_fixas(llng_contador)("str_situacao")) = "AGENDADO") Then
                'cor da fonte da linha em preto
                psub_ajustar_cor_linha_grade msf_grade_receitas_despesas_fixas, llng_contador, vbButtonText
            'se está pago/recebido
            Else
                'cor da fonte da linha em azul
                psub_ajustar_cor_linha_grade msf_grade_receitas_despesas_fixas, llng_contador, vbBlue
            End If
            'se ainda houver registros
            If (llng_contador < llng_registros) Then
                'adiciona mais uma linha
                msf_grade_receitas_despesas_fixas.Rows = msf_grade_receitas_despesas_fixas.Rows + 1
            End If
        Next
        msf_grade_receitas_despesas_fixas.Redraw = True
        msf_grade_receitas_despesas_fixas.Row = 1
    Else
        With msf_grade_receitas_despesas_fixas
            .Clear
            .Cols = 1
            .Rows = 2
            .ColWidth(enm_ultimas_baixas.col_pagamento) = .Width - 100
            .TextMatrix(0, enm_ultimas_baixas.col_pagamento) = " Mensagem"
            .TextMatrix(1, enm_ultimas_baixas.col_pagamento) = " Não há receitas/despesas marcadas como fixas. "
        End With
        GoTo fim_lsub_preencher_grade_receitas_despesas_fixas
    End If
fim_lsub_preencher_grade_receitas_despesas_fixas:
    Set lobj_receitas_despesas_fixas = Nothing
    Exit Sub
erro_lsub_preencher_grade_receitas_despesas_fixas:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "lsub_preencher_grade_receitas_despesas_fixas"
    GoTo fim_lsub_preencher_grade_receitas_despesas_fixas
End Sub

Private Sub lsub_preencher_grade_saldos_atual()
    On Error GoTo erro_lsub_preencher_grade_saldos_atual
    Dim lobj_saldos As Object
    Dim lstr_sql As String
    Dim llng_contador As Long
    Dim llng_registros As Long
    Dim lcur_saldo_total As Currency
    'configura a grade de contas
    With msf_grade_contas
        .Clear
        .Cols = 2
        .Rows = 2
        .ColWidth(enm_saldo_atual.col_conta) = 2200
        .ColWidth(enm_saldo_atual.col_valor) = 1500
        .ColAlignment(enm_saldo_atual.col_valor) = flexAlignRightCenter
        .TextMatrix(0, enm_saldo_atual.col_conta) = " Conta"
        .TextMatrix(0, enm_saldo_atual.col_valor) = " Saldo atual"
    End With
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select * from [tb_contas] where "
    lstr_sql = lstr_sql & " [chr_ativo] = 'S' "
    lstr_sql = lstr_sql & " order by [str_descricao] asc"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_saldos, lstr_sql, "frm_agenda_financeira", "lsub_preencher_grade_saldos_atual")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_preencher_grade_saldos_atual
    End If
    'caso existam registros
    llng_registros = lobj_saldos.Count
    If (llng_registros > 0) Then
        msf_grade_contas.Redraw = False
        For llng_contador = 1 To llng_registros
            msf_grade_contas.Row = llng_contador
            msf_grade_contas.Col = enm_saldo_atual.col_conta
            msf_grade_contas.RowData(llng_contador) = lobj_saldos(llng_contador)("int_codigo")
            msf_grade_contas.TextMatrix(llng_contador, enm_saldo_atual.col_conta) = " " & lobj_saldos(llng_contador)("str_descricao")
            msf_grade_contas.TextMatrix(llng_contador, enm_saldo_atual.col_valor) = " " & Format$(lobj_saldos(llng_contador)("num_saldo"), pcst_formato_numerico)
            lcur_saldo_total = lcur_saldo_total + (CCur(lobj_saldos(llng_contador)("num_saldo")))
            msf_grade_contas.Row = llng_contador
            msf_grade_contas.Col = 1
            'se o saldo da conta for negativo
            If (CCur(lobj_saldos(llng_contador)("num_saldo")) < 0) Then
                'cor da fonte da linha em vermelho
                psub_ajustar_cor_linha_grade msf_grade_contas, llng_contador, vbRed
            End If
            'se o saldo da conta for positivo
            If (CCur(lobj_saldos(llng_contador)("num_saldo")) > 0) Then
                'cor da fonte da linha em azul
                psub_ajustar_cor_linha_grade msf_grade_contas, llng_contador, vbBlue
            End If
            'se o saldo da conta for zero
            If (CCur(lobj_saldos(llng_contador)("num_saldo")) = 0) Then
                'cor da fonte da linha em preto
                psub_ajustar_cor_linha_grade msf_grade_contas, llng_contador, vbWindowText
            End If
            'se ainda houver registros
            If (llng_contador < llng_registros) Then
                'adiciona mais uma linha
                msf_grade_contas.Rows = msf_grade_contas.Rows + 1
            End If
            'insere linha com os totais
            If (llng_contador = llng_registros) Then
                'incrementa mais uma linha
                msf_grade_contas.Rows = msf_grade_contas.Rows + 1
                'incrementa mais um no contador
                llng_contador = llng_contador + 1
                'atribui os valores totais
                msf_grade_contas.RowData(llng_contador) = 99999
                msf_grade_contas.TextMatrix(llng_contador, enm_saldo_atual.col_conta) = " -- TOTAL -- "
                msf_grade_contas.TextMatrix(llng_contador, enm_saldo_atual.col_valor) = " " & Format$(lcur_saldo_total, pcst_formato_numerico)
                'se o total da soma for negativo
                If (lcur_saldo_total < 0) Then
                    'cor da fonte da linha em vermelho
                    psub_ajustar_cor_linha_grade msf_grade_contas, llng_contador, vbRed
                End If
                'se o saldo da conta for positivo
                If (lcur_saldo_total > 0) Then
                    'cor da fonte da linha em azul
                    psub_ajustar_cor_linha_grade msf_grade_contas, llng_contador, vbBlue
                End If
                'se o saldo da conta for zero
                If (lcur_saldo_total = 0) Then
                    'cor da fonte da linha em preto
                    psub_ajustar_cor_linha_grade msf_grade_contas, llng_contador, vbWindowText
                End If
            End If
        Next
        msf_grade_contas.Redraw = True
        msf_grade_contas.Col = 0
        msf_grade_contas.Row = 1
    Else
        With msf_grade_contas
            .Clear
            .Cols = 1
            .Rows = 2
            .ColWidth(enm_saldo_atual.col_conta) = .Width - 100
            .TextMatrix(0, enm_saldo_atual.col_conta) = " Mensagem"
            .TextMatrix(1, enm_saldo_atual.col_conta) = " Não há contas cadastradas."
        End With
        GoTo fim_lsub_preencher_grade_saldos_atual
    End If
fim_lsub_preencher_grade_saldos_atual:
    'destrói os objetos
    Set lobj_saldos = Nothing
    Exit Sub
erro_lsub_preencher_grade_saldos_atual:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_agenda_financeira", "lsub_preencher_grade_saldos_atual"
    GoTo fim_lsub_preencher_grade_saldos_atual
End Sub

Private Sub lsub_preencher_grade_ultimas_baixas()
    On Error GoTo Erro_lsub_preencher_grade_ultimas_baixas
    Dim lobj_ultimas_baixas As Object
    Dim lstr_sql As String
    Dim llng_contador As Long
    Dim llng_registros As Long
    'monta a grade de últimas baixas
    With msf_grade_ultimas_baixas
        .Clear
        .Cols = 5
        .Rows = 2
        .ColWidth(enm_ultimas_baixas.col_pagamento) = 1200
        .ColWidth(enm_ultimas_baixas.col_conta) = 2205
        .ColWidth(enm_ultimas_baixas.col_receita_despesa) = 3940
        .ColWidth(enm_ultimas_baixas.col_descricao) = 3940
        .ColWidth(enm_ultimas_baixas.col_valor) = 1200
        .TextMatrix(0, enm_ultimas_baixas.col_pagamento) = " Pagamento"
        .TextMatrix(0, enm_ultimas_baixas.col_conta) = " Conta"
        .TextMatrix(0, enm_ultimas_baixas.col_receita_despesa) = " Receita/Despesa"
        .TextMatrix(0, enm_ultimas_baixas.col_descricao) = " Descrição"
        .TextMatrix(0, enm_ultimas_baixas.col_valor) = " Valor"
        .ColAlignment(enm_ultimas_baixas.col_valor) = flexAlignRightCenter
    End With
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select "
    lstr_sql = lstr_sql & "     [tb_movimentacao].[int_codigo], "
    lstr_sql = lstr_sql & "     [tb_movimentacao].[chr_tipo], "
    lstr_sql = lstr_sql & "     [tb_movimentacao].[dt_vencimento], "
    lstr_sql = lstr_sql & "     [tb_movimentacao].[dt_pagamento], "
    lstr_sql = lstr_sql & "     [tb_contas].[str_descricao] AS [str_descricao_conta], "
    lstr_sql = lstr_sql & "     ( "
    lstr_sql = lstr_sql & "         case "
    lstr_sql = lstr_sql & "             when [tb_movimentacao].[chr_tipo] = 'E' then "
    lstr_sql = lstr_sql & "                 (select [str_descricao] from [tb_receitas] where [tb_receitas].[int_codigo] = [tb_movimentacao].[int_receita]) "
    lstr_sql = lstr_sql & "             when [tb_movimentacao].[chr_tipo] = 'S' then "
    lstr_sql = lstr_sql & "                 (select [str_descricao] from [tb_despesas] where [tb_despesas].[int_codigo] = [tb_movimentacao].[int_despesa]) "
    lstr_sql = lstr_sql & "         end "
    lstr_sql = lstr_sql & "     ) as [str_receita_despesa], "
    lstr_sql = lstr_sql & "     [tb_movimentacao].[str_descricao], "
    lstr_sql = lstr_sql & "     [tb_movimentacao].[num_valor] "
    lstr_sql = lstr_sql & " from "
    lstr_sql = lstr_sql & "     [tb_movimentacao] "
    lstr_sql = lstr_sql & " inner join "
    lstr_sql = lstr_sql & "     [tb_contas] on "
    lstr_sql = lstr_sql & "         [tb_contas].[int_codigo] = [tb_movimentacao].[int_conta] "
    lstr_sql = lstr_sql & " order by "
    lstr_sql = lstr_sql & "     [tb_movimentacao].[int_codigo] desc "
    lstr_sql = lstr_sql & " limit 50 "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_ultimas_baixas, lstr_sql, "frm_agenda_financeira", "lsub_preencher_grade_ultimas_baixas")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo Fim_lsub_preencher_grade_ultimas_baixas
    End If
    'caso existam registros
    llng_registros = lobj_ultimas_baixas.Count
    If (llng_registros > 0) Then
        msf_grade_ultimas_baixas.Redraw = False
        For llng_contador = 1 To llng_registros
            msf_grade_ultimas_baixas.Row = llng_contador
            msf_grade_ultimas_baixas.Col = enm_ultimas_baixas.col_pagamento
            msf_grade_ultimas_baixas.RowData(llng_contador) = lobj_ultimas_baixas(llng_contador)("int_codigo")
            msf_grade_ultimas_baixas.TextMatrix(llng_contador, enm_ultimas_baixas.col_pagamento) = " " & Format$(lobj_ultimas_baixas(llng_contador)("dt_pagamento"), pcst_formato_data)
            msf_grade_ultimas_baixas.TextMatrix(llng_contador, enm_ultimas_baixas.col_conta) = " " & lobj_ultimas_baixas(llng_contador)("str_descricao_conta")
            msf_grade_ultimas_baixas.TextMatrix(llng_contador, enm_ultimas_baixas.col_valor) = " " & IIf(lobj_ultimas_baixas(llng_contador)("chr_tipo") = "E", "+", "-") & Format$(lobj_ultimas_baixas(llng_contador)("num_valor"), pcst_formato_numerico)
            msf_grade_ultimas_baixas.TextMatrix(llng_contador, enm_ultimas_baixas.col_descricao) = " " & lobj_ultimas_baixas(llng_contador)("str_descricao")
            msf_grade_ultimas_baixas.TextMatrix(llng_contador, enm_ultimas_baixas.col_receita_despesa) = " " & lobj_ultimas_baixas(llng_contador)("str_receita_despesa")
            'se a data de pagamento for maior que a data de vencimento
            If (CDate(lobj_ultimas_baixas(llng_contador)("dt_pagamento")) > CDate(lobj_ultimas_baixas(llng_contador)("dt_vencimento"))) Then
                'cor da fonte da linha em vermelho
                psub_ajustar_cor_linha_grade msf_grade_ultimas_baixas, llng_contador, vbRed
            End If
            'se a data de pagamento for menor que a data de vencimento
            If (CDate(lobj_ultimas_baixas(llng_contador)("dt_pagamento")) < CDate(lobj_ultimas_baixas(llng_contador)("dt_vencimento"))) Then
                'cor da fonte da linha em azul
                psub_ajustar_cor_linha_grade msf_grade_ultimas_baixas, llng_contador, vbBlue
            End If
            'se a data de pagamento for igual à data de vencimento
            If (CDate(lobj_ultimas_baixas(llng_contador)("dt_pagamento")) = CDate(lobj_ultimas_baixas(llng_contador)("dt_vencimento"))) Then
                'cor da fonte da linha em preto
                psub_ajustar_cor_linha_grade msf_grade_ultimas_baixas, llng_contador, vbWindowText
            End If
            'se ainda houver registros
            If (llng_contador < llng_registros) Then
                'adiciona mais uma linha
                msf_grade_ultimas_baixas.Rows = msf_grade_ultimas_baixas.Rows + 1
            End If
        Next
        msf_grade_ultimas_baixas.Redraw = True
        msf_grade_ultimas_baixas.Row = 1
    Else
        With msf_grade_ultimas_baixas
            .Clear
            .Cols = 1
            .Rows = 2
            .ColWidth(enm_ultimas_baixas.col_pagamento) = .Width - 100
            .TextMatrix(0, enm_ultimas_baixas.col_pagamento) = " Mensagem"
            .TextMatrix(1, enm_ultimas_baixas.col_pagamento) = " Não há baixas realizadas. "
        End With
        GoTo Fim_lsub_preencher_grade_ultimas_baixas
    End If
Fim_lsub_preencher_grade_ultimas_baixas:
    Set lobj_ultimas_baixas = Nothing
    Exit Sub
Erro_lsub_preencher_grade_ultimas_baixas:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "lsub_preencher_grade_ultimas_baixas"
    GoTo Fim_lsub_preencher_grade_ultimas_baixas
End Sub

Private Sub mnu_msf_grade_contas_cadastro_Click()
    On Error GoTo erro_mnu_msf_grade_contas_cadastro_Click
    'mesma lógica do menu principal
    frm_cadastro_contas.Show
    frm_cadastro_contas.ZOrder 0
fim_mnu_msf_grade_contas_cadastro_Click:
    Exit Sub
erro_mnu_msf_grade_contas_cadastro_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_contas_cadastro_Click"
    GoTo fim_mnu_msf_grade_contas_cadastro_Click
End Sub

Private Sub mnu_msf_grade_contas_copiar_Click()
    On Error GoTo erro_mnu_msf_grade_contas_copiar_Click
    pfct_copiar_conteudo_grade msf_grade_contas
fim_mnu_msf_grade_contas_copiar_Click:
    Exit Sub
erro_mnu_msf_grade_contas_copiar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_contas_copiar_Click"
    GoTo fim_mnu_msf_grade_contas_copiar_Click
End Sub

Private Sub mnu_msf_grade_contas_exportar_Click()
    On Error GoTo erro_mnu_msf_grade_contas_exportar_Click
    pfct_exportar_conteudo_grade msf_grade_contas, "saldo_atual_contas"
fim_mnu_msf_grade_contas_exportar_Click:
    Exit Sub
erro_mnu_msf_grade_contas_exportar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_contas_exportar_Click"
    GoTo fim_mnu_msf_grade_contas_exportar_Click
End Sub

Private Sub mnu_msf_grade_contas_pagar_cadastro_Click()
    On Error GoTo erro_mnu_msf_grade_contas_pagar_cadastro_Click
    'mesma lógica do menu principal
    frm_cadastro_contas_pagar.Show
    frm_cadastro_contas_pagar.ZOrder 0
fim_mnu_msf_grade_contas_pagar_cadastro_Click:
    Exit Sub
erro_mnu_msf_grade_contas_pagar_cadastro_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_contas_pagar_cadastro_Click"
    GoTo fim_mnu_msf_grade_contas_pagar_cadastro_Click
End Sub

Private Sub mnu_msf_grade_contas_pagar_copiar_Click()
    On Error GoTo erro_mnu_msf_grade_contas_pagar_copiar_Click
    pfct_copiar_conteudo_grade msf_grade_contas_pagar
fim_mnu_msf_grade_contas_pagar_copiar_Click:
    Exit Sub
erro_mnu_msf_grade_contas_pagar_copiar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_contas_pagar_copiar_Click"
    GoTo fim_mnu_msf_grade_contas_pagar_copiar_Click
End Sub

Private Sub mnu_msf_grade_contas_pagar_exportar_Click()
    On Error GoTo erro_mnu_msf_grade_contas_pagar_exportar_Click
    pfct_exportar_conteudo_grade msf_grade_contas_pagar, "contas_a_pagar"
fim_mnu_msf_grade_contas_pagar_exportar_Click:
    Exit Sub
erro_mnu_msf_grade_contas_pagar_exportar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_contas_pagar_exportar_Click"
    GoTo fim_mnu_msf_grade_contas_pagar_exportar_Click
End Sub

Private Sub mnu_msf_grade_contas_receber_cadastro_Click()
    On Error GoTo erro_mnu_msf_grade_contas_receber_cadastro_Click
    'mesma lógica do menu principal
    frm_cadastro_contas_receber.Show
    frm_cadastro_contas_receber.ZOrder 0
fim_mnu_msf_grade_contas_receber_cadastro_Click:
    Exit Sub
erro_mnu_msf_grade_contas_receber_cadastro_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_contas_receber_cadastro_Click"
    GoTo fim_mnu_msf_grade_contas_receber_cadastro_Click
End Sub

Private Sub mnu_msf_grade_contas_receber_copiar_Click()
    On Error GoTo erro_mnu_msf_grade_contas_receber_copiar_Click
    pfct_copiar_conteudo_grade msf_grade_contas_receber
fim_mnu_msf_grade_contas_receber_copiar_Click:
    Exit Sub
erro_mnu_msf_grade_contas_receber_copiar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_contas_receber_copiar_Click"
    GoTo fim_mnu_msf_grade_contas_receber_copiar_Click
End Sub

Private Sub mnu_msf_grade_contas_receber_exportar_Click()
    On Error GoTo erro_mnu_msf_grade_contas_receber_exportar_Click
    pfct_exportar_conteudo_grade msf_grade_contas_receber, "contas_a_receber"
fim_mnu_msf_grade_contas_receber_exportar_Click:
    Exit Sub
erro_mnu_msf_grade_contas_receber_exportar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_contas_receber_exportar_Click"
    GoTo fim_mnu_msf_grade_contas_receber_exportar_Click
End Sub

Private Sub mnu_msf_grade_receitas_despesas_fixas_cadastro_Click()
    On Error GoTo erro_mnu_msf_grade_receitas_despesas_fixas_cadastro_Click
    'mesma lógica do menu principal
    frm_movimentacao_geral.Show
    frm_movimentacao_geral.ZOrder 0
fim_mnu_msf_grade_receitas_despesas_fixas_cadastro_Click:
    Exit Sub
erro_mnu_msf_grade_receitas_despesas_fixas_cadastro_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_receitas_despesas_fixas_cadastro_Click"
    GoTo fim_mnu_msf_grade_receitas_despesas_fixas_cadastro_Click
End Sub

Private Sub mnu_msf_grade_receitas_despesas_fixas_copiar_Click()
    On Error GoTo erro_mnu_msf_grade_receitas_despesas_fixas_copiar_Click
    pfct_copiar_conteudo_grade msf_grade_receitas_despesas_fixas
fim_mnu_msf_grade_receitas_despesas_fixas_copiar_Click:
    Exit Sub
erro_mnu_msf_grade_receitas_despesas_fixas_copiar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_receitas_despesas_fixas_copiar_Click"
    GoTo fim_mnu_msf_grade_receitas_despesas_fixas_copiar_Click
End Sub

Private Sub mnu_msf_grade_receitas_despesas_fixas_exportar_Click()
    On Error GoTo erro_mnu_msf_grade_receitas_despesas_fixas_exportar_Click
    pfct_exportar_conteudo_grade msf_grade_receitas_despesas_fixas, "receitas_despesas_fixas"
fim_mnu_msf_grade_receitas_despesas_fixas_exportar_Click:
    Exit Sub
erro_mnu_msf_grade_receitas_despesas_fixas_exportar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_receitas_despesas_fixas_exportar_Click"
    GoTo fim_mnu_msf_grade_receitas_despesas_fixas_exportar_Click
End Sub

Private Sub mnu_msf_grade_ultimas_baixas_cadastro_Click()
    On Error GoTo erro_mnu_msf_grade_ultimas_baixas_cadastro_Click
    'mesma lógica do menu principal
    frm_movimentacao_geral.Show
    frm_movimentacao_geral.ZOrder 0
fim_mnu_msf_grade_ultimas_baixas_cadastro_Click:
    Exit Sub
erro_mnu_msf_grade_ultimas_baixas_cadastro_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_ultimas_baixas_cadastro_Click"
    GoTo fim_mnu_msf_grade_ultimas_baixas_cadastro_Click
End Sub

Private Sub mnu_msf_grade_ultimas_baixas_copiar_Click()
    On Error GoTo erro_mnu_msf_grade_ultimas_baixas_copiar_Click
    pfct_copiar_conteudo_grade msf_grade_ultimas_baixas
fim_mnu_msf_grade_ultimas_baixas_copiar_Click:
    Exit Sub
erro_mnu_msf_grade_ultimas_baixas_copiar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_ultimas_baixas_copiar_Click"
    GoTo fim_mnu_msf_grade_ultimas_baixas_copiar_Click
End Sub

Private Sub mnu_msf_grade_ultimas_baixas_exportar_Click()
    On Error GoTo erro_mnu_msf_grade_ultimas_baixas_exportar_Click
    pfct_exportar_conteudo_grade msf_grade_ultimas_baixas, "ultimas_baixas"
fim_mnu_msf_grade_ultimas_baixas_exportar_Click:
    Exit Sub
erro_mnu_msf_grade_ultimas_baixas_exportar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "mnu_msf_grade_ultimas_baixas_exportar_Click"
    GoTo fim_mnu_msf_grade_ultimas_baixas_exportar_Click
End Sub

Private Sub msf_grade_contas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_msf_grade_contas_MouseUp
    If (Button = 2) Then 'botão direito do mouse
        PopupMenu mnu_msf_grade_contas 'exibimos o popup
    End If
fim_msf_grade_contas_MouseUp:
    Exit Sub
erro_msf_grade_contas_MouseUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "msf_grade_contas_MouseUp"
    GoTo fim_msf_grade_contas_MouseUp
End Sub

Private Sub msf_grade_contas_pagar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_msf_grade_contas_pagar_MouseUp
    If (Button = 2) Then 'botão direito do mouse
        PopupMenu mnu_msf_grade_contas_pagar 'exibimos o popup
    End If
fim_msf_grade_contas_pagar_MouseUp:
    Exit Sub
erro_msf_grade_contas_pagar_MouseUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "msf_grade_contas_pagar_MouseUp"
    GoTo fim_msf_grade_contas_pagar_MouseUp
End Sub

Private Sub msf_grade_contas_receber_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_msf_grade_contas_receber_MouseUp
    If (Button = 2) Then 'botão direito do mouse
        PopupMenu mnu_msf_grade_contas_receber 'exibimos o popup
    End If
fim_msf_grade_contas_receber_MouseUp:
    Exit Sub
erro_msf_grade_contas_receber_MouseUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "msf_grade_contas_receber_MouseUp"
    GoTo fim_msf_grade_contas_receber_MouseUp
End Sub

Private Sub msf_grade_receitas_despesas_fixas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_msf_grade_receitas_despesas_fixas_MouseUp
    If (Button = 2) Then 'botão direito do mouse
        PopupMenu mnu_msf_grade_receitas_despesas_fixas 'exibimos o popup
    End If
fim_msf_grade_receitas_despesas_fixas_MouseUp:
    Exit Sub
erro_msf_grade_receitas_despesas_fixas_MouseUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "msf_grade_receitas_despesas_fixas_MouseUp"
    GoTo fim_msf_grade_receitas_despesas_fixas_MouseUp
End Sub

Private Sub msf_grade_ultimas_baixas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_msf_grade_ultimas_baixas_MouseUp
    If (Button = 2) Then 'botão direito do mouse
        PopupMenu mnu_msf_grade_ultimas_baixas 'exibimos o popup
    End If
fim_msf_grade_ultimas_baixas_MouseUp:
    Exit Sub
erro_msf_grade_ultimas_baixas_MouseUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_financeiro_agenda", "msf_grade_ultimas_baixas_MouseUp"
    GoTo fim_msf_grade_ultimas_baixas_MouseUp
End Sub

