VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_cadastro_contas_pagar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contas a Pagar"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11775
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
   ScaleHeight     =   5370
   ScaleWidth      =   11775
   Begin VB.CommandButton cmd_filtrar 
      Caption         =   "&Filtrar (F7)"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   1275
   End
   Begin VB.Frame fme_filtros 
      Caption         =   " Filtros "
      Height          =   1155
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   11535
      Begin VB.ComboBox cbo_ordem 
         Height          =   315
         ItemData        =   "frm_cadastro_contas_pagar.frx":0000
         Left            =   6780
         List            =   "frm_cadastro_contas_pagar.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   660
         Width           =   2295
      End
      Begin VB.ComboBox cbo_ordenar_por 
         Height          =   315
         ItemData        =   "frm_cadastro_contas_pagar.frx":0004
         Left            =   4380
         List            =   "frm_cadastro_contas_pagar.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   660
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtp_de 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   660
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   39591
      End
      Begin MSComCtl2.DTPicker dtp_ate 
         Height          =   315
         Left            =   2220
         TabIndex        =   11
         Top             =   660
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   39591
      End
      Begin VB.Label lbl_ordem 
         AutoSize        =   -1  'True
         Caption         =   "&Em ordem:"
         Height          =   195
         Left            =   6780
         TabIndex        =   9
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lbl_ordenar_por 
         AutoSize        =   -1  'True
         Caption         =   "&Ordenar por:"
         Height          =   195
         Left            =   4380
         TabIndex        =   8
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lbl_periodo 
         AutoSize        =   -1  'True
         Caption         =   "Exibir de:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lbl_ate 
         AutoSize        =   -1  'True
         Caption         =   "até:"
         Height          =   195
         Left            =   2220
         TabIndex        =   7
         Top             =   360
         Width           =   300
      End
   End
   Begin VB.CommandButton cmd_baixar 
      Caption         =   "&Baixar (F4)"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton cmd_fechar 
      Caption         =   "&Fechar (F8)"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton cmd_lancar 
      Caption         =   "&Lançar (F2)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar (F3)"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1275
   End
   Begin MSComctlLib.StatusBar stb_status 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   15
      Top             =   5055
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20717
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid msf_grade 
      Height          =   3135
      Left            =   120
      TabIndex        =   14
      Top             =   1860
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      ScrollBars      =   2
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
Attribute VB_Name = "frm_cadastro_contas_pagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enm_conta_pagar
    col_vencimento = 0
    col_baixa_automatica = 1
    col_despesa = 2
    col_valor = 3
    col_parcela = 4
    col_descricao = 5
    col_documento = 6
End Enum

Private Enum enm_status
    pnl_mensagem = 1
End Enum

Public Sub lsub_ajustar_grade(ByRef pgrd_grade As MSFlexGrid)
    On Error GoTo erro_lsub_ajustar_grade
    Dim llng_contador As Long
    For llng_contador = 0 To pgrd_grade.Rows - 1
        pgrd_grade.RowData(llng_contador) = 0
    Next
    With pgrd_grade
        .Clear
        .Cols = 7
        .Rows = 2
        .ColWidth(enm_conta_pagar.col_vencimento) = 1110
        .ColWidth(enm_conta_pagar.col_baixa_automatica) = 1450
        .ColWidth(enm_conta_pagar.col_despesa) = 2340
        .ColWidth(enm_conta_pagar.col_valor) = 900
        .ColWidth(enm_conta_pagar.col_parcela) = 900
        .ColWidth(enm_conta_pagar.col_descricao) = 3495
        .ColWidth(enm_conta_pagar.col_documento) = 1005
        .TextMatrix(0, enm_conta_pagar.col_vencimento) = " Vencimento"
        .TextMatrix(0, enm_conta_pagar.col_baixa_automatica) = " Baixa automática"
        .TextMatrix(0, enm_conta_pagar.col_despesa) = " Despesa"
        .TextMatrix(0, enm_conta_pagar.col_valor) = " Valor"
        .TextMatrix(0, enm_conta_pagar.col_parcela) = " Parcela"
        .TextMatrix(0, enm_conta_pagar.col_descricao) = " Descrição"
        .TextMatrix(0, enm_conta_pagar.col_documento) = " Documento"
    End With
    stb_status.Panels(enm_status.pnl_mensagem).Text = "" 'limpa a barra de status
fim_lsub_ajustar_grade:
    Exit Sub
erro_lsub_ajustar_grade:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "lsub_ajustar_grade"
    GoTo fim_lsub_ajustar_grade
End Sub

Public Sub lsub_preencher_combos()
    On Error GoTo erro_lsub_preencher_combos
    psub_ajustar_combos_data dtp_de, dtp_ate
    With cbo_ordenar_por
        .Clear
        .AddItem "- Selecione o campo -", 0
        .AddItem "- Lançamento", 1
        .AddItem "- Vencimento", 2
        .AddItem "- Despesa", 3
        .AddItem "- Valor", 4
        .AddItem "- Descrição", 5
        .AddItem "- Documento", 6
        .ListIndex = 0
    End With
    psub_preencher_ordem cbo_ordem
fim_lsub_preencher_combos:
    Exit Sub
erro_lsub_preencher_combos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "lsub_preencher_combos"
    GoTo fim_lsub_preencher_combos
End Sub

Private Sub lsub_cancelar_conta_pagar(ByVal plng_codigo As Long)
    On Error GoTo erro_lsub_cancelar_conta_pagar
    Dim lobj_cancelar_conta_pagar As Object
    Dim lobj_ocorrencias As Object
    Dim lstr_sql As String
    Dim lstr_chave As String
    Dim lstr_mensagem As String
    Dim llng_registros As Long
    Dim llng_ocorrencias As Long
    Dim lint_resposta As Integer
    'monta o comando sql
    lstr_sql = "select * from [tb_contas_pagar] where [int_codigo] = " & pfct_tratar_numero_sql(plng_codigo)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_cancelar_conta_pagar, lstr_sql, "frm_cadastro_contas_pagar", "lsub_cancelar_conta_pagar")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_cancelar_conta_pagar
    End If
    llng_registros = lobj_cancelar_conta_pagar.Count
    If (llng_registros > 0) Then
        lstr_chave = lobj_cancelar_conta_pagar(1)("str_chave")
        If (lstr_chave <> "") Then
            'monta o comando sql
            lstr_sql = "select count(*) as [int_quantidade] from [tb_contas_pagar] where [str_chave] = '" & lstr_chave & "'"
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_ocorrencias, lstr_sql, "frm_cadastro_contas_pagar", "lsub_cancelar_conta_pagar")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_lsub_cancelar_conta_pagar
            End If
            llng_registros = lobj_ocorrencias.Count
            If (llng_registros > 0) Then
                llng_ocorrencias = CLng(lobj_ocorrencias(1)("int_quantidade"))
                If (llng_ocorrencias > 0) Then
                    'monta mensagem
                    lstr_mensagem = ""
                    lstr_mensagem = lstr_mensagem & "Foram encontradas '"
                    lstr_mensagem = lstr_mensagem & CStr(llng_ocorrencias)
                    lstr_mensagem = lstr_mensagem & "' ocorrências desta conta a pagar." & vbCrLf
                    lstr_mensagem = lstr_mensagem & "Excluir todas as ocorrências?" & vbCrLf & vbCrLf
                    lstr_mensagem = lstr_mensagem & "- Sim, excluir todas as ocorrências encontradas." & vbCrLf
                    lstr_mensagem = lstr_mensagem & "- Não, excluir apenas a conta a pagar selecionada." & vbCrLf
                    lstr_mensagem = lstr_mensagem & "- Cancelar a operação de exclusão."
                    lint_resposta = MsgBox(lstr_mensagem, vbYesNoCancel + vbQuestion + vbDefaultButton3, pcst_nome_aplicacao)
                    If (lint_resposta = vbYes) Then
                        'monta o comando sql
                        lstr_sql = "delete from [tb_contas_pagar] where [str_chave] = '" & lstr_chave & "'"
                        'executa o comando sql e devolve o objeto
                        If (Not pfct_executar_comando_sql(lobj_cancelar_conta_pagar, lstr_sql, "frm_cadastro_contas_pagar", "lsub_cancelar_conta_pagar")) Then
                            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                            GoTo fim_lsub_cancelar_conta_pagar
                        Else
                            MsgBox "Operação de exclusão múltipla executada com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                            GoTo fim_lsub_cancelar_conta_pagar
                        End If
                    ElseIf (lint_resposta = vbNo) Then
                        'monta o comando sql
                         lstr_sql = "delete from [tb_contas_pagar] where [int_codigo] = " & pfct_tratar_numero_sql(plng_codigo)
                        'executa o comando sql e devolve o objeto
                        If (Not pfct_executar_comando_sql(lobj_cancelar_conta_pagar, lstr_sql, "frm_cadastro_contas_pagar", "lsub_cancelar_conta_pagar")) Then
                            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                            GoTo fim_lsub_cancelar_conta_pagar
                        Else
                            MsgBox "Operação de exclusão executada com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                            GoTo fim_lsub_cancelar_conta_pagar
                        End If
                    ElseIf (lint_resposta = vbCancel) Then
                        MsgBox "Operação de exclusão cancelada.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                        GoTo fim_lsub_cancelar_conta_pagar
                    End If
                End If
            End If
        ElseIf (lstr_chave = "") Then
            lstr_mensagem = "Deseja excluir a conta selecionada?"
            lint_resposta = MsgBox(lstr_mensagem, vbYesNo + vbQuestion + vbDefaultButton2, pcst_nome_aplicacao)
            If (lint_resposta = vbYes) Then
                'monta o comando sql
                lstr_sql = "delete from [tb_contas_pagar] where [int_codigo] = " & pfct_tratar_numero_sql(plng_codigo)
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_cancelar_conta_pagar, lstr_sql, "frm_cadastro_contas_pagar", "lsub_cancelar_conta_pagar")) Then
                    MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                    GoTo fim_lsub_cancelar_conta_pagar
                Else
                    MsgBox "Operação de exclusão executada com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                    GoTo fim_lsub_cancelar_conta_pagar
                End If
            End If
        End If
    End If
fim_lsub_cancelar_conta_pagar:
    lsub_preencher_combos
    lsub_ajustar_grade msf_grade
    stb_status.Panels(enm_status.pnl_mensagem).Text = "" 'limpa a barra de status
    'destrói os objetos
    Set lobj_cancelar_conta_pagar = Nothing
    Set lobj_ocorrencias = Nothing
    Exit Sub
erro_lsub_cancelar_conta_pagar:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "lsub_cancelar_conta_pagar"
    GoTo fim_lsub_cancelar_conta_pagar
    Resume 0
End Sub

Private Sub lsub_preencher_grade(ByVal pstr_data_de As String, _
                                 ByVal pstr_data_ate As String, _
                                 ByVal pstr_ordenar_por As String, _
                                 ByVal pstr_ordem As String)
    On Error GoTo erro_lsub_preencher_grade
    Dim lobj_contas_pagar As Object
    Dim lstr_sql As String
    Dim llng_contador As Long
    Dim llng_registros As Long
    Dim lcur_valor_total As Currency
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select "
    lstr_sql = lstr_sql & " [tb_despesas].[str_descricao] as [str_descricao_despesa],"
    lstr_sql = lstr_sql & " [tb_contas_pagar].*"
    lstr_sql = lstr_sql & " from "
    lstr_sql = lstr_sql & " [tb_contas_pagar]"
    lstr_sql = lstr_sql & " inner join "
    lstr_sql = lstr_sql & " [tb_despesas] on [tb_contas_pagar].[int_despesa] = [tb_despesas].[int_codigo]"
    lstr_sql = lstr_sql & " where "
    lstr_sql = lstr_sql & " [dt_vencimento] between '" & pstr_data_de & "' and '" & pstr_data_ate & "'"
    lstr_sql = lstr_sql & " order by "
    lstr_sql = lstr_sql & " " & pstr_ordenar_por & " " & pstr_ordem
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_contas_pagar, lstr_sql, "frm_cadastro_contas_pagar", "lsub_preencher_grade")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_preencher_grade
    End If
    llng_registros = lobj_contas_pagar.Count
    If (llng_registros > 0) Then
        msf_grade.Redraw = False
        For llng_contador = 1 To llng_registros
            msf_grade.Row = llng_contador
            msf_grade.Col = enm_conta_pagar.col_vencimento
            msf_grade.RowData(llng_contador) = lobj_contas_pagar(llng_contador)("int_codigo")
            msf_grade.TextMatrix(llng_contador, enm_conta_pagar.col_vencimento) = " " & Format$(lobj_contas_pagar(llng_contador)("dt_vencimento"), pcst_formato_data)
            msf_grade.TextMatrix(llng_contador, enm_conta_pagar.col_baixa_automatica) = " " & IIf(lobj_contas_pagar(llng_contador)("chr_baixa_automatica") = "S", "Sim", "Não")
            msf_grade.TextMatrix(llng_contador, enm_conta_pagar.col_despesa) = " " & lobj_contas_pagar(llng_contador)("str_descricao_despesa")
            msf_grade.TextMatrix(llng_contador, enm_conta_pagar.col_valor) = " " & Format$(lobj_contas_pagar(llng_contador)("num_valor"), pcst_formato_numerico)
            'ini parcelas
            msf_grade.TextMatrix(llng_contador, enm_conta_pagar.col_parcela) = " " & _
                Format$(lobj_contas_pagar(llng_contador)("int_parcela"), pcst_formato_numerico_parcela) & "/" & _
                Format$(lobj_contas_pagar(llng_contador)("int_total_parcelas"), pcst_formato_numerico_parcela)
            'fim parcelas
            msf_grade.TextMatrix(llng_contador, enm_conta_pagar.col_descricao) = " " & lobj_contas_pagar(llng_contador)("str_descricao")
            msf_grade.TextMatrix(llng_contador, enm_conta_pagar.col_documento) = " " & lobj_contas_pagar(llng_contador)("str_documento")
            'se a data de vencimento for menor que a data de hoje, conta está atrasada
            If (CDate(lobj_contas_pagar(llng_contador)("dt_vencimento")) < Date) Then
                'cor da fonte da linha em vermelho
                psub_ajustar_cor_linha_grade msf_grade, llng_contador, vbRed
            End If
            'se a data de vencimento for maior que a data de hoje, conta ainda vai vencer
            If (CDate(lobj_contas_pagar(llng_contador)("dt_vencimento")) > Date) Then
                'cor da fonte da linha em azul
                psub_ajustar_cor_linha_grade msf_grade, llng_contador, vbBlue
            End If
            'se a data de vencimento for igual a data de hoje, conta vence no dia
            If (CDate(lobj_contas_pagar(llng_contador)("dt_vencimento")) = Date) Then
                'cor da fonte da linha em preto
                psub_ajustar_cor_linha_grade msf_grade, llng_contador, vbWindowText
            End If
            'se ainda houver registros
            If (llng_contador < llng_registros) Then
                'adiciona mais uma linha
                msf_grade.Rows = msf_grade.Rows + 1
            End If
            'acumula o valor total
            lcur_valor_total = lcur_valor_total + CCur(lobj_contas_pagar(llng_contador)("num_valor"))
        Next
        msf_grade.Redraw = True
        msf_grade.Row = 1
        stb_status.Panels(enm_status.pnl_mensagem).Text = "Total de contas a pagar de [" & Format$(dtp_de.Value, pcst_formato_data) & "] a [" & Format$(dtp_ate.Value, pcst_formato_data) & "] -> [" & Format$(llng_registros, "00") & "]" & " " & pfct_retorna_simbolo_moeda() & " " & Format$(lcur_valor_total, pcst_formato_numerico)
    Else
        MsgBox "Atenção!" & vbCrLf & "Não foram encontradas [contas a pagar] no período selecionado.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        stb_status.Panels(enm_status.pnl_mensagem).Text = "Não há contas a pagar para o período selecionado."
        GoTo fim_lsub_preencher_grade
    End If
fim_lsub_preencher_grade:
    'destrói os objetos
    Set lobj_contas_pagar = Nothing
    Exit Sub
erro_lsub_preencher_grade:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "lsub_preencher_grade"
    GoTo fim_lsub_preencher_grade
End Sub

Private Sub cbo_ordem_DropDown()
    On Error GoTo erro_cbo_ordem_DropDown
    psub_campo_got_focus cbo_ordem
fim_cbo_ordem_DropDown:
    Exit Sub
erro_cbo_ordem_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "cbo_ordem_DropDown"
    GoTo fim_cbo_ordem_DropDown
End Sub

Private Sub cbo_ordem_GotFocus()
    On Error GoTo erro_cbo_ordem_GotFocus
    psub_campo_got_focus cbo_ordem
fim_cbo_ordem_GotFocus:
    Exit Sub
erro_cbo_ordem_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "cbo_ordem_GotFocus"
    GoTo fim_cbo_ordem_GotFocus
End Sub

Private Sub cbo_ordem_LostFocus()
    On Error GoTo erro_cbo_ordem_LostFocus
    psub_campo_lost_focus cbo_ordem
fim_cbo_ordem_LostFocus:
    Exit Sub
erro_cbo_ordem_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "cbo_ordem_LostFocus"
    GoTo fim_cbo_ordem_LostFocus
End Sub

Private Sub cbo_ordenar_por_DropDown()
    On Error GoTo erro_cbo_ordenar_por_DropDown
    psub_campo_got_focus cbo_ordenar_por
fim_cbo_ordenar_por_DropDown:
    Exit Sub
erro_cbo_ordenar_por_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "cbo_ordenar_por_DropDown"
    GoTo fim_cbo_ordenar_por_DropDown
End Sub

Private Sub cbo_ordenar_por_GotFocus()
    On Error GoTo erro_cbo_ordenar_por_GotFocus
    psub_campo_got_focus cbo_ordenar_por
fim_cbo_ordenar_por_GotFocus:
    Exit Sub
erro_cbo_ordenar_por_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "cbo_ordenar_por_GotFocus"
    GoTo fim_cbo_ordenar_por_GotFocus
End Sub

Private Sub cbo_ordenar_por_LostFocus()
    On Error GoTo erro_cbo_ordenar_por_LostFocus
    psub_campo_lost_focus cbo_ordenar_por
fim_cbo_ordenar_por_LostFocus:
    Exit Sub
erro_cbo_ordenar_por_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "cbo_ordenar_por_LostFocus"
    GoTo fim_cbo_ordenar_por_LostFocus
End Sub

Private Sub cmd_cancelar_Click()
    On Error GoTo erro_cmd_cancelar_Click
    Dim llng_codigo_item As Long
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_cancelar.Enabled) Then
        Exit Sub
    End If
    llng_codigo_item = msf_grade.RowData(msf_grade.Row)
    If (llng_codigo_item = 0) Then
        MsgBox "Selecione um item na grade.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        GoTo fim_cmd_cancelar_Click
    Else
        lsub_cancelar_conta_pagar llng_codigo_item
    End If
fim_cmd_cancelar_Click:
    Exit Sub
erro_cmd_cancelar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "cmd_cancelar_Click"
    GoTo fim_cmd_cancelar_Click
End Sub

Private Sub cmd_baixar_Click()
    On Error GoTo erro_cmd_baixar_Click
    Dim llng_codigo_item As Long
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_baixar.Enabled) Then
        Exit Sub
    End If
    llng_codigo_item = msf_grade.RowData(msf_grade.Row)
    If (llng_codigo_item = 0) Then
        MsgBox "Selecione um item na grade.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        GoTo fim_cmd_baixar_Click
    Else
        frm_cadastro_contas_pagar.Enabled = False
        frm_baixar_contas_pagar.Left = frm_cadastro_contas_pagar.Left + 250
        frm_baixar_contas_pagar.Top = frm_cadastro_contas_pagar.Top + 250
        frm_baixar_contas_pagar.Show
    End If
fim_cmd_baixar_Click:
    Exit Sub
erro_cmd_baixar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "cmd_baixar_Click"
    GoTo fim_cmd_baixar_Click
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "cmd_fechar_Click"
    GoTo fim_cmd_fechar_Click
End Sub

Private Sub cmd_filtrar_Click()
    On Error GoTo erro_cmd_filtrar_Click
    Dim lstr_ordenar_por As String
    Dim lstr_ordem As String
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_filtrar.Enabled) Then
        Exit Sub
    End If
    Select Case cbo_ordenar_por.Text
        Case "- Selecione o campo -"
            MsgBox "Atenção!" & vbCrLf & "É necessário selecionar o campo para ordenar a consulta.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            cbo_ordenar_por.SetFocus
            GoTo fim_cmd_filtrar_Click
        Case "- Lançamento"
            lstr_ordenar_por = "[tb_contas_pagar].[int_codigo]"
        Case "- Vencimento"
            lstr_ordenar_por = "[dt_vencimento]"
        Case "- Despesa"
            lstr_ordenar_por = "[tb_despesas].[int_codigo]"
        Case "- Descrição"
            lstr_ordenar_por = "[str_descricao_despesa]"
        Case "- Valor"
            lstr_ordenar_por = "[num_valor]"
        Case "- Documento"
            lstr_ordenar_por = "[str_documento]"
    End Select
    Select Case cbo_ordem.Text
        Case "- Selecione a ordem -"
            MsgBox "Atenção!" & vbCrLf & "É necessário selecionar a ordem da consulta.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            cbo_ordem.SetFocus
            GoTo fim_cmd_filtrar_Click
        Case "- Crescente"
            lstr_ordem = "asc"
        Case "- Decrescente"
            lstr_ordem = "desc"
    End Select
    lsub_ajustar_grade msf_grade
    lsub_preencher_grade Format$(dtp_de.Value, pcst_formato_data_sql), Format$(dtp_ate.Value, pcst_formato_data_sql), lstr_ordenar_por, lstr_ordem
fim_cmd_filtrar_Click:
    Exit Sub
erro_cmd_filtrar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "cmd_filtrar_Click"
    GoTo fim_cmd_filtrar_Click
End Sub

Private Sub cmd_lancar_Click()
    On Error GoTo erro_cmd_lancar_Click
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_lancar.Enabled) Then
        Exit Sub
    End If
    frm_cadastro_contas_pagar.Enabled = False
    frm_lancar_contas_pagar.Left = frm_cadastro_contas_pagar.Left + 250
    frm_lancar_contas_pagar.Top = frm_cadastro_contas_pagar.Top + 250
    frm_lancar_contas_pagar.Show
fim_cmd_lancar_Click:
    Exit Sub
erro_cmd_lancar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "cmd_lancar_Click"
    GoTo fim_cmd_lancar_Click
End Sub

Private Sub dtp_ate_DropDown()
    On Error GoTo erro_dtp_ate_DropDown
    psub_campo_got_focus dtp_ate
fim_dtp_ate_DropDown:
    Exit Sub
erro_dtp_ate_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "dtp_ate_DropDown"
    GoTo fim_dtp_ate_DropDown
End Sub

Private Sub dtp_ate_GotFocus()
    On Error GoTo erro_dtp_ate_GotFocus
    psub_campo_got_focus dtp_ate
fim_dtp_ate_GotFocus:
    Exit Sub
erro_dtp_ate_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "dtp_ate_GotFocus"
    GoTo fim_dtp_ate_GotFocus
End Sub

Private Sub dtp_ate_LostFocus()
    On Error GoTo erro_dtp_ate_LostFocus
    psub_campo_lost_focus dtp_ate
fim_dtp_ate_LostFocus:
    Exit Sub
erro_dtp_ate_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "dtp_ate_LostFocus"
    GoTo fim_dtp_ate_LostFocus
End Sub

Private Sub dtp_de_DropDown()
    On Error GoTo erro_dtp_de_DropDown
    psub_campo_got_focus dtp_de
fim_dtp_de_DropDown:
    Exit Sub
erro_dtp_de_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "dtp_de_DropDown"
    GoTo fim_dtp_de_DropDown
End Sub

Private Sub dtp_de_GotFocus()
    On Error GoTo erro_dtp_de_GotFocus
    psub_campo_got_focus dtp_de
fim_dtp_de_GotFocus:
    Exit Sub
erro_dtp_de_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "dtp_de_GotFocus"
    GoTo fim_dtp_de_GotFocus
End Sub

Private Sub dtp_de_LostFocus()
    On Error GoTo erro_dtp_de_LostFocus
    psub_campo_lost_focus dtp_de
fim_dtp_de_LostFocus:
    Exit Sub
erro_dtp_de_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "dtp_de_LostFocus"
    GoTo fim_dtp_de_LostFocus
End Sub

Private Sub Form_Initialize()
    On Error GoTo Erro_Form_Initialize
    InitCommonControls
Fim_Form_Initialize:
    Exit Sub
Erro_Form_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "Form_Initialize"
    GoTo Fim_Form_Initialize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/financeiro_contas_pagar.htm", 0
        Case vbKeyF2
            cmd_lancar_Click
        Case vbKeyF3
            cmd_cancelar_Click
        Case vbKeyF4
            cmd_baixar_Click
        Case vbKeyF7
            cmd_filtrar_Click
        Case vbKeyF8
            cmd_fechar_Click
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
    lsub_preencher_combos
    lsub_ajustar_grade msf_grade
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "Form_Load"
    GoTo fim_Form_Load
End Sub

Private Sub msf_grade_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_msf_grade_MouseUp
    If (Button = 2) Then 'botão direito do mouse
        PopupMenu mnu_msf_grade 'exibimos o popup
    End If
fim_msf_grade_MouseUp:
    Exit Sub
erro_msf_grade_MouseUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "msf_grade_MouseUp"
    GoTo fim_msf_grade_MouseUp
End Sub

Private Sub mnu_msf_grade_copiar_Click()
    On Error GoTo erro_mnu_msf_grade_copiar_Click
    pfct_copiar_conteudo_grade msf_grade
fim_mnu_msf_grade_copiar_Click:
    Exit Sub
erro_mnu_msf_grade_copiar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "mnu_msf_grade_copiar_Click"
    GoTo fim_mnu_msf_grade_copiar_Click
End Sub

Private Sub mnu_msf_grade_exportar_Click()
    On Error GoTo erro_mnu_msf_grade_exportar_Click
    pfct_exportar_conteudo_grade msf_grade, "contas_a_pagar"
fim_mnu_msf_grade_exportar_Click:
    Exit Sub
erro_mnu_msf_grade_exportar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_contas_pagar", "mnu_msf_grade_exportar_Click"
    GoTo fim_mnu_msf_grade_exportar_Click
End Sub
