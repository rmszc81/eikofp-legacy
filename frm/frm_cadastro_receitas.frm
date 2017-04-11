VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_cadastro_receitas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Receitas"
   ClientHeight    =   4140
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   7875
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
   ScaleHeight     =   4140
   ScaleWidth      =   7875
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "Ca&ncelar (F8)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3300
      TabIndex        =   3
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmd_salvar 
      Caption         =   "&Salvar (F7)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   60
      Width           =   975
   End
   Begin VB.Frame fme_campos 
      Enabled         =   0   'False
      Height          =   3315
      Left            =   3660
      TabIndex        =   7
      Top             =   480
      Width           =   4155
      Begin VB.CheckBox chk_fixa 
         Caption         =   "&Fixa?"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   180
         Width           =   780
      End
      Begin VB.CheckBox chk_ativo 
         Caption         =   "&Ativo?"
         Height          =   255
         Left            =   3300
         TabIndex        =   10
         Top             =   180
         Width           =   780
      End
      Begin VB.TextBox txt_receita 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   3915
      End
      Begin VB.TextBox txt_observacoes 
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1260
         Width           =   3915
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
         Left            =   2160
         TabIndex        =   13
         Top             =   1020
         Width           =   1875
      End
      Begin VB.Label lbl_receita 
         AutoSize        =   -1  'True
         Caption         =   "&Receita:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Width           =   600
      End
      Begin VB.Label lbl_observacoes 
         AutoSize        =   -1  'True
         Caption         =   "&Observações:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "E&xcluir (F4)"
      Height          =   375
      Left            =   4500
      TabIndex        =   4
      Top             =   60
      Width           =   1005
   End
   Begin VB.CommandButton cmd_fechar 
      Caption         =   "&Fechar (F6)"
      Height          =   375
      Left            =   6780
      TabIndex        =   6
      Top             =   60
      Width           =   1005
   End
   Begin VB.CommandButton cmd_atualizar 
      Caption         =   "&Atualizar (F5)"
      Height          =   375
      Left            =   5580
      TabIndex        =   5
      Top             =   60
      Width           =   1125
   End
   Begin MSComctlLib.StatusBar stb_status 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   16
      Top             =   3855
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13838
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid msf_grade 
      Height          =   3285
      Left            =   60
      TabIndex        =   15
      Top             =   540
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   5794
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmd_editar 
      Caption         =   "&Editar (F3)"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   60
      Width           =   1005
   End
   Begin VB.CommandButton cmd_inserir 
      Caption         =   "&Inserir (F2)"
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1095
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
Attribute VB_Name = "frm_cadastro_receitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enm_receita
    col_receita = 0
End Enum

Private Enum enm_status
    pnl_mensagem = 1
End Enum

Private Const mcst_inserir As Byte = 1
Private Const mcst_editar As Byte = 2

Private mbte_ac As Byte
Private mlng_registro_selecionado As Long

Private Sub cmd_atualizar_Click()
    On Error GoTo erro_cmd_atualizar_Click
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_atualizar.Enabled) Then
        Exit Sub
    End If
    lsub_preencher_grade
fim_cmd_atualizar_Click:
    Exit Sub
erro_cmd_atualizar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "cmd_atualizar_Click"
    GoTo fim_cmd_atualizar_Click
End Sub

Private Sub cmd_cancelar_Click()
    On Error GoTo erro_cmd_cancelar_Click
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_cancelar.Enabled) Then
        Exit Sub
    End If
    lsub_alterar_estado_controles True
    lsub_limpar_campos
    lsub_preencher_grade
fim_cmd_cancelar_Click:
    Exit Sub
erro_cmd_cancelar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "cmd_cancelar_Click"
    GoTo fim_cmd_cancelar_Click
End Sub

Private Sub cmd_editar_Click()
    On Error GoTo erro_cmd_editar_Click
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_editar.Enabled) Then
        Exit Sub
    End If
    If (lfct_verificar_selecao) Then
        mbte_ac = mcst_editar
        lsub_alterar_estado_controles False
        txt_receita.SetFocus
    End If
fim_cmd_editar_Click:
    Exit Sub
erro_cmd_editar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "cmd_editar_Click"
    GoTo fim_cmd_editar_Click
End Sub

Private Sub cmd_excluir_Click()
    On Error GoTo erro_cmd_excluir_click
    Dim llng_resposta As Long
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_excluir.Enabled) Then
        Exit Sub
    End If
    If (lfct_verificar_selecao) Then
        llng_resposta = MsgBox("Deseja excluir o registro selecionado?", vbYesNo + vbQuestion + vbDefaultButton2, pcst_nome_aplicacao)
        If (llng_resposta = vbYes) Then
            If (lfct_excluir_registro(mlng_registro_selecionado)) Then
                MsgBox "Registro excluído com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            End If
        End If
        lsub_limpar_campos
        lsub_preencher_grade
        lsub_preencher_campos
    End If
fim_cmd_excluir_click:
    Exit Sub
erro_cmd_excluir_click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "cmd_excluir_Click"
    GoTo fim_cmd_excluir_click
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "cmd_fechar_Click"
    GoTo fim_cmd_fechar_Click
End Sub

Private Sub cmd_inserir_Click()
    On Error GoTo erro_cmd_inserir_Click
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_inserir.Enabled) Then
        Exit Sub
    End If
    mbte_ac = mcst_inserir
    lsub_alterar_estado_controles False
    lsub_limpar_campos
    txt_receita.SetFocus
fim_cmd_inserir_Click:
    Exit Sub
erro_cmd_inserir_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "cmd_inserir_Click"
    GoTo fim_cmd_inserir_Click
End Sub

Private Sub cmd_salvar_Click()
    On Error GoTo erro_cmd_salvar_Click
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_salvar.Enabled) Then
        Exit Sub
    End If
    Me.ValidateControls
    If (txt_receita.Text = "") Then
        MsgBox "Atenção!" & vbCrLf & "O campo [receita] não pode estar em branco.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_receita.SetFocus
        GoTo fim_cmd_salvar_Click
    ElseIf (Len(txt_receita.Text) < 2) Then
        MsgBox "Atenção!" & vbCrLf & "O campo [receita] não pode conter menos de 02 (dois) caracteres.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_receita.SetFocus
        txt_receita.SelStart = 0
        txt_receita.SelLength = Len(txt_receita.Text)
        GoTo fim_cmd_salvar_Click
    End If
    If (mbte_ac = mcst_inserir) Then
        lfct_salvar_registro True
    ElseIf (mbte_ac = mcst_editar) Then
        lfct_salvar_registro False, mlng_registro_selecionado
    End If
    lsub_alterar_estado_controles True
    lsub_limpar_campos
    lsub_preencher_grade
fim_cmd_salvar_Click:
    Exit Sub
erro_cmd_salvar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "cmd_salvar_Click"
    GoTo fim_cmd_salvar_Click
End Sub

Private Sub Form_Initialize()
    On Error GoTo Erro_Form_Initialize
    InitCommonControls
Fim_Form_Initialize:
    Exit Sub
Erro_Form_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "Form_Initialize"
    GoTo Fim_Form_Initialize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/cadastros_receitas.htm", 0
        Case vbKeyF2
            cmd_inserir_Click
        Case vbKeyF3
            cmd_editar_Click
        Case vbKeyF4
            cmd_excluir_Click
        Case vbKeyF5
            cmd_atualizar_Click
        Case vbKeyF6
            cmd_fechar_Click
        Case vbKeyF7
            cmd_salvar_Click
        Case vbKeyF8
            cmd_cancelar_Click
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
    lsub_preencher_grade
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "Form_Load"
    GoTo fim_Form_Load
End Sub

Private Sub lsub_alterar_estado_controles(ByVal pbln_habilitar As Boolean)
    On Error GoTo erro_lsub_alterar_estado_controles
    chk_ativo.Enabled = IIf(mbte_ac = mcst_inserir, False, True)
    cmd_inserir.Enabled = pbln_habilitar
    cmd_editar.Enabled = pbln_habilitar
    cmd_salvar.Enabled = Not pbln_habilitar
    cmd_cancelar.Enabled = Not pbln_habilitar
    cmd_excluir.Enabled = pbln_habilitar
    cmd_atualizar.Enabled = pbln_habilitar
    cmd_fechar.Enabled = pbln_habilitar
    fme_campos.Enabled = Not pbln_habilitar
    msf_grade.Enabled = pbln_habilitar
fim_lsub_alterar_estado_controles:
    Exit Sub
erro_lsub_alterar_estado_controles:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "lsub_alterar_estado_controles"
    GoTo fim_lsub_alterar_estado_controles
End Sub

Private Sub lsub_limpar_campos()
    On Error GoTo erro_lsub_limpar_campos
    chk_fixa.Value = vbUnchecked
    chk_ativo.Value = IIf(mbte_ac = mcst_inserir, vbChecked, vbUnchecked)
    txt_receita.Text = ""
    txt_observacoes.Text = ""
fim_lsub_limpar_campos:
    Exit Sub
erro_lsub_limpar_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "lsub_limpar_campos"
    GoTo fim_lsub_limpar_campos
End Sub

Private Function lfct_salvar_registro(ByVal pbln_novo As Boolean, Optional ByVal plng_codigo As Long = 0) As Boolean
    On Error GoTo erro_lfct_salvar_registro
    Dim lobj_salvar As Object
    Dim lstr_sql As String
    If (pbln_novo) Then
        'monta o comando sql
        lstr_sql = ""
        lstr_sql = lstr_sql & " insert into [tb_receitas] "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " [str_descricao], "
        lstr_sql = lstr_sql & " [str_observacoes], "
        lstr_sql = lstr_sql & " [chr_ativo], "
        lstr_sql = lstr_sql & " [chr_fixa] "
        lstr_sql = lstr_sql & " ) values ( "
        lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(txt_receita.Text) & "', "
        lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(txt_observacoes.Text) & "', "
        lstr_sql = lstr_sql & " '" & IIf(chk_ativo.Value = vbChecked, "S", "N") & "', "
        lstr_sql = lstr_sql & " '" & IIf(chk_fixa.Value = vbChecked, "S", "N") & "' "
        lstr_sql = lstr_sql & " ) "
    Else
        lstr_sql = ""
        lstr_sql = lstr_sql & " update [tb_receitas] set "
        lstr_sql = lstr_sql & " [str_descricao] = '" & pfct_tratar_texto_sql(txt_receita.Text) & "', "
        lstr_sql = lstr_sql & " [str_observacoes] = '" & pfct_tratar_texto_sql(txt_observacoes.Text) & "', "
        lstr_sql = lstr_sql & " [chr_ativo] = '" & IIf(chk_ativo.Value = vbChecked, "S", "N") & "', "
        lstr_sql = lstr_sql & " [chr_fixa] = '" & IIf(chk_fixa.Value = vbChecked, "S", "N") & "' "
        lstr_sql = lstr_sql & " where "
        lstr_sql = lstr_sql & " [int_codigo] = " & pfct_tratar_numero_sql(plng_codigo) & " "
    End If
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_salvar, lstr_sql, "frm_cadastro_receitas", "lfct_salvar_registro")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lfct_salvar_registro
    End If
    lfct_salvar_registro = True
fim_lfct_salvar_registro:
    'destrói os objetos
    Set lobj_salvar = Nothing
    Exit Function
erro_lfct_salvar_registro:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "lfct_salvar_registro"
    GoTo fim_lfct_salvar_registro
End Function

Private Function lfct_excluir_registro(ByVal plng_codigo As Long) As Boolean
    On Error GoTo erro_lfct_excluir_registro
    Dim lobj_excluir As Object
    Dim lobj_movimentacao As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    'monta o comando sql
    lstr_sql = "select * from [tb_movimentacao] where int_receita = " & pfct_tratar_numero_sql(plng_codigo)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_movimentacao, lstr_sql, "frm_cadastro_receitas", "lfct_excluir_registro")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lfct_excluir_registro
    End If
    llng_registros = lobj_movimentacao.Count
    If (llng_registros > 0) Then
        MsgBox "Atenção!" & vbCrLf & "Não é possível excluir a receita selecionada pois existem [" & CStr(llng_registros) & "] movimentações relacionadas.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        GoTo fim_lfct_excluir_registro
    Else
        'monta o comando sql
        lstr_sql = "delete from [tb_receitas] where int_codigo = " & pfct_tratar_numero_sql(plng_codigo)
        'executa o comando sql e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_excluir, lstr_sql, "frm_cadastro_receitas", "lfct_excluir_registro")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_lfct_excluir_registro
        End If
    End If
    'retorna true
    lfct_excluir_registro = True
fim_lfct_excluir_registro:
    'destrói os objetos
    Set lobj_excluir = Nothing
    Set lobj_movimentacao = Nothing
    Exit Function
erro_lfct_excluir_registro:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "lfct_excluir_registro"
    GoTo fim_lfct_excluir_registro
End Function

Private Function lfct_verificar_selecao() As Boolean
    On Error GoTo erro_lfct_verificar_selecao
    If (mlng_registro_selecionado = 0) Then
        MsgBox "Selecione um item na grade.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        GoTo fim_lfct_verificar_selecao
    End If
    lfct_verificar_selecao = True
fim_lfct_verificar_selecao:
    Exit Function
erro_lfct_verificar_selecao:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "lfct_verificar_selecao"
    GoTo fim_lfct_verificar_selecao
End Function

Private Sub lsub_preencher_grade()
    On Error GoTo erro_lsub_preencher_grade
    Dim lobj_receitas As Object
    Dim lstr_sql As String
    Dim llng_contador As Long
    Dim llng_registros As Long
    'monta a grade de receitas
    With msf_grade
        .Clear
        .Rows = 2
        .Row = 0
        .Col = enm_receita.col_receita
        .ColWidth(enm_receita.col_receita) = 3150
        .Text = "Receitas"
    End With
    'monta o comando sql
    lstr_sql = " select * from [tb_receitas] order by [str_descricao] asc "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_receitas, lstr_sql, "frm_cadastro_receitas", "lsub_preencher_grade")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_preencher_grade
    End If
    llng_registros = lobj_receitas.Count
    If (llng_registros > 0) Then
        msf_grade.Redraw = False
        For llng_contador = 1 To llng_registros
            msf_grade.Row = llng_contador
            msf_grade.Col = enm_receita.col_receita
            msf_grade.RowData(llng_contador) = lobj_receitas(llng_contador)("int_codigo")
            msf_grade.TextMatrix(llng_contador, enm_receita.col_receita) = lobj_receitas(llng_contador)("str_descricao")
            If (llng_contador < llng_registros) Then
                msf_grade.Rows = msf_grade.Rows + 1
            End If
        Next
        msf_grade.Redraw = True
        msf_grade.Row = 1
        stb_status.Panels(enm_status.pnl_mensagem).Text = "Total de receitas cadastradas: " & llng_registros
        lsub_preencher_campos
        msf_grade_Click
    Else
        stb_status.Panels(enm_status.pnl_mensagem).Text = "Não há receitas cadastradas"
        msf_grade.TextMatrix(1, enm_receita.col_receita) = "Não há receitas cadastradas"
    End If
fim_lsub_preencher_grade:
    'destrói os objetos
    Set lobj_receitas = Nothing
    Exit Sub
erro_lsub_preencher_grade:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "lsub_preencher_grade"
    GoTo fim_lsub_preencher_grade
End Sub

Private Sub lsub_preencher_campos()
    On Error GoTo erro_lsub_preencher_campos
    Dim lobj_campos As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    'monta o comando sql
    lstr_sql = " select * from [tb_receitas] where [int_codigo] = " & pfct_tratar_numero_sql(msf_grade.RowData(msf_grade.Row))
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_campos, lstr_sql, "frm_cadastro_receitas", "lsub_preencher_campos")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_lsub_preencher_campos
    End If
    llng_registros = lobj_campos.Count
    If (llng_registros > 0) Then
        chk_fixa.Value = IIf(lobj_campos(1)("chr_fixa") = "S", vbChecked, vbUnchecked)
        chk_ativo.Value = IIf(lobj_campos(1)("chr_ativo") = "S", vbChecked, vbUnchecked)
        txt_receita.Text = lobj_campos(1)("str_descricao")
        txt_observacoes.Text = lobj_campos(1)("str_observacoes")
    End If
fim_lsub_preencher_campos:
    'destrói os objetos
    Set lobj_campos = Nothing
    Exit Sub
erro_lsub_preencher_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "lsub_preencher_campos"
    GoTo fim_lsub_preencher_campos
End Sub

Private Sub mnu_msf_grade_copiar_Click()
    On Error GoTo erro_mnu_msf_grade_copiar_Click
    pfct_copiar_conteudo_grade msf_grade
fim_mnu_msf_grade_copiar_Click:
    Exit Sub
erro_mnu_msf_grade_copiar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "mnu_msf_grade_copiar_Click"
    GoTo fim_mnu_msf_grade_copiar_Click
End Sub

Private Sub mnu_msf_grade_exportar_Click()
    On Error GoTo erro_mnu_msf_grade_exportar_Click
    pfct_exportar_conteudo_grade msf_grade, "receitas"
fim_mnu_msf_grade_exportar_Click:
    Exit Sub
erro_mnu_msf_grade_exportar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "mnu_msf_grade_exportar_Click"
    GoTo fim_mnu_msf_grade_exportar_Click
End Sub

Private Sub msf_grade_Click()
    On Error GoTo erro_msf_grade_Click
    mlng_registro_selecionado = msf_grade.RowData(msf_grade.Row)
    lsub_preencher_campos
fim_msf_grade_Click:
    Exit Sub
erro_msf_grade_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "msf_grade_Click"
    GoTo fim_msf_grade_Click
End Sub

Private Sub msf_grade_EnterCell()
    On Error GoTo erro_msf_grade_EnterCell
    psub_campo_got_focus msf_grade
fim_msf_grade_EnterCell:
    Exit Sub
erro_msf_grade_EnterCell:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "msf_grade_EnterCell"
    GoTo fim_msf_grade_EnterCell
End Sub

Private Sub msf_grade_LeaveCell()
    On Error GoTo erro_msf_grade_LeaveCell
    psub_campo_lost_focus msf_grade
fim_msf_grade_LeaveCell:
    Exit Sub
erro_msf_grade_LeaveCell:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "msf_grade_LeaveCell"
    GoTo fim_msf_grade_LeaveCell
End Sub

Private Sub msf_grade_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_msf_grade_MouseUp
    If (Button = 2) Then 'botão direito do mouse
        PopupMenu mnu_msf_grade 'exibimos o popup
    End If
fim_msf_grade_MouseUp:
    Exit Sub
erro_msf_grade_MouseUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "msf_grade_MouseUp"
    GoTo fim_msf_grade_MouseUp
End Sub
Private Sub txt_observacoes_GotFocus()
    On Error GoTo erro_txt_observacoes_gotFocus
    psub_campo_got_focus txt_observacoes
fim_txt_observacoes_gotFocus:
    Exit Sub
erro_txt_observacoes_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "txt_observacoes_GotFocus"
    GoTo fim_txt_observacoes_gotFocus
End Sub

Private Sub txt_observacoes_LostFocus()
    On Error GoTo erro_txt_observacoes_LostFocus
    psub_campo_lost_focus txt_observacoes
fim_txt_observacoes_LostFocus:
    Exit Sub
erro_txt_observacoes_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "txt_observacoes_LostFocus"
    GoTo fim_txt_observacoes_LostFocus
End Sub

Private Sub txt_observacoes_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_observacoes_validate
    psub_tratar_campo txt_observacoes
fim_txt_observacoes_validate:
    Exit Sub
erro_txt_observacoes_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "txt_observacoes_validate"
    GoTo fim_txt_observacoes_validate
End Sub

Private Sub txt_receita_GotFocus()
    On Error GoTo erro_txt_receita_gotFocus
    psub_campo_got_focus txt_receita
fim_txt_receita_gotFocus:
    Exit Sub
erro_txt_receita_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "txt_receita_GotFocus"
    GoTo fim_txt_receita_gotFocus
End Sub

Private Sub txt_receita_LostFocus()
    On Error GoTo erro_txt_receita_LostFocus
    psub_campo_lost_focus txt_receita
fim_txt_receita_LostFocus:
    Exit Sub
erro_txt_receita_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "txt_receita_LostFocus"
    GoTo fim_txt_receita_LostFocus
End Sub

Private Sub txt_receita_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_receita_validate
    psub_tratar_campo txt_receita
fim_txt_receita_validate:
    Exit Sub
erro_txt_receita_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_cadastro_receitas", "txt_receita_validate"
    GoTo fim_txt_receita_validate
End Sub
