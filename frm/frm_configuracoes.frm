VERSION 5.00
Begin VB.Form frm_configuracoes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Configurações "
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4995
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
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra_configuracoes_gerais 
      Caption         =   " Configurações g&erais "
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   3420
      Width           =   4755
      Begin VB.CheckBox chk_carregar_agenda_financeira_login 
         Caption         =   "C&arregar agenda financeira no Login"
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Width           =   4455
      End
      Begin VB.CheckBox chk_nao_permitir_lancamentos_duplicados 
         Caption         =   "&Não permitir lançamentos duplicados"
         Height          =   315
         Left            =   180
         TabIndex        =   16
         Top             =   900
         Width           =   4455
      End
      Begin VB.CheckBox chk_considerar_data_vencimento_baixa_imediata 
         Caption         =   "&Considerar data de vencimento na baixa imediata"
         Height          =   315
         Left            =   180
         TabIndex        =   18
         Top             =   1500
         Width           =   4455
      End
      Begin VB.CheckBox chk_permitir_alterar_dados_detalhes 
         Caption         =   "P&ermitir alterações de detalhes (movimentação)"
         Height          =   315
         Left            =   180
         TabIndex        =   17
         Top             =   1200
         Width           =   4455
      End
      Begin VB.CheckBox chk_permitir_lancamentos_retroativos 
         Caption         =   "&Permitir lançamentos retroativos"
         Height          =   315
         Left            =   180
         TabIndex        =   15
         Top             =   600
         Width           =   4455
      End
   End
   Begin VB.Frame fra_configuracoes_data 
      Caption         =   " &Data "
      Height          =   1155
      Left            =   2520
      TabIndex        =   10
      Top             =   2160
      Width           =   2355
      Begin VB.ComboBox cbo_intervalo_data 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   660
         Width           =   2115
      End
      Begin VB.Label lbl_intervalo_data 
         AutoSize        =   -1  'True
         Caption         =   "Intervalo de data:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Width           =   1320
      End
   End
   Begin VB.Frame fra_configuracoes_moeda 
      Caption         =   " &Moeda "
      Height          =   1155
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
      Begin VB.ComboBox cbo_moeda 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label lbl_moeda 
         AutoSize        =   -1  'True
         Caption         =   "&Moeda:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame fra_banco_dados 
      Caption         =   " &Banco de dados "
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4755
      Begin VB.CheckBox chk_ativar_backup 
         Caption         =   "&Ativar backup do banco"
         Height          =   315
         Left            =   2655
         TabIndex        =   3
         Top             =   660
         Width           =   1995
      End
      Begin VB.CommandButton cmd_pesquisar 
         Caption         =   "..."
         Height          =   315
         Left            =   4200
         TabIndex        =   6
         Top             =   1440
         Width           =   435
      End
      Begin VB.TextBox txt_pasta_backup 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   4035
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   660
         Width           =   2475
      End
      Begin VB.Label lbl_pasta_backup 
         AutoSize        =   -1  'True
         Caption         =   "&Pasta para backup:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1140
         Width           =   1395
      End
      Begin VB.Label lbl_periodo 
         AutoSize        =   -1  'True
         Caption         =   "&Período:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar (F3)"
      Height          =   435
      Left            =   3660
      TabIndex        =   20
      Top             =   5460
      Width           =   1215
   End
   Begin VB.CommandButton cmd_aplicar 
      Caption         =   "&Aplicar (F2)"
      Height          =   435
      Left            =   2400
      TabIndex        =   19
      Top             =   5460
      Width           =   1215
   End
End
Attribute VB_Name = "frm_configuracoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lsub_habilitar_campos_backup(ByVal pbln_habilitar As Boolean)
    On Error GoTo erro_lsub_habilitar_campos_backup
    lbl_periodo.Enabled = pbln_habilitar
    lbl_pasta_backup.Enabled = pbln_habilitar
    cbo_periodo.Enabled = pbln_habilitar
    txt_pasta_backup.Enabled = pbln_habilitar
    cmd_pesquisar.Enabled = pbln_habilitar
fim_lsub_habilitar_campos_backup:
    Exit Sub
erro_lsub_habilitar_campos_backup:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "lsub_habilitar_campos_backup"
    GoTo fim_lsub_habilitar_campos_backup
End Sub

Private Sub lsub_preencher_configuracoes()
    On Error GoTo erro_lsub_preencher_configuracoes
    'backup
    chk_ativar_backup.Value = IIf(p_backup.bln_ativar, vbChecked, vbUnchecked)
    cbo_periodo.ListIndex = p_backup.pb_periodo_backup
    txt_pasta_backup.Text = p_backup.str_caminho
    'moeda
    cbo_moeda.ListIndex = p_usuario.sm_simbolo_moeda
    'intervalo de data
    cbo_intervalo_data.ListIndex = p_usuario.id_intervalo_data
    'geral
    chk_carregar_agenda_financeira_login.Value = IIf(p_usuario.bln_carregar_agenda_financeira_login, vbChecked, vbUnchecked)
    chk_permitir_lancamentos_retroativos.Value = IIf(p_usuario.bln_lancamentos_retroativos, vbChecked, vbUnchecked)
    chk_nao_permitir_lancamentos_duplicados.Value = IIf(p_usuario.bln_lancamentos_duplicados, vbChecked, vbUnchecked)
    chk_permitir_alterar_dados_detalhes.Value = IIf(p_usuario.bln_alteracoes_detalhes, vbChecked, vbUnchecked)
    chk_considerar_data_vencimento_baixa_imediata.Value = IIf(p_usuario.bln_data_vencimento_baixa_imediata, vbChecked, vbUnchecked)
    'dispara o evento click do componente
    chk_ativar_backup_Click
fim_lsub_preencher_configuracoes:
    Exit Sub
erro_lsub_preencher_configuracoes:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "lsub_preencher_configuracoes"
    GoTo fim_lsub_preencher_configuracoes
End Sub

Private Sub lsub_atribuir_configuracoes()
    On Error GoTo erro_lsub_atribuir_configuracoes
    Dim ldt_nova_data As Date
    'backup
    p_backup.bln_ativar = IIf(chk_ativar_backup.Value = vbChecked, True, False)
    'se foi alterado o período
    If (p_backup.pb_periodo_backup <> cbo_periodo.ListIndex) Then
        'altera o período de backup
        p_backup.pb_periodo_backup = cbo_periodo.ListIndex
        'verifica qual período foi selecionado
        Select Case p_backup.pb_periodo_backup
            Case enm_periodo_backup.pb_diario
                ldt_nova_data = DateAdd("d", 1, Now)
            Case enm_periodo_backup.pb_semanal
                ldt_nova_data = DateAdd("d", 7, Now)
            Case enm_periodo_backup.pb_quinzenal
                ldt_nova_data = DateAdd("d", 15, Now)
            Case enm_periodo_backup.pb_mensal
                ldt_nova_data = DateAdd("m", 1, Now)
        End Select
        'altera a data do próximo backup
        p_backup.dt_proximo_backup = ldt_nova_data
    End If
    p_backup.str_caminho = txt_pasta_backup.Text
    'moeda
    p_usuario.sm_simbolo_moeda = cbo_moeda.ListIndex
    'intervalo de data
    p_usuario.id_intervalo_data = cbo_intervalo_data.ListIndex
    'geral
    p_usuario.bln_carregar_agenda_financeira_login = IIf(chk_carregar_agenda_financeira_login.Value = vbChecked, True, False)
    p_usuario.bln_lancamentos_retroativos = IIf(chk_permitir_lancamentos_retroativos.Value = vbChecked, True, False)
    p_usuario.bln_lancamentos_duplicados = IIf(chk_nao_permitir_lancamentos_duplicados.Value = vbChecked, True, False)
    p_usuario.bln_alteracoes_detalhes = IIf(chk_permitir_alterar_dados_detalhes.Value = vbChecked, True, False)
    p_usuario.bln_data_vencimento_baixa_imediata = IIf(chk_considerar_data_vencimento_baixa_imediata.Value = vbChecked, True, False)
fim_lsub_atribuir_configuracoes:
    Exit Sub
erro_lsub_atribuir_configuracoes:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "lsub_atribuir_configuracoes"
    GoTo fim_lsub_atribuir_configuracoes
End Sub

Private Function lfct_validar_campos() As Boolean
    On Error GoTo erro_lfct_validar_campos
    If (chk_ativar_backup.Value = vbChecked) Then
        'período de backup
        If (cbo_periodo.ListIndex = enm_periodo_backup.pb_selecione) Then
            MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [período].", vbOKOnly + vbInformation, pcst_nome_aplicacao
            cbo_periodo.SetFocus
            GoTo fim_lfct_validar_campos
        End If
        'pasta de backup
        If (txt_pasta_backup.Text = "") Then
            MsgBox "Atenção!" & vbCrLf & "Selecione uma [pasta para backup].", vbOKOnly + vbInformation, pcst_nome_aplicacao
            txt_pasta_backup.SetFocus
            GoTo fim_lfct_validar_campos
        End If
    End If
    'moeda
    If (cbo_moeda.ListIndex = enm_simbolo_moeda.sm_selecione) Then
        MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [moeda].", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_moeda.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    'intervalo de data
    If (cbo_intervalo_data.ListIndex = enm_intervalo_data.id_selecione) Then
        MsgBox "Atenção!" & vbCrLf & "Selecione um item no campo [intervalo de data].", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_intervalo_data.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    lfct_validar_campos = True
fim_lfct_validar_campos:
    Exit Function
erro_lfct_validar_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "lfct_validar_campos"
    GoTo fim_lfct_validar_campos
End Function

Private Sub cbo_periodo_DropDown()
    On Error GoTo erro_cbo_periodo_DropDown
    psub_campo_got_focus cbo_periodo
fim_cbo_periodo_DropDown:
    Exit Sub
erro_cbo_periodo_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "cbo_periodo_DropDown"
    GoTo fim_cbo_periodo_DropDown
End Sub

Private Sub cbo_periodo_GotFocus()
    On Error GoTo erro_cbo_periodo_gotFocus
    psub_campo_got_focus cbo_periodo
fim_cbo_periodo_gotFocus:
    Exit Sub
erro_cbo_periodo_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "cbo_periodo_GotFocus"
    GoTo fim_cbo_periodo_gotFocus
End Sub

Private Sub cbo_periodo_LostFocus()
    On Error GoTo erro_cbo_periodo_LostFocus
    psub_campo_lost_focus cbo_periodo
fim_cbo_periodo_LostFocus:
    Exit Sub
erro_cbo_periodo_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "cbo_periodo_LostFocus"
    GoTo fim_cbo_periodo_LostFocus
End Sub

Private Sub cbo_moeda_DropDown()
    On Error GoTo erro_cbo_moeda_DropDown
    psub_campo_got_focus cbo_moeda
fim_cbo_moeda_DropDown:
    Exit Sub
erro_cbo_moeda_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "cbo_moeda_DropDown"
    GoTo fim_cbo_moeda_DropDown
End Sub

Private Sub cbo_moeda_GotFocus()
    On Error GoTo erro_cbo_moeda_gotFocus
    psub_campo_got_focus cbo_moeda
fim_cbo_moeda_gotFocus:
    Exit Sub
erro_cbo_moeda_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "cbo_moeda_GotFocus"
    GoTo fim_cbo_moeda_gotFocus
End Sub

Private Sub cbo_moeda_LostFocus()
    On Error GoTo erro_cbo_moeda_LostFocus
    psub_campo_lost_focus cbo_moeda
fim_cbo_moeda_LostFocus:
    Exit Sub
erro_cbo_moeda_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "cbo_moeda_LostFocus"
    GoTo fim_cbo_moeda_LostFocus
End Sub

Private Sub cbo_intervalo_data_DropDown()
    On Error GoTo erro_cbo_intervalo_data_DropDown
    psub_campo_got_focus cbo_intervalo_data
fim_cbo_intervalo_data_DropDown:
    Exit Sub
erro_cbo_intervalo_data_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "cbo_intervalo_data_DropDown"
    GoTo fim_cbo_intervalo_data_DropDown
End Sub

Private Sub cbo_intervalo_data_GotFocus()
    On Error GoTo erro_cbo_intervalo_data_gotFocus
    psub_campo_got_focus cbo_intervalo_data
fim_cbo_intervalo_data_gotFocus:
    Exit Sub
erro_cbo_intervalo_data_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "cbo_intervalo_data_GotFocus"
    GoTo fim_cbo_intervalo_data_gotFocus
End Sub

Private Sub cbo_intervalo_data_LostFocus()
    On Error GoTo erro_cbo_intervalo_data_LostFocus
    psub_campo_lost_focus cbo_intervalo_data
fim_cbo_intervalo_data_LostFocus:
    Exit Sub
erro_cbo_intervalo_data_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "cbo_intervalo_data_LostFocus"
    GoTo fim_cbo_intervalo_data_LostFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/configuracoes.htm", 0
        Case vbKeyF2
            cmd_aplicar_Click
        Case vbKeyF3
            cmd_cancelar_Click
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub txt_pasta_backup_GotFocus()
    On Error GoTo erro_txt_pasta_backup_GotFocus
    psub_campo_got_focus txt_pasta_backup
fim_txt_pasta_backup_GotFocus:
    Exit Sub
erro_txt_pasta_backup_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "txt_pasta_backup_GotFocus"
    GoTo fim_txt_pasta_backup_GotFocus
End Sub

Private Sub txt_pasta_backup_LostFocus()
    On Error GoTo erro_txt_pasta_backup_LostFocus
    psub_campo_lost_focus txt_pasta_backup
fim_txt_pasta_backup_LostFocus:
    Exit Sub
erro_txt_pasta_backup_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "txt_pasta_backup_LostFocus"
    GoTo fim_txt_pasta_backup_LostFocus
End Sub

Private Sub chk_ativar_backup_Click()
    On Error GoTo erro_chk_ativar_backup_Click
    If (chk_ativar_backup.Value = vbChecked) Then
        lsub_habilitar_campos_backup True
    Else
        lsub_habilitar_campos_backup False
    End If
fim_chk_ativar_backup_Click:
    Exit Sub
erro_chk_ativar_backup_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "chk_ativar_backup_Click"
    GoTo fim_chk_ativar_backup_Click
End Sub

Private Sub cmd_aplicar_Click()
    On Error GoTo erro_cmd_aplicar_Click
    Dim lbln_retorno As Boolean
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_aplicar.Enabled) Then
        Exit Sub
    End If
    If (lfct_validar_campos()) Then
        lsub_atribuir_configuracoes
        'ajusta o tipo de banco de dados para config
        p_banco.tb_tipo_banco = tb_config
        'salva as configurações do usuário
        lbln_retorno = pfct_salvar_configuracoes_usuario(p_usuario.lng_codigo)
        'ajusta o tipo de banco de dados para dados
        p_banco.tb_tipo_banco = tb_dados
        'se salvou os dados corretamente
        If (lbln_retorno) Then
            Unload Me
        End If
    End If
fim_cmd_aplicar_Click:
    Exit Sub
erro_cmd_aplicar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes_data", "cmd_aplicar_Click"
    GoTo fim_cmd_aplicar_Click
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes_data", "cmd_cancelar_Click"
    GoTo fim_cmd_cancelar_Click
End Sub

Private Sub cmd_pesquisar_Click()
    On Error GoTo erro_cmd_pesquisar_Click
    Dim llng_retorno As Long
    Dim lstr_caminho As String
    Dim lstr_titulo As String
    Dim tBrowseInfo As BrowseInfo
    lstr_titulo = "Selecione uma pasta para backup:"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(lstr_titulo, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    llng_retorno = SHBrowseForFolder(tBrowseInfo)
    If (llng_retorno) Then
        lstr_caminho = Space$(MAX_LENGTH)
        SHGetPathFromIDList llng_retorno, lstr_caminho
        lstr_caminho = Left$(lstr_caminho, InStr(lstr_caminho, vbNullChar) - 1)
        'verifica o último caracter do caminho
        If (Right$(lstr_caminho, 1) <> "\") Then
            lstr_caminho = lstr_caminho & "\"
        End If
        'verifica se é um caminho de rede
        If (Left$(lstr_caminho, 2) = "\\") Then
            'exibe mensagem
            MsgBox "Atenção!" & vbCrLf & "Caminhos de rede não são suportados pelo sistema.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            'desvia ao fim do método
            GoTo fim_cmd_pesquisar_Click
        End If
        'atribui o caminho ao text-box
        txt_pasta_backup.Text = UCase$(lstr_caminho)
    End If
fim_cmd_pesquisar_Click:
    Exit Sub
erro_cmd_pesquisar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "cmd_pesquisar_Click"
    GoTo fim_cmd_pesquisar_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes_data", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
    psub_preencher_periodo_backup cbo_periodo
    psub_preencher_simbolos_moeda cbo_moeda
    psub_preencher_intervalo_data cbo_intervalo_data
    'carrega as configurações
    lsub_preencher_configuracoes
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_configuracoes", "Form_Load"
    GoTo fim_Form_Load
End Sub
