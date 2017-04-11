VERSION 5.00
Begin VB.Form frm_usuario_alterar_senha 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Alterar Senha"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6435
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
   ScaleHeight     =   1635
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_lembrete_senha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   40
      TabIndex        =   7
      Top             =   1200
      Width           =   4125
   End
   Begin VB.TextBox txt_confirmar_nova_senha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4320
      MaxLength       =   32
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   420
      Width           =   2025
   End
   Begin VB.TextBox txt_nova_senha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2220
      MaxLength       =   32
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   420
      Width           =   2025
   End
   Begin VB.TextBox txt_senha_atual 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   32
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   420
      Width           =   2025
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5355
      TabIndex        =   9
      Top             =   1165
      Width           =   990
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   375
      Left            =   4290
      TabIndex        =   8
      Top             =   1165
      Width           =   1020
   End
   Begin VB.Label lbl_lembrete_senha 
      AutoSize        =   -1  'True
      Caption         =   "&Lembrete de senha:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   900
      Width           =   1440
   End
   Begin VB.Label lbl_confirmar_nova_senha 
      AutoSize        =   -1  'True
      Caption         =   "&Confirmar nova senha:"
      Height          =   195
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   1650
   End
   Begin VB.Label lbl_nova_senha 
      AutoSize        =   -1  'True
      Caption         =   "&Nova senha:"
      Height          =   195
      Left            =   2220
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lbl_senha_atual 
      AutoSize        =   -1  'True
      Caption         =   "&Senha atual:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "frm_usuario_alterar_senha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function lfct_valida_campos() As Boolean
    On Error GoTo erro_lfct_valida_campos
    Dim lstr_senha As String
    Dim lstr_nova_senha As String
    Dim lstr_confirmar_nova_senha As String
    Me.ValidateControls
    lstr_senha = txt_senha_atual.Text
    lstr_nova_senha = txt_nova_senha.Text
    lstr_confirmar_nova_senha = txt_confirmar_nova_senha.Text
    'senha atual
    If (lstr_senha = "") Then
        MsgBox "Campo [senha atual] não pode estar em branco.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_senha_atual.Text = ""
        txt_senha_atual.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    If (Len(lstr_senha) < 4) Then
        MsgBox "Campo [senha atual] deve conter no mínimo 04 (quatro) caracteres.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_senha_atual.Text = ""
        txt_senha_atual.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    'nova senha
    If (lstr_nova_senha = "") Then
        MsgBox "Campo [nova senha] não pode estar em branco.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_nova_senha.Text = ""
        txt_nova_senha.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    If (Len(lstr_nova_senha) < 4) Then
        MsgBox "Campo [nova senha] deve conter no mínimo 04 (quatro) caracteres.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_nova_senha.Text = ""
        txt_nova_senha.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    'confirmar senha
    If (lstr_confirmar_nova_senha = "") Then
        MsgBox "Campo [confirmar nova senha] não pode estar em branco.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_confirmar_nova_senha.Text = ""
        txt_confirmar_nova_senha.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    If (Len(lstr_confirmar_nova_senha) < 4) Then
        MsgBox "Campo [confirmar senha] deve conter no mínimo 04 (quatro) caracteres.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_confirmar_nova_senha.Text = ""
        txt_confirmar_nova_senha.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    'nova senha <> confirmar nova senha?
    If (txt_nova_senha.Text <> txt_confirmar_nova_senha.Text) Then
        MsgBox "Campos [nova senha] e [confirmar nova senha] não podem ser diferentes.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_nova_senha.Text = ""
        txt_confirmar_nova_senha.Text = ""
        txt_nova_senha.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    'senha digitada <> senha atual?
    If (txt_senha_atual.Text <> p_usuario.str_senha) Then
        MsgBox "Campo [senha atual] digitado inválido.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_senha_atual.Text = ""
        txt_nova_senha.Text = ""
        txt_confirmar_nova_senha.Text = ""
        txt_senha_atual.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    lfct_valida_campos = True
fim_lfct_valida_campos:
    Exit Function
erro_lfct_valida_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "lfct_valida_campos"
    GoTo fim_lfct_valida_campos
End Function

Private Sub lsub_preencher_dados()
    On Error GoTo erro_lsub_preencher_dados
    txt_lembrete_senha.Text = p_usuario.str_lembrete_senha
fim_lsub_preencher_dados:
    Exit Sub
erro_lsub_preencher_dados:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "lsub_preencher_dados"
    GoTo fim_lsub_preencher_dados
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "cmd_cancelar_click"
    GoTo fim_cmd_cancelar_Click
End Sub

Private Sub cmd_alterar_Click()
    On Error GoTo erro_cmd_alterar_click
    Dim lstr_nova_senha As String
    Dim lstr_lembrete_senha As String
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_alterar.Enabled) Then
        Exit Sub
    End If
    lstr_nova_senha = txt_nova_senha.Text
    lstr_lembrete_senha = txt_lembrete_senha.Text
    If (lfct_valida_campos()) Then
        p_usuario.str_senha = lstr_nova_senha
        p_usuario.str_lembrete_senha = lstr_lembrete_senha
        '
        p_banco.tb_tipo_banco = tb_config
        pfct_ajustar_caminho_banco tb_config
        '
        psub_atualizar_usuario p_usuario.lng_codigo, True
        '
        p_banco.tb_tipo_banco = tb_dados
        pfct_ajustar_caminho_banco tb_dados
        '
        MsgBox "Senha do usuário [" & p_usuario.str_login & "] alterada com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        Unload Me
    End If
fim_cmd_alterar_click:
    Exit Sub
erro_cmd_alterar_click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "cmd_login_click"
    GoTo fim_cmd_alterar_click
End Sub

Private Sub Form_Initialize()
    On Error GoTo Erro_Form_Initialize
    InitCommonControls
Fim_Form_Initialize:
    Exit Sub
Erro_Form_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "form_initialize"
    GoTo Fim_Form_Initialize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/usuarios_alterar_senha.htm", 0
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
    lsub_preencher_dados
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "Form_Load"
    GoTo fim_Form_Load
End Sub

Private Sub txt_lembrete_senha_GotFocus()
    On Error GoTo erro_txt_lembrete_senha_gotFocus
    psub_campo_got_focus txt_lembrete_senha
fim_txt_lembrete_senha_gotFocus:
    Exit Sub
erro_txt_lembrete_senha_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "txt_lembrete_senha_GotFocus"
    GoTo fim_txt_lembrete_senha_gotFocus
End Sub

Private Sub txt_lembrete_senha_LostFocus()
    On Error GoTo erro_txt_lembrete_senha_LostFocus
    psub_campo_lost_focus txt_lembrete_senha
fim_txt_lembrete_senha_LostFocus:
    Exit Sub
erro_txt_lembrete_senha_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "txt_lembrete_senha_LostFocus"
    GoTo fim_txt_lembrete_senha_LostFocus
End Sub

Private Sub txt_lembrete_senha_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_lembrete_senha_validate
    psub_tratar_campo txt_lembrete_senha
    Cancel = Not pfct_validar_campo(txt_lembrete_senha, tc_texto)
fim_txt_lembrete_senha_validate:
    Exit Sub
erro_txt_lembrete_senha_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "txt_lembrete_senha_validate"
    GoTo fim_txt_lembrete_senha_validate
End Sub

Private Sub txt_nova_senha_GotFocus()
    On Error GoTo erro_txt_nova_senha_gotFocus
    psub_campo_got_focus txt_nova_senha
fim_txt_nova_senha_gotFocus:
    Exit Sub
erro_txt_nova_senha_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "txt_nova_senha_GotFocus"
    GoTo fim_txt_nova_senha_gotFocus
End Sub

Private Sub txt_nova_senha_LostFocus()
    On Error GoTo erro_txt_nova_senha_LostFocus
    psub_campo_lost_focus txt_nova_senha
fim_txt_nova_senha_LostFocus:
    Exit Sub
erro_txt_nova_senha_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "txt_nova_senha_LostFocus"
    GoTo fim_txt_nova_senha_LostFocus
End Sub

Private Sub txt_confirmar_nova_senha_GotFocus()
    On Error GoTo erro_txt_confirmar_nova_senha_gotFocus
    psub_campo_got_focus txt_confirmar_nova_senha
fim_txt_confirmar_nova_senha_gotFocus:
    Exit Sub
erro_txt_confirmar_nova_senha_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "txt_confirmar_nova_senha_GotFocus"
    GoTo fim_txt_confirmar_nova_senha_gotFocus
End Sub

Private Sub txt_confirmar_nova_senha_LostFocus()
    On Error GoTo erro_txt_confirmar_nova_senha_LostFocus
    psub_campo_lost_focus txt_confirmar_nova_senha
fim_txt_confirmar_nova_senha_LostFocus:
    Exit Sub
erro_txt_confirmar_nova_senha_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "txt_confirmar_nova_senha_LostFocus"
    GoTo fim_txt_confirmar_nova_senha_LostFocus
End Sub

Private Sub txt_nova_senha_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_nova_senha_validate
    Cancel = Not pfct_validar_campo(txt_nova_senha, tc_texto)
fim_txt_nova_senha_validate:
    Exit Sub
erro_txt_nova_senha_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "txt_nova_senha_validate"
    GoTo fim_txt_nova_senha_validate
End Sub

Private Sub txt_confirmar_nova_senha_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_confirmar_nova_senha_validate
    Cancel = Not pfct_validar_campo(txt_confirmar_nova_senha, tc_texto)
fim_txt_confirmar_nova_senha_validate:
    Exit Sub
erro_txt_confirmar_nova_senha_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "txt_confirmar_nova_senha_validate"
    GoTo fim_txt_confirmar_nova_senha_validate
End Sub

Private Sub txt_senha_atual_GotFocus()
    On Error GoTo erro_txt_senha_atual_gotFocus
    psub_campo_got_focus txt_senha_atual
fim_txt_senha_atual_gotFocus:
    Exit Sub
erro_txt_senha_atual_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "txt_senha_atual_GotFocus"
    GoTo fim_txt_senha_atual_gotFocus
End Sub

Private Sub txt_senha_atual_LostFocus()
    On Error GoTo erro_txt_senha_atual_LostFocus
    psub_campo_lost_focus txt_senha_atual
fim_txt_senha_atual_LostFocus:
    Exit Sub
erro_txt_senha_atual_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "txt_senha_atual_LostFocus"
    GoTo fim_txt_senha_atual_LostFocus
End Sub

Private Sub txt_senha_atual_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_senha_atual_validate
    Cancel = Not pfct_validar_campo(txt_senha_atual, tc_texto)
fim_txt_senha_atual_validate:
    Exit Sub
erro_txt_senha_atual_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_alterar_senha", "txt_senha_atual_validate"
    GoTo fim_txt_senha_atual_validate
End Sub

