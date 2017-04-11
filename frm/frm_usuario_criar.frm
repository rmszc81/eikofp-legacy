VERSION 5.00
Begin VB.Form frm_usuario_criar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Criar Usuário"
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
   Begin VB.TextBox txt_confirmar_senha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4320
      MaxLength       =   32
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   420
      Width           =   2025
   End
   Begin VB.TextBox txt_senha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2220
      MaxLength       =   32
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   420
      Width           =   2025
   End
   Begin VB.TextBox txt_usuario 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   32
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
   Begin VB.CommandButton cmd_criar 
      Caption         =   "&Criar"
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
   Begin VB.Label lbl_confirmar_senha 
      AutoSize        =   -1  'True
      Caption         =   "&Confirmar senha:"
      Height          =   195
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label lbl_senha 
      AutoSize        =   -1  'True
      Caption         =   "&Senha:"
      Height          =   195
      Left            =   2220
      TabIndex        =   1
      Top             =   120
      Width           =   510
   End
   Begin VB.Label lbl_usuario 
      AutoSize        =   -1  'True
      Caption         =   "&Nome de usuário:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1260
   End
End
Attribute VB_Name = "frm_usuario_criar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function lfct_valida_campos() As Boolean
    On Error GoTo erro_lfct_valida_campos
    Dim lstr_usuario As String
    Dim lstr_senha As String
    Dim lstr_confirmar_senha As String
    lstr_usuario = txt_usuario.Text
    lstr_senha = txt_senha.Text
    lstr_confirmar_senha = txt_confirmar_senha.Text
    Me.ValidateControls
    'usuário
    If (lstr_usuario = "") Then
        MsgBox "Campo [usuário] não pode estar em branco.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_usuario.Text = ""
        txt_usuario.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    If (Len(lstr_usuario) < 2) Then
        MsgBox "Campo [usuário] deve conter no mínimo 02 (dois) caracteres.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_usuario.Text = ""
        txt_usuario.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    'senha
    If (lstr_senha = "") Then
        MsgBox "Campo [senha] não pode estar em branco.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_senha.Text = ""
        txt_senha.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    If (Len(lstr_senha) < 4) Then
        MsgBox "Campo [senha] deve conter no mínimo 04 (quatro) caracteres.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_senha.Text = ""
        txt_confirmar_senha.Text = ""
        txt_senha.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    'confirmar senha
    If (lstr_confirmar_senha = "") Then
        MsgBox "Campo [confirmar senha] não pode estar em branco.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_confirmar_senha.Text = ""
        txt_confirmar_senha.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    If (Len(lstr_confirmar_senha) < 4) Then
        MsgBox "Campo [confirmar senha] deve conter no mínimo 04 (quatro) caracteres.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_confirmar_senha.Text = ""
        txt_confirmar_senha.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    'senha <> confirmar senha
    If (txt_senha.Text <> txt_confirmar_senha.Text) Then
        MsgBox "Campos [senha] e [confirmar senha] não podem ser diferentes.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_senha.Text = ""
        txt_confirmar_senha.Text = ""
        txt_senha.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    'usuário já existe?
    If (pfct_verificar_usuario_existe(lstr_usuario)) Then
        MsgBox "Usuário [" & lstr_usuario & "] já existe na base de dados." & vbCrLf & _
               "Insira outro nome de usuário e tente novamente.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_usuario.Text = ""
        txt_senha.Text = ""
        txt_confirmar_senha.Text = ""
        txt_usuario.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    lfct_valida_campos = True
fim_lfct_valida_campos:
    Exit Function
erro_lfct_valida_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "lfct_valida_campos"
    GoTo fim_lfct_valida_campos
End Function

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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "cmd_cancelar_click"
    GoTo fim_cmd_cancelar_Click
End Sub

Private Sub cmd_criar_Click()
    On Error GoTo erro_cmd_criar_click
    Dim lstr_usuario As String
    Dim lstr_senha As String
    Dim lstr_lembrete_senha As String
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_criar.Enabled) Then
        Exit Sub
    End If
    lstr_usuario = txt_usuario.Text
    lstr_senha = txt_senha.Text
    lstr_lembrete_senha = txt_lembrete_senha.Text
    If (lfct_valida_campos()) Then
        If (pfct_criar_usuario(lstr_usuario, pfct_criptografia(lstr_senha), lstr_lembrete_senha)) Then
            MsgBox "Usuário [" & lstr_usuario & "] criado com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            Unload Me
        Else
            MsgBox "Erro ao criar o usuário [" & lstr_usuario & "].", vbOKOnly + vbCritical, pcst_nome_aplicacao
            txt_usuario.SetFocus
        End If
    End If
fim_cmd_criar_click:
    Exit Sub
erro_cmd_criar_click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "cmd_login_click"
    GoTo fim_cmd_criar_click
End Sub

Private Sub Form_Initialize()
    On Error GoTo Erro_Form_Initialize
    InitCommonControls
Fim_Form_Initialize:
    Exit Sub
Erro_Form_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "form_initialize"
    GoTo Fim_Form_Initialize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/usuarios_criar_usuario.htm", 0
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub txt_lembrete_senha_GotFocus()
    On Error GoTo erro_txt_lembrete_senha_gotFocus
    psub_campo_got_focus txt_lembrete_senha
fim_txt_lembrete_senha_gotFocus:
    Exit Sub
erro_txt_lembrete_senha_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "txt_lembrete_senha_GotFocus"
    GoTo fim_txt_lembrete_senha_gotFocus
End Sub

Private Sub txt_lembrete_senha_LostFocus()
    On Error GoTo erro_txt_lembrete_senha_LostFocus
    psub_campo_lost_focus txt_lembrete_senha
fim_txt_lembrete_senha_LostFocus:
    Exit Sub
erro_txt_lembrete_senha_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "txt_lembrete_senha_LostFocus"
    GoTo fim_txt_lembrete_senha_LostFocus
End Sub

Private Sub txt_lembrete_senha_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_lembrete_senha_validate
    psub_tratar_campo txt_lembrete_senha
    Cancel = Not pfct_validar_campo(txt_lembrete_senha, tc_texto)
fim_txt_lembrete_senha_validate:
    Exit Sub
erro_txt_lembrete_senha_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "txt_lembrete_senha_validate"
    GoTo fim_txt_lembrete_senha_validate
End Sub

Private Sub txt_senha_GotFocus()
    On Error GoTo erro_txt_senha_gotFocus
    psub_campo_got_focus txt_senha
fim_txt_senha_gotFocus:
    Exit Sub
erro_txt_senha_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "txt_senha_GotFocus"
    GoTo fim_txt_senha_gotFocus
End Sub

Private Sub txt_senha_LostFocus()
    On Error GoTo erro_txt_senha_LostFocus
    psub_campo_lost_focus txt_senha
fim_txt_senha_LostFocus:
    Exit Sub
erro_txt_senha_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "txt_senha_LostFocus"
    GoTo fim_txt_senha_LostFocus
End Sub

Private Sub txt_confirmar_senha_GotFocus()
    On Error GoTo erro_txt_confirmar_senha_gotFocus
    psub_campo_got_focus txt_confirmar_senha
fim_txt_confirmar_senha_gotFocus:
    Exit Sub
erro_txt_confirmar_senha_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "txt_confirmar_senha_GotFocus"
    GoTo fim_txt_confirmar_senha_gotFocus
End Sub

Private Sub txt_confirmar_senha_LostFocus()
    On Error GoTo erro_txt_confirmar_senha_LostFocus
    psub_campo_lost_focus txt_confirmar_senha
fim_txt_confirmar_senha_LostFocus:
    Exit Sub
erro_txt_confirmar_senha_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "txt_confirmar_senha_LostFocus"
    GoTo fim_txt_confirmar_senha_LostFocus
End Sub

Private Sub txt_senha_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_senha_validate
    Cancel = Not pfct_validar_campo(txt_senha, tc_texto)
fim_txt_senha_validate:
    Exit Sub
erro_txt_senha_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "txt_senha_validate"
    GoTo fim_txt_senha_validate
End Sub

Private Sub txt_confirmar_senha_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_confirmar_senha_validate
    Cancel = Not pfct_validar_campo(txt_confirmar_senha, tc_texto)
fim_txt_confirmar_senha_validate:
    Exit Sub
erro_txt_confirmar_senha_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "txt_confirmar_senha_validate"
    GoTo fim_txt_confirmar_senha_validate
End Sub

Private Sub txt_usuario_GotFocus()
    On Error GoTo erro_txt_usuario_gotFocus
    psub_campo_got_focus txt_usuario
fim_txt_usuario_gotFocus:
    Exit Sub
erro_txt_usuario_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "txt_usuario_GotFocus"
    GoTo fim_txt_usuario_gotFocus
End Sub

Private Sub txt_usuario_LostFocus()
    On Error GoTo erro_txt_usuario_LostFocus
    psub_campo_lost_focus txt_usuario
fim_txt_usuario_LostFocus:
    Exit Sub
erro_txt_usuario_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "txt_usuario_LostFocus"
    GoTo fim_txt_usuario_LostFocus
End Sub

Private Sub txt_usuario_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_usuario_validate
    psub_tratar_campo txt_usuario
    Cancel = Not pfct_validar_campo(txt_usuario, tc_texto)
fim_txt_usuario_validate:
    Exit Sub
erro_txt_usuario_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_criar", "txt_usuario_validate"
    GoTo fim_txt_usuario_validate
End Sub

