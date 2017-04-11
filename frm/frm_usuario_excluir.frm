VERSION 5.00
Begin VB.Form frm_usuario_excluir 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Excluir Usuário"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   315
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
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_senha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   840
      MaxLength       =   32
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1320
      Width           =   2025
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6795
      TabIndex        =   6
      Top             =   1275
      Width           =   990
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   375
      Left            =   5790
      TabIndex        =   5
      Top             =   1275
      Width           =   965
   End
   Begin VB.Label lbl_senha 
      AutoSize        =   -1  'True
      Caption         =   "&Senha:"
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   1020
      Width           =   510
   End
   Begin VB.Image img_mensagem 
      Height          =   480
      Left            =   180
      Picture         =   "frm_usuario_excluir.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lbl_mensagem_02 
      AutoSize        =   -1  'True
      Caption         =   "Para excluir seu usuário definitivamente, insira sua senha no campo abaixo e clique 'excluir'."
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   6600
   End
   Begin VB.Label lbl_mensagem_01 
      AutoSize        =   -1  'True
      Caption         =   "Ao excluir um usuário, todos os dados referentes ao mesmo serão eliminados permanentemente."
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   6960
   End
   Begin VB.Label lbl_atencao 
      AutoSize        =   -1  'True
      Caption         =   "Atenção!"
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
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "frm_usuario_excluir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function lfct_valida_campos() As Boolean
    On Error GoTo erro_lfct_valida_campos
    Dim lstr_senha As String
    lstr_senha = txt_senha.Text
    Me.ValidateControls
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
        txt_senha.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    'valida senha digitada
    If (txt_senha.Text <> p_usuario.str_senha) Then
        MsgBox "Campo [senha] inválido.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        txt_senha.Text = ""
        txt_senha.SetFocus
        GoTo fim_lfct_valida_campos
    End If
    lfct_valida_campos = True
fim_lfct_valida_campos:
    Exit Function
erro_lfct_valida_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_excluir", "lfct_valida_campos"
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_excluir", "cmd_cancelar_click"
    GoTo fim_cmd_cancelar_Click
End Sub

Private Sub cmd_excluir_Click()
    On Error GoTo erro_cmd_excluir_click
    Dim lint_resposta As Integer
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_excluir.Enabled) Then
        Exit Sub
    End If
    If (lfct_valida_campos()) Then
        'exibe mensagem ao usuário
        lint_resposta = MsgBox("Confirma a exclusão do usuário [" & p_usuario.str_login & "] ?" & vbCrLf & "Esta operação não poderá ser desfeita.", vbYesNo + vbQuestion + vbDefaultButton2, pcst_nome_aplicacao)
        'se usuário confirmou exclusão
        If (lint_resposta = vbYes) Then
            'ajusta o tipo de banco de dados
            p_banco.tb_tipo_banco = tb_config
            'configura o banco de dados
            pfct_ajustar_caminho_banco tb_config
            If (pfct_excluir_usuario(p_usuario.lng_codigo)) Then 'executa exclusão
                MsgBox "Usuário [" & p_usuario.str_login & "] excluído com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                'usuário
                With p_usuario
                    .lng_codigo = 0
                    .str_login = ""
                    .str_senha = ""
                    .str_lembrete_senha = ""
                    .dt_criado_em = "00:00:00"
                    .dt_ultimo_acesso = "00:00:00"
                    .id_intervalo_data = id_selecione
                    .sm_simbolo_moeda = sm_selecione
                End With
                'backup
                With p_backup
                    .bln_ativar = False
                    .pb_periodo_backup = pb_selecione
                    .str_caminho = ""
                    .dt_ultimo_backup = "00:00:00"
                    .dt_proximo_backup = "00:00:00"
                End With
                Unload Me
            End If
        End If
    End If
fim_cmd_excluir_click:
    Exit Sub
erro_cmd_excluir_click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_excluir", "cmd_login_click"
    GoTo fim_cmd_excluir_click
End Sub

Private Sub Form_Activate()
    On Error GoTo Erro_Form_Activate
    Beep
Fim_Form_Activate:
    Exit Sub
Erro_Form_Activate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_excluir", "form_activate"
    GoTo Fim_Form_Activate
End Sub

Private Sub Form_Initialize()
    On Error GoTo Erro_Form_Initialize
    InitCommonControls
Fim_Form_Initialize:
    Exit Sub
Erro_Form_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_excluir", "form_initialize"
    GoTo Fim_Form_Initialize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_excluir", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/usuarios_excluir.htm", 0
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_excluir", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub txt_senha_GotFocus()
    On Error GoTo erro_txt_senha_gotFocus
    psub_campo_got_focus txt_senha
fim_txt_senha_gotFocus:
    Exit Sub
erro_txt_senha_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_excluir", "txt_senha_GotFocus"
    GoTo fim_txt_senha_gotFocus
End Sub

Private Sub txt_senha_LostFocus()
    On Error GoTo erro_txt_senha_LostFocus
    psub_campo_lost_focus txt_senha
fim_txt_senha_LostFocus:
    Exit Sub
erro_txt_senha_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_excluir", "txt_senha_LostFocus"
    GoTo fim_txt_senha_LostFocus
End Sub

Private Sub txt_senha_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_senha_validate
    Cancel = Not pfct_validar_campo(txt_senha, tc_texto)
fim_txt_senha_validate:
    Exit Sub
erro_txt_senha_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_excluir", "txt_senha_validate"
    GoTo fim_txt_senha_validate
End Sub

