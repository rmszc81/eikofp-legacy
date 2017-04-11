VERSION 5.00
Begin VB.Form frm_usuario_login 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Login"
   ClientHeight    =   1320
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
   ScaleHeight     =   1320
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_criar_novo_usuario 
      Caption         =   "&Criar Usuário"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   855
      Width           =   1305
   End
   Begin VB.TextBox txt_senha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2565
      MaxLength       =   32
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   420
      Width           =   2325
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3870
      TabIndex        =   7
      Top             =   855
      Width           =   990
   End
   Begin VB.CommandButton cmd_login 
      Caption         =   "&Login"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   855
      Width           =   945
   End
   Begin VB.ComboBox cbo_usuarios 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   420
      Width           =   2355
   End
   Begin VB.Label lbl_lembrete_senha 
      AutoSize        =   -1  'True
      Caption         =   "[ ? ]"
      Height          =   195
      Left            =   4620
      TabIndex        =   2
      Top             =   120
      Width           =   285
   End
   Begin VB.Label lbl_senha 
      AutoSize        =   -1  'True
      Caption         =   "&Senha:"
      Height          =   195
      Left            =   2565
      TabIndex        =   1
      Top             =   120
      Width           =   510
   End
   Begin VB.Label lbl_usuarios 
      AutoSize        =   -1  'True
      Caption         =   "&Selecione o usuário:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frm_usuario_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mint_tentativas As Integer
Private mbln_login As Boolean

'variáveis para uso com classe ToolTip
Private mobj_tooltip As CToolTip
Private mstr_titulo As String
Private mstr_mensagem As String
'

Private Function lfct_valida_campos() As Boolean
    On Error GoTo erro_lfct_valida_campos
    Dim lint_codigo_usuario As Integer
    Dim lstr_senha As String
    lint_codigo_usuario = cbo_usuarios.ItemData(cbo_usuarios.ListIndex)
    lstr_senha = txt_senha.Text
    Me.ValidateControls
    If (lint_codigo_usuario = 0) Then
        MsgBox "Selecione um [usuário] na lista.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        cbo_usuarios.SetFocus
        GoTo fim_lfct_valida_campos
    End If
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
    lfct_valida_campos = True
fim_lfct_valida_campos:
    Exit Function
erro_lfct_valida_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "lfct_valida_campos"
    GoTo fim_lfct_valida_campos
End Function

Private Sub cbo_usuarios_Click()
    On Error GoTo erro_cbo_usuarios_Click
    Dim llng_codigo_usuario As Long
    Dim lstr_lembrete_senha As String
    'retorna código do usuário selecionado
    llng_codigo_usuario = cbo_usuarios.ItemData(cbo_usuarios.ListIndex)
    'se houver usuário selecionado
    If (llng_codigo_usuario > 0) Then
        'busca o lembrete de senha
        lstr_lembrete_senha = pfct_carregar_lembrete_usuario(llng_codigo_usuario)
    Else
        'senão, limpa a variável
        lstr_lembrete_senha = "Não há usuário selecionado na lista."
    End If
    'exibe o lembrete de senha caso haja
    If (lstr_lembrete_senha = "") Then
        lbl_lembrete_senha.Tag = "Não há lembrete de senha para este usuário."
    Else
        lbl_lembrete_senha.Tag = lstr_lembrete_senha
    End If
    lbl_lembrete_senha.Refresh
fim_cbo_usuarios_Click:
    Exit Sub
erro_cbo_usuarios_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "cbo_usuarios_Click"
    GoTo fim_cbo_usuarios_Click
End Sub

Private Sub cbo_usuarios_DropDown()
    On Error GoTo erro_cbo_usuarios_DropDown
    psub_campo_got_focus cbo_usuarios
fim_cbo_usuarios_DropDown:
    Exit Sub
erro_cbo_usuarios_DropDown:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "cbo_usuarios_DropDown"
    GoTo fim_cbo_usuarios_DropDown
End Sub

Private Sub cbo_usuarios_GotFocus()
    On Error GoTo erro_cbo_usuarios_gotFocus
    psub_campo_got_focus cbo_usuarios
fim_cbo_usuarios_gotFocus:
    Exit Sub
erro_cbo_usuarios_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "cbo_usuarios_GotFocus"
    GoTo fim_cbo_usuarios_gotFocus
End Sub

Private Sub cbo_usuarios_LostFocus()
    On Error GoTo erro_cbo_usuarios_LostFocus
    psub_campo_lost_focus cbo_usuarios
fim_cbo_usuarios_LostFocus:
    Exit Sub
erro_cbo_usuarios_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "cbo_usuarios_LostFocus"
    GoTo fim_cbo_usuarios_LostFocus
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "cmd_cancelar_click"
    GoTo fim_cmd_cancelar_Click
End Sub

Private Sub cmd_cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_cmd_cancelar_MouseMove
    'atribui textos às variáveis
    mstr_titulo = "Cancelar [F3]"
    mstr_mensagem = "Cancela a autenticação e encerra a aplicação."
    'exibe o tooltip
    psub_exibir_tooltip cmd_cancelar, mobj_tooltip, mstr_titulo, mstr_mensagem
fim_cmd_cancelar_MouseMove:
    Exit Sub
erro_cmd_cancelar_MouseMove:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "cmd_cancelar_MouseMove"
    GoTo fim_cmd_cancelar_MouseMove
End Sub

Private Sub cmd_criar_novo_usuario_Click()
    On Error GoTo erro_cmd_criar_novo_usuario_click
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_criar_novo_usuario.Enabled) Then
        Exit Sub
    End If
    frm_usuario_criar.Show vbModal, frm_usuario_login
fim_cmd_criar_novo_usuario_click:
    Exit Sub
erro_cmd_criar_novo_usuario_click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "cmd_criar_novo_usuario_click"
    GoTo fim_cmd_criar_novo_usuario_click
End Sub

Private Sub cmd_criar_novo_usuario_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_cmd_criar_novo_usuario_MouseMove
    'atribui textos às variáveis
    mstr_titulo = "Criar Usuário [F4]"
    mstr_mensagem = "Abre a tela para criação de novos usuários."
    'exibe o tooltip
    psub_exibir_tooltip cmd_criar_novo_usuario, mobj_tooltip, mstr_titulo, mstr_mensagem
fim_cmd_criar_novo_usuario_MouseMove:
    Exit Sub
erro_cmd_criar_novo_usuario_MouseMove:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "cmd_criar_novo_usuario_MouseMove"
    GoTo fim_cmd_criar_novo_usuario_MouseMove
End Sub

Private Sub cmd_login_Click()
    On Error GoTo erro_cmd_login_click
    Dim llng_codigo_usuario As Long
    Dim lstr_senha As String
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_login.Enabled) Then
        Exit Sub
    End If
    llng_codigo_usuario = cbo_usuarios.ItemData(cbo_usuarios.ListIndex)
    lstr_senha = txt_senha.Text
    'valida os campos
    If (lfct_valida_campos()) Then
        'valida a senha do usuário
        If (pfct_validar_senha(llng_codigo_usuario, pfct_criptografia(lstr_senha))) Then
            'carrega os dados do usuário
            If (pfct_carregar_dados_usuario(llng_codigo_usuario)) Then
                'carrega as configurações do usuário
                If (pfct_carregar_configuracoes_usuario(llng_codigo_usuario)) Then
                    'ajusta o tipo de banco
                     p_banco.tb_tipo_banco = tb_dados
                     'configura o tipo de banco de dados
                     pfct_ajustar_caminho_banco tb_dados
                     'cria as tabelas do usuário
                     If (pfct_criar_tabelas_usuario()) Then
                        'ajusta a variável para true
                        mbln_login = True
                        'informa ao sistema que estamos logando
                        frm_principal.esta_logando = True
                        'descarrega o formulário
                        Unload Me
                     End If
                End If
            End If
        Else
            mbln_login = False
            mint_tentativas = mint_tentativas + 1
            MsgBox "Senha inválida.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            txt_senha.Text = ""
            txt_senha.SetFocus
        End If
    End If
fim_cmd_login_click:
    If (mint_tentativas >= 3) Then
        MsgBox "Tentativas de login esgotadas." & vbCrLf & _
               "A aplicação será encerrada agora.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        End
    End If
    Exit Sub
erro_cmd_login_click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "cmd_login_click"
    GoTo fim_cmd_login_click
End Sub

Private Sub cmd_login_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_cmd_login_MouseMove
    'atribui textos às variáveis
    mstr_titulo = "Login [F2]"
    mstr_mensagem = "Realiza a validação dos campos e em seguida, entra na aplicação."
    'exibe o tooltip
    psub_exibir_tooltip cmd_login, mobj_tooltip, mstr_titulo, mstr_mensagem
fim_cmd_login_MouseMove:
    Exit Sub
erro_cmd_login_MouseMove:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "cmd_login_MouseMove"
    GoTo fim_cmd_login_MouseMove
End Sub

Private Sub Form_Activate()
    On Error GoTo Erro_Form_Activate
    Beep
    psub_preencher_usuarios cbo_usuarios
    If (cbo_usuarios.ListCount > 1) Then
        lbl_usuarios.Enabled = True
        lbl_senha.Enabled = True
        cbo_usuarios.Enabled = True
        txt_senha.Enabled = True
        cmd_login.Enabled = True
    Else
        lbl_usuarios.Enabled = False
        lbl_senha.Enabled = False
        cbo_usuarios.Enabled = False
        txt_senha.Enabled = False
        cmd_login.Enabled = False
    End If
    If (cbo_usuarios.Enabled) Then
        cbo_usuarios.SetFocus
    End If
    txt_senha.Text = ""
Fim_Form_Activate:
    Exit Sub
Erro_Form_Activate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "form_activate"
    GoTo Fim_Form_Activate
End Sub

Private Sub Form_Initialize()
    On Error GoTo Erro_Form_Initialize
    InitCommonControls
Fim_Form_Initialize:
    Exit Sub
Erro_Form_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "form_initialize"
    GoTo Fim_Form_Initialize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/usuarios_login.htm", 0
        Case vbKeyF2
            cmd_login_Click
        Case vbKeyF3
            cmd_cancelar_Click
        Case vbKeyF4
            cmd_criar_novo_usuario_Click
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_Form_MouseMove
    'limpa variáveis
    mstr_titulo = ""
    mstr_mensagem = ""
    'chama o destrutor do tooltip
    psub_destruir_tooltip mobj_tooltip
fim_Form_MouseMove:
    Exit Sub
erro_Form_MouseMove:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "Form_MouseMove"
    GoTo fim_Form_MouseMove
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo erro_Form_Unload
    If (Not mbln_login) Then
        Set mobj_tooltip = Nothing
        End
    End If
    mbln_login = False
fim_Form_Unload:
    Exit Sub
erro_Form_Unload:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "form_unload"
    GoTo fim_Form_Unload
End Sub

Private Sub lbl_lembrete_senha_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_lbl_lembrete_senha_MouseMove
    'atribui textos às variáveis
    mstr_titulo = "Lembrete de senha"
    mstr_mensagem = lbl_lembrete_senha.Tag
    'exibe o tooltip
    psub_exibir_tooltip frm_usuario_login, mobj_tooltip, mstr_titulo, mstr_mensagem
fim_lbl_lembrete_senha_MouseMove:
    Exit Sub
erro_lbl_lembrete_senha_MouseMove:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "lbl_lembrete_senha_MouseMove"
    GoTo fim_lbl_lembrete_senha_MouseMove
End Sub

Private Sub lbl_senha_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_lbl_senha_MouseMove
    'atribui textos às variáveis
    mstr_titulo = "Senha"
    mstr_mensagem = "Senha do usuário para autenticar na aplicação."
    'exibe o tooltip
    psub_exibir_tooltip frm_usuario_login, mobj_tooltip, mstr_titulo, mstr_mensagem
fim_lbl_senha_MouseMove:
    Exit Sub
erro_lbl_senha_MouseMove:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "lbl_senha_MouseMove"
    GoTo fim_lbl_senha_MouseMove
End Sub

Private Sub lbl_usuarios_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_lbl_usuarios_MouseMove
    'atribui textos às variáveis
    mstr_titulo = "Selecione o usuário"
    mstr_mensagem = "Permite selecionar um usuário na lista para autenticação."
    'exibe o tooltip
    psub_exibir_tooltip frm_usuario_login, mobj_tooltip, mstr_titulo, mstr_mensagem
fim_lbl_usuarios_MouseMove:
    Exit Sub
erro_lbl_usuarios_MouseMove:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "lbl_usuarios_MouseMove"
    GoTo fim_lbl_usuarios_MouseMove
End Sub

Private Sub txt_senha_GotFocus()
    On Error GoTo erro_txt_senha_gotFocus
    psub_campo_got_focus txt_senha
fim_txt_senha_gotFocus:
    Exit Sub
erro_txt_senha_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "txt_senha_GotFocus"
    GoTo fim_txt_senha_gotFocus
End Sub

Private Sub txt_senha_LostFocus()
    On Error GoTo erro_txt_senha_LostFocus
    psub_campo_lost_focus txt_senha
fim_txt_senha_LostFocus:
    Exit Sub
erro_txt_senha_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "txt_senha_LostFocus"
    GoTo fim_txt_senha_LostFocus
End Sub

Private Sub txt_senha_Validate(Cancel As Boolean)
    On Error GoTo erro_txt_senha_validate
    Cancel = Not pfct_validar_campo(txt_senha, tc_texto)
fim_txt_senha_validate:
    Exit Sub
erro_txt_senha_validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_usuario_login", "txt_senha_validate"
    GoTo fim_txt_senha_validate
End Sub

