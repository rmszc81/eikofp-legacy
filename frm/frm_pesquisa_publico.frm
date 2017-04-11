VERSION 5.00
Begin VB.Form frm_pesquisa_publico 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Pesquisa de Público"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5775
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
   ScaleHeight     =   7275
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_opiniao 
      Height          =   1215
      Left            =   120
      MaxLength       =   512
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   5640
      Width           =   5535
   End
   Begin VB.TextBox txt_origem 
      Height          =   315
      Left            =   120
      MaxLength       =   60
      TabIndex        =   19
      Top             =   4860
      Width           =   5535
   End
   Begin VB.TextBox txt_cidade 
      Height          =   315
      Left            =   120
      MaxLength       =   35
      TabIndex        =   12
      Top             =   3300
      Width           =   3195
   End
   Begin VB.TextBox txt_estado 
      Height          =   315
      Left            =   3480
      MaxLength       =   30
      TabIndex        =   9
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txt_pais 
      Height          =   315
      Left            =   120
      MaxLength       =   50
      TabIndex        =   8
      Top             =   2520
      Width           =   3195
   End
   Begin VB.CommandButton cmd_participar 
      Caption         =   "&Participar (F2)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton cmd_fechar 
      Caption         =   "&Fechar (F8)"
      Height          =   375
      Left            =   1500
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txt_nome 
      Height          =   315
      Left            =   120
      MaxLength       =   60
      TabIndex        =   3
      Top             =   960
      Width           =   5535
   End
   Begin VB.TextBox txt_email 
      Height          =   315
      Left            =   120
      MaxLength       =   60
      TabIndex        =   5
      Top             =   1740
      Width           =   5535
   End
   Begin VB.TextBox txt_data_nascimento 
      Height          =   315
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   13
      Top             =   3300
      Width           =   2175
   End
   Begin VB.ComboBox cbo_sexo 
      Height          =   315
      ItemData        =   "frm_pesquisa_publico.frx":0000
      Left            =   3480
      List            =   "frm_pesquisa_publico.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox txt_profissao 
      Height          =   315
      Left            =   120
      MaxLength       =   60
      TabIndex        =   16
      Top             =   4080
      Width           =   3195
   End
   Begin VB.Label lbl_opiniao 
      AutoSize        =   -1  'True
      Caption         =   "Tem algo mais a nos dizer?"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   5340
      Width           =   1905
   End
   Begin VB.Label lbl_origem 
      AutoSize        =   -1  'True
      Caption         =   "Como ficou sabendo de nós?"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lbl_cidade 
      AutoSize        =   -1  'True
      Caption         =   "Cidade:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   555
   End
   Begin VB.Label lbl_estado 
      AutoSize        =   -1  'True
      Caption         =   "Estado:"
      Height          =   195
      Left            =   3480
      TabIndex        =   7
      Top             =   2220
      Width           =   555
   End
   Begin VB.Label lbl_pais 
      AutoSize        =   -1  'True
      Caption         =   "País:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2220
      Width           =   345
   End
   Begin VB.Label lbl_nome 
      AutoSize        =   -1  'True
      Caption         =   "Nome ( * ) :"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   855
   End
   Begin VB.Label lbl_email 
      AutoSize        =   -1  'True
      Caption         =   "E-mail ( * ) :"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   870
   End
   Begin VB.Label lbl_nascimento 
      AutoSize        =   -1  'True
      Caption         =   "Data de nascimento:"
      Height          =   195
      Left            =   3480
      TabIndex        =   11
      Top             =   3000
      Width           =   1485
   End
   Begin VB.Label lbl_sexo 
      AutoSize        =   -1  'True
      Caption         =   "Sexo:"
      Height          =   195
      Left            =   3480
      TabIndex        =   15
      Top             =   3780
      Width           =   420
   End
   Begin VB.Label lbl_profissao 
      AutoSize        =   -1  'True
      Caption         =   "Profissão:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3780
      Width           =   720
   End
   Begin VB.Label lbl_campos_obrigatorios 
      AutoSize        =   -1  'True
      Caption         =   "* Campos marcados com asterisco ( * ) são de preenchimento obrigatório."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   165
      Left            =   120
      TabIndex        =   22
      Top             =   7020
      Width           =   5550
   End
End
Attribute VB_Name = "frm_pesquisa_publico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enm_sexo
    op_selecione = 0
    op_masculino = 1
    op_feminino = 2
End Enum

Private mbln_campos_alterados As Boolean

Private Sub lsub_desabilitar_campos()
    On Error GoTo Erro_lsub_desabilitar_campos
    If (p_registro.dt_data_liberacao <> CDate(0)) Then '30/12/1899 00:00:00
        txt_nome.Enabled = False
        txt_email.Enabled = False
    End If
Fim_lsub_desabilitar_campos:
    Exit Sub
Erro_lsub_desabilitar_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "lsub_desabilitar_campos"
    GoTo Fim_lsub_desabilitar_campos
End Sub

Private Function lfct_validar_campos() As Boolean
    On Error GoTo erro_lfct_validar_campos
    
    If (Trim$(txt_nome.Text) = Empty) Then
        MsgBox "Campo [nome] é obrigatório.", vbOKOnly + vbExclamation, pcst_nome_aplicacao
        txt_nome.Text = Empty
        txt_nome.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    
    If (Trim$(txt_email.Text) = Empty) Then
        MsgBox "Campo [email] é obrigatório.", vbOKOnly + vbExclamation, pcst_nome_aplicacao
        txt_email.Text = Empty
        txt_email.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    
    If (Not pfct_verificar_email(txt_email.Text)) Then
        MsgBox "Digite um [email] válido.", vbOKOnly + vbExclamation, pcst_nome_aplicacao
        txt_email.Text = Empty
        txt_email.SetFocus
        GoTo fim_lfct_validar_campos
    End If
    
    If (Trim$(txt_data_nascimento.Text) <> Empty) Then
        If (Not IsDate(txt_data_nascimento.Text)) Then
            MsgBox "Digite uma [data de nascimento] válida.", vbOKOnly + vbExclamation, pcst_nome_aplicacao
            txt_data_nascimento.Text = Empty
            txt_data_nascimento.SetFocus
            GoTo fim_lfct_validar_campos
        End If
    End If
    
    lfct_validar_campos = True
fim_lfct_validar_campos:
    Exit Function
erro_lfct_validar_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "lfct_validar_campos"
    GoTo fim_lfct_validar_campos
End Function

Private Sub lsub_preencher_campos()
    On Error GoTo erro_lsub_preencher_campos
    If (p_usuario.bln_participou_pesquisa) Then
        With p_registro
            'textbox
            txt_nome.Text = .str_nome
            txt_email.Text = .str_email
            txt_pais.Text = .str_pais
            txt_estado.Text = .str_estado
            txt_cidade.Text = .str_cidade
            If (Format$(p_registro.dt_data_nascimento, "dd/mm/yyyy hh:mm:ss") <> "30/12/1899 00:00:00") Then
                txt_data_nascimento.Text = Format$(.dt_data_nascimento, "dd/mm/yyyy")
            End If
            txt_profissao.Text = .str_profissao
            txt_origem.Text = .str_origem
            txt_opiniao.Text = .str_opiniao
            'combo
            Select Case .chr_sexo
                Case Empty
                    cbo_sexo.ListIndex = op_selecione
                Case "M"
                    cbo_sexo.ListIndex = op_masculino
                Case "F"
                    cbo_sexo.ListIndex = op_feminino
            End Select
        End With
    End If
    mbln_campos_alterados = False
fim_lsub_preencher_campos:
    Exit Sub
erro_lsub_preencher_campos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "lsub_preencher_campos"
    GoTo fim_lsub_preencher_campos
End Sub

Private Sub cbo_sexo_Change()
    On Error GoTo Erro_cbo_sexo_Change
    mbln_campos_alterados = True
Fim_cbo_sexo_Change:
    Exit Sub
Erro_cbo_sexo_Change:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "cbo_sexo_Change"
    GoTo Fim_cbo_sexo_Change
End Sub

Private Sub cbo_sexo_GotFocus()
    On Error GoTo Erro_cbo_sexo_GotFocus
    psub_campo_got_focus cbo_sexo
Fim_cbo_sexo_GotFocus:
    Exit Sub
Erro_cbo_sexo_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "cbo_sexo_GotFocus"
    GoTo Fim_cbo_sexo_GotFocus
End Sub

Private Sub cbo_sexo_LostFocus()
    On Error GoTo Erro_cbo_sexo_LostFocus
    psub_campo_lost_focus cbo_sexo
Fim_cbo_sexo_LostFocus:
    Exit Sub
Erro_cbo_sexo_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "cbo_sexo_LostFocus"
    GoTo Fim_cbo_sexo_LostFocus
End Sub

Private Sub cbo_sexo_Validate(Cancel As Boolean)
    On Error GoTo Erro_cbo_sexo_Validate
    psub_tratar_campo cbo_sexo
    Cancel = Not pfct_validar_campo(cbo_sexo, tc_texto)
Fim_cbo_sexo_Validate:
    Exit Sub
Erro_cbo_sexo_Validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "cbo_sexo_Validate"
    GoTo Fim_cbo_sexo_Validate
End Sub

Private Sub cmd_fechar_Click()
    On Error GoTo erro_cmd_fechar_Click
    Unload Me
fim_cmd_fechar_Click:
    Exit Sub
erro_cmd_fechar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "cmd_fechar_Click"
    GoTo fim_cmd_fechar_Click
End Sub

Private Sub cmd_participar_Click()
    On Error GoTo Erro_cmd_participar_Click
    Dim lbln_retorno As Boolean
    Dim ldt_data_servidor As Date
    Me.ValidateControls
    If (lfct_validar_campos) Then
        'retorna a data do servidor mysql
        ldt_data_servidor = pfct_retorna_data_hora_mysql()
        'se não houve retorno
        If (ldt_data_servidor = CDate(0)) Then
            ldt_data_servidor = Now
        End If
        'preenche os dados do objeto
        With p_registro
            .str_nome = txt_nome.Text
            .str_email = txt_email.Text
            .str_pais = txt_pais.Text
            .str_estado = txt_estado.Text
            .str_cidade = txt_cidade.Text
            If (txt_data_nascimento.Text <> Empty) Then
                .dt_data_nascimento = CDate(txt_data_nascimento.Text)
            Else
                .dt_data_nascimento = CDate(0)
            End If
            .str_profissao = txt_profissao.Text
            .chr_sexo = IIf(cbo_sexo.ListIndex <> 0, IIf(cbo_sexo.ListIndex = 1, "M", "F"), "")
            .str_origem = txt_origem.Text
            .str_opiniao = txt_opiniao.Text
            .bln_newsletter = False
            .str_id_cpu = p_pc.str_id_cpu
            .str_id_hd = p_pc.str_id_hd
            .dt_data_registro = IIf(.dt_data_registro <> CDate(0), .dt_data_registro, ldt_data_servidor)
            .dt_data_liberacao = IIf(.dt_data_liberacao <> CDate(0), .dt_data_liberacao, CDate(0))
            .bln_banido = IIf(.bln_banido, True, False)
        End With
        'preenche os dados do objeto
        With p_usuario
            'ajusta o flag para true
            .bln_participou_pesquisa = True
        End With
        'ajusta o tipo de banco
        p_banco.tb_tipo_banco = tb_config
        'salva as configurações do usuário
        lbln_retorno = pfct_salvar_configuracoes_usuario(p_usuario.lng_codigo)
        'se gravou com sucesso
        If (lbln_retorno) Then
            'se algum dos campos foi alterado, sinaliza para atualização do registro online
            frm_principal.nao_verificar_registro = Not mbln_campos_alterados
            'exibe mensagem ao usuário
            MsgBox "Obrigado por participar da nossa pesquisa.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            'descarrega o form
            Unload Me
        End If
    End If
Fim_cmd_participar_Click:
    p_banco.tb_tipo_banco = tb_dados
    Exit Sub
Erro_cmd_participar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "cmd_participar_Click"
    GoTo Fim_cmd_participar_Click
End Sub

Private Sub Form_Activate()
    On Error GoTo Erro_Form_Activate
    If ((txt_nome.Visible) And (txt_nome.Enabled)) Then
        txt_nome.SetFocus
    End If
Fim_Form_Activate:
    Exit Sub
Erro_Form_Activate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "Form_Activate"
    GoTo Fim_Form_Activate
End Sub

Private Sub Form_Initialize()
    On Error GoTo Erro_Form_Initialize
    InitCommonControls
Fim_Form_Initialize:
    Exit Sub
Erro_Form_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "Form_Initialize"
    GoTo Fim_Form_Initialize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Erro_Form_KeyPress
    psub_campo_keypress KeyAscii
Fim_Form_KeyPress:
    Exit Sub
Erro_Form_KeyPress:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "Form_KeyPress"
    GoTo Fim_Form_KeyPress
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/ajuda_pesquisa_publico.htm", 0
        Case vbKeyF2
            cmd_participar_Click
        Case vbKeyF8
            cmd_fechar_Click
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
    'preenche o combo sexo
    With cbo_sexo
        .Clear
        .AddItem "- selecione -", op_selecione
        .AddItem "- MASCULINO", op_masculino
        .AddItem "- FEMININO", op_feminino
        .ListIndex = 0
    End With
    lsub_desabilitar_campos
    lsub_preencher_campos
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "Form_Load"
    GoTo fim_Form_Load
End Sub

Private Sub txt_cidade_Change()
    On Error GoTo Erro_txt_cidade_Change
    mbln_campos_alterados = True
Fim_txt_cidade_Change:
    Exit Sub
Erro_txt_cidade_Change:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_cidade_Change"
    GoTo Fim_txt_cidade_Change
End Sub

Private Sub txt_cidade_GotFocus()
    On Error GoTo Erro_txt_cidade_GotFocus
    psub_campo_got_focus txt_cidade
Fim_txt_cidade_GotFocus:
    Exit Sub
Erro_txt_cidade_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_cidade_GotFocus"
    GoTo Fim_txt_cidade_GotFocus
End Sub

Private Sub txt_cidade_LostFocus()
    On Error GoTo Erro_txt_cidade_LostFocus
    psub_campo_lost_focus txt_cidade
Fim_txt_cidade_LostFocus:
    Exit Sub
Erro_txt_cidade_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_cidade_LostFocus"
    GoTo Fim_txt_cidade_LostFocus
End Sub

Private Sub txt_cidade_Validate(Cancel As Boolean)
    On Error GoTo Erro_txt_cidade_Validate
    psub_tratar_campo txt_cidade
    Cancel = Not pfct_validar_campo(txt_cidade, tc_texto)
Fim_txt_cidade_Validate:
    Exit Sub
Erro_txt_cidade_Validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_cidade_Validate"
    GoTo Fim_txt_cidade_Validate
End Sub

Private Sub txt_data_nascimento_Change()
    On Error GoTo Erro_txt_data_nascimento_Change
    mbln_campos_alterados = True
Fim_txt_data_nascimento_Change:
    Exit Sub
Erro_txt_data_nascimento_Change:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_data_nascimento_Change"
    GoTo Fim_txt_data_nascimento_Change
End Sub

Private Sub txt_data_nascimento_GotFocus()
    On Error GoTo Erro_txt_data_nascimento_GotFocus
    psub_campo_got_focus txt_data_nascimento
Fim_txt_data_nascimento_GotFocus:
    Exit Sub
Erro_txt_data_nascimento_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_data_nascimento_GotFocus"
    GoTo Fim_txt_data_nascimento_GotFocus
End Sub

Private Sub txt_data_nascimento_LostFocus()
    On Error GoTo Erro_txt_data_nascimento_LostFocus
    psub_campo_lost_focus txt_data_nascimento
Fim_txt_data_nascimento_LostFocus:
    Exit Sub
Erro_txt_data_nascimento_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_data_nascimento_LostFocus"
    GoTo Fim_txt_data_nascimento_LostFocus
End Sub

Private Sub txt_data_nascimento_Validate(Cancel As Boolean)
    On Error GoTo Erro_txt_data_nascimento_Validate
    psub_tratar_campo txt_data_nascimento
    If (Trim$(txt_data_nascimento.Text) <> Empty) Then
        If (Not IsDate(txt_data_nascimento.Text)) Then
            txt_data_nascimento.Text = Empty
        End If
    End If
    Cancel = Not pfct_validar_campo(txt_data_nascimento, tc_texto)
Fim_txt_data_nascimento_Validate:
    Exit Sub
Erro_txt_data_nascimento_Validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_data_nascimento_Validate"
    GoTo Fim_txt_data_nascimento_Validate
End Sub

Private Sub txt_email_Change()
    On Error GoTo Erro_txt_email_Change
    mbln_campos_alterados = True
Fim_txt_email_Change:
    Exit Sub
Erro_txt_email_Change:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_email_Change"
    GoTo Fim_txt_email_Change
End Sub

Private Sub txt_email_GotFocus()
    On Error GoTo Erro_txt_email_GotFocus
    psub_campo_got_focus txt_email
Fim_txt_email_GotFocus:
    Exit Sub
Erro_txt_email_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_email_GotFocus"
    GoTo Fim_txt_email_GotFocus
End Sub

Private Sub txt_email_LostFocus()
    On Error GoTo Erro_txt_email_LostFocus
    psub_campo_lost_focus txt_email
Fim_txt_email_LostFocus:
    Exit Sub
Erro_txt_email_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_email_LostFocus"
    GoTo Fim_txt_email_LostFocus
End Sub

Private Sub txt_email_Validate(Cancel As Boolean)
    On Error GoTo Erro_txt_email_Validate
    psub_tratar_campo txt_email
    Cancel = Not pfct_validar_campo(txt_email, tc_texto)
Fim_txt_email_Validate:
    Exit Sub
Erro_txt_email_Validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_email_Validate"
    GoTo Fim_txt_email_Validate
End Sub

Private Sub txt_estado_Change()
    On Error GoTo Erro_txt_estado_Change
    mbln_campos_alterados = True
Fim_txt_estado_Change:
    Exit Sub
Erro_txt_estado_Change:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_estado_Change"
    GoTo Fim_txt_estado_Change
End Sub

Private Sub txt_estado_GotFocus()
    On Error GoTo Erro_txt_estado_GotFocus
    psub_campo_got_focus txt_estado
Fim_txt_estado_GotFocus:
    Exit Sub
Erro_txt_estado_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_estado_GotFocus"
    GoTo Fim_txt_estado_GotFocus
End Sub

Private Sub txt_estado_LostFocus()
    On Error GoTo Erro_txt_estado_LostFocus
    psub_campo_lost_focus txt_estado
Fim_txt_estado_LostFocus:
    Exit Sub
Erro_txt_estado_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_estado_LostFocus"
    GoTo Fim_txt_estado_LostFocus
End Sub

Private Sub txt_estado_Validate(Cancel As Boolean)
    On Error GoTo Erro_txt_estado_Validate
    psub_tratar_campo txt_estado
    Cancel = Not pfct_validar_campo(txt_estado, tc_texto)
Fim_txt_estado_Validate:
    Exit Sub
Erro_txt_estado_Validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_estado_Validate"
    GoTo Fim_txt_estado_Validate
End Sub

Private Sub txt_nome_Change()
    On Error GoTo Erro_txt_nome_Change
    mbln_campos_alterados = True
Fim_txt_nome_Change:
    Exit Sub
Erro_txt_nome_Change:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_nome_Change"
    GoTo Fim_txt_nome_Change
End Sub

Private Sub txt_nome_GotFocus()
    On Error GoTo Erro_txt_nome_GotFocus
    psub_campo_got_focus txt_nome
Fim_txt_nome_GotFocus:
    Exit Sub
Erro_txt_nome_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_nome_GotFocus"
    GoTo Fim_txt_nome_GotFocus
End Sub

Private Sub txt_nome_LostFocus()
    On Error GoTo Erro_txt_nome_LostFocus
    psub_campo_lost_focus txt_nome
Fim_txt_nome_LostFocus:
    Exit Sub
Erro_txt_nome_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_nome_LostFocus"
    GoTo Fim_txt_nome_LostFocus
End Sub

Private Sub txt_nome_Validate(Cancel As Boolean)
    On Error GoTo Erro_txt_nome_Validate
    psub_tratar_campo txt_nome
    Cancel = Not pfct_validar_campo(txt_nome, tc_texto)
Fim_txt_nome_Validate:
    Exit Sub
Erro_txt_nome_Validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_nome_Validate"
    GoTo Fim_txt_nome_Validate
End Sub

Private Sub txt_opiniao_Change()
    On Error GoTo Erro_txt_opiniao_Change
    mbln_campos_alterados = True
Fim_txt_opiniao_Change:
    Exit Sub
Erro_txt_opiniao_Change:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_opiniao_Change"
    GoTo Fim_txt_opiniao_Change
End Sub

Private Sub txt_opiniao_GotFocus()
    On Error GoTo Erro_txt_opiniao_GotFocus
    psub_campo_got_focus txt_opiniao
Fim_txt_opiniao_GotFocus:
    Exit Sub
Erro_txt_opiniao_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_opiniao_GotFocus"
    GoTo Fim_txt_opiniao_GotFocus
End Sub

Private Sub txt_opiniao_LostFocus()
    On Error GoTo Erro_txt_opiniao_LostFocus
    psub_campo_lost_focus txt_opiniao
Fim_txt_opiniao_LostFocus:
    Exit Sub
Erro_txt_opiniao_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_opiniao_LostFocus"
    GoTo Fim_txt_opiniao_LostFocus
End Sub

Private Sub txt_opiniao_Validate(Cancel As Boolean)
    On Error GoTo Erro_txt_opiniao_Validate
    psub_tratar_campo txt_opiniao
    Cancel = Not pfct_validar_campo(txt_opiniao, tc_texto)
Fim_txt_opiniao_Validate:
    Exit Sub
Erro_txt_opiniao_Validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_opiniao_Validate"
    GoTo Fim_txt_opiniao_Validate
End Sub

Private Sub txt_origem_Change()
    On Error GoTo Erro_txt_origem_Change
    mbln_campos_alterados = True
Fim_txt_origem_Change:
    Exit Sub
Erro_txt_origem_Change:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_origem_Change"
    GoTo Fim_txt_origem_Change
End Sub

Private Sub txt_origem_GotFocus()
    On Error GoTo Erro_txt_origem_GotFocus
    psub_campo_got_focus txt_origem
Fim_txt_origem_GotFocus:
    Exit Sub
Erro_txt_origem_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_origem_GotFocus"
    GoTo Fim_txt_origem_GotFocus
End Sub

Private Sub txt_origem_LostFocus()
    On Error GoTo Erro_txt_origem_LostFocus
    psub_campo_lost_focus txt_origem
Fim_txt_origem_LostFocus:
    Exit Sub
Erro_txt_origem_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_origem_LostFocus"
    GoTo Fim_txt_origem_LostFocus
End Sub

Private Sub txt_origem_Validate(Cancel As Boolean)
    On Error GoTo Erro_txt_origem_Validate
    psub_tratar_campo txt_origem
    Cancel = Not pfct_validar_campo(txt_origem, tc_texto)
Fim_txt_origem_Validate:
    Exit Sub
Erro_txt_origem_Validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_origem_Validate"
    GoTo Fim_txt_origem_Validate
End Sub

Private Sub txt_pais_Change()
    On Error GoTo Erro_txt_pais_Change
    mbln_campos_alterados = True
Fim_txt_pais_Change:
    Exit Sub
Erro_txt_pais_Change:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_pais_Change"
    GoTo Fim_txt_pais_Change
End Sub

Private Sub txt_pais_GotFocus()
    On Error GoTo Erro_txt_pais_GotFocus
    psub_campo_got_focus txt_pais
Fim_txt_pais_GotFocus:
    Exit Sub
Erro_txt_pais_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_pais_GotFocus"
    GoTo Fim_txt_pais_GotFocus
End Sub

Private Sub txt_pais_LostFocus()
    On Error GoTo Erro_txt_pais_LostFocus
    psub_campo_lost_focus txt_pais
Fim_txt_pais_LostFocus:
    Exit Sub
Erro_txt_pais_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_pais_LostFocus"
    GoTo Fim_txt_pais_LostFocus
End Sub

Private Sub txt_pais_Validate(Cancel As Boolean)
    On Error GoTo Erro_txt_pais_Validate
    psub_tratar_campo txt_pais
    Cancel = Not pfct_validar_campo(txt_pais, tc_texto)
Fim_txt_pais_Validate:
    Exit Sub
Erro_txt_pais_Validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_pais_Validate"
    GoTo Fim_txt_pais_Validate
End Sub

Private Sub txt_profissao_Change()
    On Error GoTo Erro_txt_profissao_Change
    mbln_campos_alterados = True
Fim_txt_profissao_Change:
    Exit Sub
Erro_txt_profissao_Change:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_profissao_Change"
    GoTo Fim_txt_profissao_Change
End Sub

Private Sub txt_profissao_GotFocus()
    On Error GoTo Erro_txt_profissao_GotFocus
    psub_campo_got_focus txt_profissao
Fim_txt_profissao_GotFocus:
    Exit Sub
Erro_txt_profissao_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_profissao_GotFocus"
    GoTo Fim_txt_profissao_GotFocus
End Sub

Private Sub txt_profissao_LostFocus()
    On Error GoTo Erro_txt_profissao_LostFocus
    psub_campo_lost_focus txt_profissao
Fim_txt_profissao_LostFocus:
    Exit Sub
Erro_txt_profissao_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_profissao_LostFocus"
    GoTo Fim_txt_profissao_LostFocus
End Sub

Private Sub txt_profissao_Validate(Cancel As Boolean)
    On Error GoTo Erro_txt_profissao_Validate
    psub_tratar_campo txt_profissao
    Cancel = Not pfct_validar_campo(txt_profissao, tc_texto)
Fim_txt_profissao_Validate:
    Exit Sub
Erro_txt_profissao_Validate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_pesquisa_publico", "txt_profissao_Validate"
    GoTo Fim_txt_profissao_Validate
End Sub

