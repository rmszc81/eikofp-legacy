VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0F0877EF-2A93-4AE6-8BA8-4129832C32C3}#230.0#0"; "SMARTMENUXP.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frm_principal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "#"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   9825
   Icon            =   "frm_principal.frx":0000
   LockControls    =   -1  'True
   NegotiateToolbars=   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd_arquivos 
      Left            =   90
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmr_baixas_automaticas 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   90
      Top             =   615
   End
   Begin VBSmartXPMenu.SmartMenuXP smxp_principal 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      SmoothMenuBar   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shadow          =   0   'False
   End
   Begin VB.Timer tmr_timer 
      Interval        =   1000
      Left            =   90
      Top             =   90
   End
   Begin MSComctlLib.StatusBar stb_status 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   7905
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4180
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4180
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4180
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4180
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frm_principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enm_status
    pnl_data_hora = 1
    pnl_usuario = 2
    pnl_ajuda = 3
    pnl_versao = 4
End Enum

Private mbln_exibir_introducao As Boolean
Private mbln_nao_verificar_registro As Boolean
Private mbln_ja_respondeu_atualizacao As Boolean
Private mbln_esta_logando As Boolean

Public Property Get esta_logando() As Boolean
    On Error GoTo Erro_esta_logando
    esta_logando = mbln_esta_logando
Fim_esta_logando:
    Exit Property
Erro_esta_logando:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "esta_logando"
    GoTo Fim_esta_logando
End Property

Public Property Let esta_logando(ByVal pbln_valor As Boolean)
    On Error GoTo Erro_esta_logando
    mbln_esta_logando = pbln_valor
Fim_esta_logando:
    Exit Property
Erro_esta_logando:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "esta_logando"
    GoTo Fim_esta_logando
End Property

Public Property Get nao_verificar_registro() As Boolean
    On Error GoTo Erro_nao_verificar_registro
    nao_verificar_registro = mbln_nao_verificar_registro
Fim_nao_verificar_registro:
    Exit Property
Erro_nao_verificar_registro:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "nao_verificar_registro"
    GoTo Fim_nao_verificar_registro
End Property

Public Property Let nao_verificar_registro(ByVal pbln_valor As Boolean)
    On Error GoTo Erro_nao_verificar_registro
    mbln_nao_verificar_registro = pbln_valor
Fim_nao_verificar_registro:
    Exit Property
Erro_nao_verificar_registro:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "nao_verificar_registro"
    GoTo Fim_nao_verificar_registro
End Property

Private Sub lsub_logoff()
    On Error GoTo erro_lsub_logoff
    Dim lint_resposta As Integer
    lint_resposta = MsgBox("Deseja fazer o logoff do sistema?", vbYesNo + vbQuestion + vbDefaultButton2, pcst_nome_aplicacao)
    If (lint_resposta = vbYes) Then
        'desabilita o timer de baixas automáticas
        tmr_baixas_automaticas.Enabled = False
        'config
        p_banco.tb_tipo_banco = tb_config
        'atualiza as configurações do usuário
        If (pfct_salvar_configuracoes_usuario(p_usuario.lng_codigo)) Then
            'ajusta a data do último acesso do usuário
            p_usuario.dt_ultimo_acesso = Now
            'atualiza o usuário
            psub_atualizar_usuario p_usuario.lng_codigo, False
            'faz manutenção da base de dados
            psub_limpar_banco
        End If
        'ajusta variáveis para false
        mbln_exibir_introducao = False
        mbln_nao_verificar_registro = False
        'fechar todos os forms
        psub_fechar_forms
        'ajusta os dados do usuário
        With p_usuario
            .lng_codigo = 0
            .str_login = ""
            .str_senha = ""
            .dt_criado_em = "00:00:00"
            .dt_ultimo_acesso = "00:00:00"
            .id_intervalo_data = id_selecione
            .sm_simbolo_moeda = sm_selecione
            .bln_carregar_agenda_financeira_login = True
            .bln_lancamentos_retroativos = False
            .bln_alteracoes_detalhes = False
            .bln_data_vencimento_baixa_imediata = False
            .bln_participou_pesquisa = False
        End With
        'ajusta os dados do backup
        With p_backup
            .bln_ativar = False
            .pb_periodo_backup = pb_selecione
            .str_caminho = ""
            .dt_ultimo_backup = "00:00:00"
            .dt_proximo_backup = "00:00:00"
        End With
        'ajusta os dados do registro
        With p_registro
            .str_nome = Empty
            .str_email = Empty
            .str_pais = Empty
            .str_estado = Empty
            .str_cidade = Empty
            .dt_data_nascimento = CDate(0)  'retorna [30/12/1899 00:00:00]
            .str_profissao = Empty
            .chr_sexo = Empty
            .str_origem = Empty
            .str_opiniao = Empty
            .bln_newsletter = False
            .str_id_cpu = Empty
            .str_id_hd = Empty
            .dt_data_registro = CDate(0)    'retorna [30/12/1899 00:00:00]
            .dt_data_liberacao = CDate(0)   'retorna [30/12/1899 00:00:00]
            .bln_banido = False
        End With
        'ajusta o tipo de banco de dados
        p_banco.tb_tipo_banco = tb_config
        'configura o banco de dados
        pfct_ajustar_caminho_banco tb_config
        'dispara o evento do form
        MDIForm_Activate
    End If
fim_lsub_logoff:
    Exit Sub
erro_lsub_logoff:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "lsub_logoff"
    GoTo fim_lsub_logoff
End Sub

Private Sub lsub_excluir_usuario()
    On Error GoTo erro_lsub_excluir_usuario
    frm_usuario_excluir.Show vbModal, frm_principal
    If (p_usuario.lng_codigo = 0) Then
        'desabilita o timer de baixas automáticas
        tmr_baixas_automaticas.Enabled = False
        'ajusta variáveis para false
        mbln_exibir_introducao = False
        mbln_nao_verificar_registro = False
        'fechar todos os forms
        psub_fechar_forms
        'dispara o evento
        MDIForm_Activate
    End If
fim_lsub_excluir_usuario:
    Exit Sub
erro_lsub_excluir_usuario:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "lsub_excluir_usuario"
    GoTo fim_lsub_excluir_usuario
End Sub

Private Sub lsub_criar_menu_principal()
    On Error GoTo erro_lsub_criar_menu_principal
    With smxp_principal.MenuItems
        'menu usuários
        .Add 0, "k_usuarios", smiNone, "&Usuários"
        .Add "k_usuarios", "k_usuarios_logoff", smiNone, "&Logoff..."
        .Add "k_usuarios", "k_usuarios_separador_01", smiSeparator
        .Add "k_usuarios", "k_usuarios_alterar_senha", smiNone, "&Alterar senha"
        .Add "k_usuarios", "k_usuarios_excluir_usuario", smiNone, "Excluir meu usuário"
        'menu cadastros
        .Add 0, "k_cadastros", smiNone, "&Cadastros"
        .Add "k_cadastros", "k_cadastros_contas", smiNone, "&Contas"
        .Add "k_cadastros", "k_cadastros_receitas", smiNone, "&Receitas"
        .Add "k_cadastros", "k_cadastros_despesas", smiNone, "&Despesas"
        .Add "k_cadastros", "k_cadastros_formas_pagamento", smiNone, "&Formas de Pagamento"
        'menu financeiro
        .Add 0, "k_financeiro", smiNone, "&Financeiro"
        .Add "k_financeiro", "k_financeiro_agenda_financeira", smiNone, "&Agenda financeira"
        .Add "k_financeiro", "k_financeiro_contas_pagar", smiNone, "Contas a &Pagar"
        .Add "k_financeiro", "k_financeiro_contas_receber", smiNone, "Contas a &Receber"
        'menu movimentação
        .Add 0, "k_movimentacao", smiNone, "&Movimentação"
        .Add "k_movimentacao", "k_movimentacao_geral", smiNone, "Movimentação &Geral"
        .Add "k_movimentacao", "k_movimentacao_receitas_despesas", smiNone, "Receitas x Despesas"
        .Add "k_movimentacao", "k_movimentacao_por_receitas_despesas", smiNone, "p&or Receitas/Despesas"
        'menu simulação
        .Add 0, "k_simulacao", smiNone, "&Simulação"
        .Add "k_simulacao", "k_simulacao_receitas_despesas", smiNone, "Receitas x Despesas"
        'menu gráficos
        .Add 0, "k_graficos", smiNone, "&Gráficos"
        .Add "k_graficos", "k_graficos_geral_conta", smiNone, "Geral por &Conta"
        'menu configurações
        .Add 0, "k_configuracoes", smiNone, "&Configurações"
        'menu backup
        .Add 0, "k_backup", smiNone, "&Backup"
        .Add "k_backup", "k_backup_realizar", smiNone, "Realizar &backup agora..."
        .Add "k_backup", "k_backup_restaurar", smiNone, "&Restaurar backup anterior..."
        'menu ajuda
        .Add 0, "k_ajuda", smiNone, "&Ajuda"
        .Add "k_ajuda", "k_ajuda_introducao", smiNone, "&Introdução"
        .Add "k_ajuda", "k_ajuda_tutorial", smiNone, "&Tutorial"
        .Add "k_ajuda", "k_ajuda_suporte", smiNone, "S&uporte técnico"
        .Add "k_ajuda", "k_ajuda_pesquisa", smiNone, "&Pesquisa"
        .Add "k_ajuda", "k_ajuda_change_log", smiNone, "&Change Log"
        .Add "k_ajuda", "k_ajuda_sobre", smiNone, "&Sobre..."
        'menu sair
        .Add 0, "k_sair", smiNone, "&Sair"
    End With
fim_lsub_criar_menu_principal:
    Exit Sub
erro_lsub_criar_menu_principal:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "lsub_criar_menu_principal"
    GoTo fim_lsub_criar_menu_principal
End Sub

Private Sub smxp_principal_Click(ByVal ID As Long)
    On Error GoTo erro_smxp_principal_Click
    Dim lint_resposta As Integer
    With smxp_principal.MenuItems
        Select Case .Key(ID)
        Case "k_usuarios_logoff"
            lsub_logoff
        Case "k_usuarios_alterar_senha"
            frm_usuario_alterar_senha.Show vbModal, frm_principal
        Case "k_usuarios_excluir_usuario"
            lsub_excluir_usuario
        Case "k_cadastros_contas"
            frm_cadastro_contas.Show
            frm_cadastro_contas.ZOrder 0
        Case "k_cadastros_receitas"
            frm_cadastro_receitas.Show
            frm_cadastro_receitas.ZOrder 0
        Case "k_cadastros_despesas"
            frm_cadastro_despesas.Show
            frm_cadastro_despesas.ZOrder 0
        Case "k_cadastros_formas_pagamento"
            frm_cadastro_formas_pagamento.Show
            frm_cadastro_formas_pagamento.ZOrder 0
        Case "k_financeiro_agenda_financeira"
            frm_financeiro_agenda.Left = 100
            frm_financeiro_agenda.Top = 100
            frm_financeiro_agenda.Show
            frm_financeiro_agenda.ZOrder 0
        Case "k_financeiro_contas_pagar"
            frm_cadastro_contas_pagar.Show
            frm_cadastro_contas_pagar.ZOrder 0
        Case "k_financeiro_contas_receber"
            frm_cadastro_contas_receber.Show
            frm_cadastro_contas_receber.ZOrder 0
        Case "k_movimentacao_geral"
            frm_movimentacao_geral.Show
            frm_movimentacao_geral.ZOrder 0
        Case "k_movimentacao_receitas_despesas"
            frm_movimentacao_receitas_despesas.Show
            frm_movimentacao_receitas_despesas.ZOrder 0
        Case "k_movimentacao_por_receitas_despesas"
            frm_movimentacao_por_receitas_despesas.Show
            frm_movimentacao_por_receitas_despesas.ZOrder 0
        Case "k_simulacao_receitas_despesas"
            frm_simulacao_receitas_despesas.Show
            frm_simulacao_receitas_despesas.ZOrder 0
        Case "k_graficos_geral_conta"
            frm_graficos_geral_conta.Show
            frm_graficos_geral_conta.ZOrder 0
        Case "k_configuracoes"
            'desabilita o timer de baixas automáticas
            tmr_baixas_automaticas.Enabled = False
            'exibe o form de configurações
            frm_configuracoes.Show vbModal, Me
            'habilita o timer de baixas automáticas
            tmr_baixas_automaticas.Enabled = True
        Case "k_backup_realizar"
            lint_resposta = MsgBox("Atenção!" & vbCrLf & "Essa operação pode levar até alguns minutos." & vbCrLf & "Continuar?", vbYesNo + vbQuestion + vbDefaultButton2, pcst_nome_aplicacao)
            If (lint_resposta = vbYes) Then
                'desabilita o timer de baixas automáticas
                tmr_baixas_automaticas.Enabled = False
                'ajusta propriedade do form de backup
                frm_backup_realizar.tipo_backup = "M"
                'exibe o form de backup
                frm_backup_realizar.Show vbModal, Me
                'habilita o timer de baixas automáticas
                tmr_baixas_automaticas.Enabled = True
            End If
        Case "k_backup_restaurar"
            'desabilita o timer de baixas automáticas
            tmr_baixas_automaticas.Enabled = False
            'exibe o form de restauração de backup
            frm_backup_restaurar.Show vbModal, Me
            'habilita o timer de baixas automáticas
            tmr_baixas_automaticas.Enabled = True
        Case "k_ajuda_introducao"
            psub_exibir_ajuda Me, "html/introducao.htm", 0
        Case "k_ajuda_tutorial"
            psub_exibir_ajuda Me, "html/ajuda_tutorial.htm", 0
        Case "k_ajuda_suporte"
            psub_exibir_ajuda Me, "html/ajuda_suporte.htm", 0
        Case "k_ajuda_pesquisa"
            'desabilita o timer de baixas automáticas
            tmr_baixas_automaticas.Enabled = False
            'exibe o form de pesquisa de publico
            frm_pesquisa_publico.Show vbModal, Me
            'habilita o timer de baixas automáticas
            tmr_baixas_automaticas.Enabled = True
        Case "k_ajuda_change_log"
            psub_exibir_ajuda Me, "html/ajuda_changelog.htm", 0
        Case "k_ajuda_sobre"
            frm_splash_sobre.Show vbModal, Me
        Case "k_sair"
            Unload Me
        End Select
    End With
fim_smxp_principal_Click:
    Exit Sub
erro_smxp_principal_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "smxp_principal_Click"
    GoTo fim_smxp_principal_Click
End Sub

Private Sub MDIForm_Activate()
    On Error GoTo erro_MDIForm_Activate
    Dim lint_resposta As Integer
    Dim lstr_mensagem As String
    Dim ltpe_registro As tpe_registro
    'ajusta o caption da janela principal
    frm_principal.Caption = pcst_nome_aplicacao
    'verifica se há usuário logado
    If (p_usuario.lng_codigo = 0) Then
        'registro (força o form a voltar ao estado inicial)
        frm_imagem_fundo.registrado = True
        'dispara o evento de redimensionamento do form
        MDIForm_Resize
        'continua o processo
        If (Not frm_usuario_login.Visible) Then
            frm_usuario_login.Show vbModal, frm_principal
        End If
    End If
    If (p_usuario.str_login <> "") Then
        'ajusta o caption da aplicação
        frm_principal.Caption = pcst_nome_aplicacao & " - [" & UCase$(p_usuario.str_login) & "]"
        'se o backup está desativado
        If (Not p_backup.bln_ativar) Then
            'desativa o menu de backup realizar
            smxp_principal.MenuItems.Enabled("k_backup_realizar") = False
        Else
            'ativa o menu backup realizar
            smxp_principal.MenuItems.Enabled("k_backup_realizar") = True
        End If
        'se é o primeiro acesso do usuário
        If (Not mbln_exibir_introducao) Then
            mbln_exibir_introducao = True
            If (Format$(p_usuario.dt_ultimo_acesso, "dd/mm/yyyy hh:mm:ss") = "30/12/1899 00:00:00") Then
                lstr_mensagem = ""
                lstr_mensagem = lstr_mensagem & "Olá " & p_usuario.str_login & ", seja bem vindo(a)." & vbCrLf
                lstr_mensagem = lstr_mensagem & "Este é o seu primeiro acesso ao " & pcst_nome_aplicacao & "." & vbCrLf
                lstr_mensagem = lstr_mensagem & "Gostaria de visualizar a introdução da aplicação?" & vbCrLf
                lint_resposta = MsgBox(lstr_mensagem, vbYesNo + vbQuestion + vbDefaultButton2, pcst_nome_aplicacao)
                If (lint_resposta = vbYes) Then
                    psub_exibir_ajuda Me, "html/introducao.htm", 0
                End If
            End If
        End If
        'se a aplicação está em modo offline, não precisamos verificar nada disso
        If (Not p_modo_offline) Then
            'se o registro ainda não foi verificado
            If (Not mbln_nao_verificar_registro) Then
                'se o usuário participou da pesquisa
                If (p_usuario.bln_participou_pesquisa) Then
                    'verifica se os dados foram gravados online
                    If (pfct_carrega_dados_pesquisa_publico(p_usuario.str_login, p_pc.str_id_cpu, p_pc.str_id_hd, ltpe_registro)) Then
                        'se o código for zero, não houve gravação online
                        If (ltpe_registro.int_codigo = 0) Then
                            'faz a inserção dos dados
                            pfct_inserir_dados_pesquisa_publico
                        Else
                            'se os dados forem diferentes
                            If (Not pfct_comparar_registros(p_registro, ltpe_registro)) Then
                                'faz a atualização dos dados
                                pfct_atualizar_dados_pesquisa_publico ltpe_registro
                            End If
                        End If
                    End If
                Else    'usuário não está registrado localmente
                    'se existe registro online
                    If (pfct_carrega_dados_pesquisa_publico(p_usuario.str_login, p_pc.str_id_cpu, p_pc.str_id_hd, ltpe_registro)) Then
                        'se o código for diferente de zero, já houve gravação online
                        If (ltpe_registro.int_codigo <> 0) Then
                            'atualiza o objeto de registro
                            If (pfct_copiar_registros(p_registro, ltpe_registro)) Then
                                'sinaliza o usuário como participante da pesquisa
                                p_usuario.bln_participou_pesquisa = True
                                'altera o tipo de banco
                                p_banco.tb_tipo_banco = tb_config
                                'salva as configurações do usuário
                                pfct_salvar_configuracoes_usuario p_usuario.lng_codigo
                                'altera o tipo de banco
                                p_banco.tb_tipo_banco = tb_dados
                            End If
                        End If
                    End If
                End If
                'insere o registro de acesso do usuário
                pfct_inserir_historico_acesso p_usuario.str_login, p_pc.str_id_cpu, p_pc.str_id_hd
                'sinaliza que o registro da aplicação já foi verificado
                mbln_nao_verificar_registro = True
            End If
            'registro
            frm_imagem_fundo.registrado = p_usuario.bln_participou_pesquisa
            'dispara o evento de redimensionamento do form
            MDIForm_Resize
            'verifica se há atualizações
            If (pfct_verificar_atualizacao() And Not (mbln_ja_respondeu_atualizacao)) Then
                'exibe mensagem ao usuário
                If (MsgBox("Há uma atualização do Eiko Finanças Pessoais disponível!" & vbCrLf & "Deseja atualizar agora?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes) Then
                    'chama a api do windows
                    ShellExecute 0&, vbNullString, app_setup, vbNullString, vbNullString, SW_SHOWNORMAL
                End If
                'marca como pergunta já respondida
                mbln_ja_respondeu_atualizacao = True
            End If
        Else
            'se estamos em modo offline, desativamos a pesquisa de público
            smxp_principal.MenuItems.Enabled(smxp_principal.MenuItems.Key2ID("k_ajuda_pesquisa")) = False
        End If
        'se a configuração está definida para carregar a agenda financeira
        If ((p_usuario.bln_carregar_agenda_financeira_login) And (mbln_esta_logando)) Then
            'ativa a agenda financeira
            smxp_principal_Click smxp_principal.MenuItems.Key2ID("k_financeiro_agenda_financeira")
            'informamos ao sistema que já terminamos o login
            mbln_esta_logando = False
        End If
        'ativa o timer de baixas automáticas
        tmr_baixas_automaticas.Enabled = True
    End If
fim_MDIForm_Activate:
    Exit Sub
erro_MDIForm_Activate:
    'altera o tipo de banco
    p_banco.tb_tipo_banco = tb_dados
    'gera o log de erros
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "MDIForm_Activate"
    GoTo fim_MDIForm_Activate
End Sub

Private Sub MDIForm_Initialize()
    On Error GoTo erro_MDIForm_Initialize
    InitCommonControls
fim_MDIForm_Initialize:
    Exit Sub
erro_MDIForm_Initialize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "MDIForm_Initialize"
    GoTo fim_MDIForm_Initialize
End Sub

Private Sub MDIForm_Load()
    On Error GoTo erro_MDIForm_Load
    lsub_criar_menu_principal
    frm_imagem_fundo.Show
    'se não estamos em modo offline
    If (Not p_modo_offline) Then
        'mostramos o form de doação
        frm_doe.Show
    End If
fim_MDIForm_Load:
    Exit Sub
erro_MDIForm_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "MDIForm_Load"
    GoTo fim_MDIForm_Load
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim lint_resposta As Integer
    On Error GoTo erro_MDIForm_QueryUnload
    lint_resposta = MsgBox("Deseja encerrar a aplicação?", vbYesNo + vbQuestion + vbDefaultButton2, pcst_nome_aplicacao)
    If (lint_resposta = vbNo) Then
        Cancel = True
    Else
    
        'desabilita o timer de baixas automáticas
        tmr_baixas_automaticas.Enabled = False
    
        'fecha todos os forms abertos
        psub_fechar_forms
    
        'carrega o form de backup
        Load frm_backup_realizar
        'passa o parâmetro tipo_backup
        frm_backup_realizar.tipo_backup = "A"
        'exibe o form
        frm_backup_realizar.Show 1, Me
        
        'config
        p_banco.tb_tipo_banco = tb_config
        'atualiza as configurações do usuário
        If (pfct_salvar_configuracoes_usuario(p_usuario.lng_codigo)) Then
            'ajusta a data do último acesso do usuário
            p_usuario.dt_ultimo_acesso = Now
            'atualiza o usuário
            psub_atualizar_usuario p_usuario.lng_codigo, False
            'faz manutenção da base de dados
            psub_limpar_banco
        End If
        
        'usuário
        p_banco.tb_tipo_banco = tb_dados
        psub_limpar_banco
        
    End If
fim_MDIForm_QueryUnload:
    Exit Sub
erro_MDIForm_QueryUnload:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "MDIForm_QueryUnload"
    GoTo fim_MDIForm_QueryUnload
End Sub

Private Sub MDIForm_Resize()
    On Error GoTo erro_MDIForm_Resize
    'posiciona a imagem de fundo
    frm_imagem_fundo.Move (Me.Width - frm_imagem_fundo.Width - 450), (Me.Height - frm_imagem_fundo.Height - 1350)
    'se não estamos em modo offline
    If (Not p_modo_offline) Then
        'posiciona a imagem de doe
        frm_doe.Move (Me.Width - frm_doe.Width - 350), 150
    End If
fim_MDIForm_Resize:
    Exit Sub
erro_MDIForm_Resize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "MDIForm_Resize"
    GoTo fim_MDIForm_Resize
End Sub

Private Sub tmr_baixas_automaticas_Timer()
    On Error GoTo Erro_tmr_baixas_automaticas_Timer
    Dim lobj_atualizar_saldo_conta As Object
    Dim lobj_contas_receber As Object
    Dim lobj_contas_pagar As Object
    Dim lobj_contas_receber_excluir As Object
    Dim lobj_contas_pagar_excluir As Object
    Dim lobj_movimentacao As Object
    Dim lobj_contas As Object
    Dim lstr_sql As String
    Dim llng_contador As Long
    Dim llng_registros As Long
    Dim llng_conta As Long
    Dim llng_codigo_conta_pagar As Long
    Dim llng_codigo_conta_receber As Long
    Dim ldbl_saldo_atual As Double
    Dim ldbl_limite_negativo As Double
    Dim ldbl_valor_movimentacao As Double
    'desabilita o timer
    tmr_baixas_automaticas.Enabled = False
    ' -- início baixa automática de contas a receber -- '
    'ajusta para banco tipo dados
    p_banco.tb_tipo_banco = tb_dados
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = "select * from [tb_contas_receber] where [chr_baixa_automatica] = 'S' and [dt_vencimento] <= '" & pfct_tratar_data_sql(Now) & "'"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_contas_receber, lstr_sql, "frm_principal", "tmr_baixas_automaticas_Timer")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo Fim_tmr_baixas_automaticas_Timer
    End If
    llng_registros = lobj_contas_receber.Count
    If (llng_registros > 0) Then
        'percorre os registros encontrados
        For llng_contador = 1 To llng_registros
            'retorna o código do registro na tabela
            llng_codigo_conta_receber = CLng(lobj_contas_receber(llng_contador)("int_codigo"))
            'retorna o código da conta
            llng_conta = CLng(lobj_contas_receber(llng_contador)("int_conta_baixa_automatica"))
            'seleciona o saldo da conta vinculada
            lstr_sql = ""
            lstr_sql = "select * from [tb_contas] where [int_codigo] = " & pfct_tratar_numero_sql(llng_conta)
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_contas, lstr_sql, "frm_principal", "tmr_baixas_automaticas_Timer")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo Fim_tmr_baixas_automaticas_Timer
            End If
            'retorna os valores para o cálculo
            ldbl_saldo_atual = CDbl(lobj_contas(1)("num_saldo"))
            ldbl_valor_movimentacao = CDbl(lobj_contas_receber(llng_contador)("num_valor"))
            '-- atualiza o saldo da conta --'
            'monta o comando sql
            lstr_sql = ""
            lstr_sql = lstr_sql & " update "
            lstr_sql = lstr_sql & " [tb_contas] "
            lstr_sql = lstr_sql & " set "
            lstr_sql = lstr_sql & " [num_saldo] = " & pfct_tratar_numero_sql((ldbl_saldo_atual + ldbl_valor_movimentacao))
            lstr_sql = lstr_sql & " where "
            lstr_sql = lstr_sql & " [int_codigo] = " & pfct_tratar_numero_sql(llng_conta)
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_atualizar_saldo_conta, lstr_sql, "frm_principal", "tmr_baixas_automaticas_Timer")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo Fim_tmr_baixas_automaticas_Timer
            End If
            '-- insere uma nova movimentação --'
            'monta o comando sql
            lstr_sql = ""
            lstr_sql = lstr_sql & " insert into [tb_movimentacao] "
            lstr_sql = lstr_sql & " ( "
            lstr_sql = lstr_sql & " [int_conta], "
            lstr_sql = lstr_sql & " [int_receita], "
            lstr_sql = lstr_sql & " [int_despesa], "
            lstr_sql = lstr_sql & " [int_forma_pagamento], "
            lstr_sql = lstr_sql & " [chr_tipo], "
            lstr_sql = lstr_sql & " [dt_vencimento], "
            lstr_sql = lstr_sql & " [dt_pagamento], "
            lstr_sql = lstr_sql & " [num_valor], "
            lstr_sql = lstr_sql & " [int_parcela], "
            lstr_sql = lstr_sql & " [int_total_parcelas], "
            lstr_sql = lstr_sql & " [str_descricao], "
            lstr_sql = lstr_sql & " [str_documento], "
            lstr_sql = lstr_sql & " [str_codigo_barras], "
            lstr_sql = lstr_sql & " [str_observacoes] "
            lstr_sql = lstr_sql & " ) values ( "
            lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_conta) & ", "
            lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_contas_receber(llng_contador)("int_receita")) & ", "
            lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(0) & ", "
            lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_contas_receber(llng_contador)("int_forma_pagamento")) & ", "
            lstr_sql = lstr_sql & " 'E', "
            lstr_sql = lstr_sql & " '" & pfct_tratar_data_sql(lobj_contas_receber(llng_contador)("dt_vencimento")) & "', "
            lstr_sql = lstr_sql & " '" & pfct_tratar_data_sql(lobj_contas_receber(llng_contador)("dt_vencimento")) & "', "    'como é uma baixa automática, a data de pagamento será igual à data de vencimento
            lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_contas_receber(llng_contador)("num_valor")) & ", "
            lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_contas_receber(llng_contador)("int_parcela")) & ", "
            lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_contas_receber(llng_contador)("int_total_parcelas")) & ", "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_contas_receber(llng_contador)("str_descricao")) & "', "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_contas_receber(llng_contador)("str_documento")) & "', "
            lstr_sql = lstr_sql & " '', " '[str_codigo_barras]
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_contas_receber(llng_contador)("str_observacoes")) & "' "
            lstr_sql = lstr_sql & " ) "
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_movimentacao, lstr_sql, "frm_principal", "tmr_baixas_automaticas_Timer")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo Fim_tmr_baixas_automaticas_Timer
            End If
            '-- exclui a conta a receber da tabela --'
            'monta o comando sql
            lstr_sql = "delete from [tb_contas_receber] where [int_codigo] = " & pfct_tratar_numero_sql(llng_codigo_conta_receber)
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_contas_receber_excluir, lstr_sql, "frm_principal", "tmr_baixas_automaticas_Timer")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo Fim_tmr_baixas_automaticas_Timer
            End If
        Next llng_contador
    End If
    ' -- início baixa automática de contas a pagar -- '
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = "select * from [tb_contas_pagar] where [chr_baixa_automatica] = 'S' and [dt_vencimento] <= '" & pfct_tratar_data_sql(Now) & "'"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_contas_pagar, lstr_sql, "frm_principal", "tmr_baixas_automaticas_Timer")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo Fim_tmr_baixas_automaticas_Timer
    End If
    llng_registros = lobj_contas_pagar.Count
    If (llng_registros > 0) Then
        'percorre os registros encontrados
        For llng_contador = 1 To llng_registros
            'retorna o código do registro na tabela
            llng_codigo_conta_pagar = CLng(lobj_contas_pagar(llng_contador)("int_codigo"))
            'retorna o código da conta
            llng_conta = CLng(lobj_contas_pagar(llng_contador)("int_conta_baixa_automatica"))
            'seleciona o saldo da conta vinculada
            lstr_sql = ""
            lstr_sql = "select * from [tb_contas] where [int_codigo] = " & pfct_tratar_numero_sql(llng_conta)
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_contas, lstr_sql, "frm_principal", "tmr_baixas_automaticas_Timer")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo Fim_tmr_baixas_automaticas_Timer
            End If
            'retorna os valores para o cálculo
            ldbl_saldo_atual = CDbl(lobj_contas(1)("num_saldo"))
            ldbl_limite_negativo = CDbl(lobj_contas(1)("num_limite_negativo"))
            ldbl_valor_movimentacao = CDbl(lobj_contas_pagar(llng_contador)("num_valor"))
            'se houver saldo na conta, realiza a baixa
            If ((ldbl_saldo_atual - ldbl_valor_movimentacao) >= (ldbl_limite_negativo * -1)) Then
                '-- atualiza o saldo da conta --'
                'monta o comando sql
                lstr_sql = ""
                lstr_sql = lstr_sql & " update "
                lstr_sql = lstr_sql & " [tb_contas] "
                lstr_sql = lstr_sql & " set "
                lstr_sql = lstr_sql & " [num_saldo] = " & pfct_tratar_numero_sql((ldbl_saldo_atual - ldbl_valor_movimentacao))
                lstr_sql = lstr_sql & " where "
                lstr_sql = lstr_sql & " [int_codigo] = " & pfct_tratar_numero_sql(llng_conta)
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_atualizar_saldo_conta, lstr_sql, "frm_principal", "tmr_baixas_automaticas_Timer")) Then
                    MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                    GoTo Fim_tmr_baixas_automaticas_Timer
                End If
                '-- insere uma nova movimentação --'
                'monta o comando sql
                lstr_sql = ""
                lstr_sql = lstr_sql & " insert into [tb_movimentacao] "
                lstr_sql = lstr_sql & " ( "
                lstr_sql = lstr_sql & " [int_conta], "
                lstr_sql = lstr_sql & " [int_receita], "
                lstr_sql = lstr_sql & " [int_despesa], "
                lstr_sql = lstr_sql & " [int_forma_pagamento], "
                lstr_sql = lstr_sql & " [chr_tipo], "
                lstr_sql = lstr_sql & " [dt_vencimento], "
                lstr_sql = lstr_sql & " [dt_pagamento], "
                lstr_sql = lstr_sql & " [num_valor], "
                lstr_sql = lstr_sql & " [int_parcela], "
                lstr_sql = lstr_sql & " [int_total_parcelas], "
                lstr_sql = lstr_sql & " [str_descricao], "
                lstr_sql = lstr_sql & " [str_documento], "
                lstr_sql = lstr_sql & " [str_codigo_barras], "
                lstr_sql = lstr_sql & " [str_observacoes] "
                lstr_sql = lstr_sql & " ) values ( "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(llng_conta) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(0) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_contas_pagar(llng_contador)("int_despesa")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_contas_pagar(llng_contador)("int_forma_pagamento")) & ", "
                lstr_sql = lstr_sql & " 'S', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_data_sql(lobj_contas_pagar(llng_contador)("dt_vencimento")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_data_sql(lobj_contas_pagar(llng_contador)("dt_vencimento")) & "', "    'como é uma baixa automática, a data de pagamento será igual à data de vencimento
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_contas_pagar(llng_contador)("num_valor")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_contas_pagar(llng_contador)("int_parcela")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_contas_pagar(llng_contador)("int_total_parcelas")) & ", "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_contas_pagar(llng_contador)("str_descricao")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_contas_pagar(llng_contador)("str_documento")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_contas_pagar(llng_contador)("str_codigo_barras")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_contas_pagar(llng_contador)("str_observacoes")) & "' "
                lstr_sql = lstr_sql & " ) "
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_movimentacao, lstr_sql, "frm_principal", "tmr_baixas_automaticas_Timer")) Then
                    MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                    GoTo Fim_tmr_baixas_automaticas_Timer
                End If
                '-- exclui a conta a pagar da tabela --'
                'monta o comando sql
                lstr_sql = "delete from [tb_contas_pagar] where [int_codigo] = " & pfct_tratar_numero_sql(llng_codigo_conta_pagar)
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_contas_pagar_excluir, lstr_sql, "frm_principal", "tmr_baixas_automaticas_Timer")) Then
                    MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                    GoTo Fim_tmr_baixas_automaticas_Timer
                End If
            End If
        Next llng_contador
    End If
    'se não estamos logando
    If (pfct_form_esta_carregado("frm_financeiro_agenda")) Then
        'se o form financeiro agenda estiver visível
        If (frm_financeiro_agenda.Visible) Then
            'força a atualização das informações
            frm_financeiro_agenda.Form_Activate
        End If
    End If
    'habilita o timer após a conclusão das baixas
    tmr_baixas_automaticas.Enabled = True
Fim_tmr_baixas_automaticas_Timer:
    Set lobj_atualizar_saldo_conta = Nothing
    Set lobj_contas_receber = Nothing
    Set lobj_contas_pagar = Nothing
    Set lobj_contas_receber_excluir = Nothing
    Set lobj_contas_pagar_excluir = Nothing
    Set lobj_movimentacao = Nothing
    Set lobj_contas = Nothing
    Exit Sub
Erro_tmr_baixas_automaticas_Timer:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "tmr_baixas_automaticas_Timer"
    GoTo Fim_tmr_baixas_automaticas_Timer
    Resume 0
End Sub

Private Sub tmr_timer_Timer()
    On Error GoTo erro_tmr_timer_timer
    stb_status.Panels(pnl_data_hora).Text = "Data/Hora: " & Format$(Now, "dd/mm/yyyy hh:mm:ss")
    If (p_usuario.str_login <> "") Then
        stb_status.Panels(pnl_usuario).Text = "Usuário: [" & UCase$(p_usuario.str_login) & "]"
    Else
        stb_status.Panels(pnl_usuario).Text = "Aguardando Login"
    End If
    stb_status.Panels(pnl_ajuda).Text = "Tecle F1 para Ajuda"
    stb_status.Panels(pnl_versao).Text = "Versão -> app: " & pcst_app_ver & " - bd: " & Replace$(pcst_dba_ver, ",", ".")
fim_tmr_timer_timer:
    Exit Sub
erro_tmr_timer_timer:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_principal", "tmr_timer_timer"
    GoTo fim_tmr_timer_timer
End Sub
