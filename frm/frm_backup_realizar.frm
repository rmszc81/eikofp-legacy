VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_backup_realizar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Backup "
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7860
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
   ScaleHeight     =   3750
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pb_progresso_geral 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView lst_log_operacoes 
      Height          =   2475
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   4366
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ProgressBar pb_progresso_individual 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Label lbl_log_operacoes 
      AutoSize        =   -1  'True
      Caption         =   "&Log de operações:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frm_backup_realizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr_tipo_backup As String * 1
Private mlng_codigo_usuario_atual As Long
Private mbln_backup_concluido As Boolean

Public Property Let tipo_backup(ByVal pstrValor As String)
    On Error GoTo erro_tipo_backup
    mstr_tipo_backup = pstrValor
fim_tipo_backup:
    Exit Property
erro_tipo_backup:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_realizar", "tipo_backup"
    GoTo fim_tipo_backup
End Property

Public Property Get tipo_backup() As String
    On Error GoTo erro_tipo_backup
    tipo_backup = mstr_tipo_backup
fim_tipo_backup:
    Exit Property
erro_tipo_backup:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_realizar", "tipo_backup"
    GoTo fim_tipo_backup
End Property

Private Sub lsub_ajustar_barra_progresso_geral(ByVal pint_valor_inicial As Integer, ByVal pint_valor_atual As Integer, ByVal pint_valor_maximo As Integer)
    On Error GoTo erro_lsub_ajustar_barra_progresso_geral
    'se todos os valores forem zero
    If (pint_valor_inicial = 0) And (pint_valor_atual = 0) And (pint_valor_maximo = 0) Then
        'ajustamos o valor da barra para zero
        pb_progresso_geral.Value = 0
    Else
        'ajustamos o valor inicial
        pb_progresso_geral.Min = pint_valor_inicial
        'ajustamos o valor atual
        pb_progresso_geral.Value = pint_valor_atual
        'ajustamos o valor máximo
        pb_progresso_geral.Max = pint_valor_maximo
    End If
fim_lsub_ajustar_barra_progresso_geral:
    Exit Sub
erro_lsub_ajustar_barra_progresso_geral:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_realizar", "lsub_ajustar_barra_progresso_geral"
    GoTo fim_lsub_ajustar_barra_progresso_geral
End Sub

Private Sub lsub_ajustar_barra_progresso_individual(ByVal pint_valor_inicial As Integer, ByVal pint_valor_atual As Integer, ByVal pint_valor_maximo As Integer)
    On Error GoTo erro_lsub_ajustar_barra_progresso_individual
    'se todos os valores forem zero
    If (pint_valor_inicial = 0) And (pint_valor_atual = 0) And (pint_valor_maximo = 0) Then
        'ajustamos o valor da barra para zero
        pb_progresso_individual.Value = 0
    Else
        'ajustamos o valor inicial
        pb_progresso_individual.Min = pint_valor_inicial
        'ajustamos o valor atual
        pb_progresso_individual.Value = pint_valor_atual
        'ajustamos o valor máximo
        pb_progresso_individual.Max = pint_valor_maximo
    End If
fim_lsub_ajustar_barra_progresso_individual:
    Exit Sub
erro_lsub_ajustar_barra_progresso_individual:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_realizar", "lsub_ajustar_barra_progresso_individual"
    GoTo fim_lsub_ajustar_barra_progresso_individual
End Sub

Private Sub lsub_ajustar_lista_log()
    On Error GoTo erro_lsub_ajustar_lista_log
    With lst_log_operacoes
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Key:="k_contagem", Text:=" # ", Width:=500, Alignment:=lvwColumnLeft
        .ColumnHeaders.Add Key:="k_data", Text:=" Data ", Width:=1100, Alignment:=lvwColumnLeft
        .ColumnHeaders.Add Key:="k_hora", Text:=" Hora ", Width:=1100, Alignment:=lvwColumnLeft
        .ColumnHeaders.Add Key:="k_evento", Text:=" Evento ", Width:=4640, Alignment:=lvwColumnLeft
    End With
fim_lsub_ajustar_lista_log:
    Exit Sub
erro_lsub_ajustar_lista_log:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_realizar", "lsub_ajustar_lista_log"
    GoTo fim_lsub_ajustar_lista_log
End Sub

Private Sub lsub_ajustar_status_operacao_mensagem(ByVal pstr_mensagem As String)
    On Error GoTo erro_lsub_ajustar_status_operacao_mensagem
    Dim lst_item As ListItem
    'força o processamento de mensagens
    DoEvents
    'adiciona os dados às colunas
    Set lst_item = lst_log_operacoes.ListItems.Add(Text:=CStr(lst_log_operacoes.ListItems.Count + 1))
    lst_item.ListSubItems.Add Text:=Format$(Now, "dd/mm/yyyy")
    lst_item.ListSubItems.Add Text:=Format$(Now, "hh:mm:ss")
    lst_item.ListSubItems.Add Text:=pstr_mensagem
    'posiciona o foco no último item da lista
    lst_log_operacoes.ListItems(lst_log_operacoes.ListItems.Count).Selected = True
    'força o componente a ajustar o foco no item selecionado
    lst_log_operacoes.SelectedItem.EnsureVisible
    'atualiza o componente
    lst_log_operacoes.Refresh
fim_lsub_ajustar_status_operacao_mensagem:
    Set lst_item = Nothing
    Exit Sub
erro_lsub_ajustar_status_operacao_mensagem:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_realizar", "lsub_ajustar_status_operacao_mensagem"
    GoTo fim_lsub_ajustar_status_operacao_mensagem
End Sub

Private Function lfct_executar_backup() As Boolean
    On Error GoTo erro_lfct_executar_backup
    Dim lobj_usuarios As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    'ajusta a variável para false
    mbln_backup_concluido = False
    'atualiza o form
    Me.Refresh
    'ajusta para banco tipo config
    p_banco.tb_tipo_banco = tb_config
    'tipo de backup automático
    If (mstr_tipo_backup = "A") Then
        'monta o comando sql
        lstr_sql = "select * from [tb_backup] where [chr_ativar] = 'S' and [dt_proximo_backup] <= '" & Format$(Now, "yyyy-mm-dd") & "'"
        'executa o comando sql e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_usuarios, lstr_sql, "frm_backup_realizar", "lfct_executar_backup")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_lfct_executar_backup
        Else
            'contagem de registros
            llng_registros = lobj_usuarios.Count
            'se houver registros
            If (llng_registros > 0) Then
                'executa o backup individual de cada usuário
                For llng_contador = 1 To llng_registros
                    'processa as mensagens do windows
                    DoEvents
                    'executa a função
                    If (Not lfct_executar_backup_usuario(lobj_usuarios(llng_contador)("int_usuario"))) Then
                        GoTo fim_lfct_executar_backup
                    End If
                Next
                'ajusta variável para true
                mbln_backup_concluido = True
            Else 'se não houver registros
                'ajusta variável para true
                mbln_backup_concluido = True
                'retorna true
                lfct_executar_backup = True
                'desvia ao fim do método
                GoTo fim_lfct_executar_backup
            End If
        End If
    End If
    'tipo de backup manual
    If (mstr_tipo_backup = "M") Then
        If (Not lfct_executar_backup_usuario(p_usuario.lng_codigo)) Then
            GoTo fim_lfct_executar_backup
        End If
        mbln_backup_concluido = True
    End If
    
    'limpa o nome do arquivo de backup
    p_backup.str_nome = Empty
    
    'recarrega os dados do usuário logado
    If (pfct_carregar_dados_usuario(mlng_codigo_usuario_atual)) Then
        'recarrega as configurações do usuário logado
        If (pfct_carregar_configuracoes_usuario(mlng_codigo_usuario_atual)) Then
            ' -- backup -- '
            p_banco.tb_tipo_banco = tb_backup
            pfct_ajustar_caminho_banco tb_backup
            ' -- config -- '
            p_banco.tb_tipo_banco = tb_config
            pfct_ajustar_caminho_banco tb_config
            ' -- dados -- '
            p_banco.tb_tipo_banco = tb_dados
            pfct_ajustar_caminho_banco tb_dados
        End If
    End If
    
    'retorna true
    lfct_executar_backup = True
fim_lfct_executar_backup:
    Set lobj_usuarios = Nothing
    Exit Function
erro_lfct_executar_backup:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_realizar", "lfct_executar_backup"
    GoTo fim_lfct_executar_backup
End Function

Private Function lfct_executar_backup_usuario(ByVal plng_usuario As Long) As Boolean
    On Error GoTo erro_lfct_executar_backup_usuario
    Dim lobj_backup As Object
    Dim lobj_dados As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
        
    'ajusta o banco de dados para config
    p_banco.tb_tipo_banco = tb_config
    
    'carrega os dados do usuário
    If (pfct_carregar_dados_usuario(plng_usuario)) Then
    
        'carrega as configurações do usuário
        If (pfct_carregar_configuracoes_usuario(plng_usuario)) Then
        
            ' -- config -- '
            p_banco.tb_tipo_banco = tb_config
            pfct_ajustar_caminho_banco tb_config
            
            ' -- dados -- '
            p_banco.tb_tipo_banco = tb_dados
            pfct_ajustar_caminho_banco tb_dados
        
            'atualiza o form
            Me.Refresh
            
            'ajusta a mensagem inicia
            lsub_ajustar_status_operacao_mensagem "Iniciando backup do usuário [" & p_usuario.str_login & "] ..."
            
            'força a aplicação a reprocessar as mensagens
            DoEvents
        
            'aguarda 2 segundos antes de iniciar...
            Sleep (2000)
            
            'verificando se o diretório existe.
            lsub_ajustar_status_operacao_mensagem "Verificando se a pasta de backup existe..."
            If (Not pfct_verificar_pasta_existe(p_banco.str_caminho_backup)) Then
                lsub_ajustar_status_operacao_mensagem "Pasta de backup não encontrada."
                lsub_ajustar_status_operacao_mensagem "Criando a pasta de backup..."
                If (Not pfct_criar_pasta(p_backup.str_caminho)) Then
                    lsub_ajustar_status_operacao_mensagem "Erro ao criar a pasta de backup. Operação cancelada."
                    GoTo fim_lfct_executar_backup_usuario
                Else
                    lsub_ajustar_status_operacao_mensagem "Pasta de backup criada com sucesso."
                End If
            Else
                lsub_ajustar_status_operacao_mensagem "Pasta de backup encontrada."
            End If
            
            'atualizamos a barra de progresso geral
            lsub_ajustar_barra_progresso_geral 0, 1, 17 'total de 17 passos
            
            'concatena o nome do arquivo a gerar
            'formato - ano;mes;dia;usuario.dbbkp
            p_backup.str_nome = ""
            p_backup.str_nome = p_backup.str_nome & Format$(Year(Now), "0000") & ";"        'ano
            p_backup.str_nome = p_backup.str_nome & Format$(Month(Now), "00") & ";"         'mês
            p_backup.str_nome = p_backup.str_nome & Format$(Day(Now), "00") & ";"           'dia
            p_backup.str_nome = p_backup.str_nome & LCase$(p_usuario.str_login)             'usuario
            p_backup.str_nome = p_backup.str_nome & ".dbbkp"                                'extensão
            
            ' -- backup -- '
            p_banco.tb_tipo_banco = tb_backup
            pfct_ajustar_caminho_banco tb_backup
            
            lsub_ajustar_status_operacao_mensagem "Verificando se existe arquivo de backup anterior..."
            
            'verifica se o arquivo existe
            If (pfct_verificar_arquivo_existe(p_banco.str_caminho_dados_backup)) Then
                lsub_ajustar_status_operacao_mensagem "Arquivo encontrado. Excluindo arquivo..."
                'se existe exclui
                If (Not pfct_excluir_arquivo(p_banco.str_caminho_dados_backup)) Then
                    lsub_ajustar_status_operacao_mensagem "Erro ao excluir arquivo. Operação cancelada."
                    GoTo fim_lfct_executar_backup_usuario
                Else
                    lsub_ajustar_status_operacao_mensagem "Arquivo excluído com sucesso."
                End If
            Else
                lsub_ajustar_status_operacao_mensagem "Arquivo não foi localizado."
            End If
            
            'atualizamos a barra de progresso geral
            lsub_ajustar_barra_progresso_geral 0, 2, 17 'total de 17 passos
            
            lsub_ajustar_status_operacao_mensagem "Criando tabelas de configuração..."
            
            'cria as tabelas de config
            If (Not pfct_criar_tabelas_config) Then
                lsub_ajustar_status_operacao_mensagem "Erro ao criar tabelas. Operação cancelada."
                GoTo fim_lfct_executar_backup_usuario
            Else
            
                lsub_ajustar_status_operacao_mensagem "Tabelas criadas com sucesso."
                lsub_ajustar_status_operacao_mensagem "Criando tabelas de dados..."
                
                'cria as tabelas de dados
                If (Not pfct_criar_tabelas_usuario) Then
                    lsub_ajustar_status_operacao_mensagem "Erro ao criar tabelas. Operação cancelada."
                    GoTo fim_lfct_executar_backup_usuario
                Else
                    lsub_ajustar_status_operacao_mensagem "Tabelas criadas com sucesso."
                    
                    'atualizamos a barra de progresso geral
                    lsub_ajustar_barra_progresso_geral 0, 3, 17 'total de 17 passos
                    
                    'mensagem
                    lsub_ajustar_status_operacao_mensagem "Processo de cópia será iniciado em 5 segundos."
                    lsub_ajustar_status_operacao_mensagem "Isso pode levar alguns minutos, aguarde..."
        
                    'aguarda 5 segundos antes de iniciar a cópia...
                    Sleep (5000)
        
                    '----- ini config -----'
                    '
                    
                    'ajusta o banco para config
                    p_banco.tb_tipo_banco = tb_config
        
                    'usuários
                    
                    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de usuários..."
        
                    'monta o comando sql
                    lstr_sql = "select * from [tb_usuarios] where [int_codigo] = " & pfct_tratar_numero_sql(p_usuario.lng_codigo)
                    'executa o comando sql e devolve o objeto
                    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de usuários. Operação cancelada."
                        GoTo fim_lfct_executar_backup_usuario
                    Else
                        llng_registros = lobj_dados.Count
                        If (llng_registros = 0) Then
                            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
                        ElseIf (llng_registros > 0) Then
                            'exibe mensagem
                            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
                            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
                            'ajusta o banco para dados
                            p_banco.tb_tipo_banco = tb_backup
                            'percorre o objeto
                            For llng_contador = 1 To llng_registros
                                'processa as mensagens do windows
                                DoEvents
                                'monta o comando sql
                                lstr_sql = ""
                                lstr_sql = lstr_sql & " insert into [tb_usuarios] "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " [int_codigo], "
                                lstr_sql = lstr_sql & " [str_usuario], "
                                lstr_sql = lstr_sql & " [str_senha], "
                                lstr_sql = lstr_sql & " [dt_criado_em], "
                                lstr_sql = lstr_sql & " [tm_criado_em], "
                                lstr_sql = lstr_sql & " [dt_ultimo_acesso], "
                                lstr_sql = lstr_sql & " [tm_ultimo_acesso] "
                                lstr_sql = lstr_sql & " ) "
                                lstr_sql = lstr_sql & " values "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_codigo")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_usuario")) & "', "
                                lstr_sql = lstr_sql & " '" & lobj_dados(llng_contador)("str_senha") & "', " 'sem tratamento, pois já é criptografado
                                lstr_sql = lstr_sql & " '" & lobj_dados(llng_contador)("dt_criado_em") & "', "
                                lstr_sql = lstr_sql & " '" & lobj_dados(llng_contador)("tm_criado_em") & "', "
                                lstr_sql = lstr_sql & " '" & lobj_dados(llng_contador)("dt_ultimo_acesso") & "', "
                                lstr_sql = lstr_sql & " '" & lobj_dados(llng_contador)("tm_ultimo_acesso") & "' "
                                lstr_sql = lstr_sql & " ) "
                                'executa o comando sql e devolve o objeto
                                If (Not pfct_executar_comando_sql(lobj_backup, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                                    GoTo fim_lfct_executar_backup_usuario
                                End If
                                'atualizamos a barra de progresso
                                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
                            Next
                            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
                        End If
                    End If
                    
                    'atualizamos a barra de progresso geral
                    lsub_ajustar_barra_progresso_geral 0, 4, 17 'total de 17 passos
        
                    'ajusta o banco para config
                    p_banco.tb_tipo_banco = tb_config
        
                    'config
                    
                    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de configuração..."
        
                    'monta o comando sql
                    lstr_sql = "select * from [tb_config] where [int_usuario] = " & pfct_tratar_numero_sql(p_usuario.lng_codigo)
                    'executa o comando sql e devolve o objeto
                    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de configuração. Operação cancelada."
                        GoTo fim_lfct_executar_backup_usuario
                    Else
                        llng_registros = lobj_dados.Count
                        If (llng_registros = 0) Then
                            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
                        ElseIf (llng_registros > 0) Then
                            'exibe mensagem
                            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
                            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
                            'ajusta o banco para dados
                            p_banco.tb_tipo_banco = tb_backup
                            'percorre o objeto
                            For llng_contador = 1 To llng_registros
                                'processa as mensagens do windows
                                DoEvents
                                'monta o comando sql
                                lstr_sql = ""
                                lstr_sql = lstr_sql & " insert into [tb_config] "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " [int_codigo], "
                                lstr_sql = lstr_sql & " [int_usuario], "
                                lstr_sql = lstr_sql & " [int_moeda], "
                                lstr_sql = lstr_sql & " [int_intervalo_data], "
                                lstr_sql = lstr_sql & " [chr_carregar_agenda_financeira_login], "
                                lstr_sql = lstr_sql & " [chr_lancamentos_retroativos], "
                                lstr_sql = lstr_sql & " [chr_alteracoes_detalhes], "
                                lstr_sql = lstr_sql & " [chr_data_vencimento_baixa_imediata], "
                                lstr_sql = lstr_sql & " [chr_lancamentos_duplicados], "
                                lstr_sql = lstr_sql & " [chr_participou_pesquisa] "
                                lstr_sql = lstr_sql & " ) "
                                lstr_sql = lstr_sql & " values "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_codigo")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_usuario")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_moeda")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_intervalo_data")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_carregar_agenda_financeira_login")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_lancamentos_retroativos")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_alteracoes_detalhes")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_data_vencimento_baixa_imediata")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_lancamentos_duplicados")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_participou_pesquisa")) & "' "
                                lstr_sql = lstr_sql & " ) "
                                'executa o comando sql e devolve o objeto
                                If (Not pfct_executar_comando_sql(lobj_backup, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                                    GoTo fim_lfct_executar_backup_usuario
                                End If
                                'atualizamos a barra de progresso
                                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
                            Next
                            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
                        End If
                    End If
                    
                    'atualizamos a barra de progresso geral
                    lsub_ajustar_barra_progresso_geral 0, 5, 17 'total de 17 passos
        
                    'ajusta o banco para config
                    p_banco.tb_tipo_banco = tb_config
        
                    'backup
                    
                    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de backup..."
        
                    'monta o comando sql
                    lstr_sql = "select * from [tb_backup] where [int_usuario] = " & pfct_tratar_numero_sql(p_usuario.lng_codigo)
                    'executa o comando sql e devolve o objeto
                    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de backup. Operação cancelada."
                        GoTo fim_lfct_executar_backup_usuario
                    Else
                        llng_registros = lobj_dados.Count
                        If (llng_registros = 0) Then
                            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
                        ElseIf (llng_registros > 0) Then
                            'exibe mensagem
                            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
                            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
                            'ajusta o banco para dados
                            p_banco.tb_tipo_banco = tb_backup
                            'percorre o objeto
                            For llng_contador = 1 To llng_registros
                                'processa as mensagens do windows
                                DoEvents
                                'monta o comando sql
                                lstr_sql = ""
                                lstr_sql = lstr_sql & " insert into [tb_backup] "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " [int_codigo], "
                                lstr_sql = lstr_sql & " [int_usuario], "
                                lstr_sql = lstr_sql & " [chr_ativar], "
                                lstr_sql = lstr_sql & " [int_periodo], "
                                lstr_sql = lstr_sql & " [str_caminho], "
                                lstr_sql = lstr_sql & " [dt_ultimo_backup], "
                                lstr_sql = lstr_sql & " [dt_proximo_backup] "
                                lstr_sql = lstr_sql & " ) "
                                lstr_sql = lstr_sql & " values "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_codigo")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_usuario")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_ativar")) & "', "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_periodo")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_caminho")) & "', "
                                lstr_sql = lstr_sql & " '" & lobj_dados(llng_contador)("dt_ultimo_backup") & "', "
                                lstr_sql = lstr_sql & " '" & lobj_dados(llng_contador)("dt_proximo_backup") & "' "
                                lstr_sql = lstr_sql & " ) "
                                'executa o comando sql e devolve o objeto
                                If (Not pfct_executar_comando_sql(lobj_backup, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                                    GoTo fim_lfct_executar_backup_usuario
                                End If
                                'atualizamos a barra de progresso
                                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
                            Next
                            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
                        End If
                    End If
                    
                    'atualizamos a barra de progresso geral
                    lsub_ajustar_barra_progresso_geral 0, 6, 17 'total de 17 passos
                    
                    '
                    '----- fim config -----'

                    '----- ini dados -----'
                    '
                    
                    'ajusta o banco para dados
                    p_banco.tb_tipo_banco = tb_dados
        
                    'contas
                    
                    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de contas..."
        
                    'monta o comando sql
                    lstr_sql = "select * from [tb_contas]"
                    'executa o comando sql e devolve o objeto
                    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de contas. Operação cancelada."
                        GoTo fim_lfct_executar_backup_usuario
                    Else
                        llng_registros = lobj_dados.Count
                        If (llng_registros = 0) Then
                            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
                        ElseIf (llng_registros > 0) Then
                            'exibe mensagem
                            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
                            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
                            'ajusta o banco para dados
                            p_banco.tb_tipo_banco = tb_backup
                            'percorre o objeto
                            For llng_contador = 1 To llng_registros
                                'processa as mensagens do windows
                                DoEvents
                                'monta o comando sql
                                lstr_sql = ""
                                lstr_sql = lstr_sql & " insert into [tb_contas] "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " [int_codigo], "
                                lstr_sql = lstr_sql & " [str_descricao], "
                                lstr_sql = lstr_sql & " [num_saldo], "
                                lstr_sql = lstr_sql & " [num_limite_negativo], "
                                lstr_sql = lstr_sql & " [str_observacoes], "
                                lstr_sql = lstr_sql & " [chr_ativo] "
                                lstr_sql = lstr_sql & " ) "
                                lstr_sql = lstr_sql & " values "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_codigo")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_descricao")) & "', "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("num_saldo")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("num_limite_negativo")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_observacoes")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_ativo")) & "' "
                                lstr_sql = lstr_sql & " ) "
                                'executa o comando sql e devolve o objeto
                                If (Not pfct_executar_comando_sql(lobj_backup, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                                    GoTo fim_lfct_executar_backup_usuario
                                End If
                                'atualizamos a barra de progresso
                                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
                            Next
                            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
                        End If
                    End If
                    
                    'atualizamos a barra de progresso geral
                    lsub_ajustar_barra_progresso_geral 0, 7, 17 'total de 17 passos
        
                    'ajusta o banco para dados
                    p_banco.tb_tipo_banco = tb_dados
        
                    'receitas
                    
                    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de receitas..."
        
                    'monta o comando sql
                    lstr_sql = "select * from [tb_receitas]"
                    'executa o comando sql e devolve o objeto
                    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de receitas. Operação cancelada."
                        GoTo fim_lfct_executar_backup_usuario
                    Else
                        llng_registros = lobj_dados.Count
                        If (llng_registros = 0) Then
                            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
                        ElseIf (llng_registros > 0) Then
                            'exibe mensagem
                            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
                            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
                            'ajusta o banco para dados
                            p_banco.tb_tipo_banco = tb_backup
                            'percorre o objeto
                            For llng_contador = 1 To llng_registros
                                'processa as mensagens do windows
                                DoEvents
                                'monta o comando sql
                                lstr_sql = ""
                                lstr_sql = lstr_sql & " insert into [tb_receitas]"
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " [int_codigo], "
                                lstr_sql = lstr_sql & " [str_descricao], "
                                lstr_sql = lstr_sql & " [str_observacoes], "
                                lstr_sql = lstr_sql & " [chr_fixa], "
                                lstr_sql = lstr_sql & " [chr_ativo] "
                                lstr_sql = lstr_sql & " ) "
                                lstr_sql = lstr_sql & " values "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_codigo")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_descricao")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_observacoes")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_fixa")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_ativo")) & "' "
                                lstr_sql = lstr_sql & " ) "
                                'executa o comando sql e devolve o objeto
                                If (Not pfct_executar_comando_sql(lobj_backup, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                                    GoTo fim_lfct_executar_backup_usuario
                                End If
                                'atualizamos a barra de progresso
                                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
                            Next
                            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
                        End If
                    End If
                    
                    'atualizamos a barra de progresso geral
                    lsub_ajustar_barra_progresso_geral 0, 8, 17 'total de 17 passos
        
                    'ajusta o banco para dados
                    p_banco.tb_tipo_banco = tb_dados
        
                    'despesas
                    
                    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de despesas..."
        
                    'monta o comando sql
                    lstr_sql = "select * from [tb_despesas]"
                    'executa o comando sql e devolve o objeto
                    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de despesas. Operação cancelada."
                        GoTo fim_lfct_executar_backup_usuario
                    Else
                        llng_registros = lobj_dados.Count
                        If (llng_registros = 0) Then
                            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
                        ElseIf (llng_registros > 0) Then
                            'exibe mensagem
                            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
                            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
                            'ajusta o banco para dados
                            p_banco.tb_tipo_banco = tb_backup
                            'percorre o objeto
                            For llng_contador = 1 To llng_registros
                                'processa as mensagens do windows
                                DoEvents
                                'monta o comando sql
                                lstr_sql = ""
                                lstr_sql = lstr_sql & " insert into [tb_despesas]"
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " [int_codigo], "
                                lstr_sql = lstr_sql & " [str_descricao], "
                                lstr_sql = lstr_sql & " [str_observacoes], "
                                lstr_sql = lstr_sql & " [chr_fixa], "
                                lstr_sql = lstr_sql & " [chr_ativo] "
                                lstr_sql = lstr_sql & " ) "
                                lstr_sql = lstr_sql & " values "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_codigo")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_descricao")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_observacoes")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_fixa")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_ativo")) & "' "
                                lstr_sql = lstr_sql & " ) "
                                'executa o comando sql e devolve o objeto
                                If (Not pfct_executar_comando_sql(lobj_backup, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                                    GoTo fim_lfct_executar_backup_usuario
                                End If
                                'atualizamos a barra de progresso
                                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
                            Next
                            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
                        End If
                    End If
                    
                    'atualizamos a barra de progresso geral
                    lsub_ajustar_barra_progresso_geral 0, 9, 17 'total de 17 passos
        
                    'ajusta o banco para dados
                    p_banco.tb_tipo_banco = tb_dados
        
                    'formas de pagamento
                    
                    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de formas de pagamento..."
        
                    'monta o comando sql
                    lstr_sql = "select * from [tb_formas_pagamento]"
                    'executa o comando sql e devolve o objeto
                    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de formas de pagamento. Operação cancelada."
                        GoTo fim_lfct_executar_backup_usuario
                    Else
                        llng_registros = lobj_dados.Count
                        If (llng_registros = 0) Then
                            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
                        ElseIf (llng_registros > 0) Then
                            'exibe mensagem
                            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
                            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
                            'ajusta o banco para dados
                            p_banco.tb_tipo_banco = tb_backup
                            'percorre o objeto
                            For llng_contador = 1 To llng_registros
                                'processa as mensagens do windows
                                DoEvents
                                'monta o comando sql
                                lstr_sql = ""
                                lstr_sql = lstr_sql & " insert into [tb_formas_pagamento]"
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " [int_codigo], "
                                lstr_sql = lstr_sql & " [str_descricao], "
                                lstr_sql = lstr_sql & " [str_observacoes], "
                                lstr_sql = lstr_sql & " [chr_ativo] "
                                lstr_sql = lstr_sql & " ) "
                                lstr_sql = lstr_sql & " values "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_codigo")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_descricao")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_observacoes")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_ativo")) & "' "
                                lstr_sql = lstr_sql & " ) "
                                'executa o comando sql e devolve o objeto
                                If (Not pfct_executar_comando_sql(lobj_backup, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                                    GoTo fim_lfct_executar_backup_usuario
                                End If
                                'atualizamos a barra de progresso
                                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
                            Next
                            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
                        End If
                    End If
                    
                    'atualizamos a barra de progresso geral
                    lsub_ajustar_barra_progresso_geral 0, 10, 17 'total de 17 passos
        
                    'ajusta o banco para dados
                    p_banco.tb_tipo_banco = tb_dados
        
                    'formas de pagamento
                    
                    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de contas a pagar..."
        
                    'monta o comando sql
                    lstr_sql = "select * from [tb_contas_pagar]"
                    'executa o comando sql e devolve o objeto
                    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de contas a pagar. Operação cancelada."
                        GoTo fim_lfct_executar_backup_usuario
                    Else
                        llng_registros = lobj_dados.Count
                        If (llng_registros = 0) Then
                            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
                        ElseIf (llng_registros > 0) Then
                            'exibe mensagem
                            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
                            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
                            'ajusta o banco para dados
                            p_banco.tb_tipo_banco = tb_backup
                            'percorre o objeto
                            For llng_contador = 1 To llng_registros
                                'processa as mensagens do windows
                                DoEvents
                                'monta o comando sql
                                lstr_sql = ""
                                lstr_sql = lstr_sql & " insert into [tb_contas_pagar] "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " [int_codigo], "
                                lstr_sql = lstr_sql & " [chr_baixa_automatica], "
                                lstr_sql = lstr_sql & " [int_conta_baixa_automatica], "
                                lstr_sql = lstr_sql & " [int_despesa], "
                                lstr_sql = lstr_sql & " [int_forma_pagamento], "
                                lstr_sql = lstr_sql & " [dt_vencimento], "
                                lstr_sql = lstr_sql & " [int_parcela], "
                                lstr_sql = lstr_sql & " [int_total_parcelas], "
                                lstr_sql = lstr_sql & " [num_valor], "
                                lstr_sql = lstr_sql & " [str_descricao], "
                                lstr_sql = lstr_sql & " [str_documento], "
                                lstr_sql = lstr_sql & " [str_chave], "
                                lstr_sql = lstr_sql & " [str_codigo_barras], "
                                lstr_sql = lstr_sql & " [str_observacoes] "
                                lstr_sql = lstr_sql & " ) "
                                lstr_sql = lstr_sql & " values "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_codigo")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_baixa_automatica")) & "', "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_conta_baixa_automatica")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_despesa")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_forma_pagamento")) & ", "
                                lstr_sql = lstr_sql & " '" & lobj_dados(llng_contador)("dt_vencimento") & "', "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_parcela")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_total_parcelas")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("num_valor")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_descricao")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_documento")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_chave")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_codigo_barras")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_observacoes")) & "' "
                                lstr_sql = lstr_sql & " ) "
                                'executa o comando sql e devolve o objeto"
                                If (Not pfct_executar_comando_sql(lobj_backup, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                                    GoTo fim_lfct_executar_backup_usuario
                                End If
                                'atualizamos a barra de progresso
                                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
                            Next
                            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
                        End If
                    End If
                    
                    'atualizamos a barra de progresso geral
                    lsub_ajustar_barra_progresso_geral 0, 11, 17 'total de 17 passos
        
                    'ajusta o banco para dados
                    p_banco.tb_tipo_banco = tb_dados
        
                    'formas de pagamento
                    
                    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de contas a receber..."
        
                    'monta o comando sql
                    lstr_sql = "select * from [tb_contas_receber]"
                    'executa o comando sql e devolve o objeto
                    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de contas a receber. Operação cancelada."
                        GoTo fim_lfct_executar_backup_usuario
                    Else
                        llng_registros = lobj_dados.Count
                        If (llng_registros = 0) Then
                            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
                        ElseIf (llng_registros > 0) Then
                            'exibe mensagem
                            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
                            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
                            'ajusta o banco para dados
                            p_banco.tb_tipo_banco = tb_backup
                            'percorre o objeto
                            For llng_contador = 1 To llng_registros
                                'processa as mensagens do windows
                                DoEvents
                                'monta o comando sql
                                lstr_sql = ""
                                lstr_sql = lstr_sql & " insert into [tb_contas_receber] "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " [int_codigo], "
                                lstr_sql = lstr_sql & " [chr_baixa_automatica], "
                                lstr_sql = lstr_sql & " [int_conta_baixa_automatica], "
                                lstr_sql = lstr_sql & " [int_receita], "
                                lstr_sql = lstr_sql & " [int_forma_pagamento], "
                                lstr_sql = lstr_sql & " [dt_vencimento], "
                                lstr_sql = lstr_sql & " [int_parcela], "
                                lstr_sql = lstr_sql & " [int_total_parcelas], "
                                lstr_sql = lstr_sql & " [num_valor], "
                                lstr_sql = lstr_sql & " [str_descricao], "
                                lstr_sql = lstr_sql & " [str_documento], "
                                lstr_sql = lstr_sql & " [str_chave], "
                                lstr_sql = lstr_sql & " [str_observacoes] "
                                lstr_sql = lstr_sql & " ) "
                                lstr_sql = lstr_sql & " values "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_codigo")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_baixa_automatica")) & "', "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_conta_baixa_automatica")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_receita")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_forma_pagamento")) & ", "
                                lstr_sql = lstr_sql & " '" & lobj_dados(llng_contador)("dt_vencimento") & "', "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_parcela")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_total_parcelas")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("num_valor")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_descricao")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_documento")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_chave")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_observacoes")) & "' "
                                lstr_sql = lstr_sql & " ) "
                                'executa o comando sql e devolve o objeto"
                                If (Not pfct_executar_comando_sql(lobj_backup, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                                    GoTo fim_lfct_executar_backup_usuario
                                End If
                                'atualizamos a barra de progresso
                                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
                            Next
                            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
                        End If
                    End If
                    
                    'atualizamos a barra de progresso geral
                    lsub_ajustar_barra_progresso_geral 0, 13, 17 'total de 17 passos
        
                    'ajusta o banco para dados
                    p_banco.tb_tipo_banco = tb_dados
        
                    'movimentação
                    
                    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de movimentação..."
        
                    'monta o comando sql
                    lstr_sql = "select * from [tb_movimentacao]"
                    'executa o comando sql e devolve o objeto
                    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de movimentação. Operação cancelada."
                        GoTo fim_lfct_executar_backup_usuario
                    Else
                        llng_registros = lobj_dados.Count
                        If (llng_registros = 0) Then
                            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
                        ElseIf (llng_registros > 0) Then
                            'exibe mensagem
                            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
                            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
                            'ajusta o banco para dados
                            p_banco.tb_tipo_banco = tb_backup
                            'percorre o objeto
                            For llng_contador = 1 To llng_registros
                                'processa as mensagens do windows
                                DoEvents
                                'monta o comando sql
                                lstr_sql = ""
                                lstr_sql = lstr_sql & " insert into [tb_movimentacao] "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " [int_codigo], "
                                lstr_sql = lstr_sql & " [int_conta], "
                                lstr_sql = lstr_sql & " [int_receita], "
                                lstr_sql = lstr_sql & " [int_despesa], "
                                lstr_sql = lstr_sql & " [int_forma_pagamento], "
                                lstr_sql = lstr_sql & " [chr_tipo], "
                                lstr_sql = lstr_sql & " [dt_vencimento], "
                                lstr_sql = lstr_sql & " [dt_pagamento], "
                                lstr_sql = lstr_sql & " [int_parcela], "
                                lstr_sql = lstr_sql & " [int_total_parcelas], "
                                lstr_sql = lstr_sql & " [num_valor], "
                                lstr_sql = lstr_sql & " [str_descricao], "
                                lstr_sql = lstr_sql & " [str_documento], "
                                lstr_sql = lstr_sql & " [str_codigo_barras], "
                                lstr_sql = lstr_sql & " [str_observacoes] "
                                lstr_sql = lstr_sql & " ) "
                                lstr_sql = lstr_sql & " values "
                                lstr_sql = lstr_sql & " ( "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_codigo")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_conta")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_receita")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_despesa")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_forma_pagamento")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("chr_tipo")) & "', "
                                lstr_sql = lstr_sql & " '" & lobj_dados(llng_contador)("dt_vencimento") & "', "
                                lstr_sql = lstr_sql & " '" & lobj_dados(llng_contador)("dt_pagamento") & "', "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_parcela")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("int_total_parcelas")) & ", "
                                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_dados(llng_contador)("num_valor")) & ", "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_descricao")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_documento")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_codigo_barras")) & "', "
                                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_dados(llng_contador)("str_observacoes")) & "' "
                                lstr_sql = lstr_sql & " ) "
                                'executa o comando sql e devolve o objeto
                                If (Not pfct_executar_comando_sql(lobj_backup, lstr_sql, "frm_backup_realizar", "lfct_executar_backup_usuario")) Then
                                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                                    GoTo fim_lfct_executar_backup_usuario
                                End If
                                'atualizamos a barra de progresso
                                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
                            Next
                            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
                        End If
                    End If
                    
                    'atualizamos a barra de progresso geral
                    lsub_ajustar_barra_progresso_geral 0, 14, 17 'total de 17 passos
                    
                    'atualizamos a barra de progresso (zeramos pois os progressos individuais se encerraram)
                    lsub_ajustar_barra_progresso_individual 0, 0, 0
                    
                    'ajusta a data do último backup
                    p_backup.dt_ultimo_backup = Now
                    
                    'ajusta a data do próximo backup
                    Select Case p_backup.pb_periodo_backup
                        Case enm_periodo_backup.pb_diario
                            p_backup.dt_proximo_backup = DateAdd("d", 1, Now)
                        Case enm_periodo_backup.pb_semanal
                            p_backup.dt_proximo_backup = DateAdd("d", 7, Now)
                        Case enm_periodo_backup.pb_quinzenal
                            p_backup.dt_proximo_backup = DateAdd("d", 15, Now)
                        Case enm_periodo_backup.pb_mensal
                            p_backup.dt_proximo_backup = DateAdd("m", 1, Now)
                    End Select
                    
                    'ajusta o banco para backup
                    p_banco.tb_tipo_banco = tb_backup
                    
                    'mensagem
                    lsub_ajustar_status_operacao_mensagem "Salvando as configurações do backup..."
                    
                    'salva as configurações
                    If (Not pfct_salvar_configuracoes_usuario(plng_usuario)) Then
                        lsub_ajustar_status_operacao_mensagem "Erro ao salvar as configurações do backup. Operação cancelada."
                        GoTo fim_lfct_executar_backup_usuario
                    Else
                        lsub_ajustar_status_operacao_mensagem "Configurações salvas com sucesso."
                    End If
                    
                    'atualizamos a barra de progresso geral
                    lsub_ajustar_barra_progresso_geral 0, 15, 17 'total de 17 passos
                    
                    'limpeza da base de backup
                    psub_limpar_banco
                    
                    'exibe mensagem ao usuário
                    lsub_ajustar_status_operacao_mensagem "Movendo o arquivo gerado para a pasta de destino."
                    
                    'move o arquivo de backup gerado para a pasta de destino
                    If (Not pfct_mover_arquivo(p_banco.str_caminho_dados_backup, (p_backup.str_caminho & p_backup.str_nome))) Then
                        lsub_ajustar_status_operacao_mensagem "Erro ao mover o arquivo gerado para a pasta de destino. Operação cancelada."
                        GoTo fim_lfct_executar_backup_usuario
                    Else
                        lsub_ajustar_status_operacao_mensagem "Arquivo movido com sucesso."
                    End If
                    
                    'atualizamos a barra de progresso geral
                    lsub_ajustar_barra_progresso_geral 0, 16, 17 'total de 17 passos
                    
                    '
                    '----- fim dados -----'
                    
                    'ajusta o tipo de banco para config
                    p_banco.tb_tipo_banco = tb_config
                               
                    'mensagem
                    lsub_ajustar_status_operacao_mensagem "Salvando as configurações do usuário..."
                    
                    'salva as configurações
                    If (Not pfct_salvar_configuracoes_usuario(plng_usuario)) Then
                        lsub_ajustar_status_operacao_mensagem "Erro ao salvar as configurações do usuário. Operação cancelada."
                        GoTo fim_lfct_executar_backup_usuario
                    Else
                        lsub_ajustar_status_operacao_mensagem "Configurações salvas com sucesso."
                    End If
                    
                    'atualizamos a barra de progresso geral
                    lsub_ajustar_barra_progresso_geral 0, 17, 17 'total de 17 passos
                    
                    'limpeza da base de backup
                    psub_limpar_banco
                    
                    'atualizamos a barra de progresso geral (zeramos pois o progresso geral se encerrou)
                    lsub_ajustar_barra_progresso_geral 0, 0, 0
                    
                End If
            End If
        End If
    End If
    lfct_executar_backup_usuario = True
fim_lfct_executar_backup_usuario:
    Set lobj_backup = Nothing
    Set lobj_dados = Nothing
    Exit Function
erro_lfct_executar_backup_usuario:
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 0, 0
    'atualizamos a barra de progresso individual
    lsub_ajustar_barra_progresso_geral 0, 0, 0
    'continuamos o tratamento de erros
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_realizar", "lfct_executar_backup_usuario"
    GoTo fim_lfct_executar_backup_usuario
End Function

Private Sub Form_Activate()
    On Error GoTo Erro_Form_Activate
    
    'se o tipo de backup for automático
    If (mstr_tipo_backup = "A") Then
        Me.Caption = " Backup Automático "
        lsub_ajustar_status_operacao_mensagem "Tipo de backup: Automático."
    End If
    
    'se o tipo de backup for manual
    If (mstr_tipo_backup = "M") Then
        Me.Caption = " Backup Manual "
        lsub_ajustar_status_operacao_mensagem "Tipo de backup: Manual."
    End If
    
    'chama a função de backup
    lfct_executar_backup
    
    'se o tipo de backup for manual
    If (mstr_tipo_backup = "M") Then
        If (mbln_backup_concluido) Then
            lsub_ajustar_status_operacao_mensagem "Operação de backup manual completa com sucesso."
        End If
        'atribui true à variável modular
        mbln_backup_concluido = True
    End If
    
    'descarrega o formulário
    If (mstr_tipo_backup = "A") Then
        If (mbln_backup_concluido) Then
            lsub_ajustar_status_operacao_mensagem "Operação de backup automático completo com sucesso."
        End If
        'atribui true à variável modular
        mbln_backup_concluido = True
        'descarrega o formulário
        Unload Me
    End If
    
Fim_Form_Activate:
    Exit Sub
Erro_Form_Activate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_realizar", "Form_Activate"
    GoTo Fim_Form_Activate
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            If (mstr_tipo_backup = "A") Then
                psub_exibir_ajuda Me, "html/backup_automatico.htm", 0
            ElseIf (mstr_tipo_backup = "M") Then
                psub_exibir_ajuda Me, "html/backup_manual.htm", 0
            End If
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_realizar", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
    'armazena o código do usuário atual
    mlng_codigo_usuario_atual = p_usuario.lng_codigo
    'ajusta a grade de log
    lsub_ajustar_lista_log
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_realizar", "Form_Load"
    GoTo fim_Form_Load
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo erro_Form_QueryUnload
    Cancel = Not mbln_backup_concluido
fim_Form_QueryUnload:
    Exit Sub
erro_Form_QueryUnload:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_realizar", "Form_QueryUnload"
    GoTo fim_Form_QueryUnload
End Sub
