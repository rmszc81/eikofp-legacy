VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_backup_restaurar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Restaurar Backup "
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   210
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
   ScaleHeight     =   855
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlg_restaurar_backup 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_restaurar 
      Caption         =   "&Restaurar (F2)"
      Height          =   375
      Left            =   5220
      TabIndex        =   9
      Top             =   1740
      Width           =   1275
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar (F3)"
      Height          =   375
      Left            =   6540
      TabIndex        =   10
      Top             =   1740
      Width           =   1215
   End
   Begin VB.TextBox txt_data_backup 
      Height          =   315
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1260
      Width           =   2595
   End
   Begin VB.TextBox txt_senha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5220
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1260
      Width           =   2535
   End
   Begin VB.TextBox txt_usuario 
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1260
      Width           =   2295
   End
   Begin VB.TextBox txt_backup_restaurar 
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   7095
   End
   Begin VB.CommandButton cmd_selecionar 
      Caption         =   "..."
      Height          =   315
      Left            =   7320
      TabIndex        =   2
      Top             =   420
      Width           =   435
   End
   Begin MSComctlLib.ListView lst_log_operacoes 
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   4048
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
   Begin MSComctlLib.ProgressBar pb_progresso_geral 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   5040
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pb_progresso_individual 
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   5400
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
      TabIndex        =   11
      Top             =   2340
      Width           =   1335
   End
   Begin VB.Label lbl_data_backup 
      AutoSize        =   -1  'True
      Caption         =   "&Data do backup selecionado:"
      Height          =   195
      Left            =   2520
      TabIndex        =   4
      Top             =   960
      Width           =   2070
   End
   Begin VB.Label lbl_senha 
      AutoSize        =   -1  'True
      Caption         =   "&Senha:"
      Height          =   195
      Left            =   5220
      TabIndex        =   5
      Top             =   960
      Width           =   510
   End
   Begin VB.Label lbl_usuario 
      AutoSize        =   -1  'True
      Caption         =   "&Usuário:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   600
   End
   Begin VB.Label lbl_backup_restaurar 
      AutoSize        =   -1  'True
      Caption         =   "Backup a &restaurar:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1425
   End
End
Attribute VB_Name = "frm_backup_restaurar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'armazena a senha do usuário do backup selecionado
Private mstr_senha As String

'controle para o fechamento do form
Private mbln_pode_fechar As Boolean

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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "lsub_ajustar_barra_progresso_geral"
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "lsub_ajustar_barra_progresso_individual"
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "lsub_ajustar_lista_log"
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "lsub_ajustar_status_operacao_mensagem"
    GoTo fim_lsub_ajustar_status_operacao_mensagem
End Sub

Private Sub lsub_configura_dialog()
    On Error GoTo erro_lsub_configura_dialog
    With dlg_restaurar_backup
        .CancelError = False
        .DialogTitle = pcst_nome_aplicacao & " - restaurar backup "
        'parâmetros
        .Flags = cdlOFNFileMustExist + _
                 cdlOFNExplorer + _
                 cdlOFNLongNames + _
                 cdlOFNHideReadOnly
        'filtro de arquivo
        .Filter = "Backup EikoFP (*.dbbkp)|*.dbbkp"
        .FilterIndex = 1
        If (p_backup.str_caminho <> "") Then
            .InitDir = p_backup.str_caminho
        Else
            .InitDir = p_banco.str_caminho_comum
        End If
    End With
fim_lsub_configura_dialog:
    Exit Sub
erro_lsub_configura_dialog:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "lsub_configura_dialog"
    GoTo fim_lsub_configura_dialog
End Sub

Private Function lfct_limpa_bancos() As Boolean
    On Error GoTo erro_lfct_limpa_bancos
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Realizando manutenção nas bases de dados."
    lsub_ajustar_status_operacao_mensagem "Por favor, aguarde..."
    'limpa o banco config
    p_banco.tb_tipo_banco = tb_config
    psub_limpar_banco
    'limpa o banco dados
    p_banco.tb_tipo_banco = tb_dados
    psub_limpar_banco
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Manutenção realizada com sucesso."
    'retorna true
    lfct_limpa_bancos = True
fim_lfct_limpa_bancos:
    Exit Function
erro_lfct_limpa_bancos:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "lfct_limpa_bancos"
    GoTo fim_lfct_limpa_bancos
End Function

Private Function lfct_validar_senha() As Boolean
    On Error GoTo erro_lfct_validar_senha
    'se está em branco
    If (txt_senha.Text = "") Then
        'exibe mensagem
        MsgBox "Atenção!" & vbCrLf & "Digite a senha do usuário.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        'ajusta o foco no campo
        txt_senha.SetFocus
        'desvia ao fim do método
        GoTo fim_lfct_validar_senha
    End If
    'se tem menos que 4 caracteres
    If (Len(txt_senha.Text) < 4) Then
        'exibe mensagem
        MsgBox "Atenção!" & vbCrLf & "Campo [senha] deve conter no mínimo 04 (quatro) caracteres.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        'limpa o campo
        txt_senha.Text = ""
        'ajusta o foco no campo
        txt_senha.SetFocus
        'desvia ao fim do método
        GoTo fim_lfct_validar_senha
    End If
    'se a senha digitada é a mesma do backup
    If (pfct_criptografia(txt_senha) <> mstr_senha) Then
        'exibe mensagem
        MsgBox "Atenção!" & vbCrLf & "Senha inválida.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        'limpa o campo
        txt_senha.Text = ""
        'ajusta o foco no campo
        txt_senha.SetFocus
        'desvia ao fim do método
        GoTo fim_lfct_validar_senha
    End If
    lfct_validar_senha = True
fim_lfct_validar_senha:
    Exit Function
erro_lfct_validar_senha:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "lfct_validar_senha"
    GoTo fim_lfct_validar_senha
End Function

Private Function lfct_restaurar_backup() As Boolean
    On Error GoTo erro_lfct_restaurar_backup
    Dim lobj_restaurar As Object
    Dim lobj_dados As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    
    'atualiza o form
    Me.Refresh
    
    'ajusta a mensagem inicial
    lsub_ajustar_status_operacao_mensagem "Iniciando restauração do backup..."
    
    'força a aplicação a reprocessar as mensagens
    DoEvents

    'aguarda 2 segundos antes de iniciar...
    Sleep (2000)
    
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Movendo o arquivo para a pasta de destino."
    
    'movemos o arquivo de origem para a pasta de destino
    If (Not pfct_mover_arquivo(p_banco.str_caminho_dados_restaurar, (p_banco.str_caminho_backup & pfct_retorna_nome_arquivo(p_banco.str_caminho_dados_restaurar)))) Then
        lsub_ajustar_status_operacao_mensagem "Erro ao mover o arquivo de origem para a pasta de destino. Operação cancelada."
        GoTo fim_lfct_restaurar_backup
    Else
        lsub_ajustar_status_operacao_mensagem "Arquivo movido com sucesso."
    End If
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 1, 21 'total de 21 passos
    
    'ajustamos o nome do arquivo de origem do backup
    p_backup.str_nome = pfct_retorna_nome_arquivo(p_banco.str_caminho_dados_restaurar)
    
    'ajusta o banco para restaurar
    p_banco.tb_tipo_banco = tb_restaurar
    
    'configura o banco de restauração
    pfct_ajustar_caminho_banco tb_restaurar
    
    'ajusta o banco para config
    p_banco.tb_tipo_banco = tb_config
    
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Limpando as tabelas da base de dados atual..."
    
    ' -- ini config -- '
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de configuração..."
    'monta o comando sql (contagem de registros)
    lstr_sql = "select * from [tb_config] where [int_usuario] = " & pfct_tratar_numero_sql(p_usuario.lng_codigo)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        'exibe mensagem
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de configuração. Operação cancelada."
        'desvia ao fim do método
        GoTo fim_lfct_restaurar_backup
    Else
        'quantidade de registros
        llng_registros = lobj_dados.Count
        If (llng_registros = 0) Then
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        Else
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Excluindo dados, aguarde..."
            'monta o comando sql (exclusão de registros)
            lstr_sql = "delete from [tb_config] where [int_usuario] = " & pfct_tratar_numero_sql(p_usuario.lng_codigo)
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Erro na exclusão dos dados. Operação cancelada."
                'desvia ao fim do método
                GoTo fim_lfct_restaurar_backup
            Else
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Dados excluídos com sucesso."
            End If
        End If
    End If
    ' -- fim config -- '
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 2, 21 'total de 21 passos
    
    ' -- ini backup -- '
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de backup..."
    'monta o comando sql (contagem de registros)
    lstr_sql = "select * from [tb_backup] where [int_usuario] = " & pfct_tratar_numero_sql(p_usuario.lng_codigo)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        'exibe mensagem
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de backup. Operação cancelada."
        'desvia ao fim do método
        GoTo fim_lfct_restaurar_backup
    Else
        'quantidade de registros
        llng_registros = lobj_dados.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        Else
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Excluindo dados, aguarde..."
            'monta o comando sql (exclusão de registros)
            lstr_sql = "delete from [tb_backup] where [int_usuario] = " & pfct_tratar_numero_sql(p_usuario.lng_codigo)
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Erro na exclusão dos dados. Operação cancelada."
                'desvia ao fim do método
                GoTo fim_lfct_restaurar_backup
            Else
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Dados excluídos com sucesso."
            End If
        End If
    End If
    ' -- fim backup -- '
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 3, 21 'total de 21 passos
    
    'ajusta o banco para dados
    p_banco.tb_tipo_banco = tb_dados
    
    ' -- ini contas -- '
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de contas..."
    'monta o comando sql (contagem de registros)
    lstr_sql = "select * from [tb_contas]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        'exibe mensagem
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de contas. Operação cancelada."
        'desvia ao fim do método
        GoTo fim_lfct_restaurar_backup
    Else
        'quantidade de registros
        llng_registros = lobj_dados.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        Else
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Excluindo dados, aguarde..."
            'monta o comando sql (exclusão de registros)
            lstr_sql = "delete from [tb_contas]"
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Erro na exclusão dos dados. Operação cancelada."
                'desvia ao fim do método
                GoTo fim_lfct_restaurar_backup
            Else
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Dados excluídos com sucesso."
            End If
        End If
    End If
    ' -- fim contas -- '
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 4, 21 'total de 21 passos

    ' -- ini despesas -- '
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de despesas..."
    'monta o comando sql (contagem de registros)
    lstr_sql = "select * from [tb_despesas]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        'exibe mensagem
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de despesas. Operação cancelada."
        'desvia ao fim do método
        GoTo fim_lfct_restaurar_backup
    Else
        'quantidade de registros
        llng_registros = lobj_dados.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        Else
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Excluindo dados, aguarde..."
            'monta o comando sql (exclusão de registros)
            lstr_sql = "delete from [tb_despesas]"
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Erro na exclusão dos dados. Operação cancelada."
                'desvia ao fim do método
                GoTo fim_lfct_restaurar_backup
            Else
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Dados excluídos com sucesso."
            End If
        End If
    End If
    ' -- fim despesas -- '
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 5, 21 'total de 21 passos
    
    ' -- ini receitas -- '
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de receitas..."
    'monta o comando sql (contagem de registros)
    lstr_sql = "select * from [tb_receitas]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        'exibe mensagem
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de receitas. Operação cancelada."
        'desvia ao fim do método
        GoTo fim_lfct_restaurar_backup
    Else
        'quantidade de registros
        llng_registros = lobj_dados.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        Else
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Excluindo dados, aguarde..."
            'monta o comando sql (exclusão de registros)
            lstr_sql = "delete from [tb_receitas]"
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Erro na exclusão dos dados. Operação cancelada."
                'desvia ao fim do método
                GoTo fim_lfct_restaurar_backup
            Else
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Dados excluídos com sucesso."
            End If
        End If
    End If
    ' -- fim receitas -- '
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 6, 21 'total de 21 passos
    
    ' -- ini formas pagamento -- '
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de formas de pagamento..."
    'monta o comando sql (contagem de registros)
    lstr_sql = "select * from [tb_formas_pagamento]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        'exibe mensagem
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de formas de pagamento. Operação cancelada."
        'desvia ao fim do método
        GoTo fim_lfct_restaurar_backup
    Else
        'quantidade de registros
        llng_registros = lobj_dados.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        Else
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Excluindo dados, aguarde..."
            'monta o comando sql (exclusão de registros)
            lstr_sql = "delete from [tb_formas_pagamento]"
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Erro na exclusão dos dados. Operação cancelada."
                'desvia ao fim do método
                GoTo fim_lfct_restaurar_backup
            Else
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Dados excluídos com sucesso."
            End If
        End If
    End If
    ' -- fim formas pagamento -- '
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 7, 21 'total de 21 passos
    
    ' -- ini formas pagamento -- '
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de formas de pagamento..."
    'monta o comando sql (contagem de registros)
    lstr_sql = "select * from [tb_formas_pagamento]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        'exibe mensagem
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de formas de pagamento. Operação cancelada."
        'desvia ao fim do método
        GoTo fim_lfct_restaurar_backup
    Else
        'quantidade de registros
        llng_registros = lobj_dados.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        Else
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Excluindo dados, aguarde..."
            'monta o comando sql (exclusão de registros)
            lstr_sql = "delete from [tb_formas_pagamento]"
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Erro na exclusão dos dados. Operação cancelada."
                'desvia ao fim do método
                GoTo fim_lfct_restaurar_backup
            Else
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Dados excluídos com sucesso."
            End If
        End If
    End If
    ' -- fim formas pagamento -- '
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 8, 21 'total de 21 passos

    ' -- ini contas pagar -- '
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de formas de contas a pagar..."
    'monta o comando sql (contagem de registros)
    lstr_sql = "select * from [tb_contas_pagar]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        'exibe mensagem
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de formas de contas a pagar. Operação cancelada."
        'desvia ao fim do método
        GoTo fim_lfct_restaurar_backup
    Else
        'quantidade de registros
        llng_registros = lobj_dados.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        Else
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Excluindo dados, aguarde..."
            'monta o comando sql (exclusão de registros)
            lstr_sql = "delete from [tb_contas_pagar]"
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Erro na exclusão dos dados. Operação cancelada."
                'desvia ao fim do método
                GoTo fim_lfct_restaurar_backup
            Else
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Dados excluídos com sucesso."
            End If
        End If
    End If
    ' -- fim contas pagar -- '
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 9, 21 'total de 21 passos

    ' -- ini contas receber -- '
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de formas de contas a receber..."
    'monta o comando sql (contagem de registros)
    lstr_sql = "select * from [tb_contas_receber]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        'exibe mensagem
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de formas de contas a receber. Operação cancelada."
        'desvia ao fim do método
        GoTo fim_lfct_restaurar_backup
    Else
        'quantidade de registros
        llng_registros = lobj_dados.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        Else
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Excluindo dados, aguarde..."
            'monta o comando sql (exclusão de registros)
            lstr_sql = "delete from [tb_contas_receber]"
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Erro na exclusão dos dados. Operação cancelada."
                'desvia ao fim do método
                GoTo fim_lfct_restaurar_backup
            Else
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Dados excluídos com sucesso."
            End If
        End If
    End If
    ' -- fim contas receber -- '
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 10, 21 'total de 21 passos
    
    ' -- ini movimentação -- '
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de movimentação..."
    'monta o comando sql (contagem de registros)
    lstr_sql = "select * from [tb_movimentacao]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        'exibe mensagem
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela de movimentação. Operação cancelada."
        'desvia ao fim do método
        GoTo fim_lfct_restaurar_backup
    Else
        'quantidade de registros
        llng_registros = lobj_dados.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        Else
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Excluindo dados, aguarde..."
            'monta o comando sql (exclusão de registros)
            lstr_sql = "delete from [tb_movimentacao]"
            'executa o comando sql e devolve o objeto
            If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Erro na exclusão dos dados. Operação cancelada."
                'desvia ao fim do método
                GoTo fim_lfct_restaurar_backup
            Else
                'exibe mensagem
                lsub_ajustar_status_operacao_mensagem "Dados excluídos com sucesso."
            End If
        End If
    End If
    ' -- fim movimentação -- '
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 11, 21 'total de 21 passos

    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Limpeza das tabelas concluída. Iniciando cópia dos dados."
    lsub_ajustar_status_operacao_mensagem "Este processo pode levar alguns minutos."
    lsub_ajustar_status_operacao_mensagem "Por favor, aguarde..."
    
    'aguarda 5 segundos antes de continuar...
    Sleep (5000)
    
    '-- ini config --'
    'ajusta o banco para restaurar
    p_banco.tb_tipo_banco = tb_restaurar
    pfct_ajustar_caminho_banco tb_restaurar
    
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de configuração..."
    'monta o comando sql
    lstr_sql = "select * from [tb_config]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_restaurar, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela. Operação cancelada."
        GoTo fim_lfct_restaurar_backup
    Else
        llng_registros = lobj_restaurar.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        ElseIf (llng_registros > 0) Then
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
            'ajusta o banco para config
            p_banco.tb_tipo_banco = tb_config
            'percorre o objeto
            For llng_contador = 1 To llng_registros
                'processa as mensagens do windows
                DoEvents
                'monta o comando sql
                lstr_sql = ""
                lstr_sql = lstr_sql & " insert into [tb_config] "
                lstr_sql = lstr_sql & " ( "
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
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(p_usuario.lng_codigo) & ", " 'código do usuário logado
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_moeda")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_intervalo_data")) & ", "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_carregar_agenda_financeira_login")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_lancamentos_retroativos")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_alteracoes_detalhes")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_data_vencimento_baixa_imediata")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_lancamentos_duplicados")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_participou_pesquisa")) & "' "
                lstr_sql = lstr_sql & " ) "
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                    GoTo fim_lfct_restaurar_backup
                End If
                'atualizamos a barra de progresso
                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
            Next
            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
        End If
    End If
    '-- fim config --'
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 12, 21 'total de 21 passos
    
    '-- ini backup --'
    'ajusta o banco para restaurar
    p_banco.tb_tipo_banco = tb_restaurar
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de backup..."
    'monta o comando sql
    lstr_sql = "select * from [tb_backup]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_restaurar, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela. Operação cancelada."
        GoTo fim_lfct_restaurar_backup
    Else
        llng_registros = lobj_restaurar.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        ElseIf (llng_registros > 0) Then
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
            'ajusta o banco para config
            p_banco.tb_tipo_banco = tb_config
            'percorre o objeto
            For llng_contador = 1 To llng_registros
                'processa as mensagens do windows
                DoEvents
                'monta o comando sql
                lstr_sql = ""
                lstr_sql = lstr_sql & " insert into [tb_backup] "
                lstr_sql = lstr_sql & " ( "
                lstr_sql = lstr_sql & " [int_usuario], "
                lstr_sql = lstr_sql & " [chr_ativar], "
                lstr_sql = lstr_sql & " [int_periodo], "
                lstr_sql = lstr_sql & " [str_caminho], "
                lstr_sql = lstr_sql & " [dt_ultimo_backup], "
                lstr_sql = lstr_sql & " [dt_proximo_backup] "
                lstr_sql = lstr_sql & " ) "
                lstr_sql = lstr_sql & " values "
                lstr_sql = lstr_sql & " ( "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(p_usuario.lng_codigo) & ", " 'código do usuário logado
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_ativar")) & "', "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_periodo")) & ", "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_caminho")) & "', "
                lstr_sql = lstr_sql & " '" & lobj_restaurar(llng_contador)("dt_ultimo_backup") & "', "
                lstr_sql = lstr_sql & " '" & lobj_restaurar(llng_contador)("dt_proximo_backup") & "' "
                lstr_sql = lstr_sql & " ) "
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                    GoTo fim_lfct_restaurar_backup
                End If
                'atualizamos a barra de progresso
                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
            Next
            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
        End If
    End If
    '-- fim backup --'
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 13, 21 'total de 21 passos
    
    '-- ini contas --'
    'ajusta o banco para restaurar
    p_banco.tb_tipo_banco = tb_restaurar
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de contas..."
    'monta o comando sql
    lstr_sql = "select * from [tb_contas]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_restaurar, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela. Operação cancelada."
        GoTo fim_lfct_restaurar_backup
    Else
        llng_registros = lobj_restaurar.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        ElseIf (llng_registros > 0) Then
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
            'ajusta o banco para dados
            p_banco.tb_tipo_banco = tb_dados
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
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_codigo")) & ", "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_descricao")) & "', "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("num_saldo")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("num_limite_negativo")) & ", "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_observacoes")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_ativo")) & "' "
                lstr_sql = lstr_sql & " ) "
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                    GoTo fim_lfct_restaurar_backup
                End If
                'atualizamos a barra de progresso
                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
            Next
            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
        End If
    End If
    '-- fim contas --'
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 14, 21 'total de 21 passos
    
    '-- ini despesas --'
    'ajusta o banco para restaurar
    p_banco.tb_tipo_banco = tb_restaurar
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de despesas..."
    'monta o comando sql
    lstr_sql = "select * from [tb_despesas]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_restaurar, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela. Operação cancelada."
        GoTo fim_lfct_restaurar_backup
    Else
        llng_registros = lobj_restaurar.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        ElseIf (llng_registros > 0) Then
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
            'ajusta o banco para dados
            p_banco.tb_tipo_banco = tb_dados
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
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_codigo")) & ", "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_descricao")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_observacoes")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_fixa")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_ativo")) & "' "
                lstr_sql = lstr_sql & " ) "
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                    GoTo fim_lfct_restaurar_backup
                End If
                'atualizamos a barra de progresso
                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
            Next
            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
        End If
    End If
    '-- fim despesas --'

    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 15, 21 'total de 21 passos

    '-- ini receitas --'
    'ajusta o banco para restaurar
    p_banco.tb_tipo_banco = tb_restaurar
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de receitas..."
    'monta o comando sql
    lstr_sql = "select * from [tb_receitas]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_restaurar, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela. Operação cancelada."
        GoTo fim_lfct_restaurar_backup
    Else
        llng_registros = lobj_restaurar.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        ElseIf (llng_registros > 0) Then
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
            'ajusta o banco para dados
            p_banco.tb_tipo_banco = tb_dados
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
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_codigo")) & ", "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_descricao")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_observacoes")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_fixa")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_ativo")) & "' "
                lstr_sql = lstr_sql & " ) "
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                    GoTo fim_lfct_restaurar_backup
                End If
                'atualizamos a barra de progresso
                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
            Next
            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
        End If
    End If
    '-- fim receitas --'

    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 16, 21 'total de 21 passos

    '-- ini formas pagamento --'
    'ajusta o banco para restaurar
    p_banco.tb_tipo_banco = tb_restaurar
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de formas de pagamento..."
    'monta o comando sql
    lstr_sql = "select * from [tb_formas_pagamento]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_restaurar, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela. Operação cancelada."
        GoTo fim_lfct_restaurar_backup
    Else
        llng_registros = lobj_restaurar.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        ElseIf (llng_registros > 0) Then
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
            'ajusta o banco para dados
            p_banco.tb_tipo_banco = tb_dados
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
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_codigo")) & ", "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_descricao")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_observacoes")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_ativo")) & "' "
                lstr_sql = lstr_sql & " ) "
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                    GoTo fim_lfct_restaurar_backup
                End If
                'atualizamos a barra de progresso
                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
            Next
            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
        End If
    End If
    '-- fim formas pagamento --'
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 17, 21 'total de 21 passos
    
    '-- ini contas pagar --'
    'ajusta o banco para restaurar
    p_banco.tb_tipo_banco = tb_restaurar
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de contas a pagar..."
    'monta o comando sql
    lstr_sql = "select * from [tb_contas_pagar]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_restaurar, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela. Operação cancelada."
        GoTo fim_lfct_restaurar_backup
    Else
        llng_registros = lobj_restaurar.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        ElseIf (llng_registros > 0) Then
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
            'ajusta o banco para dados
            p_banco.tb_tipo_banco = tb_dados
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
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_codigo")) & ", "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_baixa_automatica")) & "', "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_conta_baixa_automatica")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_despesa")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_forma_pagamento")) & ", "
                lstr_sql = lstr_sql & " '" & lobj_restaurar(llng_contador)("dt_vencimento") & "', "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_parcela")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_total_parcelas")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("num_valor")) & ", "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_descricao")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_documento")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_chave")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_codigo_barras")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_observacoes")) & "' "
                lstr_sql = lstr_sql & " ) "
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                    GoTo fim_lfct_restaurar_backup
                End If
                'atualizamos a barra de progresso
                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
            Next
            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
        End If
    End If
    '-- fim contas pagar --'

    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 18, 21 'total de 21 passos

    '-- ini contas receber --'
    'ajusta o banco para restaurar
    p_banco.tb_tipo_banco = tb_restaurar
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de contas a receber..."
    'monta o comando sql
    lstr_sql = "select * from [tb_contas_receber]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_restaurar, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela. Operação cancelada."
        GoTo fim_lfct_restaurar_backup
    Else
        llng_registros = lobj_restaurar.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        ElseIf (llng_registros > 0) Then
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
            'ajusta o banco para dados
            p_banco.tb_tipo_banco = tb_dados
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
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_codigo")) & ", "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_baixa_automatica")) & "', "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_conta_baixa_automatica")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_receita")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_forma_pagamento")) & ", "
                lstr_sql = lstr_sql & " '" & lobj_restaurar(llng_contador)("dt_vencimento") & "', "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_parcela")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_total_parcelas")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("num_valor")) & ", "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_descricao")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_documento")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_chave")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_observacoes")) & "' "
                lstr_sql = lstr_sql & " ) "
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                    GoTo fim_lfct_restaurar_backup
                End If
                'atualizamos a barra de progresso
                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
            Next
            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
        End If
    End If
    '-- fim contas receber --'

    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 19, 21 'total de 21 passos

    '-- ini movimentação --'
    'ajusta o banco para restaurar
    p_banco.tb_tipo_banco = tb_restaurar
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Lendo a tabela de movimentação..."
    'monta o comando sql
    lstr_sql = "select * from [tb_movimentacao]"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_restaurar, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
        lsub_ajustar_status_operacao_mensagem "Erro na leitura da tabela. Operação cancelada."
        GoTo fim_lfct_restaurar_backup
    Else
        llng_registros = lobj_restaurar.Count
        If (llng_registros = 0) Then
            lsub_ajustar_status_operacao_mensagem "Não foram localizados registros."
        ElseIf (llng_registros > 0) Then
            'exibe mensagem
            lsub_ajustar_status_operacao_mensagem "Foram localizados " & CStr(llng_registros) & " registros."
            lsub_ajustar_status_operacao_mensagem "Copiando dados, aguarde..."
            'ajusta o banco para dados
            p_banco.tb_tipo_banco = tb_dados
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
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_codigo")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_conta")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_receita")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_despesa")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_forma_pagamento")) & ", "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("chr_tipo")) & "', "
                lstr_sql = lstr_sql & " '" & lobj_restaurar(llng_contador)("dt_vencimento") & "', "
                lstr_sql = lstr_sql & " '" & lobj_restaurar(llng_contador)("dt_pagamento") & "', "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_parcela")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("int_total_parcelas")) & ", "
                lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(lobj_restaurar(llng_contador)("num_valor")) & ", "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_descricao")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_documento")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_codigo_barras")) & "', "
                lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(lobj_restaurar(llng_contador)("str_observacoes")) & "' "
                lstr_sql = lstr_sql & " ) "
                'executa o comando sql e devolve o objeto
                If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_restaurar_backup")) Then
                    lsub_ajustar_status_operacao_mensagem "Erro na gravação dos dados. Operação cancelada."
                    GoTo fim_lfct_restaurar_backup
                End If
                'atualizamos a barra de progresso
                lsub_ajustar_barra_progresso_individual 0, llng_contador, llng_registros
            Next
            lsub_ajustar_status_operacao_mensagem "Dados copiados com sucesso."
        End If
    End If
    '-- fim movimentação --'
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 20, 21 'total de 21 passos
    
    'atualizamos a barra de progresso (zeramos pois os progressos individuais se encerraram)
    lsub_ajustar_barra_progresso_individual 0, 0, 0
                    
    'exibe mensagem
    lsub_ajustar_status_operacao_mensagem "Excluindo arquivo de backup."
    
    'excluímos o arquivo de backup
    If Not (pfct_excluir_arquivo(p_banco.str_caminho_dados_restaurar)) Then
        lsub_ajustar_status_operacao_mensagem "Erro ao excluir o arquivo de backup. Operação cancelada."
        GoTo fim_lfct_restaurar_backup
    Else
        lsub_ajustar_status_operacao_mensagem "Arquivo excluído com sucesso."
    End If
    
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 21, 21 'total de 21 passos
    
    'atualizamos a barra de progresso geral (zeramos pois o progresso geral se encerrou)
    lsub_ajustar_barra_progresso_geral 0, 0, 0

    'retorna true
    lfct_restaurar_backup = True
fim_lfct_restaurar_backup:
    Set lobj_restaurar = Nothing
    Set lobj_dados = Nothing
    Exit Function
erro_lfct_restaurar_backup:
    'atualizamos a barra de progresso geral
    lsub_ajustar_barra_progresso_geral 0, 0, 0
    'atualizamos a barra de progresso individual
    lsub_ajustar_barra_progresso_geral 0, 0, 0
    'continuamos o tratamento de erros
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "lfct_restaurar_backup"
    GoTo fim_lfct_restaurar_backup
    Resume 0
End Function

Private Function lfct_verificar_backup() As Boolean
    On Error GoTo erro_lfct_verificar_backup
    Dim lobj_dados As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select "
    lstr_sql = lstr_sql & " [tb_usuarios].[int_codigo], "
    lstr_sql = lstr_sql & " [tb_usuarios].[str_usuario], "
    lstr_sql = lstr_sql & " [tb_usuarios].[str_senha], "
    lstr_sql = lstr_sql & " [tb_usuarios].[str_lembrete_senha], "
    lstr_sql = lstr_sql & " [tb_backup].[dt_ultimo_backup], "
    lstr_sql = lstr_sql & " [tb_backup].[tm_ultimo_backup] "
    lstr_sql = lstr_sql & " from "
    lstr_sql = lstr_sql & " [tb_usuarios] "
    lstr_sql = lstr_sql & " inner join "
    lstr_sql = lstr_sql & " [tb_backup] "
    lstr_sql = lstr_sql & " on [tb_backup].[int_usuario] = [tb_usuarios].[int_codigo] "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "frm_backup_restaurar", "lfct_verificar_backup")) Then
        'mensagem diferenciada para este processo
        MsgBox "Atenção!" & vbCrLf & "Arquivo de backup selecionado inválido.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        'ajusta o foco no campo
        txt_backup_restaurar.SetFocus
        'desvia ao fim do método
        GoTo fim_lfct_verificar_backup
    Else
        'quantidade de registros
        llng_registros = lobj_dados.Count
        If (llng_registros = 1) Then
            'verifica se o usuário do banco selecionado
            'é igual ao usuário atual
            If (lobj_dados(1)("str_usuario") = p_usuario.str_login) Then
                'atribui os dados nos campos
                txt_usuario.Text = lobj_dados(1)("str_usuario")
                txt_data_backup.Text = Format$(lobj_dados(1)("dt_ultimo_backup"), "dd/mm/yyyy") & " " & Format$(lobj_dados(1)("tm_ultimo_backup"), "hh:mm:ss")
                mstr_senha = lobj_dados(1)("str_senha")
                lfct_verificar_backup = True
            Else
                MsgBox "Atenção!" & vbCrLf & "Não é possível restaurar backup de outro usuário.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                GoTo fim_lfct_verificar_backup
            End If
        Else
            MsgBox "Atenção!" & vbCrLf & "Arquivo de backup selecionado inválido.", vbOKOnly + vbInformation, pcst_nome_aplicacao
            GoTo fim_lfct_verificar_backup
        End If
    End If
    'retorna true
    lfct_verificar_backup = True
fim_lfct_verificar_backup:
    Set lobj_dados = Nothing
    Exit Function
erro_lfct_verificar_backup:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_verificar", "lfct_verificar_backup"
    GoTo fim_lfct_verificar_backup
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
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "cmd_cancelar_Click"
    GoTo fim_cmd_cancelar_Click
End Sub

Private Sub cmd_restaurar_Click()
    On Error GoTo erro_cmd_restaurar_Click
    Dim lint_resposta As Integer
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_restaurar.Enabled) Then
        Exit Sub
    End If
    'valida a senha do usuário
    If (lfct_validar_senha) Then
        lint_resposta = MsgBox("Atenção!" & vbCrLf & _
                               "Todos os dados da base atual serão substituídos pelo conteúdo do backup." & vbCrLf & _
                               "Deseja continuar?", vbYesNo + vbQuestion + vbDefaultButton2, pcst_nome_aplicacao)
        If (lint_resposta = vbYes) Then
            'desabilita os campos
            lbl_senha.Enabled = False
            txt_senha.Enabled = False
            'command button
            cmd_restaurar.Enabled = False
            cmd_cancelar.Enabled = False
            'esconde o form
            Me.Hide
            'ajusta a altura da janela
            Me.Height = 6135
            'mostra o form
            Me.Show
            'ajusta valor variável
            mbln_pode_fechar = False
            'inicia o processo de restauração
            If (lfct_restaurar_backup()) Then
                If (Not lfct_limpa_bancos) Then
                    MsgBox "Atenção!" & vbCrLf & "Erro ao realizar a manutenção da base de dados.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                    GoTo fim_cmd_restaurar_Click
                Else
                    'ajusta o banco para config
                    p_banco.tb_tipo_banco = tb_config
                    'recarrega os dados do usuário
                    If (pfct_carregar_dados_usuario(p_usuario.lng_codigo)) Then
                        'recarrega as configurações do usuário
                        If (Not pfct_carregar_configuracoes_usuario(p_usuario.lng_codigo)) Then
                            MsgBox "Atenção!" & vbCrLf & "Erro ao recarregar os dados do usuário.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                            GoTo fim_cmd_restaurar_Click
                        End If
                    Else
                        MsgBox "Atenção!" & vbCrLf & "Erro ao recarregar os dados do usuário.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                        GoTo fim_cmd_restaurar_Click
                    End If
                    'exibe mensagem
                    lsub_ajustar_status_operacao_mensagem "Restauração de backup concluído com sucesso."
                    'ajusta o banco para dados
                    p_banco.tb_tipo_banco = tb_dados
                    'ajusta valor variável
                    mbln_pode_fechar = True
                    'exibe mensagem
                    MsgBox "Atenção!" & vbCrLf & "Backup restaurado com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
                    'descarrega o form
                    Unload Me
                    'desvia ao fim do método
                    'GoTo fim_cmd_restaurar_Click
                End If
            Else
                'ajusta valor variável
                mbln_pode_fechar = True
                'exibe mensagem
                MsgBox "Atenção!" & vbCrLf & "Erro ao restaurar backup.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                'desvia ao fim do método
                GoTo fim_cmd_restaurar_Click
            End If
        End If
    Else
        GoTo fim_cmd_restaurar_Click
    End If
fim_cmd_restaurar_Click:
    Exit Sub
erro_cmd_restaurar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "cmd_restaurar_Click"
    GoTo fim_cmd_restaurar_Click
End Sub

Private Sub cmd_selecionar_Click()
    On Error GoTo erro_cmd_selecionar_Click
    Dim lint_resposta As Integer
    Dim lstr_arquivo As String
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_selecionar.Enabled) Then
        Exit Sub
    End If
    'verifica se a aplicação está sendo executada
    'com privilégios de administrador
    If (Not pfct_verificar_administrador()) Then
        'caso não esteja, perguntamos ao usuário se quer continuar
        lint_resposta = MsgBox("A aplicação está sendo executada sem privilégios de administrador e podem ocorrer erros ao tentar restaurar o backup." & vbCrLf & "Deseja continuar?", vbYesNo + vbQuestion + vbDefaultButton2, pcst_nome_aplicacao)
        'se a resposta for não
        If (lint_resposta = vbNo) Then
            'desvia ao fim do método
            GoTo fim_cmd_selecionar_Click
        End If
    End If
    lsub_configura_dialog
    With dlg_restaurar_backup
        .ShowOpen
        If (.FileName <> "") Then
            'atribui nome do arquivo selecionado à variável
            lstr_arquivo = .FileName
            'verificamos se podemos abrir o arquivo
            If (pfct_pode_abrir_arquivo(lstr_arquivo)) Then
                'atribui na caixa de texto o nome do arquivo
                txt_backup_restaurar.Text = lstr_arquivo
                'ajusta o tipo de banco
                p_banco.tb_tipo_banco = tb_restaurar
                'configura o caminho do arquivo de restauração
                p_banco.str_caminho_dados_restaurar = lstr_arquivo
                'chama a função de verificação do banco
                If (lfct_verificar_backup) Then
                    'desabilita os campos
                    'label
                    lbl_backup_restaurar.Enabled = False
                    lbl_usuario.Enabled = False
                    lbl_data_backup.Enabled = False
                    'textbox
                    txt_backup_restaurar.Enabled = False
                    txt_usuario.Enabled = False
                    txt_data_backup.Enabled = False
                    'command button
                    cmd_selecionar.Enabled = False
                    'aumenta a altura da janela
                    Me.Height = 2610
                    'limpa o campo
                    txt_senha.Text = ""
                    'ajusta o foco no campo
                    txt_senha.SetFocus
                Else
                    'limpa a caixa de texto
                    txt_backup_restaurar.Text = ""
                    'desvia ao fim do método
                    GoTo fim_cmd_selecionar_Click
                End If
            Else
                'exibe mensagem ao usuário
                MsgBox "Não é possível restaurar o backup." & vbCrLf & "O arquivo informado não está acessível para leitura.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                'desvia ao fim do método
                GoTo fim_cmd_selecionar_Click
            End If
        Else
            GoTo fim_cmd_selecionar_Click
        End If
    End With
fim_cmd_selecionar_Click:
    Exit Sub
erro_cmd_selecionar_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "cmd_selecionar_Click"
    GoTo fim_cmd_selecionar_Click
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/backup_restore.htm", 0
        Case vbKeyF2
            cmd_restaurar_Click
        Case vbKeyF3
            cmd_cancelar_Click
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_Load()
    On Error GoTo erro_Form_Load
    lsub_ajustar_lista_log
    'ajusta valor variável
    mbln_pode_fechar = True
fim_Form_Load:
    Exit Sub
erro_Form_Load:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "Form_Load"
    GoTo fim_Form_Load
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo erro_Form_Unload
    If (Not mbln_pode_fechar) Then
        Cancel = True
    Else
        'se o backup está desativado
        If (Not p_backup.bln_ativar) Then
            'desativa o menu de backup realizar
            frm_principal.smxp_principal.MenuItems.Enabled("k_backup_realizar") = False
        Else
            'ativa o menu backup realizar
            frm_principal.smxp_principal.MenuItems.Enabled("k_backup_realizar") = True
        End If
        'verifica a agenda financeira
        With frm_financeiro_agenda
            If (.Visible) Then
                .Form_Activate
            End If
        End With
    End If
fim_Form_Unload:
    Exit Sub
erro_Form_Unload:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "Form_Unload"
    GoTo fim_Form_Unload
End Sub

Private Sub txt_backup_restaurar_GotFocus()
    On Error GoTo erro_txt_backup_restaurar_GotFocus
    psub_campo_got_focus txt_backup_restaurar
fim_txt_backup_restaurar_GotFocus:
    Exit Sub
erro_txt_backup_restaurar_GotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "txt_backup_restaurar_GotFocus"
    GoTo fim_txt_backup_restaurar_GotFocus
End Sub

Private Sub txt_backup_restaurar_LostFocus()
    On Error GoTo erro_txt_backup_restaurar_LostFocus
    psub_campo_lost_focus txt_backup_restaurar
fim_txt_backup_restaurar_LostFocus:
    Exit Sub
erro_txt_backup_restaurar_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "txt_backup_restaurar_LostFocus"
    GoTo fim_txt_backup_restaurar_LostFocus
End Sub

Private Sub txt_senha_GotFocus()
    On Error GoTo erro_txt_senha_gotFocus
    psub_campo_got_focus txt_senha
fim_txt_senha_gotFocus:
    Exit Sub
erro_txt_senha_gotFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "txt_senha_GotFocus"
    GoTo fim_txt_senha_gotFocus
End Sub

Private Sub txt_senha_LostFocus()
    On Error GoTo erro_txt_senha_LostFocus
    psub_campo_lost_focus txt_senha
fim_txt_senha_LostFocus:
    Exit Sub
erro_txt_senha_LostFocus:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_backup_restaurar", "txt_senha_LostFocus"
    GoTo fim_txt_senha_LostFocus
End Sub

