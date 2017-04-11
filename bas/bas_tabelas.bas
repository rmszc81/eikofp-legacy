Attribute VB_Name = "bas_tabelas"
Option Explicit

'cria as tabelas de configuração
Public Function pfct_criar_tabelas_config() As Boolean
    On Error GoTo erro_pfct_criar_tabelas_config
    Dim lobj_tabela As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    'tb_usuarios
    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_usuarios'"
    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_criar_tabelas_config
    End If
    llng_registros = lobj_tabela.Count
    If (llng_registros = 0) Then
        'monta o comando sql
        lstr_sql = ""
        lstr_sql = lstr_sql & " create table [tb_usuarios] "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " [int_codigo] integer not null primary key autoincrement, "
        lstr_sql = lstr_sql & " [str_usuario] nvarchar(32) not null, "
        lstr_sql = lstr_sql & " [str_senha] nvarchar(64) not null, "
        lstr_sql = lstr_sql & " [str_lembrete_senha] nvarchar(40) null, "
        lstr_sql = lstr_sql & " [dt_criado_em] date not null, "
        lstr_sql = lstr_sql & " [tm_criado_em] time not null, "
        lstr_sql = lstr_sql & " [dt_ultimo_acesso] date null, "
        lstr_sql = lstr_sql & " [tm_ultimo_acesso] time null "
        lstr_sql = lstr_sql & " ) "
        'executa o comando e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_config
        End If
    End If
    'tb_config
    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_config'"
    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_criar_tabelas_config
    End If
    llng_registros = lobj_tabela.Count
    If (llng_registros = 0) Then
        'monta o comando sql
        lstr_sql = ""
        lstr_sql = lstr_sql & " create table [tb_config] "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " [int_codigo] integer not null primary key autoincrement, "
        lstr_sql = lstr_sql & " [int_usuario] integer not null, "
        lstr_sql = lstr_sql & " [int_moeda] integer not null, "
        lstr_sql = lstr_sql & " [int_intervalo_data] integer not null, "
        lstr_sql = lstr_sql & " [chr_carregar_agenda_financeira_login] nvarchar(1) not null, "
        lstr_sql = lstr_sql & " [chr_lancamentos_retroativos] nvarchar(1) not null, "
        lstr_sql = lstr_sql & " [chr_alteracoes_detalhes] nvarchar(1) not null, "
        lstr_sql = lstr_sql & " [chr_data_vencimento_baixa_imediata] nvarchar(1) not null, "
        lstr_sql = lstr_sql & " [chr_lancamentos_duplicados] nvarchar(1) not null, "
        lstr_sql = lstr_sql & " [chr_participou_pesquisa] nvarchar(1) not null "
        lstr_sql = lstr_sql & " ) "
        'executa o comando e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_config
        End If
    Else
        ' ini --- [chr_data_vencimento_baixa_imediata] --- '
        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_config' and [sql] like '%chr_data_vencimento_baixa_imediata%'"
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_config
        End If
        llng_registros = lobj_tabela.Count
        If (llng_registros = 0) Then
            lstr_sql = "alter table [tb_config] add [chr_data_vencimento_baixa_imediata] nvarchar(1) not null default 'N'"
            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_pfct_criar_tabelas_config
            End If
        End If
        ' fim --- [chr_data_vencimento_baixa_imediata] --- '
        ' ini --- [chr_lancamentos_duplicados] --- '
        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_config' and [sql] like '%chr_lancamentos_duplicados%'"
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_config
        End If
        llng_registros = lobj_tabela.Count
        If (llng_registros = 0) Then
            lstr_sql = "alter table [tb_config] add [chr_lancamentos_duplicados] nvarchar(1) not null default 'N'"
            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_pfct_criar_tabelas_config
            End If
        End If
        ' fim --- [chr_lancamentos_duplicados] --- '
        ' ini --- [chr_participou_pesquisa] --- '
        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_config' and [sql] like '%chr_participou_pesquisa%'"
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_config
        End If
        llng_registros = lobj_tabela.Count
        If (llng_registros = 0) Then
            lstr_sql = "alter table [tb_config] add [chr_participou_pesquisa] nvarchar(1) not null default 'N'"
            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_pfct_criar_tabelas_config
            End If
        End If
        ' fim --- [chr_participou_pesquisa] --- '
        ' ini --- [chr_carregar_agenda_financeira_login] --- '
        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_config' and [sql] like '%chr_carregar_agenda_financeira_login%'"
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_config
        End If
        llng_registros = lobj_tabela.Count
        If (llng_registros = 0) Then
            lstr_sql = "alter table [tb_config] add [chr_carregar_agenda_financeira_login] nvarchar(1) not null default 'S'"
            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_pfct_criar_tabelas_config
            End If
        End If
        ' fim --- [chr_carregar_agenda_financeira_login] --- '
    End If
    'tb_backup
    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_backup'"
    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_criar_tabelas_config
    End If
    llng_registros = lobj_tabela.Count
    If (llng_registros = 0) Then
        'monta o comando sql
        lstr_sql = ""
        lstr_sql = lstr_sql & " create table [tb_backup] "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " [int_codigo] integer not null primary key autoincrement, "
        lstr_sql = lstr_sql & " [int_usuario] integer not null, "
        lstr_sql = lstr_sql & " [chr_ativar] nvarchar(1) not null, "
        lstr_sql = lstr_sql & " [int_periodo] integer not null, "
        lstr_sql = lstr_sql & " [str_caminho] nvarchar(512) not null, "
        lstr_sql = lstr_sql & " [dt_ultimo_backup] date null, "
        lstr_sql = lstr_sql & " [tm_ultimo_backup] time null, "
        lstr_sql = lstr_sql & " [dt_proximo_backup] date null, "
        lstr_sql = lstr_sql & " [tm_proximo_backup] time null "
        lstr_sql = lstr_sql & " ) "
        'executa o comando e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_config
        End If
    End If
    'tb_registros
    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_registros'"
    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_criar_tabelas_config
    End If
    llng_registros = lobj_tabela.Count
    If (llng_registros = 0) Then
        'monta o comando sql
        lstr_sql = ""
        lstr_sql = lstr_sql & " create table [tb_registros] "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " [int_codigo] integer not null primary key autoincrement, "
        lstr_sql = lstr_sql & " [int_usuario] integer not null, "
        lstr_sql = lstr_sql & " [str_nome] nvarchar(60) not null, "
        lstr_sql = lstr_sql & " [str_email] nvarchar(60) not null, "
        lstr_sql = lstr_sql & " [str_pais] nvarchar(50) null, "
        lstr_sql = lstr_sql & " [str_estado] nvarchar(30) null, "
        lstr_sql = lstr_sql & " [str_cidade] nvarchar(35) null, "
        lstr_sql = lstr_sql & " [dt_data_nascimento] date null, "
        lstr_sql = lstr_sql & " [str_profissao] nvarchar(60) null, "
        lstr_sql = lstr_sql & " [chr_sexo] nchar(1) null, "
        lstr_sql = lstr_sql & " [str_origem] nvarchar(60) null, "
        lstr_sql = lstr_sql & " [str_opiniao] nvarchar(512) null, "
        lstr_sql = lstr_sql & " [chr_newsletter] nchar(1) null, "
        lstr_sql = lstr_sql & " [str_id_cpu] nvarchar(32) not null, "
        lstr_sql = lstr_sql & " [str_id_hd] nvarchar(32) not null, "
        lstr_sql = lstr_sql & " [dt_data_registro] date not null, "
        lstr_sql = lstr_sql & " [tm_hora_registro] time not null, "
        lstr_sql = lstr_sql & " [dt_data_liberacao] date null, "
        lstr_sql = lstr_sql & " [tm_hora_liberacao] time null, "
        lstr_sql = lstr_sql & " [chr_banido] nchar(1) not null, "
        lstr_sql = lstr_sql & " [str_desc_banido] nvarchar(512) null "
        lstr_sql = lstr_sql & " ) "
        'executa o comando e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_config")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_config
        End If
    End If
    pfct_criar_tabelas_config = True
fim_pfct_criar_tabelas_config:
    Set lobj_tabela = Nothing
    Exit Function
erro_pfct_criar_tabelas_config:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_tabelas", "pfct_criar_tabelas_config"
    GoTo fim_pfct_criar_tabelas_config
End Function

Public Function pfct_criar_tabelas_usuario() As Boolean
    On Error GoTo erro_pfct_criar_tabelas_usuario
    Dim lobj_tabela As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    'tb_contas
    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_contas'"
    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_criar_tabelas_usuario
    End If
    llng_registros = lobj_tabela.Count
    If (llng_registros = 0) Then
        'monta o comando sql
        lstr_sql = ""
        lstr_sql = lstr_sql & " create table [tb_contas] "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " [int_codigo] integer not null primary key autoincrement, "
        lstr_sql = lstr_sql & " [str_descricao] nvarchar(50) not null, "
        lstr_sql = lstr_sql & " [num_saldo] numeric(15,2) not null, "
        lstr_sql = lstr_sql & " [num_limite_negativo] numeric(15,2) not null, "
        lstr_sql = lstr_sql & " [str_observacoes] nvarchar(512) null, "
        lstr_sql = lstr_sql & " [chr_ativo] nvarchar(1) not null "
        lstr_sql = lstr_sql & " ) "
        'executa o comando e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
    End If
    
'    'tb_cartoes_credito
'    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_cartoes_credito'"
'    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
'        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'        GoTo fim_pfct_criar_tabelas_usuario
'    End If
'    llng_registros = lobj_tabela.Count
'    If (llng_registros = 0) Then
'        'monta o comando sql
'        lstr_sql = ""
'        lstr_sql = lstr_sql & " create table [tb_cartoes_credito] "
'        lstr_sql = lstr_sql & " ( "
'        lstr_sql = lstr_sql & "     [int_codigo] integer not null primary key autoincrement, "
'        lstr_sql = lstr_sql & "     [int_conta] integer null, "
'        lstr_sql = lstr_sql & "     [str_descricao] nvarchar(50) not null, "
'        lstr_sql = lstr_sql & "     [str_observacoes] nvarchar(512) null, "
'        lstr_sql = lstr_sql & "     [chr_ativo] nvarchar(1) not null "
'        lstr_sql = lstr_sql & " ) "
'        'executa o comando e devolve o objeto
'        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
'            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'            GoTo fim_pfct_criar_tabelas_usuario
'        End If
'    End If

'    'tb_categorias
'    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_categorias'"
'    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
'        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'        GoTo fim_pfct_criar_tabelas_usuario
'    End If
'    llng_registros = lobj_tabela.Count
'    If (llng_registros = 0) Then
'        'monta o comando sql
'        lstr_sql = ""
'        lstr_sql = lstr_sql & " create table [tb_categorias] "
'        lstr_sql = lstr_sql & " ( "
'        lstr_sql = lstr_sql & "     [int_codigo] integer not null primary key autoincrement, "
'        lstr_sql = lstr_sql & "     [str_descricao] nvarchar(50) not null, "
'        lstr_sql = lstr_sql & "     [str_observacoes] nvarchar(512) null, "
'        lstr_sql = lstr_sql & "     [chr_ativo] nvarchar(1) not null "
'        lstr_sql = lstr_sql & " ) "
'        'executa o comando e devolve o objeto
'        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
'            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'            GoTo fim_pfct_criar_tabelas_usuario
'        End If
'    End If

    'tb_despesas
    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_despesas'"
    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_criar_tabelas_usuario
    End If
    llng_registros = lobj_tabela.Count
    If (llng_registros = 0) Then
        'monta o comando sql
        lstr_sql = ""
        lstr_sql = lstr_sql & " create table [tb_despesas] "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " [int_codigo] integer not null primary key autoincrement, "
        
'        lstr_sql = lstr_sql & " [int_categoria] integer null, "

        lstr_sql = lstr_sql & " [str_descricao] nvarchar(50) not null, "
        lstr_sql = lstr_sql & " [str_observacoes] nvarchar(512) null, "
        lstr_sql = lstr_sql & " [chr_fixa] nvarchar(1) not null, "
        lstr_sql = lstr_sql & " [chr_ativo] nvarchar(1) not null "
        lstr_sql = lstr_sql & " ) "
        'executa o comando e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
    Else
    
'        ' ini --- [int_categoria] --- '
'        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_despesas' and [sql] like '%int_categoria%'"
'        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
'            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'            GoTo fim_pfct_criar_tabelas_usuario
'        End If
'        llng_registros = lobj_tabela.Count
'        If (llng_registros = 0) Then
'            lstr_sql = "alter table [tb_despesas] add [int_categoria] integer null"
'            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
'                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'                GoTo fim_pfct_criar_tabelas_usuario
'            End If
'        End If
'        ' fim --- [int_categoria] --- '
    
        ' ini --- [chr_ativo] --- '
        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_despesas' and [sql] like '%chr_ativo%'"
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
        llng_registros = lobj_tabela.Count
        If (llng_registros = 0) Then
            lstr_sql = "alter table [tb_despesas] add [chr_ativo] nvarchar(1) not null default 'S'"
            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_pfct_criar_tabelas_usuario
            End If
        End If
        ' fim --- [chr_ativo] --- '
        ' ini --- [chr_fixa] --- '
        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_despesas' and [sql] like '%chr_fixa%'"
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
        llng_registros = lobj_tabela.Count
        If (llng_registros = 0) Then
            lstr_sql = "alter table [tb_despesas] add [chr_fixa] nvarchar(1) not null default 'N'"
            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_pfct_criar_tabelas_usuario
            End If
        End If
        ' fim --- [chr_fixa] --- '
    End If
    'tb_receitas
    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_receitas'"
    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_criar_tabelas_usuario
    End If
    llng_registros = lobj_tabela.Count
    If (llng_registros = 0) Then
        'monta o comando sql
        lstr_sql = ""
        lstr_sql = lstr_sql & " create table [tb_receitas] "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " [int_codigo] integer not null primary key autoincrement, "
        
'        lstr_sql = lstr_sql & " [int_categoria] integer null, "

        lstr_sql = lstr_sql & " [str_descricao] nvarchar(50) not null, "
        lstr_sql = lstr_sql & " [str_observacoes] nvarchar(512) null, "
        lstr_sql = lstr_sql & " [chr_fixa] nvarchar(1) not null, "
        lstr_sql = lstr_sql & " [chr_ativo] nvarchar(1) not null "
        lstr_sql = lstr_sql & " ) "
        'executa o comando e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
    Else
    
'        ' ini --- [int_categoria] --- '
'        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_receitas' and [sql] like '%int_categoria%'"
'        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
'            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'            GoTo fim_pfct_criar_tabelas_usuario
'        End If
'        llng_registros = lobj_tabela.Count
'        If (llng_registros = 0) Then
'            lstr_sql = "alter table [tb_receitas] add [int_categoria] integer null"
'            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
'                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'                GoTo fim_pfct_criar_tabelas_usuario
'            End If
'        End If
'        ' fim --- [int_categoria] --- '
    
        ' ini --- [chr_ativo] --- '
        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_receitas' and [sql] like '%chr_ativo%'"
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
        llng_registros = lobj_tabela.Count
        If (llng_registros = 0) Then
            lstr_sql = "alter table [tb_receitas] add [chr_ativo] nvarchar(1) not null default 'S'"
            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_pfct_criar_tabelas_usuario
            End If
        End If
        ' fim --- [chr_ativo] --- '
        ' ini --- [chr_fixa] --- '
        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_receitas' and [sql] like '%chr_fixa%'"
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
        llng_registros = lobj_tabela.Count
        If (llng_registros = 0) Then
            lstr_sql = "alter table [tb_receitas] add [chr_fixa] nvarchar(1) not null default 'N'"
            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_pfct_criar_tabelas_usuario
            End If
        End If
        ' fim --- [chr_fixa] --- '
    End If
    
'    'tb_faturas_cartao_credito_mestre
'    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_faturas_cartao_credito_mestre'"
'    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
'        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'        GoTo fim_pfct_criar_tabelas_usuario
'    End If
'    llng_registros = lobj_tabela.Count
'    If (llng_registros = 0) Then
'        'monta o comando sql
'        lstr_sql = ""
'        lstr_sql = lstr_sql & " create table [tb_faturas_cartao_credito_mestre] "
'        lstr_sql = lstr_sql & " ( "
'        lstr_sql = lstr_sql & "     [int_codigo] integer not null primary key autoincrement, "
'        lstr_sql = lstr_sql & "     [int_cartao_credito] integer not null, "
'        lstr_sql = lstr_sql & "     [int_mes] integer not null, "
'        lstr_sql = lstr_sql & "     [int_ano] integer not null, "
'        lstr_sql = lstr_sql & "     [num_valor] numeric(15,2) not null, "
'        lstr_sql = lstr_sql & "     [dt_abertura] date not null, "
'        lstr_sql = lstr_sql & "     [chr_fechada] nvarchar(1) not null, "
'        lstr_sql = lstr_sql & "     [dt_fechamento] date null, "
'        lstr_sql = lstr_sql & "     [chr_paga] nvarchar(1) not null, "
'        lstr_sql = lstr_sql & "     [dt_pagamento] date null "
'        lstr_sql = lstr_sql & " ) "
'        'executa o comando e devolve o objeto
'        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
'            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'            GoTo fim_pfct_criar_tabelas_usuario
'        End If
'    End If
    
'    'tb_faturas_cartao_credito_itens
'    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_faturas_cartao_credito_itens'"
'    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
'        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'        GoTo fim_pfct_criar_tabelas_usuario
'    End If
'    llng_registros = lobj_tabela.Count
'    If (llng_registros = 0) Then
'        'monta o comando sql
'        lstr_sql = ""
'        lstr_sql = lstr_sql & " create table [tb_faturas_cartao_credito_itens] "
'        lstr_sql = lstr_sql & " ( "
'        lstr_sql = lstr_sql & "     [int_codigo] integer not null primary key autoincrement, "
'        lstr_sql = lstr_sql & "     [int_cartao_credito] integer not null, "
'        lstr_sql = lstr_sql & "     [int_fatura] integer not null, "
'        lstr_sql = lstr_sql & "     [dt_compra] date not null, "
'        lstr_sql = lstr_sql & "     [dt_pagamento] date null, "
'        lstr_sql = lstr_sql & "     [int_despesa] integer not null, "
'        lstr_sql = lstr_sql & "     [int_parcela] integer not null, "
'        lstr_sql = lstr_sql & "     [num_valor_parcela] numeric(15,2) not null, "
'        lstr_sql = lstr_sql & "     [int_total_parcelas] integer not null, "
'        lstr_sql = lstr_sql & "     [num_valor_total] numeric(15,2) not null "
'        lstr_sql = lstr_sql & " ) "
'        'executa o comando e devolve o objeto
'        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
'            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'            GoTo fim_pfct_criar_tabelas_usuario
'        End If
'    End If
    
    'tb_formas_pagamento
    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_formas_pagamento'"
    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_criar_tabelas_usuario
    End If
    llng_registros = lobj_tabela.Count
    If (llng_registros = 0) Then
        'monta o comando sql
        lstr_sql = ""
        lstr_sql = lstr_sql & " create table [tb_formas_pagamento] "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " [int_codigo] integer not null primary key autoincrement, "
        lstr_sql = lstr_sql & " [str_descricao] nvarchar(50) not null, "
        lstr_sql = lstr_sql & " [str_observacoes] nvarchar(512) null, "
        lstr_sql = lstr_sql & " [chr_ativo] nvarchar(1) not null "
        lstr_sql = lstr_sql & " ) "
        'executa o comando e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
    Else
        ' ini --- [chr_ativo] --- '
        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_formas_pagamento' and [sql] like '%chr_ativo%'"
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
        llng_registros = lobj_tabela.Count
        If (llng_registros = 0) Then
            lstr_sql = "alter table [tb_formas_pagamento] add [chr_ativo] nvarchar(1) not null default 'S'"
            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_pfct_criar_tabelas_usuario
            End If
        End If
        ' fim --- [chr_ativo] --- '
    End If
    'tb_contas_pagar
    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_contas_pagar'"
    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_criar_tabelas_usuario
    End If
    llng_registros = lobj_tabela.Count
    If (llng_registros = 0) Then
        'monta o comando sql
        lstr_sql = ""
        lstr_sql = lstr_sql & " create table [tb_contas_pagar] "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " [int_codigo] integer not null primary key autoincrement, "
        lstr_sql = lstr_sql & " [int_conta_baixa_automatica] integer null, "
        lstr_sql = lstr_sql & " [chr_baixa_automatica] nvarchar(1) not null, "
        lstr_sql = lstr_sql & " [int_despesa] integer not null, "
        lstr_sql = lstr_sql & " [int_forma_pagamento] integer not null, "
        lstr_sql = lstr_sql & " [dt_vencimento] date not null, "
        lstr_sql = lstr_sql & " [int_parcela] integer not null, "
        lstr_sql = lstr_sql & " [int_total_parcelas] integer not null, "
        lstr_sql = lstr_sql & " [num_valor] numeric(15,2) not null, "
        lstr_sql = lstr_sql & " [str_descricao] nvarchar(50) not null, "
        lstr_sql = lstr_sql & " [str_documento] nvarchar(30) null, "
        lstr_sql = lstr_sql & " [str_chave] nvarchar(255) null, "
        lstr_sql = lstr_sql & " [str_codigo_barras] nvarchar(100) null, "
        lstr_sql = lstr_sql & " [str_observacoes] nvarchar(512) null "
        lstr_sql = lstr_sql & " ) "
        'executa o comando e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
    Else
        ' ini --- [int_conta_baixa_automatica] --- '
        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_contas_pagar' and [sql] like '%int_conta_baixa_automatica%'"
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
        llng_registros = lobj_tabela.Count
        If (llng_registros = 0) Then
            lstr_sql = "alter table [tb_contas_pagar] add [int_conta_baixa_automatica] integer null default 0"
            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_pfct_criar_tabelas_usuario
            End If
        End If
        ' fim --- [int_conta_baixa_automatica] --- '
        ' ini --- [chr_baixa_automatica] --- '
        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_contas_pagar' and [sql] like '%chr_baixa_automatica%'"
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
        llng_registros = lobj_tabela.Count
        If (llng_registros = 0) Then
            lstr_sql = "alter table [tb_contas_pagar] add [chr_baixa_automatica] nvarchar(1) not null default 'N'"
            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_pfct_criar_tabelas_usuario
            End If
        End If
        ' fim --- [chr_baixa_automatica] --- '
    End If
    'tb_contas_receber
    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_contas_receber'"
    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_criar_tabelas_usuario
    End If
    llng_registros = lobj_tabela.Count
    If (llng_registros = 0) Then
        'monta o comando sql
        lstr_sql = ""
        lstr_sql = lstr_sql & " create table [tb_contas_receber] "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " [int_codigo] integer not null primary key autoincrement, "
        lstr_sql = lstr_sql & " [int_conta_baixa_automatica] integer null, "
        lstr_sql = lstr_sql & " [chr_baixa_automatica] nvarchar(1) not null, "
        lstr_sql = lstr_sql & " [int_receita] integer not null, "
        lstr_sql = lstr_sql & " [int_forma_pagamento] integer not null, "
        lstr_sql = lstr_sql & " [dt_vencimento] date not null, "
        lstr_sql = lstr_sql & " [int_parcela] integer not null, "
        lstr_sql = lstr_sql & " [int_total_parcelas] integer not null, "
        lstr_sql = lstr_sql & " [num_valor] numeric(15,2) not null, "
        lstr_sql = lstr_sql & " [str_descricao] nvarchar(50) not null, "
        lstr_sql = lstr_sql & " [str_documento] nvarchar(30) null, "
        lstr_sql = lstr_sql & " [str_chave] nvarchar(255) null, "
        lstr_sql = lstr_sql & " [str_observacoes] nvarchar(512) null "
        lstr_sql = lstr_sql & " ) "
        'executa o comando e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
    Else
        ' ini --- [int_conta_baixa_automatica] --- '
        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_contas_receber' and [sql] like '%int_conta_baixa_automatica%'"
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
        llng_registros = lobj_tabela.Count
        If (llng_registros = 0) Then
            lstr_sql = "alter table [tb_contas_receber] add [int_conta_baixa_automatica] integer null default 0"
            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_pfct_criar_tabelas_usuario
            End If
        End If
        ' fim --- [int_conta_baixa_automatica] --- '
        ' ini --- [chr_baixa_automatica] --- '
        lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_contas_receber' and [sql] like '%chr_baixa_automatica%'"
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
        llng_registros = lobj_tabela.Count
        If (llng_registros = 0) Then
            lstr_sql = "alter table [tb_contas_receber] add [chr_baixa_automatica] nvarchar(1) not null default 'N'"
            If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
                MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
                GoTo fim_pfct_criar_tabelas_usuario
            End If
        End If
        ' fim --- [chr_baixa_automatica] --- '
    End If
    'tb_movimentacao
    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_movimentacao'"
    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_criar_tabelas_usuario
    End If
    llng_registros = lobj_tabela.Count
    If (llng_registros = 0) Then
        'monta o comando sql
        lstr_sql = ""
        lstr_sql = lstr_sql & " create table [tb_movimentacao] "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " [int_codigo] integer not null primary key autoincrement, "
        lstr_sql = lstr_sql & " [int_conta] integer not null, "
        lstr_sql = lstr_sql & " [int_receita] integer not null, "
        lstr_sql = lstr_sql & " [int_despesa] integer not null, "
        lstr_sql = lstr_sql & " [int_forma_pagamento] integer not null, "
        lstr_sql = lstr_sql & " [chr_tipo] nvarchar(1) not null, "
        lstr_sql = lstr_sql & " [dt_vencimento] date not null, "
        lstr_sql = lstr_sql & " [dt_pagamento] date not null, "
        lstr_sql = lstr_sql & " [int_parcela] integer not null, "
        lstr_sql = lstr_sql & " [int_total_parcelas] integer not null, "
        lstr_sql = lstr_sql & " [num_valor] numeric(15,2) not null, "
        lstr_sql = lstr_sql & " [str_descricao] nvarchar(50) not null, "
        lstr_sql = lstr_sql & " [str_documento] nvarchar(30) null, "
        lstr_sql = lstr_sql & " [str_codigo_barras] nvarchar(100) null, "
        lstr_sql = lstr_sql & " [str_observacoes] nvarchar(512) null "
        lstr_sql = lstr_sql & " ) "
        'executa o comando e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_criar_tabelas_usuario
        End If
    End If
    
'    'tb_movimentacao_faturas
'    lstr_sql = "select * from [sqlite_master] where [tbl_name] = 'tb_movimentacao_faturas'"
'    If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
'        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'        GoTo fim_pfct_criar_tabelas_usuario
'    End If
'    llng_registros = lobj_tabela.Count
'    If (llng_registros = 0) Then
'        'monta o comando sql
'        lstr_sql = ""
'        lstr_sql = lstr_sql & " create table [tb_movimentacao_faturas] "
'        lstr_sql = lstr_sql & " ( "
'        lstr_sql = lstr_sql & "     [int_codigo] integer not null primary key autoincrement, "
'        lstr_sql = lstr_sql & "     [int_movimentacao] integer not null, "
'        lstr_sql = lstr_sql & "     [int_fatura] integer not null "
'        lstr_sql = lstr_sql & " ) "
'        'executa o comando e devolve o objeto
'        If (Not pfct_executar_comando_sql(lobj_tabela, lstr_sql, "bas_tabelas", "pfct_criar_tabelas_usuario")) Then
'            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'            GoTo fim_pfct_criar_tabelas_usuario
'        End If
'    End If
    
    pfct_criar_tabelas_usuario = True
fim_pfct_criar_tabelas_usuario:
    'destrói os objetos
    Set lobj_tabela = Nothing
    Exit Function
erro_pfct_criar_tabelas_usuario:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_tabelas", "pfct_criar_tabelas_usuario"
    GoTo fim_pfct_criar_tabelas_usuario
End Function
