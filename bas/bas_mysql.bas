Attribute VB_Name = "bas_mysql"
Option Explicit

Private mobj_conexao As MYSQL_CONNECTION
Private mobj_conexao_estado As MYSQL_CONNECTION_STATE

Private Sub psub_configurar_parametros()
    On Error GoTo Erro_psub_configurar_parametros
    With p_mysql
        .str_servidor = ""
        .str_usuario = ""
        .str_senha = ""
        .str_banco = ""
        .lng_porta = 0
    End With
Fim_psub_configurar_parametros:
    Exit Sub
Erro_psub_configurar_parametros:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_mysql", "psub_configurar_parametros"
    GoTo Fim_psub_configurar_parametros
End Sub

Private Function pfct_conectar_banco() As Boolean
    On Error GoTo erro_pfct_conectar_banco
    'configura os parâmetros de conexão
    psub_configurar_parametros
    'instancia o objeto modular
    Set mobj_conexao = New MYSQL_CONNECTION
    'ajustamos o timeout para 5 segundos
    mobj_conexao.SetOption MYSQL_OPT_CONNECT_TIMEOUT, 5
    'abre a conexão com o banco
    mobj_conexao_estado = mobj_conexao.OpenConnection(p_mysql.str_servidor, p_mysql.str_usuario, p_mysql.str_senha, p_mysql.str_banco, p_mysql.lng_porta)
    'se a conexão foi aberta com sucesso
    If (mobj_conexao_estado <> MY_CONN_OPEN) Then
        'desvia ao bloco fim
        GoTo fim_pfct_conectar_banco
    End If
    'retorna true
    pfct_conectar_banco = True
fim_pfct_conectar_banco:
    Exit Function
erro_pfct_conectar_banco:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_mysql", "pfct_conectar_banco"
    GoTo fim_pfct_conectar_banco
End Function

Private Function pfct_desconectar_banco() As Boolean
    On Error GoTo erro_pfct_desconectar_banco
    'verifica se o objeto está instanciado
    If (Not mobj_conexao Is Nothing) Then
        'verifica se a conexão está aberta e se pode ser fechada
        If (mobj_conexao_estado = MY_CONN_OPEN) Then
            'fecha a conexão
            mobj_conexao.CloseConnection
        Else
            'desvia ao bloco fim
            GoTo fim_pfct_desconectar_banco
        End If
    End If
    'retorna true
    pfct_desconectar_banco = True
fim_pfct_desconectar_banco:
    Set mobj_conexao = Nothing
    Exit Function
erro_pfct_desconectar_banco:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_mysql", "pfct_desconectar_banco"
    GoTo fim_pfct_desconectar_banco
End Function

Private Function pfct_executar_comando_sql(ByRef pobj_objeto As MYSQL_RS, _
                                           ByVal pstr_sql As String, _
                                           ByVal pstr_modulo As String, _
                                           ByVal pstr_metodo As String) As Boolean
    On Error GoTo erro_pfct_executar_comando_sql
    Dim lstr_sql As String
    Dim lobj_retorno As MYSQL_RS
    'conecta ao banco
    If (pfct_conectar_banco()) Then
        'remove excesso de espaços da string sql
        lstr_sql = pfct_remover_excesso_espacos(pstr_sql)
        'executa o comando sql
        Set lobj_retorno = mobj_conexao.Execute(lstr_sql)
        'verifica se houve algum erro no comando
        If (mobj_conexao.Error.Description <> "") Then
            'se ocorreu erro ao executar o comando
            psub_gerar_log_sql lstr_sql, pstr_modulo, pstr_metodo, mobj_conexao.Error.Description
            GoTo fim_pfct_executar_comando_sql
        Else
            'se o comando foi executado corretamente
            psub_gerar_log_sql lstr_sql, pstr_modulo, pstr_metodo, ""
        End If
        'devolve o objeto por referência
        Set pobj_objeto = lobj_retorno
        'fecha a conexão com o banco
        If (Not pfct_desconectar_banco()) Then
            'não disparar mensagem de erro
            'MsgBox "Erro ao fechar a conexão com o banco de dados.", vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_executar_comando_sql
        End If
    Else
        'não disparar mensagem de erro
        'MsgBox "Erro ao abrir a conexão com o banco de dados.", vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_executar_comando_sql
    End If
    'retorna true
    pfct_executar_comando_sql = True
fim_pfct_executar_comando_sql:
    'destrói os objetos
    Set mobj_conexao = Nothing
    Exit Function
erro_pfct_executar_comando_sql:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_mysql", "pfct_executar_comando_sql"
    GoTo fim_pfct_executar_comando_sql
End Function

Public Function pfct_retorna_data_hora_mysql() As Date
    On Error GoTo Erro_pfct_retorna_data_hora_mysql
    Dim lobj_dados As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim dt_data_servidor As Date
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = "select sysdate() as data_hora_atual"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "bas_mysql", "pfct_retorna_data_hora_mysql")) Then
        GoTo Fim_pfct_retorna_data_hora_mysql
    End If
    llng_registros = lobj_dados.RecordCount
    'se houver registros
    If (llng_registros > 0) Then
        'retorna a data/hora atual
        dt_data_servidor = lobj_dados.Fields("data_hora_atual").Value
    End If
    'retorna a data/hora do servidor
    pfct_retorna_data_hora_mysql = dt_data_servidor
Fim_pfct_retorna_data_hora_mysql:
    Set lobj_dados = Nothing
    Exit Function
Erro_pfct_retorna_data_hora_mysql:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_mysql", "pfct_retorna_data_hora_mysql"
    GoTo Fim_pfct_retorna_data_hora_mysql
End Function

Public Function pfct_carrega_dados_pesquisa_publico(ByVal str_usuario As String, ByVal str_id_cpu As String, ByVal str_id_hd As String, ByRef obj_registro As tpe_registro) As Boolean
    On Error GoTo Erro_pfct_carrega_dados_pesquisa_publico
    Dim lobj_dados As MYSQL_RS
    Dim lstr_sql As String
    Dim llng_registros As Long
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " select * from `tb_registros` where 1 = 1 "
    lstr_sql = lstr_sql & " and `str_usuario` = '" & pfct_tratar_texto_sql(str_usuario) & "' "
    lstr_sql = lstr_sql & " and `str_id_cpu` = '" & pfct_tratar_texto_sql(str_id_cpu) & "' "
    lstr_sql = lstr_sql & " and `str_id_hd` = '" & pfct_tratar_texto_sql(str_id_hd) & "' "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "bas_mysql", "pfct_carrega_dados_pesquisa_publico")) Then
        GoTo Fim_pfct_carrega_dados_pesquisa_publico
    End If
    llng_registros = lobj_dados.RecordCount
    If (llng_registros > 0) Then
        'retorna os dados ao tipo passado por referência
        With obj_registro
            .int_codigo = lobj_dados.Fields("int_codigo").Value
            .str_usuario = lobj_dados.Fields("str_usuario").Value
            .str_nome = lobj_dados.Fields("str_nome").Value
            .str_email = lobj_dados.Fields("str_email").Value
            .str_pais = lobj_dados.Fields("str_pais").Value
            .str_estado = lobj_dados.Fields("str_estado").Value
            .str_cidade = lobj_dados.Fields("str_cidade").Value
            .dt_data_nascimento = lobj_dados.Fields("dt_data_nascimento").Value
            .str_profissao = lobj_dados.Fields("str_profissao").Value
            .chr_sexo = lobj_dados.Fields("chr_sexo").Value
            .str_origem = lobj_dados.Fields("str_origem").Value
            .str_opiniao = lobj_dados.Fields("str_opiniao").Value
            .bln_newsletter = IIf(lobj_dados.Fields("chr_newsletter").Value = "S", True, False)
            .str_id_cpu = lobj_dados.Fields("str_id_cpu").Value
            .str_id_hd = lobj_dados.Fields("str_id_hd").Value
            .dt_data_registro = lobj_dados.Fields("dt_data_registro").Value
            .dt_data_liberacao = lobj_dados.Fields("dt_data_liberacao").Value
            .bln_banido = IIf(lobj_dados.Fields("chr_banido").Value = "S", True, False)
            .str_desc_banido = lobj_dados.Fields("str_desc_banido").Value
        End With
    End If
    'retorna true
    pfct_carrega_dados_pesquisa_publico = True
Fim_pfct_carrega_dados_pesquisa_publico:
    Set lobj_dados = Nothing
    Exit Function
Erro_pfct_carrega_dados_pesquisa_publico:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_mysql", "pfct_carrega_dados_pesquisa_publico"
    GoTo Fim_pfct_carrega_dados_pesquisa_publico
End Function

Public Function pfct_inserir_dados_pesquisa_publico() As Boolean
    On Error GoTo Erro_pfct_inserir_dados_pesquisa_publico
    Dim lobj_dados As Object
    Dim lstr_sql As String
    Dim ldt_data_servidor As Date
    'retorna a data do servidor
    ldt_data_servidor = pfct_retorna_data_hora_mysql()
    'se não houve retorno de data
    If (ldt_data_servidor = CDate(0)) Then 'retorna [30/12/1899 00:00:00]
        'desvia ao bloco fim
        GoTo Fim_pfct_inserir_dados_pesquisa_publico
    End If
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " insert into tb_registros "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & "     `str_usuario`, "
    lstr_sql = lstr_sql & "     `str_nome`, "
    lstr_sql = lstr_sql & "     `str_email`, "
    lstr_sql = lstr_sql & "     `str_pais`, "
    lstr_sql = lstr_sql & "     `str_estado`, "
    lstr_sql = lstr_sql & "     `str_cidade`, "
    lstr_sql = lstr_sql & "     `dt_data_nascimento`, "
    lstr_sql = lstr_sql & "     `str_profissao`, "
    lstr_sql = lstr_sql & "     `chr_sexo`, "
    lstr_sql = lstr_sql & "     `str_origem`, "
    lstr_sql = lstr_sql & "     `str_opiniao`, "
    lstr_sql = lstr_sql & "     `chr_newsletter`, "
    lstr_sql = lstr_sql & "     `str_id_cpu`, "
    lstr_sql = lstr_sql & "     `str_id_hd`, "
    lstr_sql = lstr_sql & "     `dt_data_registro`, "
    lstr_sql = lstr_sql & "     `dt_data_liberacao`, "
    lstr_sql = lstr_sql & "     `chr_banido`, "
    lstr_sql = lstr_sql & "     `str_desc_banido` "
    lstr_sql = lstr_sql & " ) "
    lstr_sql = lstr_sql & " values "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(p_usuario.str_login) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(p_registro.str_nome) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(p_registro.str_email) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(p_registro.str_pais) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(p_registro.str_estado) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(p_registro.str_cidade) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_data_sql(p_registro.dt_data_nascimento) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(p_registro.str_profissao) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(p_registro.chr_sexo) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(p_registro.str_origem) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(p_registro.str_opiniao) & "', "
    lstr_sql = lstr_sql & "     '" & IIf(p_registro.bln_newsletter, "S", "N") & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(p_registro.str_id_cpu) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(p_registro.str_id_hd) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_data_sql(p_registro.dt_data_registro) & " " & pfct_tratar_hora_sql(p_registro.dt_data_registro) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_data_sql(ldt_data_servidor) & " " & pfct_tratar_hora_sql(ldt_data_servidor) & "', "
    lstr_sql = lstr_sql & "     '" & IIf(p_registro.bln_banido, "S", "N") & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(p_registro.str_desc_banido) & "' "
    lstr_sql = lstr_sql & " ) "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "bas_mysql", "pfct_inserir_dados_pesquisa_publico")) Then
        GoTo Fim_pfct_inserir_dados_pesquisa_publico
    End If
    'ajusta a data de liberação para o objeto local
    p_registro.dt_data_liberacao = ldt_data_servidor
    'retorna true
    pfct_inserir_dados_pesquisa_publico = True
Fim_pfct_inserir_dados_pesquisa_publico:
    Set lobj_dados = Nothing
    Exit Function
Erro_pfct_inserir_dados_pesquisa_publico:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_mysql", "pfct_inserir_dados_pesquisa_publico"
    GoTo Fim_pfct_inserir_dados_pesquisa_publico
End Function

Public Function pfct_inserir_dados_historico_pesquisa_publico(ByRef ptpe_registro As tpe_registro) As Boolean
    On Error GoTo Erro_pfct_inserir_dados_historico_pesquisa_publico
    Dim lobj_dados As Object
    Dim lstr_sql As String
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " INSERT INTO tb_registros_historico "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & "     `int_codigo_usuario`, "
    lstr_sql = lstr_sql & "     `str_usuario`, "
    lstr_sql = lstr_sql & "     `str_nome`, "
    lstr_sql = lstr_sql & "     `str_email`, "
    lstr_sql = lstr_sql & "     `str_pais`, "
    lstr_sql = lstr_sql & "     `str_estado`, "
    lstr_sql = lstr_sql & "     `str_cidade`, "
    lstr_sql = lstr_sql & "     `dt_data_nascimento`, "
    lstr_sql = lstr_sql & "     `str_profissao`, "
    lstr_sql = lstr_sql & "     `chr_sexo`, "
    lstr_sql = lstr_sql & "     `str_origem`, "
    lstr_sql = lstr_sql & "     `str_opiniao`, "
    lstr_sql = lstr_sql & "     `chr_newsletter`, "
    lstr_sql = lstr_sql & "     `str_id_cpu`, "
    lstr_sql = lstr_sql & "     `str_id_hd`, "
    lstr_sql = lstr_sql & "     `dt_data_registro`, "
    lstr_sql = lstr_sql & "     `dt_data_liberacao`, "
    lstr_sql = lstr_sql & "     `chr_banido`, "
    lstr_sql = lstr_sql & "     `str_desc_banido` "
    lstr_sql = lstr_sql & " ) "
    lstr_sql = lstr_sql & " VALUES "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & "     " & pfct_tratar_numero_sql(ptpe_registro.int_codigo) & ", "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(ptpe_registro.str_usuario) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(ptpe_registro.str_nome) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(ptpe_registro.str_email) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(ptpe_registro.str_pais) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(ptpe_registro.str_estado) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(ptpe_registro.str_cidade) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_data_sql(ptpe_registro.dt_data_nascimento) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(ptpe_registro.str_profissao) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(ptpe_registro.chr_sexo) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(ptpe_registro.str_origem) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(ptpe_registro.str_opiniao) & "', "
    lstr_sql = lstr_sql & "     '" & IIf(ptpe_registro.bln_newsletter, "S", "N") & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(ptpe_registro.str_id_cpu) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(ptpe_registro.str_id_hd) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_data_sql(ptpe_registro.dt_data_registro) & " " & pfct_tratar_hora_sql(ptpe_registro.dt_data_registro) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_data_sql(ptpe_registro.dt_data_liberacao) & " " & pfct_tratar_hora_sql(ptpe_registro.dt_data_liberacao) & "', "
    lstr_sql = lstr_sql & "     '" & IIf(ptpe_registro.bln_banido, "S", "N") & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(ptpe_registro.str_desc_banido) & "' "
    lstr_sql = lstr_sql & " ) "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "bas_mysql", "pfct_inserir_dados_historico_pesquisa_publico")) Then
        GoTo Fim_pfct_inserir_dados_historico_pesquisa_publico
    End If
    'retorna true
    pfct_inserir_dados_historico_pesquisa_publico = True
Fim_pfct_inserir_dados_historico_pesquisa_publico:
    Set lobj_dados = Nothing
    Exit Function
Erro_pfct_inserir_dados_historico_pesquisa_publico:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_mysql", "pfct_inserir_dados_historico_pesquisa_publico"
    GoTo Fim_pfct_inserir_dados_historico_pesquisa_publico
End Function

Public Function pfct_atualizar_dados_pesquisa_publico(ByRef ptpe_registro As tpe_registro) As Boolean
    On Error GoTo Erro_pfct_atualizar_dados_pesquisa_publico
    Dim lobj_dados As Object
    Dim lstr_sql As String
    'antes de atualizar os dados da pesquisa, fazer insert na tabela de histórico de alterações
    If (pfct_inserir_dados_historico_pesquisa_publico(ptpe_registro)) Then
        'monta o comando sql
        lstr_sql = ""
        lstr_sql = lstr_sql & " update "
        lstr_sql = lstr_sql & "     tb_registros "
        lstr_sql = lstr_sql & " set "
        lstr_sql = lstr_sql & "     `str_pais` = '" & pfct_tratar_texto_sql(p_registro.str_pais) & "', "
        lstr_sql = lstr_sql & "     `str_estado` = '" & pfct_tratar_texto_sql(p_registro.str_estado) & "', "
        lstr_sql = lstr_sql & "     `str_cidade` = '" & pfct_tratar_texto_sql(p_registro.str_cidade) & "', "
        lstr_sql = lstr_sql & "     `dt_data_nascimento` = '" & pfct_tratar_data_sql(p_registro.dt_data_nascimento) & "', "
        lstr_sql = lstr_sql & "     `str_profissao` = '" & pfct_tratar_texto_sql(p_registro.str_profissao) & "', "
        lstr_sql = lstr_sql & "     `chr_sexo` = '" & pfct_tratar_texto_sql(p_registro.chr_sexo) & "', "
        lstr_sql = lstr_sql & "     `str_origem` = '" & pfct_tratar_texto_sql(p_registro.str_origem) & "', "
        lstr_sql = lstr_sql & "     `str_opiniao` = '" & pfct_tratar_texto_sql(p_registro.str_opiniao) & "', "
        lstr_sql = lstr_sql & "     `chr_newsletter` = '" & IIf(p_registro.bln_newsletter, "S", "N") & "' "
        lstr_sql = lstr_sql & " where 1 = 1 "
        lstr_sql = lstr_sql & "     and `int_codigo` = " & CStr(ptpe_registro.int_codigo) & " "
        'executa o comando sql e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "bas_mysql", "pfct_atualizar_dados_pesquisa_publico")) Then
            GoTo Fim_pfct_atualizar_dados_pesquisa_publico
        End If
    End If
    'retorna true
    pfct_atualizar_dados_pesquisa_publico = True
Fim_pfct_atualizar_dados_pesquisa_publico:
    Set lobj_dados = Nothing
    Exit Function
Erro_pfct_atualizar_dados_pesquisa_publico:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_mysql", "pfct_atualizar_dados_pesquisa_publico"
    GoTo Fim_pfct_atualizar_dados_pesquisa_publico
End Function

Public Function pfct_excluir_dados_pesquisa_publico(ByVal int_codigo As Integer) As Boolean
    On Error GoTo Erro_pfct_excluir_dados_pesquisa_publico
    Dim lobj_dados As Object
    Dim lstr_sql As String
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " delete from `tb_registros` where `int_codigo` = " & CStr(int_codigo) & " "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "bas_mysql", "pfct_excluir_dados_pesquisa_publico")) Then
        GoTo Fim_pfct_excluir_dados_pesquisa_publico
    End If
    'retorna true
    pfct_excluir_dados_pesquisa_publico = True
Fim_pfct_excluir_dados_pesquisa_publico:
    Set lobj_dados = Nothing
    Exit Function
Erro_pfct_excluir_dados_pesquisa_publico:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_mysql", "pfct_excluir_dados_pesquisa_publico"
    GoTo Fim_pfct_excluir_dados_pesquisa_publico
End Function

Public Function pfct_inserir_historico_acesso(ByVal str_usuario As String, ByVal str_id_cpu As String, ByVal str_id_hd As String) As Boolean
    On Error GoTo Erro_pfct_inserir_historico_acesso
    Dim lobj_dados As Object
    Dim lstr_sql As String
    Dim ldt_data_servidor As Date
    'retorna a data do servidor
    ldt_data_servidor = pfct_retorna_data_hora_mysql()
    'se não houve retorno de data
    If (ldt_data_servidor = CDate(0)) Then 'retorna [30/12/1899 00:00:00]
        'desvia ao bloco fim
        GoTo Fim_pfct_inserir_historico_acesso
    End If
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " insert into `tb_acessos` "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & "     `str_usuario`, "
    lstr_sql = lstr_sql & "     `str_id_cpu`, "
    lstr_sql = lstr_sql & "     `str_id_hd`, "
    lstr_sql = lstr_sql & "     `str_app_versao`, "
    lstr_sql = lstr_sql & "     `dt_data_acesso` "
    lstr_sql = lstr_sql & " ) "
    lstr_sql = lstr_sql & " values "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(str_usuario) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(str_id_cpu) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(str_id_hd) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_texto_sql(pcst_app_ver & " / " & Replace$(pcst_dba_ver, ",", ".")) & "', "
    lstr_sql = lstr_sql & "     '" & pfct_tratar_data_sql(ldt_data_servidor) & " " & pfct_tratar_hora_sql(ldt_data_servidor) & "' "
    lstr_sql = lstr_sql & " ); "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "bas_mysql", "pfct_inserir_historico_acesso")) Then
        GoTo Fim_pfct_inserir_historico_acesso
    End If
    'retorna true
    pfct_inserir_historico_acesso = True
Fim_pfct_inserir_historico_acesso:
    Set lobj_dados = Nothing
    Exit Function
Erro_pfct_inserir_historico_acesso:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_mysql", "pfct_inserir_historico_acesso"
    GoTo Fim_pfct_inserir_historico_acesso
End Function
