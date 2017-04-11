Attribute VB_Name = "bas_bd"
Option Explicit

Private mobj_conexao As SQLite3COMUTF8

Private Function pfct_conectar_banco() As Boolean
    On Error GoTo erro_pfct_conectar_banco
    Dim lstr_caminho_backup As String
    'instancia o objeto modular
    Set mobj_conexao = New SQLite3COMUTF8
    'não gera erro de runtime
    mobj_conexao.NoErrorMode = True
    'verifica o tipo de banco
    If (p_banco.tb_tipo_banco = tb_config) Then 'se for tipo config
        'conecta ao arquivo
        If (Not mobj_conexao.IsOpened) Then
            mobj_conexao.Open p_banco.str_caminho_dados_config
        End If
    End If
    If (p_banco.tb_tipo_banco = tb_dados) Then 'se for do tipo usuário
        'conecta ao arquivo
        If (Not mobj_conexao.IsOpened) Then
            mobj_conexao.Open p_banco.str_caminho_dados_usuario
        End If
    End If
    If (p_banco.tb_tipo_banco = tb_backup) Then
        'concatena os caminhos
        lstr_caminho_backup = p_banco.str_caminho_dados_backup
        'conecta ao arquivo
        If (Not mobj_conexao.IsOpened) Then
            mobj_conexao.Open lstr_caminho_backup
        End If
    End If
    If (p_banco.tb_tipo_banco = tb_restaurar) Then
        'conecta ao arquivo
        If (Not mobj_conexao.IsOpened) Then
            mobj_conexao.Open p_banco.str_caminho_dados_restaurar
        End If
    End If
    'retorna true
    pfct_conectar_banco = mobj_conexao.IsOpened 'retornamos o estado da conexão
fim_pfct_conectar_banco:
    Exit Function
erro_pfct_conectar_banco:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_bd", "pfct_conectar_banco"
    GoTo fim_pfct_conectar_banco
End Function

Private Function pfct_desconectar_banco() As Boolean
    On Error GoTo erro_pfct_desconectar_banco
    'verifica se o objeto está instanciado
    If (Not mobj_conexao Is Nothing) Then
        'verifica se a conexão está aberta e se pode ser fechada
        If (mobj_conexao.IsOpened) Then
            'fecha a conexão
            mobj_conexao.Close
        End If
    End If
    'retorna true
    pfct_desconectar_banco = True
fim_pfct_desconectar_banco:
    'destrói os objetos
    Set mobj_conexao = Nothing
    Exit Function
erro_pfct_desconectar_banco:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_bd", "pfct_desconectar_banco"
    GoTo fim_pfct_desconectar_banco
End Function

Public Function pfct_executar_comando_sql(ByRef pobj_objeto As Object, _
                                          ByVal pstr_sql As String, _
                                          ByVal pstr_modulo As String, _
                                          ByVal pstr_metodo As String) As Boolean
    On Error GoTo erro_pfct_executar_comando_sql
    Dim lstr_sql As String
    Dim lobj_retorno As Object
    'conecta ao banco
    If (pfct_conectar_banco()) Then
        'remove excesso de espaços da string sql
        lstr_sql = pfct_remover_excesso_espacos(pstr_sql)
        'executa o comando sql
        Set lobj_retorno = mobj_conexao.Execute(lstr_sql)
        'verifica se houve algum erro no comando
        If (mobj_conexao.LastError <> "") Then
            'se ocorreu erro ao executar o comando
            psub_gerar_log_sql lstr_sql, pstr_modulo, pstr_metodo, mobj_conexao.LastError
            GoTo fim_pfct_executar_comando_sql
        Else
            'se o comando foi executado corretamente
            psub_gerar_log_sql lstr_sql, pstr_modulo, pstr_metodo, ""
        End If
        'devolve o objeto por referência
        Set pobj_objeto = lobj_retorno
        'fecha a conexão com o banco
        If (Not pfct_desconectar_banco()) Then
            MsgBox "Erro ao fechar a conexão com o banco de dados.", vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_executar_comando_sql
        End If
    Else
        MsgBox "Erro ao abrir a conexão com o banco de dados.", vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_executar_comando_sql
    End If
    'retorna true
    pfct_executar_comando_sql = True
fim_pfct_executar_comando_sql:
    'destrói os objetos
    Set mobj_conexao = Nothing
    Exit Function
erro_pfct_executar_comando_sql:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_bd", "pfct_executar_comando_sql"
    GoTo fim_pfct_executar_comando_sql
End Function

Public Sub psub_limpar_banco()
    On Error GoTo erro_psub_limpar_banco
    Dim lobj_banco As Object
    Dim lstr_sql As String
    'monta o comando sql
    lstr_sql = "vacuum"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_banco, lstr_sql, "bas_bd", "psub_limpar_banco")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_psub_limpar_banco
    End If
fim_psub_limpar_banco:
    'destrói os objetos
    Set lobj_banco = Nothing
    Exit Sub
erro_psub_limpar_banco:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_bd", "psub_limpar_banco"
    GoTo fim_psub_limpar_banco
End Sub
