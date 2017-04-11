Attribute VB_Name = "bas_usuarios"
Option Explicit

Public Function pfct_validar_senha(ByVal plng_usuario As Long, ByVal pstr_senha As String) As Boolean
    On Error GoTo erro_pfct_valida_senha
    Dim lobj_usuario As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    'monta o comando sql
    lstr_sql = " select * from [tb_usuarios] where [int_codigo] = " & pfct_tratar_numero_sql(plng_usuario)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_usuario, lstr_sql, "bas_usuarios", "pfct_validar_senha")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_valida_senha
    End If
    'verifica se o usuário existe
    llng_registros = lobj_usuario.Count
    If (llng_registros > 0) Then
        If (pstr_senha = lobj_usuario(1)("str_senha")) Then
            pfct_validar_senha = True
        End If
    End If
fim_pfct_valida_senha:
    'destrói os objetos
    Set lobj_usuario = Nothing
    Exit Function
erro_pfct_valida_senha:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_usuarios", "pfct_valida_senha"
    GoTo fim_pfct_valida_senha
End Function

Public Function pfct_criar_usuario(ByVal pstr_usuario As String, ByVal pstr_senha As String, ByVal pstr_lembrete_senha As String) As Boolean
    On Error GoTo erro_pfct_criar_usuario
    Dim lobj_usuario As Object
    Dim lstr_sql As String
    'monta comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " insert into [tb_usuarios] "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & " [str_usuario], "
    lstr_sql = lstr_sql & " [str_senha], "
    lstr_sql = lstr_sql & " [str_lembrete_senha], "
    lstr_sql = lstr_sql & " [dt_criado_em], "
    lstr_sql = lstr_sql & " [tm_criado_em] "
    lstr_sql = lstr_sql & " ) "
    lstr_sql = lstr_sql & " values "
    lstr_sql = lstr_sql & " ( "
    lstr_sql = lstr_sql & " '" & pstr_usuario & "', "
    lstr_sql = lstr_sql & " '" & pstr_senha & "', "
    lstr_sql = lstr_sql & " '" & pstr_lembrete_senha & "', "
    lstr_sql = lstr_sql & " '" & Format$(Date, pcst_formato_data_sql) & "', "
    lstr_sql = lstr_sql & " '" & Format$(Time, pcst_formato_hora_sql) & "' "
    lstr_sql = lstr_sql & " ) "
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_usuario, lstr_sql, "bas_usuarios", "pfct_criar_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_criar_usuario
    End If
    'retorna true
    pfct_criar_usuario = True
fim_pfct_criar_usuario:
    'destrói os objetos
    Set lobj_usuario = Nothing
    Exit Function
erro_pfct_criar_usuario:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_usuarios", "pfct_criar_usuario"
    GoTo fim_pfct_criar_usuario
End Function

Public Sub psub_atualizar_usuario(ByVal plng_codigo As Long, ByVal pbln_atualizar_senha As Boolean)
    On Error GoTo erro_psub_atualizar_usuario
    Dim lobj_usuario As Object
    Dim lstr_sql As String
    'monta o comando sql
    lstr_sql = ""
    lstr_sql = lstr_sql & " update "
    lstr_sql = lstr_sql & " [tb_usuarios] "
    lstr_sql = lstr_sql & " set "
    If (pbln_atualizar_senha) Then
        lstr_sql = lstr_sql & " [str_senha] = '" & pfct_criptografia(p_usuario.str_senha) & "',"
        lstr_sql = lstr_sql & " [str_lembrete_senha] = '" & pfct_tratar_texto_sql(p_usuario.str_lembrete_senha) & "',"
    End If
    lstr_sql = lstr_sql & " [dt_ultimo_acesso] = '" & Format$(p_usuario.dt_ultimo_acesso, pcst_formato_data_sql) & "', "
    lstr_sql = lstr_sql & " [tm_ultimo_acesso] = '" & Format$(p_usuario.dt_ultimo_acesso, pcst_formato_hora_sql) & "' "
    lstr_sql = lstr_sql & " where "
    lstr_sql = lstr_sql & " [int_codigo] = " & pfct_tratar_numero_sql(plng_codigo)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_usuario, lstr_sql, "bas_usuarios", "psub_atualizar_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_psub_atualizar_usuario
    End If
fim_psub_atualizar_usuario:
    'destrói os objetos
    Set lobj_usuario = Nothing
    Exit Sub
erro_psub_atualizar_usuario:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_usuarios", "psub_atualizar_usuario"
    GoTo fim_psub_atualizar_usuario
End Sub

Public Function pfct_verificar_usuario_existe(ByVal pstr_usuario As String) As Boolean
    On Error GoTo erro_pfct_verificar_usuario_existe:
    Dim lobj_usuario As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    'monta o comando sql
    lstr_sql = "select * from [tb_usuarios] where [str_usuario] = '" & pstr_usuario & "'"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_usuario, lstr_sql, "bas_usuarios", "pfct_verificar_usuario_existe")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_verificar_usuario_existe
    End If
    'verifica se o usuário existe
    llng_registros = lobj_usuario.Count
    If (llng_registros >= 1) Then
        pfct_verificar_usuario_existe = True
    End If
fim_pfct_verificar_usuario_existe:
    'destrói os objetos
    Set lobj_usuario = Nothing
    Exit Function
erro_pfct_verificar_usuario_existe:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_usuarios", "pfct_verificar_usuario_existe"
    GoTo fim_pfct_verificar_usuario_existe
End Function

Public Function pfct_carregar_configuracoes_usuario(ByVal plng_codigo As Long) As Boolean
    On Error GoTo erro_pfct_carregar_configuracoes_usuario
    Dim lobj_config As Object
    Dim lobj_backup As Object
    Dim lobj_registro As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    ' -- tb_config -- '
    'monta o comando sql
    lstr_sql = "select * from [tb_config] where [int_usuario] = " & pfct_tratar_numero_sql(plng_codigo)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_config, lstr_sql, "bas_usuarios", "pfct_carregar_configuracoes_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_carregar_configuracoes_usuario
    End If
    'quantidade de registros
    llng_registros = lobj_config.Count
    If (llng_registros > 0) Then
        'se houver registros, carrega as configurações
        With p_usuario
            .id_intervalo_data = lobj_config(1)("int_intervalo_data")
            .sm_simbolo_moeda = lobj_config(1)("int_moeda")
            .bln_carregar_agenda_financeira_login = IIf(lobj_config(1)("chr_carregar_agenda_financeira_login") = "S", True, False)
            .bln_lancamentos_retroativos = IIf(lobj_config(1)("chr_lancamentos_retroativos") = "S", True, False)
            .bln_alteracoes_detalhes = IIf(lobj_config(1)("chr_alteracoes_detalhes") = "S", True, False)
            .bln_data_vencimento_baixa_imediata = IIf(lobj_config(1)("chr_data_vencimento_baixa_imediata") = "S", True, False)
            .bln_lancamentos_duplicados = IIf(lobj_config(1)("chr_lancamentos_duplicados") = "S", True, False)
            .bln_participou_pesquisa = IIf(lobj_config(1)("chr_participou_pesquisa") = "S", True, False)
        End With
    Else
        'se não houver registros, aplica os padrões
        With p_usuario
            .id_intervalo_data = id_30dias
            .sm_simbolo_moeda = sm_real
            .bln_carregar_agenda_financeira_login = True
            .bln_lancamentos_retroativos = False
            .bln_alteracoes_detalhes = False
            .bln_data_vencimento_baixa_imediata = False
            .bln_lancamentos_duplicados = False
            .bln_participou_pesquisa = False
        End With
    End If
    ' -- tb_backup -- '
    'monta o comando sql
    lstr_sql = "select * from [tb_backup] where [int_usuario] = " & pfct_tratar_numero_sql(p_usuario.lng_codigo)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_backup, lstr_sql, "bas_usuarios", "pfct_carregar_configuracoes_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_carregar_configuracoes_usuario
    End If
    'quantidade de registros
    llng_registros = lobj_backup.Count
    If (llng_registros > 0) Then
        'se houver registros, carrega as configurações
        'se não houver registros, aplica os padrões
        With p_backup
            .bln_ativar = IIf(lobj_backup(1)("chr_ativar") = "S", True, False)
            .pb_periodo_backup = lobj_backup(1)("int_periodo")
            .str_caminho = lobj_backup(1)("str_caminho")
            'último backup
            If (lobj_backup(1)("dt_ultimo_backup") <> "") Then
                .dt_ultimo_backup = CDate(lobj_backup(1)("dt_ultimo_backup") & " " & lobj_backup(1)("tm_ultimo_backup"))
            End If
            'próximo backup
            If (lobj_backup(1)("dt_proximo_backup") <> "") Then
                .dt_proximo_backup = CDate(lobj_backup(1)("dt_proximo_backup") & " " & lobj_backup(1)("tm_proximo_backup"))
            End If
        End With
    Else
        'se não houver registros, aplica os padrões
        With p_backup
            .bln_ativar = False 'por padrão do sistema, o backup é desativado
            .pb_periodo_backup = pb_selecione
            .str_caminho = ""
            .dt_ultimo_backup = "00:00:00"
            .dt_proximo_backup = "00:00:00"
        End With
    End If
    ' -- tb_registros -- '
    'monta o comando sql
    lstr_sql = "select * from [tb_registros] where [int_usuario] = " & pfct_tratar_numero_sql(p_usuario.lng_codigo)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_registro, lstr_sql, "bas_usuarios", "pfct_carregar_configuracoes_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_carregar_configuracoes_usuario
    End If
    'quantidade de registros
    llng_registros = lobj_registro.Count
    If (llng_registros > 0) Then
        With p_usuario
            .bln_participou_pesquisa = True
        End With
        With p_registro
            .str_nome = lobj_registro(1)("str_nome")
            .str_email = lobj_registro(1)("str_email")
            .str_pais = lobj_registro(1)("str_pais")
            .str_estado = lobj_registro(1)("str_estado")
            .str_cidade = lobj_registro(1)("str_cidade")
            If (lobj_registro(1)("dt_data_nascimento") <> "") Then
                .dt_data_nascimento = CDate(lobj_registro(1)("dt_data_nascimento"))
            End If
            .str_profissao = lobj_registro(1)("str_profissao")
            .chr_sexo = lobj_registro(1)("chr_sexo")
            .str_origem = lobj_registro(1)("str_origem")
            .str_opiniao = lobj_registro(1)("str_opiniao")
            .bln_newsletter = IIf(lobj_registro(1)("chr_newsletter") = "S", True, False)
            .str_id_cpu = lobj_registro(1)("str_id_cpu")
            .str_id_hd = lobj_registro(1)("str_id_hd")
            If (lobj_registro(1)("dt_data_registro") <> "") Then
                .dt_data_registro = CDate(lobj_registro(1)("dt_data_registro") & " " & lobj_registro(1)("tm_hora_registro"))
            End If
            If (lobj_registro(1)("dt_data_liberacao") <> "") Then
                .dt_data_liberacao = CDate(lobj_registro(1)("dt_data_liberacao") & " " & lobj_registro(1)("tm_hora_liberacao"))
            End If
            .bln_banido = IIf(lobj_registro(1)("chr_banido") = "S", True, False)
            .str_desc_banido = lobj_registro(1)("str_desc_banido")
        End With
    Else
        'se não houver registros, aplica os padrões
        With p_usuario
            .bln_participou_pesquisa = False
        End With
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
            .str_desc_banido = Empty
        End With
    End If
    'retorna true
    pfct_carregar_configuracoes_usuario = True
fim_pfct_carregar_configuracoes_usuario:
    Set lobj_config = Nothing
    Set lobj_backup = Nothing
    Set lobj_registro = Nothing
    Exit Function
erro_pfct_carregar_configuracoes_usuario:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_usuarios", "pfct_carregar_configuracoes_usuario"
    GoTo fim_pfct_carregar_configuracoes_usuario
    Resume 0
End Function

Public Function pfct_carregar_dados_usuario(ByVal plng_codigo As Long) As Boolean
    On Error GoTo erro_pfct_carregar_dados_usuario
    Dim lobj_dados As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    'monta o comando sql
    lstr_sql = "select * from [tb_usuarios] where [int_codigo] = " & pfct_tratar_numero_sql(plng_codigo)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_dados, lstr_sql, "bas_usuarios", "pfct_carregar_dados_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_carregar_dados_usuario
    Else
        'quantidade de registros
        llng_registros = lobj_dados.Count
        'se houver registros na tabela
        If (llng_registros > 0) Then
            'preenche os dados do objeto
            With p_usuario
                .lng_codigo = lobj_dados(1)("int_codigo")
                .str_login = lobj_dados(1)("str_usuario")
                .str_senha = pfct_criptografia(lobj_dados(1)("str_senha"))
                .str_lembrete_senha = lobj_dados(1)("str_lembrete_senha")
                .dt_criado_em = CDate(lobj_dados(1)("dt_criado_em"))
                If ((Not IsNull(lobj_dados(1)("dt_ultimo_acesso"))) And (Not IsNull(lobj_dados(1)("tm_ultimo_acesso")))) Then
                    .dt_ultimo_acesso = CDate(Format$(lobj_dados(1)("dt_ultimo_acesso"), "dd/mm/yyyy") & " " & Format$(lobj_dados(1)("tm_ultimo_acesso"), "hh:mm:ss"))
                End If
            End With
        End If
    End If
    pfct_carregar_dados_usuario = True
fim_pfct_carregar_dados_usuario:
    Set lobj_dados = Nothing
    Exit Function
erro_pfct_carregar_dados_usuario:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_usuarios", "pfct_carregar_dados_usuario"
    GoTo fim_pfct_carregar_dados_usuario
End Function

Public Function pfct_carregar_lembrete_usuario(ByVal plng_codigo As Long) As String
    On Error GoTo erro_pfct_carregar_lembrete_usuario
    Dim lobj_usuario As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    'monta o comando sql
    lstr_sql = "select [str_lembrete_senha] from [tb_usuarios] where [int_codigo] = " & pfct_tratar_numero_sql(plng_codigo)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_usuario, lstr_sql, "bas_usuarios", "pfct_carregar_lembrete_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_carregar_lembrete_usuario
    End If
    'quantidade de registros
    llng_registros = lobj_usuario.Count
    'se houver registros
    If (llng_registros > 0) Then
        pfct_carregar_lembrete_usuario = lobj_usuario(1)("str_lembrete_senha")
    Else
        pfct_carregar_lembrete_usuario = ""
    End If
fim_pfct_carregar_lembrete_usuario:
    Set lobj_usuario = Nothing
    Exit Function
erro_pfct_carregar_lembrete_usuario:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_usuarios", "pfct_carregar_lembrete_usuario"
    GoTo fim_pfct_carregar_lembrete_usuario
End Function

Public Function pfct_salvar_configuracoes_usuario(ByVal plng_codigo As Long) As Boolean
    On Error GoTo erro_pfct_salvar_configuracoes_usuario
    Dim lobj_config As Object
    Dim lobj_backup As Object
    Dim lobj_registro As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    'Dim ldt_data_backup As Date
    ' -- config -- '
    'monta o comando sql
    lstr_sql = "select * from [tb_config] where [int_usuario] = " & pfct_tratar_numero_sql(plng_codigo)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_config, lstr_sql, "bas_usuarios", "pfct_salvar_configuracoes_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_salvar_configuracoes_usuario
    End If
    'quantidade de registros
    llng_registros = lobj_config.Count
    'se não houver registros
    If (llng_registros = 0) Then
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
        lstr_sql = lstr_sql & " ) values ( "
        lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(p_usuario.lng_codigo) & ", "
        lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(p_usuario.sm_simbolo_moeda) & ", "
        lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(p_usuario.id_intervalo_data) & ", "
        lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(IIf(p_usuario.bln_carregar_agenda_financeira_login, "S", "N")) & "', "
        lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(IIf(p_usuario.bln_lancamentos_retroativos, "S", "N")) & "', "
        lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(IIf(p_usuario.bln_alteracoes_detalhes, "S", "N")) & "', "
        lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(IIf(p_usuario.bln_data_vencimento_baixa_imediata, "S", "N")) & "', "
        lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(IIf(p_usuario.bln_lancamentos_duplicados, "S", "N")) & "', "
        lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(IIf(p_usuario.bln_participou_pesquisa, "S", "N")) & "' "
        lstr_sql = lstr_sql & " ) "
    Else
        lstr_sql = ""
        lstr_sql = lstr_sql & " update [tb_config] set "
        lstr_sql = lstr_sql & " [int_moeda] = " & pfct_tratar_numero_sql(p_usuario.sm_simbolo_moeda) & ", "
        lstr_sql = lstr_sql & " [int_intervalo_data] = " & pfct_tratar_numero_sql(p_usuario.id_intervalo_data) & ", "
        lstr_sql = lstr_sql & " [chr_carregar_agenda_financeira_login] = '" & pfct_tratar_texto_sql(IIf(p_usuario.bln_carregar_agenda_financeira_login, "S", "N")) & "', "
        lstr_sql = lstr_sql & " [chr_lancamentos_retroativos] = '" & pfct_tratar_texto_sql(IIf(p_usuario.bln_lancamentos_retroativos, "S", "N")) & "', "
        lstr_sql = lstr_sql & " [chr_alteracoes_detalhes] = '" & pfct_tratar_texto_sql(IIf(p_usuario.bln_alteracoes_detalhes, "S", "N")) & "', "
        lstr_sql = lstr_sql & " [chr_data_vencimento_baixa_imediata] = '" & pfct_tratar_texto_sql(IIf(p_usuario.bln_data_vencimento_baixa_imediata, "S", "N")) & "', "
        lstr_sql = lstr_sql & " [chr_lancamentos_duplicados] = '" & pfct_tratar_texto_sql(IIf(p_usuario.bln_lancamentos_duplicados, "S", "N")) & "', "
        lstr_sql = lstr_sql & " [chr_participou_pesquisa] = '" & pfct_tratar_texto_sql(IIf(p_usuario.bln_participou_pesquisa, "S", "N")) & "' "
        lstr_sql = lstr_sql & " where "
        lstr_sql = lstr_sql & " [int_usuario] = " & pfct_tratar_numero_sql(p_usuario.lng_codigo) & " "
    End If
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_config, lstr_sql, "bas_usuarios", "pfct_salvar_configuracoes_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_salvar_configuracoes_usuario
    End If
    ' -- backup -- '
    'monta o comando sql
    lstr_sql = "select * from [tb_backup] where [int_usuario] = " & pfct_tratar_numero_sql(p_usuario.lng_codigo)
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_backup, lstr_sql, "bas_usuarios", "pfct_salvar_configuracoes_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_salvar_configuracoes_usuario
    End If
    'quantidade de registros
    llng_registros = lobj_backup.Count
    'se não houver registros
    If (llng_registros = 0) Then
        lstr_sql = ""
        lstr_sql = lstr_sql & " insert into [tb_backup] "
        lstr_sql = lstr_sql & " ( "
        lstr_sql = lstr_sql & " [int_usuario], "
        lstr_sql = lstr_sql & " [chr_ativar], "
        lstr_sql = lstr_sql & " [int_periodo], "
        lstr_sql = lstr_sql & " [str_caminho], "
        lstr_sql = lstr_sql & " [dt_ultimo_backup], "
        lstr_sql = lstr_sql & " [tm_ultimo_backup], "
        lstr_sql = lstr_sql & " [dt_proximo_backup], "
        lstr_sql = lstr_sql & " [tm_proximo_backup] "
        lstr_sql = lstr_sql & " ) values ( "
        lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(p_usuario.lng_codigo) & ", "
        lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(IIf(p_backup.bln_ativar, "S", "N")) & "', "
        lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(p_backup.pb_periodo_backup) & ", "
        lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(p_backup.str_caminho) & "', "
        'último backup
        If (p_backup.dt_ultimo_backup <> "00:00:00") Then
            lstr_sql = lstr_sql & " '" & Format$(p_backup.dt_ultimo_backup, pcst_formato_data_sql) & "', "  'data
            lstr_sql = lstr_sql & " '" & Format$(p_backup.dt_ultimo_backup, pcst_formato_hora_sql) & "', "  'hora
        Else
            lstr_sql = lstr_sql & " '', "   'data
            lstr_sql = lstr_sql & " '', "   'hora
        End If
        'próximo backup
        If (p_backup.dt_proximo_backup <> "00:00:00") Then
            lstr_sql = lstr_sql & " '" & Format$(p_backup.dt_proximo_backup, pcst_formato_data_sql) & "', "    'data
            lstr_sql = lstr_sql & " '" & Format$(p_backup.dt_proximo_backup, pcst_formato_hora_sql) & "' "  'hora
        Else
            lstr_sql = lstr_sql & " '', "   'data
            lstr_sql = lstr_sql & " '' "    'hora
        End If
        lstr_sql = lstr_sql & " ) "
    Else
        lstr_sql = ""
        lstr_sql = lstr_sql & " update [tb_backup] set "
        lstr_sql = lstr_sql & " [chr_ativar] = '" & pfct_tratar_texto_sql(IIf(p_backup.bln_ativar, "S", "N")) & "', "
        lstr_sql = lstr_sql & " [int_periodo] = " & pfct_tratar_numero_sql(p_backup.pb_periodo_backup) & ", "
        lstr_sql = lstr_sql & " [str_caminho] = '" & pfct_tratar_texto_sql(p_backup.str_caminho) & "', "
        'último backup
        If (p_backup.dt_ultimo_backup <> "00:00:00") Then
            lstr_sql = lstr_sql & " [dt_ultimo_backup] = '" & Format$(p_backup.dt_ultimo_backup, pcst_formato_data_sql) & "', "    'data
            lstr_sql = lstr_sql & " [tm_ultimo_backup] = '" & Format$(p_backup.dt_ultimo_backup, pcst_formato_hora_sql) & "', "  'hora
        Else
            lstr_sql = lstr_sql & " [dt_ultimo_backup] = '', "  'data
            lstr_sql = lstr_sql & " [tm_ultimo_backup] = '', "   'hora
        End If
        'próximo backup
        If (p_backup.dt_proximo_backup <> "00:00:00") Then
            lstr_sql = lstr_sql & " [dt_proximo_backup] = '" & Format$(p_backup.dt_proximo_backup, pcst_formato_data_sql) & "', "   'data
            lstr_sql = lstr_sql & " [tm_proximo_backup] = '" & Format$(p_backup.dt_proximo_backup, pcst_formato_hora_sql) & "' "    'hora
        Else
            lstr_sql = lstr_sql & " [dt_proximo_backup] = '', "    'data
            lstr_sql = lstr_sql & " [tm_proximo_backup] = '' "  'hora
        End If
        lstr_sql = lstr_sql & " where "
        lstr_sql = lstr_sql & " [int_usuario] = " & pfct_tratar_numero_sql(p_usuario.lng_codigo) & " "
    End If
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_backup, lstr_sql, "bas_usuarios", "pfct_salvar_configuracoes_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_salvar_configuracoes_usuario
    End If
    ' -- registro -- '
    If (p_usuario.bln_participou_pesquisa) Then
        'monta o comando sql
        lstr_sql = "select * from [tb_registros] where [int_usuario] = " & pfct_tratar_numero_sql(p_usuario.lng_codigo)
        'executa o comando sql e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_registro, lstr_sql, "bas_usuarios", "pfct_salvar_configuracoes_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_salvar_configuracoes_usuario
        End If
        'quantidade de registros
        llng_registros = lobj_registro.Count
        'se não houver registros
        If (llng_registros = 0) Then
            lstr_sql = ""
            lstr_sql = lstr_sql & " insert into tb_registros "
            lstr_sql = lstr_sql & " ( "
            lstr_sql = lstr_sql & " [int_usuario], "
            lstr_sql = lstr_sql & " [str_nome], "
            lstr_sql = lstr_sql & " [str_email], "
            lstr_sql = lstr_sql & " [str_pais], "
            lstr_sql = lstr_sql & " [str_estado], "
            lstr_sql = lstr_sql & " [str_cidade], "
            lstr_sql = lstr_sql & " [dt_data_nascimento], "
            lstr_sql = lstr_sql & " [str_profissao], "
            lstr_sql = lstr_sql & " [chr_sexo], "
            lstr_sql = lstr_sql & " [str_origem], "
            lstr_sql = lstr_sql & " [str_opiniao], "
            lstr_sql = lstr_sql & " [chr_newsletter], "
            lstr_sql = lstr_sql & " [str_id_cpu], "
            lstr_sql = lstr_sql & " [str_id_hd], "
            lstr_sql = lstr_sql & " [dt_data_registro], "
            lstr_sql = lstr_sql & " [tm_hora_registro], "
            lstr_sql = lstr_sql & " [dt_data_liberacao], "
            lstr_sql = lstr_sql & " [tm_hora_liberacao], "
            lstr_sql = lstr_sql & " [chr_banido], "
            lstr_sql = lstr_sql & " [str_desc_banido] "
            lstr_sql = lstr_sql & " ) "
            lstr_sql = lstr_sql & " values "
            lstr_sql = lstr_sql & " ( "
            lstr_sql = lstr_sql & " " & pfct_tratar_numero_sql(p_usuario.lng_codigo) & ", "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(p_registro.str_nome) & "', "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(p_registro.str_email) & "', "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(p_registro.str_pais) & "', "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(p_registro.str_estado) & "', "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(p_registro.str_cidade) & "', "
            If (Format$(p_registro.dt_data_nascimento, "dd/mm/yyyy hh:mm:ss") <> "30/12/1899 00:00:00") Then
                lstr_sql = lstr_sql & " '" & Format$(p_registro.dt_data_nascimento, pcst_formato_data_sql) & "', "
            Else
                lstr_sql = lstr_sql & " '', "    'data
            End If
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(p_registro.str_profissao) & "', "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(p_registro.chr_sexo) & "', "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(p_registro.str_origem) & "', "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(p_registro.str_opiniao) & "', "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(IIf(p_registro.bln_newsletter, "S", "N")) & "', "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(p_registro.str_id_cpu) & "', "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(p_registro.str_id_hd) & "', "
            If (Format$(p_registro.dt_data_registro, "dd/mm/yyyy hh:mm:ss") <> "30/12/1899 00:00:00") Then
                lstr_sql = lstr_sql & " '" & Format$(p_registro.dt_data_registro, pcst_formato_data_sql) & "', "
                lstr_sql = lstr_sql & " '" & Format$(p_registro.dt_data_registro, pcst_formato_hora_sql) & "', "
            Else
                lstr_sql = lstr_sql & " '', "    'data
                lstr_sql = lstr_sql & " '', "    'hora
            End If
            If (Format$(p_registro.dt_data_liberacao, "dd/mm/yyyy hh:mm:ss") <> "30/12/1899 00:00:00") Then
                lstr_sql = lstr_sql & " '" & Format$(p_registro.dt_data_liberacao, pcst_formato_data_sql) & "', "
                lstr_sql = lstr_sql & " '" & Format$(p_registro.dt_data_liberacao, pcst_formato_hora_sql) & "', "
            Else
                lstr_sql = lstr_sql & " '', "    'data
                lstr_sql = lstr_sql & " '', "    'hora
            End If
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(IIf(p_registro.bln_banido, "S", "N")) & "', "
            lstr_sql = lstr_sql & " '" & pfct_tratar_texto_sql(p_registro.str_desc_banido) & "' "
            lstr_sql = lstr_sql & " ) "
        Else
            lstr_sql = ""
            lstr_sql = lstr_sql & " update [tb_registros] set "
            lstr_sql = lstr_sql & " [str_nome] = '" & pfct_tratar_texto_sql(p_registro.str_nome) & "', "
            lstr_sql = lstr_sql & " [str_email] = '" & pfct_tratar_texto_sql(p_registro.str_email) & "', "
            lstr_sql = lstr_sql & " [str_pais] = '" & pfct_tratar_texto_sql(p_registro.str_pais) & "', "
            lstr_sql = lstr_sql & " [str_estado] = '" & pfct_tratar_texto_sql(p_registro.str_estado) & "', "
            lstr_sql = lstr_sql & " [str_cidade] = '" & pfct_tratar_texto_sql(p_registro.str_cidade) & "', "
            If (Format$(p_registro.dt_data_nascimento, "dd/mm/yyyy hh:mm:ss") <> "30/12/1899 00:00:00") Then
                lstr_sql = lstr_sql & " [dt_data_nascimento] = '" & Format$(p_registro.dt_data_nascimento, pcst_formato_data_sql) & "', "
            Else
                lstr_sql = lstr_sql & " [dt_data_nascimento] = '', "
            End If
            lstr_sql = lstr_sql & " [str_profissao] = '" & pfct_tratar_texto_sql(p_registro.str_profissao) & "', "
            lstr_sql = lstr_sql & " [chr_sexo] = '" & pfct_tratar_texto_sql(p_registro.chr_sexo) & "', "
            lstr_sql = lstr_sql & " [str_origem] = '" & pfct_tratar_texto_sql(p_registro.str_origem) & "', "
            lstr_sql = lstr_sql & " [str_opiniao] = '" & pfct_tratar_texto_sql(p_registro.str_opiniao) & "', "
            lstr_sql = lstr_sql & " [chr_newsletter] = '" & pfct_tratar_texto_sql(IIf(p_registro.bln_newsletter, "S", "N")) & "', "
            lstr_sql = lstr_sql & " [str_id_cpu] = '" & pfct_tratar_texto_sql(p_registro.str_id_cpu) & "', "
            lstr_sql = lstr_sql & " [str_id_hd] = '" & pfct_tratar_texto_sql(p_registro.str_id_hd) & "', "
            If (Format$(p_registro.dt_data_registro, "dd/mm/yyyy hh:mm:ss") <> "30/12/1899 00:00:00") Then
                lstr_sql = lstr_sql & " [dt_data_registro] = '" & Format$(p_registro.dt_data_registro, pcst_formato_data_sql) & "', "
                lstr_sql = lstr_sql & " [tm_hora_registro] = '" & Format$(p_registro.dt_data_registro, pcst_formato_hora_sql) & "', "
            Else
                lstr_sql = lstr_sql & " [dt_data_registro] = '', "
                lstr_sql = lstr_sql & " [tm_hora_registro] = '', "
            End If
            If (Format$(p_registro.dt_data_liberacao, "dd/mm/yyyy hh:mm:ss") <> "30/12/1899 00:00:00") Then
                lstr_sql = lstr_sql & " [dt_data_liberacao] = '" & Format$(p_registro.dt_data_liberacao, pcst_formato_data_sql) & "', "
                lstr_sql = lstr_sql & " [tm_hora_liberacao] = '" & Format$(p_registro.dt_data_liberacao, pcst_formato_hora_sql) & "', "
            Else
                lstr_sql = lstr_sql & " [dt_data_liberacao] = '', "
                lstr_sql = lstr_sql & " [tm_hora_liberacao] = '', "
            End If
            lstr_sql = lstr_sql & " [chr_banido] = '" & pfct_tratar_texto_sql(IIf(p_registro.bln_banido, "S", "N")) & "', "
            lstr_sql = lstr_sql & " [str_desc_banido] = '" & pfct_tratar_texto_sql(p_registro.str_desc_banido) & "' "
            lstr_sql = lstr_sql & " where "
            lstr_sql = lstr_sql & " [int_usuario] = " & pfct_tratar_numero_sql(p_usuario.lng_codigo) & " "
        End If
        'executa o comando sql e devolve o objeto
        If (Not pfct_executar_comando_sql(lobj_registro, lstr_sql, "bas_usuarios", "pfct_salvar_configuracoes_usuario")) Then
            MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
            GoTo fim_pfct_salvar_configuracoes_usuario
        End If
    End If
    'retorna true
    pfct_salvar_configuracoes_usuario = True
fim_pfct_salvar_configuracoes_usuario:
    Set lobj_config = Nothing
    Set lobj_backup = Nothing
    Set lobj_registro = Nothing
    Exit Function
erro_pfct_salvar_configuracoes_usuario:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_usuarios", "pfct_salvar_configuracoes_usuario"
    GoTo fim_pfct_salvar_configuracoes_usuario
End Function

Public Function pfct_excluir_usuario(ByVal plng_usuario As Long) As Boolean
    On Error GoTo erro_pfct_excluir_usuario
    Dim lobj_usuario As Object
    Dim lstr_sql As String
    'tb_backup
    lstr_sql = "delete from [tb_backup] where [int_usuario] = " & pfct_tratar_numero_sql(plng_usuario)
    If (Not pfct_executar_comando_sql(lobj_usuario, lstr_sql, "bas_usuarios", "pfct_excluir_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_excluir_usuario
    End If
    'tb_config
    lstr_sql = "delete from [tb_config] where [int_usuario] = " & pfct_tratar_numero_sql(plng_usuario)
    If (Not pfct_executar_comando_sql(lobj_usuario, lstr_sql, "bas_usuarios", "pfct_excluir_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_excluir_usuario
    End If
    'tb_usuarios
    lstr_sql = "delete from [tb_usuarios] where [int_codigo] = " & pfct_tratar_numero_sql(plng_usuario)
    If (Not pfct_executar_comando_sql(lobj_usuario, lstr_sql, "bas_usuarios", "pfct_excluir_usuario")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_pfct_excluir_usuario
    End If
    'após excluir os dados das tabelas, elimina o arquivo
    If (pfct_excluir_arquivo(p_banco.str_caminho_dados_usuario)) Then
        'retorna true
        pfct_excluir_usuario = True
    End If
fim_pfct_excluir_usuario:
    'destrói os objetos
    Set lobj_usuario = Nothing
    Exit Function
erro_pfct_excluir_usuario:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_usuarios", "pfct_excluir_usuario"
    GoTo fim_pfct_excluir_usuario
End Function
