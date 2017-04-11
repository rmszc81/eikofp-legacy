Attribute VB_Name = "bas_combos"
Option Explicit

Public Enum enm_em_ordem
    op_selecione = 0
    op_crescente = 1
    op_decrescente = 2
End Enum

Public Sub psub_preencher_usuarios(ByRef pobj_combo As Object)
    On Error GoTo erro_psub_preencher_usuarios
    Dim lobj_usuarios As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    'monta o comando sql
    lstr_sql = "select * from [tb_usuarios] order by [str_usuario] asc"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_usuarios, lstr_sql, "bas_combos", "psub_preencher_usuarios")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_psub_preencher_usuarios
    End If
    llng_registros = lobj_usuarios.Count
    pobj_combo.Clear
    pobj_combo.AddItem "- Selecione -", 0
    pobj_combo.ItemData(0) = 0
    For llng_contador = 1 To llng_registros
        pobj_combo.AddItem lobj_usuarios(llng_contador)("str_usuario"), llng_contador
        pobj_combo.ItemData(llng_contador) = lobj_usuarios(llng_contador)("int_codigo")
    Next
    pobj_combo.ListIndex = 0
fim_psub_preencher_usuarios:
    'destrói os objetos
    Set lobj_usuarios = Nothing
    Exit Sub
erro_psub_preencher_usuarios:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_combos", "psub_preencher_usuarios"
    GoTo fim_psub_preencher_usuarios
End Sub

Public Sub psub_preencher_contas(ByRef pobj_combo As Object, Optional pstr_primeiro_item As String)
    On Error GoTo erro_psub_preencher_contas
    Dim lobj_contas As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    'monta o comando sql
    lstr_sql = "select * from [tb_contas] where [chr_ativo] = 'S' order by [str_descricao] asc"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_contas, lstr_sql, "bas_combos", "psub_preencher_contas")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_psub_preencher_contas
    End If
    llng_registros = lobj_contas.Count
    pobj_combo.Clear
    If (pstr_primeiro_item <> "") Then
        pobj_combo.AddItem pstr_primeiro_item, 0
    Else
        pobj_combo.AddItem "- Selecione -", 0
    End If
    pobj_combo.ItemData(0) = 0
    For llng_contador = 1 To llng_registros
        pobj_combo.AddItem lobj_contas(llng_contador)("str_descricao"), llng_contador
        pobj_combo.ItemData(llng_contador) = lobj_contas(llng_contador)("int_codigo")
    Next
    pobj_combo.ListIndex = 0
fim_psub_preencher_contas:
    'destrói os objetos
    Set lobj_contas = Nothing
    Exit Sub
erro_psub_preencher_contas:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_combos", "psub_preencher_contas"
    GoTo fim_psub_preencher_contas
End Sub

'Public Sub psub_preencher_categorias(ByRef pobj_combo As Object, Optional pstr_primeiro_item As String)
'    On Error GoTo erro_psub_preencher_categorias
'    Dim lobj_categorias As Object
'    Dim lstr_sql As String
'    Dim llng_registros As Long
'    Dim llng_contador As Long
'    'monta o comando sql
'    lstr_sql = "select * from [tb_categorias] where [chr_ativo] = 'S' order by [str_descricao] asc"
'    'executa o comando sql e devolve o objeto
'    If (Not pfct_executar_comando_sql(lobj_categorias, lstr_sql, "bas_combos", "psub_preencher_categorias")) Then
'        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
'        GoTo fim_psub_preencher_categorias
'    End If
'    llng_registros = lobj_categorias.Count
'    pobj_combo.Clear
'    If (pstr_primeiro_item <> "") Then
'        pobj_combo.AddItem pstr_primeiro_item, 0
'    Else
'        pobj_combo.AddItem "- Selecione -", 0
'    End If
'    pobj_combo.ItemData(0) = 0
'    For llng_contador = 1 To llng_registros
'        pobj_combo.AddItem lobj_categorias(llng_contador)("str_descricao"), llng_contador
'        pobj_combo.ItemData(llng_contador) = lobj_categorias(llng_contador)("int_codigo")
'    Next
'    pobj_combo.ListIndex = 0
'fim_psub_preencher_categorias:
'    'destrói os objetos
'    Set lobj_categorias = Nothing
'    Exit Sub
'erro_psub_preencher_categorias:
'    psub_gerar_log_erro Err.Number, Err.Description, "bas_combos", "psub_preencher_categorias"
'    GoTo fim_psub_preencher_categorias
'End Sub

Public Sub psub_preencher_receitas(ByRef pobj_componente As Object, ByVal pbln_lista As Boolean)
    On Error GoTo erro_psub_preencher_receitas
    Dim lobj_receitas As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    'monta o comando sql
    lstr_sql = "select * from [tb_receitas] where [chr_ativo] = 'S' order by [str_descricao] asc"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_receitas, lstr_sql, "bas_combos", "psub_preencher_receitas")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_psub_preencher_receitas
    End If
    llng_registros = lobj_receitas.Count
    pobj_componente.Clear
    'tratamento específico para combo box
    If (Not pbln_lista) Then
        pobj_componente.AddItem "- Selecione -", 0
        pobj_componente.ItemData(0) = 0
    End If
    'percorre o objeto e adiciona os itens
    For llng_contador = 1 To llng_registros
        pobj_componente.AddItem lobj_receitas(llng_contador)("str_descricao")
        pobj_componente.ItemData(pobj_componente.NewIndex) = lobj_receitas(llng_contador)("int_codigo")
    Next
    pobj_componente.ListIndex = 0
fim_psub_preencher_receitas:
    'destrói os objetos
    Set lobj_receitas = Nothing
    Exit Sub
erro_psub_preencher_receitas:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_combos", "psub_preencher_receitas"
    GoTo fim_psub_preencher_receitas
End Sub

Public Sub psub_preencher_despesas(ByRef pobj_componente As Object, ByVal pbln_lista As Boolean)
    On Error GoTo erro_psub_preencher_despesas
    Dim lobj_despesas As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    'monta o comando sql
    lstr_sql = "select * from [tb_despesas] where [chr_ativo] = 'S' order by [str_descricao] asc"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_despesas, lstr_sql, "bas_combos", "psub_preencher_despesas")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_psub_preencher_despesas
    End If
    llng_registros = lobj_despesas.Count
    pobj_componente.Clear
    'tratamento específico para combo box
    If (Not pbln_lista) Then
        pobj_componente.AddItem "- Selecione -", 0
        pobj_componente.ItemData(0) = 0
    End If
    'percorre o objeto e adiciona os itens
    For llng_contador = 1 To llng_registros
        pobj_componente.AddItem lobj_despesas(llng_contador)("str_descricao")
        pobj_componente.ItemData(pobj_componente.NewIndex) = lobj_despesas(llng_contador)("int_codigo")
    Next
    pobj_componente.ListIndex = 0
fim_psub_preencher_despesas:
    'destrói os objetos
    Set lobj_despesas = Nothing
    Exit Sub
erro_psub_preencher_despesas:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_combos", "psub_preencher_despesas"
    GoTo fim_psub_preencher_despesas
End Sub

Public Sub psub_preencher_formas_pagamento(ByRef pobj_combo As Object)
    On Error GoTo erro_psub_preencher_formas_pagamento
    Dim lobj_formas_pagamento As Object
    Dim lstr_sql As String
    Dim llng_registros As Long
    Dim llng_contador As Long
    'monta o comando sql
    lstr_sql = "select * from [tb_formas_pagamento] where [chr_ativo] = 'S' order by [str_descricao] asc"
    'executa o comando sql e devolve o objeto
    If (Not pfct_executar_comando_sql(lobj_formas_pagamento, lstr_sql, "bas_combos", "psub_preencher_formas_pagamento")) Then
        MsgBox "Erro ao executar comando SQL.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        GoTo fim_psub_preencher_formas_pagamento
    End If
    llng_registros = lobj_formas_pagamento.Count
    pobj_combo.Clear
    pobj_combo.AddItem "- Selecione -", 0
    pobj_combo.ItemData(0) = 0
    For llng_contador = 1 To llng_registros
        pobj_combo.AddItem lobj_formas_pagamento(llng_contador)("str_descricao"), llng_contador
        pobj_combo.ItemData(llng_contador) = lobj_formas_pagamento(llng_contador)("int_codigo")
    Next
    pobj_combo.ListIndex = 0
fim_psub_preencher_formas_pagamento:
    'destrói os objetos
    Set lobj_formas_pagamento = Nothing
    Exit Sub
erro_psub_preencher_formas_pagamento:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_combos", "psub_preencher_formas_pagamento"
    GoTo fim_psub_preencher_formas_pagamento
End Sub

Public Sub psub_preencher_tempo(ByRef pobj_combo As Object)
    On Error GoTo erro_psub_preencher_tempo
    With pobj_combo
        .Clear
        .AddItem "- Selecione - ", 0
        .AddItem "- Meses", 1
        .AddItem "- Anos", 2
        .ItemData(0) = 0 'selecione
        .ItemData(1) = 30 'meses
        .ItemData(2) = 365 'anos
        .ListIndex = 0
    End With
fim_psub_preencher_tempo:
    Exit Sub
erro_psub_preencher_tempo:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_combos", "psub_preencher_tempo"
    GoTo fim_psub_preencher_tempo
End Sub

Public Sub psub_preencher_ordem(ByRef pobj_combo As Object)
    On Error GoTo erro_psub_preencher_ordem
    With pobj_combo
        .Clear
        .AddItem "- Selecione a ordem -", enm_em_ordem.op_selecione
        .AddItem "- Crescente", enm_em_ordem.op_crescente
        .AddItem "- Decrescente", enm_em_ordem.op_decrescente
        .ListIndex = enm_em_ordem.op_selecione
    End With
fim_psub_preencher_ordem:
    Exit Sub
erro_psub_preencher_ordem:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_combos", "psub_preencher_ordem"
    GoTo fim_psub_preencher_ordem
End Sub

Public Sub psub_preencher_simbolos_moeda(ByRef pobj_combo As Object)
    On Error GoTo erro_psub_preencher_simbolos_moeda
    With pobj_combo
        .Clear
        .AddItem "- Selecione a moeda - ", enm_simbolo_moeda.sm_selecione
        .AddItem "- DÓLAR (US$)", enm_simbolo_moeda.sm_dolar
        .AddItem "- EURO (€$)", enm_simbolo_moeda.sm_euro
        .AddItem "- REAL (R$)", enm_simbolo_moeda.sm_real
        .AddItem "- IENE (¥$)", enm_simbolo_moeda.sm_iene
        .ListIndex = enm_simbolo_moeda.sm_selecione
    End With
fim_psub_preencher_simbolos_moeda:
    Exit Sub
erro_psub_preencher_simbolos_moeda:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_combos", "psub_preencher_simbolos_moeda"
    GoTo fim_psub_preencher_simbolos_moeda
End Sub

Public Sub psub_preencher_intervalo_data(ByRef pobj_combo As Object)
    On Error GoTo erro_psub_preencher_intervalo_data
    With pobj_combo
        .Clear
        .AddItem "- Selecione o intervalo - ", enm_intervalo_data.id_selecione
        .AddItem "- 30 DIAS", enm_intervalo_data.id_30dias
        .AddItem "- 60 DIAS", enm_intervalo_data.id_60dias
        .AddItem "- 90 DIAS", enm_intervalo_data.id_90dias
        .AddItem "- 120 DIAS", enm_intervalo_data.id_120dias
        .ListIndex = enm_intervalo_data.id_selecione
    End With
fim_psub_preencher_intervalo_data:
    Exit Sub
erro_psub_preencher_intervalo_data:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_combos", "psub_preencher_intervalo_data"
    GoTo fim_psub_preencher_intervalo_data
End Sub

Public Sub psub_preencher_periodo_backup(ByRef pobj_combo As Object)
    On Error GoTo erro_psub_preencher_periodo_backup
    With pobj_combo
        .Clear
        .AddItem "- Selecione o período - ", enm_periodo_backup.pb_selecione
        .AddItem "- DIÁRIO", enm_periodo_backup.pb_diario
        .AddItem "- SEMANAL", enm_periodo_backup.pb_semanal
        .AddItem "- QUINZENAL", enm_periodo_backup.pb_quinzenal
        .AddItem "- MENSAL", enm_periodo_backup.pb_mensal
        .ListIndex = enm_periodo_backup.pb_selecione
    End With
fim_psub_preencher_periodo_backup:
    Exit Sub
erro_psub_preencher_periodo_backup:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_combos", "psub_preencher_periodo_backup"
    GoTo fim_psub_preencher_periodo_backup
End Sub

Public Sub psub_ajustar_combos_data(ByRef pobj_de As DTPicker, ByRef pobj_ate As DTPicker)
    On Error GoTo erro_psub_ajustar_combos
    Dim llng_intervalo As Long
    Select Case p_usuario.id_intervalo_data
        Case enm_intervalo_data.id_30dias
            llng_intervalo = 15
        Case enm_intervalo_data.id_60dias
            llng_intervalo = 30
        Case enm_intervalo_data.id_90dias
            llng_intervalo = 45
        Case enm_intervalo_data.id_120dias
            llng_intervalo = 60
        Case Else
            'assume o padrão 30 dias
            llng_intervalo = 15
    End Select
    pobj_de.Value = CDate(Date - llng_intervalo)
    pobj_ate.Value = CDate(Date + llng_intervalo)
fim_psub_ajustar_combos:
    Exit Sub
erro_psub_ajustar_combos:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_combos", "psub_ajustar_combos_data"
    GoTo fim_psub_ajustar_combos
End Sub
