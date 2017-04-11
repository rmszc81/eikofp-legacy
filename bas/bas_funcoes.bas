Attribute VB_Name = "bas_funcoes"
Option Explicit

Public Function pfct_ajustar_caminho_banco(ByVal penm_tipo_banco As enm_tipo_banco) As Boolean
    On Error GoTo erro_pfct_ajustar_caminho_banco
    'config
    If (penm_tipo_banco = tb_config) Then
        p_banco.str_caminho_dados_config = p_banco.str_caminho_config & "config.db" '.s3db
    End If
    'usuário
    If (penm_tipo_banco = tb_dados) Then
        p_banco.str_caminho_dados_usuario = p_banco.str_caminho_dados & LCase$(p_usuario.str_login) & ".db" '.s3db
    End If
    'backup
    If (penm_tipo_banco = tb_backup) Then
        p_banco.str_caminho_dados_backup = p_banco.str_caminho_backup & LCase$(p_backup.str_nome) & ".db" '.s3db
    End If
    'restaurar
    If (penm_tipo_banco = tb_restaurar) Then
        p_banco.str_caminho_dados_restaurar = p_banco.str_caminho_backup & LCase$(p_backup.str_nome) 'não precisa da extensão
    End If
    'retorna true
    pfct_ajustar_caminho_banco = True
fim_pfct_ajustar_caminho_banco:
    Exit Function
erro_pfct_ajustar_caminho_banco:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_ajustar_caminho_banco"
    GoTo fim_pfct_ajustar_caminho_banco
End Function

Public Function pfct_comparar_registros(ByRef ptpe_local As tpe_registro, ByRef ptpe_online As tpe_registro) As Boolean
    On Error GoTo Erro_pfct_comparar_registros
    If (ptpe_local.str_nome <> ptpe_online.str_nome) Then GoTo Fim_pfct_comparar_registros
    If (ptpe_local.str_email <> ptpe_online.str_email) Then GoTo Fim_pfct_comparar_registros
    If (ptpe_local.str_pais <> ptpe_online.str_pais) Then GoTo Fim_pfct_comparar_registros
    If (ptpe_local.str_estado <> ptpe_online.str_estado) Then GoTo Fim_pfct_comparar_registros
    If (ptpe_local.str_cidade <> ptpe_online.str_cidade) Then GoTo Fim_pfct_comparar_registros
    If (ptpe_local.dt_data_nascimento <> ptpe_online.dt_data_nascimento) Then GoTo Fim_pfct_comparar_registros
    If (ptpe_local.str_profissao <> ptpe_online.str_profissao) Then GoTo Fim_pfct_comparar_registros
    If (ptpe_local.chr_sexo <> ptpe_online.chr_sexo) Then GoTo Fim_pfct_comparar_registros
    If (ptpe_local.str_origem <> ptpe_online.str_origem) Then GoTo Fim_pfct_comparar_registros
    If (ptpe_local.str_opiniao <> ptpe_online.str_opiniao) Then GoTo Fim_pfct_comparar_registros
    If (ptpe_local.bln_newsletter <> ptpe_online.bln_newsletter) Then GoTo Fim_pfct_comparar_registros
    If (ptpe_local.str_id_cpu <> ptpe_online.str_id_cpu) Then GoTo Fim_pfct_comparar_registros
    If (ptpe_local.str_id_hd <> ptpe_online.str_id_hd) Then GoTo Fim_pfct_comparar_registros
    If (ptpe_local.bln_banido <> ptpe_online.bln_banido) Then GoTo Fim_pfct_comparar_registros
    If (ptpe_local.str_desc_banido <> ptpe_online.str_desc_banido) Then GoTo Fim_pfct_comparar_registros
    pfct_comparar_registros = True
Fim_pfct_comparar_registros:
    Exit Function
Erro_pfct_comparar_registros:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_comparar_registros"
    GoTo Fim_pfct_comparar_registros
End Function

Public Function pfct_copiar_registros(ByRef ptpe_local As tpe_registro, ByRef ptpe_online As tpe_registro) As Boolean
    On Error GoTo Erro_pfct_copiar_registros
    ptpe_local.int_codigo = ptpe_online.int_codigo
    ptpe_local.str_usuario = ptpe_online.str_usuario
    ptpe_local.str_nome = ptpe_online.str_nome
    ptpe_local.str_email = ptpe_online.str_email
    ptpe_local.str_pais = ptpe_online.str_pais
    ptpe_local.str_estado = ptpe_online.str_estado
    ptpe_local.str_cidade = ptpe_online.str_cidade
    ptpe_local.dt_data_nascimento = ptpe_online.dt_data_nascimento
    ptpe_local.str_profissao = ptpe_online.str_profissao
    ptpe_local.chr_sexo = ptpe_online.chr_sexo
    ptpe_local.str_origem = ptpe_online.str_origem
    ptpe_local.str_opiniao = ptpe_online.str_opiniao
    ptpe_local.bln_newsletter = ptpe_online.bln_newsletter
    ptpe_local.str_id_cpu = ptpe_online.str_id_cpu
    ptpe_local.str_id_hd = ptpe_online.str_id_hd
    ptpe_local.dt_data_registro = ptpe_online.dt_data_registro
    ptpe_local.dt_data_liberacao = ptpe_online.dt_data_liberacao
    ptpe_local.bln_banido = ptpe_online.bln_banido
    ptpe_local.str_desc_banido = ptpe_online.str_desc_banido
    pfct_copiar_registros = True
Fim_pfct_copiar_registros:
    Exit Function
Erro_pfct_copiar_registros:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_copiar_registros"
    GoTo Fim_pfct_copiar_registros
End Function

Public Function pfct_criar_pasta(ByVal pstr_caminho As String) As Boolean
    On Error GoTo erro_pfct_criar_pasta
    MkDir pstr_caminho
    pfct_criar_pasta = True
fim_pfct_criar_pasta:
    Exit Function
erro_pfct_criar_pasta:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_criar_pasta", False
    GoTo fim_pfct_criar_pasta
End Function

Public Function pfct_criptografia(ByVal pstr_texto As String) As String
    On Error GoTo erro_pfct_criptografia
    Dim lint_contador As Integer
    Dim lint_tamanho_string As Integer
    Dim lint_codigo_ascii As Integer
    Dim lstr_retorno As String
    lint_tamanho_string = Len(pstr_texto)
    If (lint_tamanho_string > 255) Then
        lint_tamanho_string = 255
    End If
    For lint_contador = 1 To lint_tamanho_string
        lint_codigo_ascii = 0
        lint_codigo_ascii = Asc(Mid$(pstr_texto, lint_contador, 1))
        lint_codigo_ascii = lint_codigo_ascii Xor lint_tamanho_string
        lstr_retorno = lstr_retorno & Chr$(lint_codigo_ascii)
    Next
    pfct_criptografia = StrReverse(lstr_retorno)
fim_pfct_criptografia:
    Exit Function
erro_pfct_criptografia:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_criptografia"
    GoTo fim_pfct_criptografia
End Function

Public Function pfct_excluir_arquivo(ByVal pstr_arquivo As String) As Boolean
    On Error GoTo erro_pfct_excluir_arquivo
    Kill pstr_arquivo
    pfct_excluir_arquivo = True
fim_pfct_excluir_arquivo:
    Exit Function
erro_pfct_excluir_arquivo:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_excluir_arquivo"
    GoTo fim_pfct_excluir_arquivo
End Function

Public Function pfct_form_esta_carregado(ByVal pstr_nome_form As String) As Boolean
    On Error GoTo erro_pfct_form_esta_carregado
    Dim lbln_retorno As Boolean
    Dim lfrm_form As Form
    For Each lfrm_form In Forms
        If (StrComp(lfrm_form.Name, pstr_nome_form, vbTextCompare) = 0) Then
            lbln_retorno = True
            Exit For
        End If
    Next
    pfct_form_esta_carregado = lbln_retorno
fim_pfct_form_esta_carregado:
    Set lfrm_form = Nothing
    Exit Function
erro_pfct_form_esta_carregado:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_form_esta_carregado"
    GoTo fim_pfct_form_esta_carregado
End Function

Public Function pfct_gerar_chave(ByVal pint_tamanho As Byte) As String
    On Error GoTo erro_pfct_gerar_chave
    Dim lstr_caracteres As String
    Dim llng_contador As Long
    Dim lstr_retorno As String
    lstr_caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890"
    For llng_contador = 1 To pint_tamanho
        Randomize
        lstr_retorno = lstr_retorno & Mid$(lstr_caracteres, CStr(CInt((Len(lstr_caracteres) * Rnd) + 1)), 1)
    Next
    pfct_gerar_chave = lstr_retorno
fim_pfct_gerar_chave:
    Exit Function
erro_pfct_gerar_chave:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_gerar_chave"
    GoTo fim_pfct_gerar_chave
End Function

Public Function pfct_hex_texto(ByVal pstr_texto As String, ByVal pbln_inverter As Boolean) As String
    On Error GoTo erro_pfct_hex_texto
    Dim lint_tamanho As Integer
    Dim lint_contador As Integer
    Dim lstr_temp As String
    Dim lstr_retorno As String
    If (pbln_inverter) Then
        pstr_texto = StrReverse(pstr_texto)
    End If
    lint_tamanho = Len(pstr_texto)
    For lint_contador = 1 To lint_tamanho Step 2
        lstr_temp = Mid$(pstr_texto, lint_contador, 2)
        lstr_temp = Val("&H" & lstr_temp)
        lstr_temp = Chr(lstr_temp)
        lstr_retorno = lstr_retorno & lstr_temp
    Next lint_contador
    pfct_hex_texto = lstr_retorno
fim_pfct_hex_texto:
    Exit Function
erro_pfct_hex_texto:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_hex_texto"
    GoTo fim_pfct_hex_texto
End Function

''''converter essas funções para o padrão eiko

'''Private Function encodeBase64(ByRef arrData() As Byte) As String
'''    Dim objXML As MSXML2.DOMDocument
'''    Dim objNode As MSXML2.IXMLDOMElement
'''
'''    Set objXML = New MSXML2.DOMDocument
'''
'''    Set objNode = objXML.createElement("b64")
'''    objNode.DataType = "bin.base64"
'''    objNode.nodeTypedValue = arrData
'''    encodeBase64 = objNode.Text
'''
'''    Set objNode = Nothing
'''    Set objXML = Nothing
'''End Function
'''
'''Private Function decodeBase64(ByVal strData As String) As Byte()
'''    Dim objXML As MSXML2.DOMDocument
'''    Dim objNode As MSXML2.IXMLDOMElement
'''
'''    Set objXML = New MSXML2.DOMDocument
'''    Set objNode = objXML.createElement("b64")
'''    objNode.DataType = "bin.base64"
'''    objNode.Text = strData
'''    decodeBase64 = objNode.nodeTypedValue
'''
'''    Set objNode = Nothing
'''    Set objXML = Nothing
'''End Function

'Public Function pfct_gravar_arquivo_de_array_de_bytes(ByVal pstr_nome_arquivo As String, ByRef parr_array_de_bytes() As Byte) As Boolean
'    On Error GoTo erro_pfct_gravar_arquivo_de_array_de_bytes
'    Dim int_arquivo As Integer
'    'buscamos o próximo arquivo
'    int_arquivo = FreeFile
'    'abrimos o arquivo em modo binário
'    Open pstr_nome_arquivo For Binary As #int_arquivo
'    'escrevemos o conteúdo do array de bytes em um arquivo
'    Put #int_arquivo, 1, parr_array_de_bytes
'    'fechamos o arquivo
'    Close #int_arquivo
'fim_pfct_gravar_arquivo_de_array_de_bytes:
'    Exit Function
'erro_pfct_gravar_arquivo_de_array_de_bytes:
'    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_gravar_arquivo_de_array_de_bytes"
'    GoTo fim_pfct_gravar_arquivo_de_array_de_bytes
'End Function

'Public Function pfct_ler_arquivo_para_array_de_bytes(ByVal pstr_arquivo As String) As Byte()
'    On Error GoTo erro_pfct_ler_arquivo_para_array_de_bytes
'    Dim int_arquivo As Integer
'    'buscamos o próximo arquivo
'    int_arquivo = FreeFile
'    'abrimos o arquivo em modo leitura
'    Open pstr_arquivo For Input Access Read As #int_arquivo
'    'se o arquivo for maior que zero bytes
'    If (LOF(int_arquivo) > 0) Then
'        'retornamos o arquivo em array de bytes
'        pfct_ler_arquivo_para_array_de_bytes = InputB(LOF(int_arquivo), int_arquivo)
'    End If
'    'fechamos o arquivo
'    Close #int_arquivo
'fim_pfct_ler_arquivo_para_array_de_bytes:
'    Exit Function
'erro_pfct_ler_arquivo_para_array_de_bytes:
'    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_ler_arquivo_para_array_de_bytes"
'    GoTo fim_pfct_ler_arquivo_para_array_de_bytes
'End Function

'Private Type GUID
'Data1 As Long
'Data2 As Integer
'Data3 As Integer
'Data4(7) As Byte
'End Type
'
'Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
'
'Public Function GetGUID() As String
'
'Dim udtGUID As GUID
'
'If (CoCreateGuid(udtGUID) = 0) Then
'
'GetGUID = _
'String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
'String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
'String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
'IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
'IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
'IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
'IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
'IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
'IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
'IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
'IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
'End If
'
'End Function

'Option Explicit
'
'Private Type Guid
'    Data1 As Long
'    Data2 As Integer
'    Data3 As Integer
'    Data4(0 To 7) As Byte
'End Type
'
'Private Declare Sub CoCreateGuid Lib "ole32.dll" (ByRef pguid As Guid)
'Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long
'
'Private Function GetGUID() As String
'    Dim MyGUID As Guid
'    Dim GUIDByte() As Byte
'    Dim GuidLen As Long
'
'    CoCreateGuid MyGUID
'
'    ReDim GUIDByte(80)
'    GuidLen = StringFromGUID2(VarPtr(MyGUID.Data1), VarPtr(GUIDByte(0)), UBound(GUIDByte))
'
'    GetGUID = Left(GUIDByte, GuidLen)
'End Function
'
'Private Sub Form_Load()
'    Debug.Print GetGUID
'End Sub

Public Function pfct_mover_arquivo(ByVal pstr_origem As String, ByVal pstr_destino As String) As Boolean
    On Error GoTo erro_pfct_mover_arquivo
    'copia o arquivo da origem para o destino
    FileCopy pstr_origem, pstr_destino
    'apaga o arquivo de origem
    Kill pstr_origem
    'sinaliza o sucesso da operação
    pfct_mover_arquivo = True
fim_pfct_mover_arquivo:
    Exit Function
erro_pfct_mover_arquivo:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_mover_arquivo"
    GoTo fim_pfct_mover_arquivo
End Function

Public Function pfct_pode_abrir_arquivo(ByVal pstr_arquivo As String) As Boolean
    On Error GoTo erro_pfct_pode_abrir_arquivo
    Dim lint_numero As Integer
    'se o arquivo existir
    If (pfct_verificar_arquivo_existe(pstr_arquivo)) Then
        'retornamos o próximo valor disponível
        lint_numero = FreeFile
        'abrimos o arquivo
        Open pstr_arquivo For Input As #lint_numero
        'fechamos o arquivo
        Close #lint_numero
        'retornamos o resultado da operação
        pfct_pode_abrir_arquivo = True
    End If
fim_pfct_pode_abrir_arquivo:
    Exit Function
erro_pfct_pode_abrir_arquivo:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_pode_abrir_arquivo", False
    GoTo fim_pfct_pode_abrir_arquivo
End Function

Public Function pfct_remover_excesso_espacos(ByVal pstr_texto As String) As String
    On Error GoTo erro_pfct_remover_excesso_espacos
    Dim lstr_temp As String
    lstr_temp = Trim$(pstr_texto)
    lstr_temp = Replace$(lstr_temp, vbTab, " ")
    lstr_temp = Replace$(lstr_temp, Chr$(160), " ") 'este caracter está presente nos arquivos do banco do brasil
    Do While InStr(1, lstr_temp, "  ") > 0
        lstr_temp = Replace$(lstr_temp, "  ", " ")
    Loop
    pfct_remover_excesso_espacos = lstr_temp
fim_pfct_remover_excesso_espacos:
    Exit Function
erro_pfct_remover_excesso_espacos:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_remover_excesso_espacos"
    GoTo fim_pfct_remover_excesso_espacos
End Function

Public Function pfct_remover_null(ByVal pstr_texto As String) As String
    On Error GoTo erro_pfct_remover_null
    pfct_remover_null = Left$(pstr_texto, lstrlenW(StrPtr(pstr_texto)))
fim_pfct_remover_null:
    Exit Function
erro_pfct_remover_null:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_remover_null"
    GoTo fim_pfct_remover_null
End Function

Public Function pfct_retorna_caminho_sistema(ByVal plng_caminho As Long) As String
    On Error GoTo erro_pfct_retorna_caminho_sistema
    Dim lstr_buffer As String
    'preenche a string com 260 espaços
    lstr_buffer = Space$(MAX_LENGTH)
    'verifica o retorno da função
    If SHGetFolderPath(frm_splash_sobre.hWnd, plng_caminho, -1, SHGFP_TYPE_CURRENT, lstr_buffer) = S_OK Then
        'retorna o caminho desejado
        pfct_retorna_caminho_sistema = pfct_remover_null(lstr_buffer)
    End If
fim_pfct_retorna_caminho_sistema:
    Exit Function
erro_pfct_retorna_caminho_sistema:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_retorna_caminho_sistema"
    GoTo fim_pfct_retorna_caminho_sistema
End Function

Public Function pfct_retorna_in(ByVal pobj_componente As Object) As String
    On Error GoTo erro_pfct_retorna_in
    Dim llng_contador As Long
    Dim llng_quantidade As Long
    Dim lstr_retorno As String
    'quantidade de itens no list box
    llng_quantidade = pobj_componente.ListCount
    'percorre os itens da lista
    For llng_contador = 0 To llng_quantidade - 1
        If (pobj_componente.Selected(llng_contador)) Then
            lstr_retorno = lstr_retorno & pobj_componente.ItemData(llng_contador) & ","
        End If
    Next
    'remove a vírgula adicional
    If (Right$(lstr_retorno, 1) = ",") Then
        lstr_retorno = Left$(lstr_retorno, Len(lstr_retorno) - 1)
    End If
    'devolve o conteúdo à função
    pfct_retorna_in = lstr_retorno
fim_pfct_retorna_in:
    Exit Function
erro_pfct_retorna_in:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_retorna_in"
    GoTo fim_pfct_retorna_in
End Function

Public Function pfct_retorna_nome_arquivo(ByVal pstr_caminho_completo As String) As String
    On Error GoTo erro_pfct_retorna_nome_arquivo
    Dim lstr_retorno As String
    Dim lobj_arquivo As New FileSystemObject
    'retorna o nome do arquivo
    lstr_retorno = lobj_arquivo.GetFileName(pstr_caminho_completo)
    'devolve o retorno à função
    pfct_retorna_nome_arquivo = lstr_retorno
fim_pfct_retorna_nome_arquivo:
    Set lobj_arquivo = Nothing
    Exit Function
erro_pfct_retorna_nome_arquivo:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_retorna_nome_arquivo"
    GoTo fim_pfct_retorna_nome_arquivo
End Function

Public Function pfct_retorna_periodo_data() As Integer
    On Error GoTo erro_pfct_retorna_periodo_data
    Dim lint_intervalo As Integer
    Select Case p_usuario.id_intervalo_data
        Case enm_intervalo_data.id_30dias
            lint_intervalo = 15
        Case enm_intervalo_data.id_60dias
            lint_intervalo = 30
        Case enm_intervalo_data.id_90dias
            lint_intervalo = 45
        Case enm_intervalo_data.id_120dias
            lint_intervalo = 60
        Case Else
            'assume o padrão 30 dias
            lint_intervalo = 15
    End Select
    pfct_retorna_periodo_data = lint_intervalo
fim_pfct_retorna_periodo_data:
    Exit Function
erro_pfct_retorna_periodo_data:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_retorna_periodo_data"
    GoTo fim_pfct_retorna_periodo_data
End Function

Public Function pfct_retorna_primeiro_dia_mes(ByVal pdt_data As Date) As Date
    On Error GoTo erro_pfct_retorna_primeiro_dia_mes
    pfct_retorna_primeiro_dia_mes = DateSerial(Year(pdt_data), Month(pdt_data), 1)
fim_pfct_retorna_primeiro_dia_mes:
    Exit Function
erro_pfct_retorna_primeiro_dia_mes:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_retorna_primeiro_dia_mes"
    GoTo fim_pfct_retorna_primeiro_dia_mes
End Function

Public Function pfct_retorna_recordset(ByVal pobj_objeto As Object) As Recordset
    On Error GoTo erro_pfct_retorna_recordset
    'declaração de variáveis
    Dim llng_contador As Long
    Dim llng_contador_campos As Long
    Dim llng_total As Long
    Dim llng_total_campos As Long
    Dim lstr_tipo_campo As String
    Dim lobj_recordset As Recordset
    'se houver instância
    If (Not pobj_objeto Is Nothing) Then
        'atribui valor
        llng_total = pobj_objeto.Count
        'se houver registros
        If (llng_total > 0) Then
            'atribui valor
            llng_total_campos = pobj_objeto(1).Count
            'se houver campos
            If (llng_total_campos > 0) Then
                'cria nova instância do objeto
                Set lobj_recordset = New Recordset
                'configura o novo recordset
                lobj_recordset.CursorType = adOpenKeyset
                lobj_recordset.LockType = adLockBatchOptimistic
                lobj_recordset.CursorLocation = adUseClient
                'monta dinamicamente o recordset
                For llng_contador_campos = 1 To llng_total_campos
                    'devolve o tipo do campo
                    lstr_tipo_campo = LCase$(Left$(pobj_objeto(1).Key(llng_contador_campos), InStr(1, pobj_objeto(1).Key(llng_contador_campos), "_") - 1))
                    'verifica tipo do campo
                    Select Case lstr_tipo_campo
                        Case "int" 'Integer
                            lobj_recordset.Fields.Append pobj_objeto(1).Key(llng_contador_campos), adInteger, , adFldUpdatable + adFldIsNullable
                        Case "num" 'Double/Currency
                            lobj_recordset.Fields.Append pobj_objeto(1).Key(llng_contador_campos), adCurrency, , adFldUpdatable + adFldIsNullable
                        Case "str" 'String
                            lobj_recordset.Fields.Append pobj_objeto(1).Key(llng_contador_campos), adVarChar, 512, adFldUpdatable + adFldIsNullable
                        Case "chr" 'String * 1
                            lobj_recordset.Fields.Append pobj_objeto(1).Key(llng_contador_campos), adChar, 1, adFldUpdatable + adFldIsNullable
                        Case "dt" 'Date
                            lobj_recordset.Fields.Append pobj_objeto(1).Key(llng_contador_campos), adDBDate, 10, adFldUpdatable + adFldIsNullable
                        Case "tm" 'Time
                            lobj_recordset.Fields.Append pobj_objeto(1).Key(llng_contador_campos), adDBTime, 8, adFldUpdatable + adFldIsNullable
                    End Select
                Next llng_contador_campos
                'abre o recordset
                lobj_recordset.Open
                'percorre o objeto adicionando dados ao recordset
                For llng_contador = 1 To llng_total
                    'adiciona novo registro
                    lobj_recordset.AddNew
                    'percorre os campos
                    For llng_contador_campos = 1 To llng_total_campos
                        lobj_recordset.Fields.Item(pobj_objeto(llng_contador).Key(llng_contador_campos)).Value = pobj_objeto(llng_contador)(pobj_objeto(llng_contador).Key(llng_contador_campos))
                    Next llng_contador_campos
                    'atualiza o registro
                    lobj_recordset.Update
                Next llng_contador
                'move ao primeiro registro
                lobj_recordset.MoveFirst
                'devolve o recordset para a função
                Set pfct_retorna_recordset = lobj_recordset
                'desvia ao bloco fim
                GoTo fim_pfct_retorna_recordset
            End If
        End If
    End If
fim_pfct_retorna_recordset:
    'destrói objetos
    Set lobj_recordset = Nothing
    'sai da função
    Exit Function
erro_pfct_retorna_recordset:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_retorna_recordset"
    GoTo fim_pfct_retorna_recordset
End Function

Public Function pfct_retorna_serie_processador() As String
    On Error GoTo Erro_pfct_retorna_serie_processador
    Dim lobj_wmi As Object
    Dim lobj_cpu As Object
    Dim lstr_retorno As String
    
    Set lobj_wmi = GetObject("winmgmts:")

    lstr_retorno = Empty
    For Each lobj_cpu In lobj_wmi.InstancesOf("Win32_Processor")
        If (IsNull(lobj_cpu.ProcessorID)) Then
            Exit For
        End If
        lstr_retorno = lstr_retorno & ("" & lobj_cpu.ProcessorID)
    Next
    
    If (lstr_retorno <> Empty) Then
        pfct_retorna_serie_processador = lstr_retorno
    End If

Fim_pfct_retorna_serie_processador:
    Set lobj_wmi = Nothing
    Set lobj_cpu = Nothing
    Exit Function
Erro_pfct_retorna_serie_processador:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_retorna_serie_processador"
    GoTo Fim_pfct_retorna_serie_processador
End Function

Public Function pfct_retorna_serie_volume(ByVal pstr_drive As String) As String
    On Error GoTo erro_pfct_retorna_serie_volume
    Dim lobj_volume As New FileSystemObject
    Dim lobj_drive As Drive
    Dim lstr_serie As String
    'atribui ao objeto
    Set lobj_drive = lobj_volume.GetDrive(pstr_drive)
    'verifica o objeto
    If (Not lobj_drive Is Nothing) Then
        'atribui à variável
        lstr_serie = CStr(Hex(lobj_drive.SerialNumber))
    Else
        lstr_serie = ""
    End If
    'retorna o valor
    pfct_retorna_serie_volume = lstr_serie
fim_pfct_retorna_serie_volume:
    'destrói os objetos
    Set lobj_volume = Nothing
    Set lobj_drive = Nothing
    Exit Function
erro_pfct_retorna_serie_volume:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_retorna_serie_volume"
    GoTo fim_pfct_retorna_serie_volume
End Function

Public Function pfct_retorna_simbolo_moeda() As String
    On Error GoTo erro_pfct_retorna_simbolo_moeda
    Dim lstr_simbolo As String
    Select Case p_usuario.sm_simbolo_moeda
        Case enm_simbolo_moeda.sm_dolar
            lstr_simbolo = "US$"
        Case enm_simbolo_moeda.sm_euro
            lstr_simbolo = "€$"
        Case enm_simbolo_moeda.sm_real
            lstr_simbolo = "R$"
        Case enm_simbolo_moeda.sm_iene
            lstr_simbolo = "¥$"
        Case Else
            'assume o real como padrão
            lstr_simbolo = "R$"
    End Select
    pfct_retorna_simbolo_moeda = lstr_simbolo
fim_pfct_retorna_simbolo_moeda:
    Exit Function
erro_pfct_retorna_simbolo_moeda:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_retorna_simbolo_moeda"
    GoTo fim_pfct_retorna_simbolo_moeda
End Function

Public Function pfct_retorna_temp() As String
    On Error GoTo Erro_pfct_retorna_temp
    Dim lstr_temp As String
    lstr_temp = Environ$("tmp")
    If (lstr_temp = Empty) Then
        lstr_temp = Environ$("temp")
    End If
    If (Right$(lstr_temp, 1) <> "\") Then
        lstr_temp = lstr_temp & "\"
    End If
    pfct_retorna_temp = lstr_temp
Fim_pfct_retorna_temp:
    Exit Function
Erro_pfct_retorna_temp:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_retorna_temp"
    GoTo Fim_pfct_retorna_temp
End Function

Public Function pfct_retorna_tipo_volume(ByVal pstr_drive As String) As String
    On Error GoTo erro_pfct_retorna_tipo_volume
    Dim lobj_volume As New FileSystemObject
    Dim lobj_drive As Drive
    Dim lstr_tipo_volume As String
    'atribui ao objeto
    Set lobj_drive = lobj_volume.GetDrive(pstr_drive)
    'verifica o objeto
    If (Not lobj_drive Is Nothing) Then
        'atribui à variável
        lstr_tipo_volume = CStr(lobj_drive.DriveType)
    Else
        'limpa a variável
        lstr_tipo_volume = ""
        'desvia ao fim do método
        GoTo fim_pfct_retorna_tipo_volume
    End If
    If (lstr_tipo_volume <> "") Then
        Select Case lstr_tipo_volume
            Case 1 'Removable
                lstr_tipo_volume = "R"
            Case 2 'Fixed
                lstr_tipo_volume = "F"
            Case Else
                lstr_tipo_volume = ""
        End Select
    End If
    'retorna o valor
    pfct_retorna_tipo_volume = lstr_tipo_volume
fim_pfct_retorna_tipo_volume:
    'destrói os objetos
    Set lobj_volume = Nothing
    Set lobj_drive = Nothing
    Exit Function
erro_pfct_retorna_tipo_volume:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_retorna_tipo_volume"
    GoTo fim_pfct_retorna_tipo_volume
End Function

Public Function pfct_retorna_ultimo_dia_mes(ByVal pdt_data As Date) As Date
    On Error GoTo erro_pfct_retorna_ultimo_dia_mes
    pfct_retorna_ultimo_dia_mes = DateAdd("d", -1, DateSerial(Year(pdt_data), Month(pdt_data) + 1, 1))
fim_pfct_retorna_ultimo_dia_mes:
    Exit Function
erro_pfct_retorna_ultimo_dia_mes:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_retorna_ultimo_dia_mes"
    GoTo fim_pfct_retorna_ultimo_dia_mes
End Function

Public Function pfct_texto_hex(ByVal pstr_texto As String, ByVal pbln_inverter As Boolean) As String
    On Error GoTo erro_pfct_texto_hex
    Dim lint_tamanho As Integer
    Dim lint_contador As Integer
    Dim lstr_temp As String
    Dim lstr_retorno As String
    lint_tamanho = Len(pstr_texto)
    For lint_contador = 1 To lint_tamanho
        lstr_temp = Mid$(pstr_texto, lint_contador, 1)
        lstr_temp = Hex$(Asc(lstr_temp))
        lstr_retorno = lstr_retorno & lstr_temp
    Next lint_contador
    If (pbln_inverter) Then
        lstr_retorno = StrReverse(lstr_retorno)
    End If
    pfct_texto_hex = lstr_retorno
fim_pfct_texto_hex:
    Exit Function
erro_pfct_texto_hex:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_texto_hex"
    GoTo fim_pfct_texto_hex
End Function

Public Function pfct_tratar_data_sql(ByVal pdt_data As Date) As String
    On Error GoTo erro_pfct_tratar_data_sql
    pfct_tratar_data_sql = Format$(pdt_data, "yyyy-mm-dd")
fim_pfct_tratar_data_sql:
    Exit Function
erro_pfct_tratar_data_sql:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_tratar_data_sql"
    GoTo fim_pfct_tratar_data_sql
End Function

Public Function pfct_tratar_hora_sql(ByVal pdt_hora As Date) As String
    On Error GoTo erro_pfct_tratar_hora_sql
    pfct_tratar_hora_sql = Format$(pdt_hora, "hh:mm:ss")
fim_pfct_tratar_hora_sql:
    Exit Function
erro_pfct_tratar_hora_sql:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_tratar_hora_sql"
    GoTo fim_pfct_tratar_hora_sql
End Function

Public Function pfct_tratar_numero_sql(ByVal pdbl_valor As Double) As String
    On Error GoTo erro_pfct_tratar_numero_sql
    pfct_tratar_numero_sql = Replace$(CStr(pdbl_valor), ",", ".")
fim_pfct_tratar_numero_sql:
    Exit Function
erro_pfct_tratar_numero_sql:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_tratar_numero_sql"
    GoTo fim_pfct_tratar_numero_sql
End Function

Public Function pfct_tratar_texto_sql(ByVal pstr_texto As String) As String
    On Error GoTo erro_pfct_tratar_texto_sql
    Dim lstr_texto As String
    lstr_texto = pfct_remover_excesso_espacos(pstr_texto)
    lstr_texto = UCase$(lstr_texto)
    pfct_tratar_texto_sql = lstr_texto
fim_pfct_tratar_texto_sql:
    Exit Function
erro_pfct_tratar_texto_sql:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_tratar_texto_sql"
    GoTo fim_pfct_tratar_texto_sql
End Function

Public Function pfct_verificar_administrador() As Boolean
    On Error GoTo erro_pfct_verificar_administrador
    Select Case IsUserAnAdmin()
        Case 1
            pfct_verificar_administrador = True
        Case Else
            pfct_verificar_administrador = False
    End Select
fim_pfct_verificar_administrador:
    Exit Function
erro_pfct_verificar_administrador:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_verificar_administrador"
    GoTo fim_pfct_verificar_administrador
End Function

Public Function pfct_verificar_arquivo_existe(ByVal pstr_arquivo As String) As Boolean
    On Error GoTo erro_pfct_verificar_arquivo_existe
    Dim lobj_arquivo As New FileSystemObject
    Dim lbln_retorno As Boolean
    'atribui o valor
    lbln_retorno = lobj_arquivo.FileExists(pstr_arquivo)
    'retorna o valor
    pfct_verificar_arquivo_existe = lbln_retorno
fim_pfct_verificar_arquivo_existe:
    'destrói os objetos
    Set lobj_arquivo = Nothing
    Exit Function
erro_pfct_verificar_arquivo_existe:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_verificar_arquivo_existe"
    GoTo fim_pfct_verificar_arquivo_existe
End Function

Public Function pfct_verificar_atualizacao() As Boolean
    On Error GoTo Erro_pfct_verificar_atualizacao
    
    Dim lobj_arquivo As FileSystemObject
    
    Dim lbln_retorno As Boolean
    Dim llng_retorno As Long
    Dim lstr_arquivo_temp As String
    Dim lstr_arquivo_remoto As String
    
    Dim lstr_dados As String
    Dim larr_dados() As String
    
    'retorna o caminho e o nome do arquivo remoto
    lstr_arquivo_remoto = app_ini
    
    'retorna o caminho e o nome do arquivo temporário
    lstr_arquivo_temp = pfct_retorna_temp & "app.ini"

    'baixa o arquivo localmente
    llng_retorno = URLDownloadToFile(0, lstr_arquivo_remoto, lstr_arquivo_temp, 0, 0)
    
    'se o download foi feito com sucesso
    If (llng_retorno = 0) Then
    
        'instancia o objeto
        Set lobj_arquivo = New FileSystemObject
    
        'retorna o conteúdo do arquivo
        lstr_dados = lobj_arquivo.OpenTextFile(lstr_arquivo_temp).ReadAll
        
        'quebra a string num array
        larr_dados = Split(lstr_dados, " ")
        
        'verifica a versão da aplicação e banco
        If ((CDbl(larr_dados(enm_app_ver.ap_app)) > CDbl(pcst_app_ver))) Or _
           ((CDbl(larr_dados(enm_app_ver.ap_bd)) > CDbl(pcst_dba_ver))) Then
        
            'sinaliza o retorno da comparação
            lbln_retorno = True
            
        End If
        
        'exclui o arquivo temporário
        Kill lstr_arquivo_temp
        
    End If
    
    pfct_verificar_atualizacao = lbln_retorno
Fim_pfct_verificar_atualizacao:
    Set lobj_arquivo = Nothing
    Exit Function
Erro_pfct_verificar_atualizacao:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_verificar_atualizacao"
    GoTo Fim_pfct_verificar_atualizacao
End Function

Public Function pfct_verificar_caminhos_aplicacao() As Boolean
    On Error GoTo erro_pfct_verificar_caminhos_aplicacao
    p_so.str_common_app_data = pfct_retorna_caminho_sistema(CSIDL_COMMON_APPDATA)
    'se a aplicação foi disparada a partir de uma unidade de rede
    If (Left$(App.Path, 2) = "\\") Then
        'desvia ao bloco fim
        GoTo fim_pfct_verificar_caminhos_aplicacao
    Else
        'retorna o drive de onde a aplicação foi disparada, ex.: [C:]
        p_so.str_drive_app_path = UCase$(Left$(App.Path, 2))
    End If
    If (Right$(p_so.str_common_app_data, 1) <> "\") Then
        p_so.str_common_app_data = p_so.str_common_app_data & "\"
    End If
    If (pfct_verificar_pasta_existe(p_so.str_common_app_data)) Then
        p_banco.str_caminho_comum = p_so.str_common_app_data & "eikoFP\"
        If (Not pfct_verificar_pasta_existe(p_banco.str_caminho_comum)) Then
            MkDir p_banco.str_caminho_comum
        End If
        p_banco.str_caminho_backup = p_banco.str_caminho_comum & "backup\"
        p_banco.str_caminho_config = p_banco.str_caminho_comum & "config\"
        p_banco.str_caminho_dados = p_banco.str_caminho_comum & "dados\"
        p_banco.str_caminho_log = p_banco.str_caminho_comum & "log\"
        If (Not pfct_verificar_pasta_existe(p_banco.str_caminho_backup)) Then
            MkDir p_banco.str_caminho_backup
        End If
        If (Not pfct_verificar_pasta_existe(p_banco.str_caminho_config)) Then
            MkDir p_banco.str_caminho_config
        End If
        If (Not pfct_verificar_pasta_existe(p_banco.str_caminho_dados)) Then
            MkDir p_banco.str_caminho_dados
        End If
        If (Not pfct_verificar_pasta_existe(p_banco.str_caminho_log)) Then
            MkDir p_banco.str_caminho_log
        End If
        pfct_verificar_caminhos_aplicacao = True
        GoTo fim_pfct_verificar_caminhos_aplicacao
    Else
        GoTo fim_pfct_verificar_caminhos_aplicacao
    End If
fim_pfct_verificar_caminhos_aplicacao:
    Exit Function
erro_pfct_verificar_caminhos_aplicacao:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_verificar_caminhos_aplicacao"
    GoTo fim_pfct_verificar_caminhos_aplicacao
End Function

Public Function pfct_verificar_email(ByVal pstr_email As String) As Boolean
    On Error GoTo Erro_pfct_verificar_email
    
    Dim arr_nomes() As String
    Dim var_nome As Variant
    Dim int_contador As Integer
    Dim str_caracter As String
    
    arr_nomes = Split(pstr_email, "@")
    
    If (UBound(arr_nomes) <> 1) Then
        GoTo Fim_pfct_verificar_email
    End If
    
    For Each var_nome In arr_nomes
        If (Len(var_nome) <= 0) Then
            GoTo Fim_pfct_verificar_email
        End If
        For int_contador = 1 To Len(var_nome)
            str_caracter = LCase(Mid(var_nome, int_contador, 1))
            If ((InStr("abcdefghijklmnopqrstuvwxyz_-.", str_caracter) <= 0) And (Not IsNumeric(str_caracter))) Then
                GoTo Fim_pfct_verificar_email
            End If
        Next int_contador
        If (Left(var_nome, 1) = "." Or Right(var_nome, 1) = ".") Then
            GoTo Fim_pfct_verificar_email
        End If
    Next
    
    If (InStr(arr_nomes(1), ".") <= 0) Then
        GoTo Fim_pfct_verificar_email
    End If
    
    int_contador = Len(arr_nomes(1)) - InStrRev(arr_nomes(1), ".")
    
    If ((int_contador <> 1) And (int_contador <> 3)) Then
        GoTo Fim_pfct_verificar_email
    End If
    
    If (InStr(pstr_email, "..") > 0) Then
        GoTo Fim_pfct_verificar_email
    End If

    pfct_verificar_email = True
Fim_pfct_verificar_email:
    Exit Function
Erro_pfct_verificar_email:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_verificar_email"
    GoTo Fim_pfct_verificar_email
End Function

Public Function pfct_verificar_pasta_existe(ByVal pstr_caminho As String) As Boolean
    On Error GoTo erro_pfct_verificar_pasta_existe
    Dim lobj_pasta As New FileSystemObject
    Dim lbln_retorno As Boolean
    'atribui valor
    lbln_retorno = lobj_pasta.FolderExists(pstr_caminho)
    'retorno do valor
    pfct_verificar_pasta_existe = lbln_retorno
fim_pfct_verificar_pasta_existe:
    'destrói os objetos
    Set lobj_pasta = Nothing
    Exit Function
erro_pfct_verificar_pasta_existe:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_verificar_pasta_existe"
    GoTo fim_pfct_verificar_pasta_existe
End Function

Public Function pfct_verificar_resolucao() As Boolean
    On Error GoTo erro_pfct_verificar_resolucao
    Dim llng_altura As Long
    Dim llng_largura As Long
    llng_altura = ((Screen.Height) / (Screen.TwipsPerPixelY))
    llng_largura = ((Screen.Width) / (Screen.TwipsPerPixelX))
    If ((llng_largura >= 1024) And (llng_altura >= 768)) Then
        pfct_verificar_resolucao = True
    End If
fim_pfct_verificar_resolucao:
    Exit Function
erro_pfct_verificar_resolucao:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_verificar_resolucao"
    GoTo fim_pfct_verificar_resolucao
End Function

Public Sub pfct_copiar_conteudo_grade(ByVal msf_grade As MSFlexGrid)
    On Error GoTo Erro_pfct_copiar_conteudo_grade
    'declaramos as variáveis
    Dim llng_linhas As Long
    Dim llng_colunas As Long
    Dim llng_linha_atual As Long
    Dim llng_coluna_atual As Long
    Dim lstr_resultado As String
    'quantidade de linhas e colunas
    llng_linhas = msf_grade.Rows
    llng_colunas = msf_grade.Cols
    'só iremos processar se houver linhas e colunas
    If ((llng_linhas > 0) And (llng_colunas > 0)) Then
        For llng_linha_atual = 0 To llng_linhas - 1
            For llng_coluna_atual = 0 To llng_colunas - 1
                'concatenamos o conteúdo da célula
                lstr_resultado = lstr_resultado & msf_grade.TextMatrix(llng_linha_atual, llng_coluna_atual)
                'se ainda não estamos na última coluna
                If (llng_coluna_atual < (llng_colunas - 1)) Then
                    'concatenamos uma tabulação
                    lstr_resultado = lstr_resultado & vbTab
                End If
            Next llng_coluna_atual
            'se ainda não estamos na última linha
            If (llng_linha_atual < (llng_linhas - 1)) Then
                'concatenamos uma quebra de linha
                lstr_resultado = lstr_resultado & vbCrLf
            End If
        Next llng_linha_atual
        'limpamos a área de transferência
        Clipboard.Clear
        'jogamos o conteúdo montado na área de transferência
        Clipboard.SetText lstr_resultado
    Else
        'exibimos uma mensagem ao usuário
        MsgBox "Não há dados a serem copiados.", vbOKOnly + vbExclamation, pcst_nome_aplicacao
    End If
Fim_pfct_copiar_conteudo_grade:
    Exit Sub
Erro_pfct_copiar_conteudo_grade:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_copiar_conteudo_grade"
    GoTo Fim_pfct_copiar_conteudo_grade
End Sub

Public Sub pfct_exportar_conteudo_grade(ByVal msf_grade As MSFlexGrid, ByVal str_nome_arquivo As String)
    On Error GoTo Erro_pfct_exportar_conteudo_grade
    'declaramos as variáveis
    Dim lobj_arquivo As New FileSystemObject
    Dim lobj_texto As TextStream
    Dim llng_linhas As Long
    Dim llng_colunas As Long
    Dim llng_linha_atual As Long
    Dim llng_coluna_atual As Long
    Dim lstr_resultado As String
    Dim lbln_csv As Boolean
    'quantidade de linhas e colunas
    llng_linhas = msf_grade.Rows
    llng_colunas = msf_grade.Cols
    'só iremos processar se houver linhas e colunas
    If ((llng_linhas > 0) And (llng_colunas > 0)) Then
        'se o usuário clicar em cancelar, será disparada uma exceção
        frm_principal.cd_arquivos.CancelError = True
        'comportamento do componente
        frm_principal.cd_arquivos.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist + cdlOFNOverwritePrompt
        'definimos o título do diálog
        frm_principal.cd_arquivos.DialogTitle = "Exportar conteúdo para arquivo"
        'filtro de arquivos (CSV e TXT)
        frm_principal.cd_arquivos.Filter = "CSV (*.csv)|*.csv|Texto (*.txt)|*.txt|"
        'definimos o filtro padrão para csv
        frm_principal.cd_arquivos.FilterIndex = 0
        'definimos a extensão padrão
        frm_principal.cd_arquivos.DefaultExt = "csv"
        'definimos o diretório padrão
        frm_principal.cd_arquivos.InitDir = pfct_retorna_caminho_sistema(CSIDL_DESKTOPDIRECTORY)
        'definimos o nome do arquivo padrão
        frm_principal.cd_arquivos.FileName = str_nome_arquivo & Format(Now, "_yyyy_mm_dd_hh_mm_ss")
        'exibimos o diálogo ao usuário
        frm_principal.cd_arquivos.ShowSave
        'verificamos se o conteúdo a ser gravado é CSV ou não
        If (LCase$(Right$(frm_principal.cd_arquivos.FileName, 3)) = "csv") Then
            'sinaliza como verdadeiro
            lbln_csv = True
        Else
            'sinaliza como falso
            lbln_csv = False
        End If
        'processamos os dados somente se o usuário salvar o arquivo
        For llng_linha_atual = 0 To llng_linhas - 1
            For llng_coluna_atual = 0 To llng_colunas - 1
                'concatenamos o conteúdo da célula
                lstr_resultado = lstr_resultado & msf_grade.TextMatrix(llng_linha_atual, llng_coluna_atual)
                'se for csv, concatenamos um ponto e vírgula
                If (lbln_csv) Then
                    lstr_resultado = lstr_resultado & ";"
                Else 'do contrário
                    'se ainda não estamos na última coluna
                    If (llng_coluna_atual < (llng_colunas - 1)) Then
                        'concatenamos uma tabulação
                        lstr_resultado = lstr_resultado & vbTab
                    End If
                End If
            Next llng_coluna_atual
            'se ainda não estamos na última linha
            If (llng_linha_atual < (llng_linhas - 1)) Then
                'concatenamos uma quebra de linha
                lstr_resultado = lstr_resultado & vbCrLf
            End If
        Next llng_linha_atual
        'abrimos o arquivo para escrita
        Set lobj_texto = lobj_arquivo.OpenTextFile(frm_principal.cd_arquivos.FileName, ForWriting, True)
        'escrevemos o arquivo
        lobj_texto.Write lstr_resultado
        'fechamos o arquivo
        lobj_texto.Close
        'exibimos uma mensagem ao usuário
        MsgBox "Dados exportados com sucesso.", vbOKOnly + vbInformation, pcst_nome_aplicacao
    Else
        'exibimos uma mensagem ao usuário
        MsgBox "Não há dados a serem exportados.", vbOKOnly + vbExclamation, pcst_nome_aplicacao
    End If
Fim_pfct_exportar_conteudo_grade:
    Set lobj_arquivo = Nothing
    Set lobj_texto = Nothing
    Exit Sub
Erro_pfct_exportar_conteudo_grade:
    If (Err.Number <> 32755) Then 'se o error for diferente de 'o usuário cancelou o diálogo'
        psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "pfct_exportar_conteudo_grade"
    End If
    GoTo Fim_pfct_exportar_conteudo_grade
End Sub

Public Sub psub_ajustar_cor_linha_grade(ByRef pobj_grade As MSFlexGrid, ByVal plng_linha As Long, ByVal plng_cor As Long)
    On Error GoTo erro_psub_ajustar_cor_linha_grade
    Dim llng_colunas As Long
    Dim llng_contador As Long
    'se houver mais de uma linha na grade
    If (pobj_grade.Rows > 1) Then
        'atribui valor à variável
        llng_colunas = pobj_grade.Cols - 1
        'percorre as colunas
        For llng_contador = 0 To llng_colunas
            pobj_grade.Row = plng_linha
            pobj_grade.Col = llng_contador
            pobj_grade.CellForeColor = plng_cor
        Next llng_contador
    End If
fim_psub_ajustar_cor_linha_grade:
    Exit Sub
erro_psub_ajustar_cor_linha_grade:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "psub_ajustar_cor_linha_grade"
    GoTo fim_psub_ajustar_cor_linha_grade
End Sub

Public Sub psub_ajustar_janela(ByRef pobj_form As Form, ByVal plng_altura As Long, ByVal plng_largura As Long, ByVal pbln_exibir As Boolean)
    On Error GoTo erro_psub_ajustar_janela
    With pobj_form
        'esconde o form
        .Hide
        'ajusta largura
        .Width = (plng_largura * Screen.TwipsPerPixelX)
        'ajusta algura
        .Height = (plng_altura * Screen.TwipsPerPixelY)
        'exibe o form?
        If (pbln_exibir) Then
            .Show
        End If
    End With
fim_psub_ajustar_janela:
    Exit Sub
erro_psub_ajustar_janela:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "psub_ajustar_janela"
    GoTo fim_psub_ajustar_janela
End Sub

Public Sub psub_destruir_tooltip(ByRef pobj_tooltip As Object)
    On Error GoTo erro_psub_destruir_tooltip
    If (Not pobj_tooltip Is Nothing) Then
        'destrói o tool tip
        pobj_tooltip.Destroy
        Set pobj_tooltip = Nothing
    End If
fim_psub_destruir_tooltip:
    Exit Sub
erro_psub_destruir_tooltip:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "psub_destruir_tooltip"
    GoTo fim_psub_destruir_tooltip
End Sub

Public Sub psub_exibir_tooltip(ByRef pobj_objeto As Object, ByRef pobj_tooltip As Object, ByVal pstr_titulo As String, ByVal pstr_mensagem As String)
    On Error GoTo erro_psub_exibir_tooltip
    'cria a instância se for nothing
    If (pobj_tooltip Is Nothing) Then
        'cria a nova instância
        Set pobj_tooltip = New CToolTip
        'ajusta as propriedades
        With pobj_tooltip
            .Style = TTStandard
            .Icon = TTNoIcon
            .Title = pstr_titulo
            .TipText = pstr_mensagem
            .Create pobj_objeto.hWnd 'exibe o tooltip
        End With
    End If
fim_psub_exibir_tooltip:
    Exit Sub
erro_psub_exibir_tooltip:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "psub_exibir_tooltip"
    GoTo fim_psub_exibir_tooltip
End Sub

Public Sub psub_fechar_forms()
    On Error GoTo erro_psub_fechar_forms
    Dim lobj_form As Form
    For Each lobj_form In Forms
        If (lobj_form.Name <> "frm_principal") And (lobj_form.Name <> "frm_imagem_fundo") Then
            Unload lobj_form
        End If
    Next
fim_psub_fechar_forms:
    Exit Sub
erro_psub_fechar_forms:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_funcoes", "psub_fechar_forms"
    GoTo fim_psub_fechar_forms
End Sub


