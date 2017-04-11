Attribute VB_Name = "bas_tipos"
Option Explicit

'enum
Public Enum enm_tipo_banco
    tb_config = 1
    tb_dados = 2
    tb_backup = 3
    tb_restaurar = 4
End Enum

'enum
Public Enum enm_periodo_backup
    pb_selecione = 0
    pb_diario = 1
    pb_semanal = 2
    pb_quinzenal = 3
    pb_mensal = 4
End Enum

'enum
Public Enum enm_intervalo_data
    id_selecione = 0
    id_30dias = 1
    id_60dias = 2
    id_90dias = 3
    id_120dias = 4
End Enum

'enum
Public Enum enm_simbolo_moeda
    sm_selecione = 0
    sm_dolar = 1
    sm_euro = 2
    sm_real = 3
    sm_iene = 4
End Enum

'enum
Public Enum enm_tipo_campo
    tc_inteiro = 1
    tc_monetario = 2
    tc_texto = 3
    tc_data = 4
End Enum

'enum
Public Enum enm_app_ver
    ap_app = 1
    ap_bd = 3
End Enum

'tipo
Private Type tpe_so
    str_common_app_data As String
    str_drive_app_path As String
End Type

'tipo
Private Type tpe_banco
    str_caminho_comum As String
    str_caminho_backup As String
    str_caminho_config As String
    str_caminho_dados As String
    str_caminho_log As String
    str_caminho_dados_backup As String
    str_caminho_dados_config As String
    str_caminho_dados_restaurar As String
    str_caminho_dados_usuario As String
    bln_ativar_log_sql As Boolean
    tb_tipo_banco As enm_tipo_banco
End Type

'tipo
Private Type tpe_usuario
    lng_codigo As Long
    str_login As String
    str_senha As String
    str_lembrete_senha As String
    dt_criado_em As Date
    dt_ultimo_acesso As Date
    id_intervalo_data As enm_intervalo_data
    sm_simbolo_moeda As enm_simbolo_moeda
    bln_carregar_agenda_financeira_login As Boolean
    bln_lancamentos_retroativos As Boolean
    bln_alteracoes_detalhes As Boolean
    bln_data_vencimento_baixa_imediata As Boolean
    bln_lancamentos_duplicados As Boolean
    bln_participou_pesquisa As Boolean
End Type

'tipo
Private Type tpe_backup
    bln_ativar As Boolean
    pb_periodo_backup As enm_periodo_backup
    str_caminho As String
    str_nome As String
    dt_ultimo_backup As Date
    dt_proximo_backup As Date
End Type

'tipo
Public Type tpe_registro
    int_codigo As Integer
    str_usuario As String
    str_nome As String
    str_email As String
    str_pais As String
    str_estado As String
    str_cidade As String
    dt_data_nascimento As Date
    str_profissao As String
    chr_sexo As String
    str_origem As String
    str_opiniao As String
    bln_newsletter As Boolean
    str_id_cpu As String
    str_id_hd As String
    dt_data_registro As Date
    dt_data_liberacao As Date
    bln_banido As Boolean
    str_desc_banido As String
End Type

'tipo
Private Type tpe_mysql
    str_servidor As String
    str_usuario As String
    str_senha As String
    str_banco As String
    lng_porta As Long
End Type

'tipo
Private Type tpe_pc
    str_id_cpu As String
    str_id_hd As String
End Type

'variáveis públicas
Public p_so As tpe_so
Public p_banco As tpe_banco
Public p_usuario As tpe_usuario
Public p_backup As tpe_backup
Public p_registro As tpe_registro
Public p_mysql As tpe_mysql
Public p_pc As tpe_pc
