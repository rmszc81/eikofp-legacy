Attribute VB_Name = "bas_principal"
Option Explicit

'variável que define se o modo debug está ativo
Public p_modo_debug As Boolean

'variável que define se o modo offline está ativo
Public p_modo_offline As Boolean

'aparência tipo xp
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

'constante para verificação de versão
Public Const app_ini As String = "http://dl.dropbox.com/u/78753613/eiko/app.ini"

'constante para download do novo instalador
Public Const app_setup As String = "http://dl.dropbox.com/u/78753613/eiko/setup.exe"

'constante com o nome da aplicação
Public Const pcst_nome_aplicacao As String = "Eiko Finanças Pessoais"

'versão do aplicativo
Public Const pcst_app_ver As String = "0.8.7"

'versão do banco
Public Const pcst_dba_ver As String = "0.2.5"

Sub Main()
    On Error GoTo erro_sub_main
    Dim lstr_comando As String
    
    'início da aplicação
    InitCommonControls
    
    'verifica se já existe uma instância do programa em memória
    If (App.PrevInstance) Then
        MsgBox "Atenção!" & vbCrLf & _
               "Não é possível abrir mais de uma instância da aplicação ao mesmo tempo." & vbCrLf & _
               "Clique em [OK] para encerrar o programa.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        End
    End If
    
    'verifica se foi passado parâmetro ao iniciar a aplicação
    lstr_comando = LCase$(Command$)
    
    'modo offline
    If (InStr(lstr_comando, "/offline") > 0) Then
        p_modo_offline = True
        MsgBox "Modo Offline Ativado.", vbOKOnly + vbInformation, pcst_nome_aplicacao
    Else
        p_modo_offline = False
    End If
    
    'modo debug
    If (InStr(lstr_comando, "/debug") > 0) Then
        p_modo_debug = True
        MsgBox "Modo Debug Ativado.", vbOKOnly + vbInformation, pcst_nome_aplicacao
    Else
        p_modo_debug = False
    End If
    
    'verificar as pastas da aplicação
    If (pfct_verificar_caminhos_aplicacao()) Then
        'ajustar os caminhos da aplicação
        If (pfct_ajustar_caminho_banco(tb_config)) Then
            'banco do tipo configuração
            p_banco.tb_tipo_banco = tb_config
            'cria as tabelas de configuração
            If (pfct_criar_tabelas_config()) Then
                'oculta o botão do form splash
                frm_splash_sobre.cmd_ok.Visible = False
                'exibe o form splash
                frm_splash_sobre.Show
                'atualiza o form splash
                frm_splash_sobre.Refresh
                'busca as informações do computador
                With p_pc
                    .str_id_cpu = pfct_retorna_serie_processador
                    .str_id_hd = pfct_retorna_serie_volume(p_so.str_drive_app_path)
                End With
                'aguarda 3 segundos
                Sleep 3000
                'esconde o form splash
                frm_splash_sobre.Hide
                'descarrega o form splash
                Unload frm_splash_sobre
                'ajusta a largura e altura do form principal
                psub_ajustar_janela frm_principal, 600, 950, False
                'exibe o form principal
                frm_principal.Show
            End If
        End If
    Else
        MsgBox "Atenção!" & vbCrLf & _
               "Erro ao verificar os caminhos padrões da aplicação." & vbCrLf & _
               "Clique em [OK] para encerrar o programa.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        End
    End If
    
fim_sub_main:
    Exit Sub
erro_sub_main:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_principal", "Main"
    End
End Sub
