Attribute VB_Name = "bas_principal"
Option Explicit

'vari�vel que define se o modo debug est� ativo
Public p_modo_debug As Boolean

'vari�vel que define se o modo offline est� ativo
Public p_modo_offline As Boolean

'apar�ncia tipo xp
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

'constante para verifica��o de vers�o
Public Const app_ini As String = "http://dl.dropbox.com/u/78753613/eiko/app.ini"

'constante para download do novo instalador
Public Const app_setup As String = "http://dl.dropbox.com/u/78753613/eiko/setup.exe"

'constante com o nome da aplica��o
Public Const pcst_nome_aplicacao As String = "Eiko Finan�as Pessoais"

'vers�o do aplicativo
Public Const pcst_app_ver As String = "0.8.7"

'vers�o do banco
Public Const pcst_dba_ver As String = "0.2.5"

Sub Main()
    On Error GoTo erro_sub_main
    Dim lstr_comando As String
    
    'in�cio da aplica��o
    InitCommonControls
    
    'verifica se j� existe uma inst�ncia do programa em mem�ria
    If (App.PrevInstance) Then
        MsgBox "Aten��o!" & vbCrLf & _
               "N�o � poss�vel abrir mais de uma inst�ncia da aplica��o ao mesmo tempo." & vbCrLf & _
               "Clique em [OK] para encerrar o programa.", vbOKOnly + vbInformation, pcst_nome_aplicacao
        End
    End If
    
    'verifica se foi passado par�metro ao iniciar a aplica��o
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
    
    'verificar as pastas da aplica��o
    If (pfct_verificar_caminhos_aplicacao()) Then
        'ajustar os caminhos da aplica��o
        If (pfct_ajustar_caminho_banco(tb_config)) Then
            'banco do tipo configura��o
            p_banco.tb_tipo_banco = tb_config
            'cria as tabelas de configura��o
            If (pfct_criar_tabelas_config()) Then
                'oculta o bot�o do form splash
                frm_splash_sobre.cmd_ok.Visible = False
                'exibe o form splash
                frm_splash_sobre.Show
                'atualiza o form splash
                frm_splash_sobre.Refresh
                'busca as informa��es do computador
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
        MsgBox "Aten��o!" & vbCrLf & _
               "Erro ao verificar os caminhos padr�es da aplica��o." & vbCrLf & _
               "Clique em [OK] para encerrar o programa.", vbOKOnly + vbCritical, pcst_nome_aplicacao
        End
    End If
    
fim_sub_main:
    Exit Sub
erro_sub_main:
    psub_gerar_log_erro Err.Number, Err.Description, "bas_principal", "Main"
    End
End Sub
