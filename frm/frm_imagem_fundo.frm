VERSION 5.00
Begin VB.Form frm_imagem_fundo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
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
   MDIChild        =   -1  'True
   ScaleHeight     =   3975
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmr_timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   60
   End
   Begin VB.Label lbl_pesquisa_publico 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clique aqui para participar da nossa pesquisa de público"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   4080
      Width           =   2835
   End
   Begin VB.Image img_fundo 
      Height          =   3975
      Left            =   0
      Picture         =   "frm_imagem_fundo.frx":0000
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "frm_imagem_fundo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbln_inverter As Boolean

Public Property Let registrado(ByVal pbln_valor As Boolean)
    On Error GoTo Erro_registrado
    If (Not pbln_valor) Then
        Me.Height = 4605
        tmr_timer.Enabled = True
    Else
        Me.Height = 3975
        tmr_timer.Enabled = False
    End If
Fim_registrado:
    Exit Property
Erro_registrado:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_imagem_fundo", "registrado"
    GoTo Fim_registrado
End Property

Private Sub Form_Activate()
    On Error GoTo Erro_Form_Activate
    Me.ZOrder 1
Fim_Form_Activate:
    Exit Sub
Erro_Form_Activate:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_imagem_fundo", "Form_Activate"
    GoTo Fim_Form_Activate
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/menu_principal.htm", 0
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_imagem_fundo", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo erro_Form_MouseMove
    lbl_pesquisa_publico.Font.Underline = False
fim_Form_MouseMove:
    Exit Sub
erro_Form_MouseMove:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_imagem_fundo", "Form_MouseMove"
    GoTo fim_Form_MouseMove
End Sub

Private Sub img_fundo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Erro_img_fundo_MouseMove
    lbl_pesquisa_publico.Font.Underline = False
Fim_img_fundo_MouseMove:
    Exit Sub
Erro_img_fundo_MouseMove:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_imagem_fundo", "img_fundo_MouseMove"
    GoTo Fim_img_fundo_MouseMove
End Sub

Private Sub lbl_pesquisa_publico_Click()
    On Error GoTo Erro_lbl_pesquisa_publico_Click
    Dim obj_pesquisa_publico As Object
    Set obj_pesquisa_publico = New frm_pesquisa_publico
    obj_pesquisa_publico.Show vbModal, frm_principal
Fim_lbl_pesquisa_publico_Click:
    Set obj_pesquisa_publico = Nothing
    Exit Sub
Erro_lbl_pesquisa_publico_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_imagem_fundo", "lbl_pesquisa_publico_Click"
    GoTo Fim_lbl_pesquisa_publico_Click
End Sub

Private Sub lbl_pesquisa_publico_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Erro_lbl_pesquisa_publico_MouseMove
    lbl_pesquisa_publico.Font.Underline = True
Fim_lbl_pesquisa_publico_MouseMove:
    Exit Sub
Erro_lbl_pesquisa_publico_MouseMove:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_imagem_fundo", "lbl_pesquisa_publico_MouseMove"
    GoTo Fim_lbl_pesquisa_publico_MouseMove
End Sub

Private Sub tmr_timer_Timer()
    On Error GoTo erro_tmr_timer_timer
    If (mbln_inverter) Then
        lbl_pesquisa_publico.ForeColor = vbHighlight
        lbl_pesquisa_publico.Refresh
    Else
        lbl_pesquisa_publico.ForeColor = vbRed
        lbl_pesquisa_publico.Refresh
    End If
    mbln_inverter = Not mbln_inverter
fim_tmr_timer_timer:
    Exit Sub
erro_tmr_timer_timer:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_imagem_fundo", "tmr_timer_Timer"
    GoTo fim_tmr_timer_timer
End Sub
