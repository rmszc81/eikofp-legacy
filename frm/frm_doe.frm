VERSION 5.00
Begin VB.Form frm_doe 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.Label lbl_pesquisa_publico 
      BackStyle       =   0  'Transparent
      Caption         =   "O desenvolvimento do Eiko Finanças Pessoais custa dinheiro! Ajude nos, clicando no botão [Doar] ao lado!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   3135
   End
   Begin VB.Image img_doe 
      Height          =   705
      Left            =   3315
      Picture         =   "frm_doe.frx":0000
      ToolTipText     =   "Clique aqui para doar!"
      Top             =   105
      Width           =   1485
   End
End
Attribute VB_Name = "frm_doe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Erro_Form_KeyUp
    Select Case KeyCode
        Case vbKeyF1
            psub_exibir_ajuda Me, "html/menu_principal.htm", 0
    End Select
Fim_Form_KeyUp:
    Exit Sub
Erro_Form_KeyUp:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_doe", "Form_KeyUp"
    GoTo Fim_Form_KeyUp
End Sub

Private Sub img_doe_Click()
    On Error GoTo erro_img_doe_Click
    'chama a api do windows
    ShellExecute 0&, vbNullString, "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=6W5CUB2SVX3Z4", vbNullString, vbNullString, SW_SHOWNORMAL
fim_img_doe_Click:
    Exit Sub
erro_img_doe_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_doe", "img_doe_Click"
    GoTo fim_img_doe_Click
End Sub
