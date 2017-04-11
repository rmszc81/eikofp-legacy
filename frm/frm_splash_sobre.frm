VERSION 5.00
Begin VB.Form frm_splash_sobre 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6300
      TabIndex        =   2
      Top             =   4380
      Width           =   915
   End
   Begin VB.Label lbl_desenvolvedor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rossano M. Szczepanski"
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
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   2040
   End
   Begin VB.Label lbl_email 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "rm.szc81@gmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   1785
   End
   Begin VB.Image img_logo 
      Height          =   1815
      Left            =   60
      Picture         =   "frm_splash_sobre.frx":0000
      Top             =   60
      Width           =   3405
   End
   Begin VB.Shape shp_borda 
      Height          =   1155
      Left            =   60
      Top             =   2520
      Width           =   1395
   End
   Begin VB.Label lbl_desenvolvido_por 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desenvolvido por:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   1515
   End
   Begin VB.Image img_fundo 
      Height          =   5415
      Left            =   3540
      Picture         =   "frm_splash_sobre.frx":1438E
      Top             =   -600
      Width           =   3750
   End
End
Attribute VB_Name = "frm_splash_sobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ok_Click()
    On Error GoTo erro_cmd_ok_Click
    'impede que o comando seja executado
    'se o botão estiver desabilitado
    If (Not cmd_ok.Enabled) Then
        Exit Sub
    End If
    Unload Me
fim_cmd_ok_Click:
    Exit Sub
erro_cmd_ok_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_splash_sobre", "cmd_ok_Click"
    GoTo fim_cmd_ok_Click
End Sub

Private Sub Form_Resize()
    On Error GoTo erro_Form_Resize
    'configura a borda
    shp_borda.Move 0, 0, ScaleWidth, ScaleHeight
    shp_borda.ZOrder 0
fim_Form_Resize:
    Exit Sub
erro_Form_Resize:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_splash_sobre", "Form_Resize"
    GoTo fim_Form_Resize
End Sub

Private Sub lbl_email_Click()
    On Error GoTo erro_lbl_email_Click
    'se estivermos em modo online
    If (Not p_modo_offline) Then
        'chama a api do windows
        ShellExecute 0&, vbNullString, "mailto:rm.szc81@gmail.com", vbNullString, vbNullString, SW_SHOWNORMAL
    End If
fim_lbl_email_Click:
    Exit Sub
erro_lbl_email_Click:
    psub_gerar_log_erro Err.Number, Err.Description, "frm_splash_sobre", "lbl_email_Click"
    GoTo fim_lbl_email_Click
End Sub
