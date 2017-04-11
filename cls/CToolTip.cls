VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const TTS_NOPREFIX = &H2
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TTF_IDISHWND = &H1
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    hwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type

Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum
Public Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum

Private mlngBackColor As Long
Private mstrTitle As String
Private mlngForeColor As Long
Private mtpeIcon As ttIconType
Private mblnCentered As Boolean
Private mtpeStyle As ttStyleEnum
Private mstrTipText As String
Private mlngVisibleTime As Long
Private mlngDelayTime As Long

Private m_lTTHwnd As Long
Private m_lParentHwnd As Long
Private ti As TOOLINFO

Public Property Let Style(ByVal vData As ttStyleEnum)
    mtpeStyle = vData
End Property

Public Property Get Style() As ttStyleEnum
    Style = mtpeStyle
End Property

Public Property Let Centered(ByVal vData As Boolean)
    mblnCentered = vData
End Property

Public Property Get Centered() As Boolean
    Centered = mblnCentered
End Property

Public Function Create(ByVal ParentHwnd As Long) As Boolean

    Dim lWinStyle As Long
    
    If m_lTTHwnd <> 0 Then
        DestroyWindow m_lTTHwnd
    End If
    
    m_lParentHwnd = ParentHwnd
    
    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    
    If mtpeStyle = TTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON
       
    m_lTTHwnd = CreateWindowEx(0&, _
        TOOLTIPS_CLASSA, _
        vbNullString, _
        lWinStyle, _
        CW_USEDEFAULT, _
        CW_USEDEFAULT, _
        CW_USEDEFAULT, _
        CW_USEDEFAULT, _
        0&, _
        0&, _
        App.hInstance, _
        0&)
               
    With ti
        If mblnCentered Then
            .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP Or TTF_IDISHWND
        Else
            .lFlags = TTF_SUBCLASS Or TTF_IDISHWND
        End If
       
        .hwnd = m_lParentHwnd
        .lId = m_lParentHwnd
        .hInstance = App.hInstance
        .lSize = Len(ti)
    End With
    
    SendMessage m_lTTHwnd, TTM_ADDTOOLA, 0&, ti
    
    If mstrTitle <> vbNullString Or mtpeIcon <> TTNoIcon Then
        SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mtpeIcon), ByVal mstrTitle
    End If
    If mlngForeColor <> Empty Then
        SendMessage m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mlngForeColor, 0&
    End If
    If mlngBackColor <> Empty Then
        SendMessage m_lTTHwnd, TTM_SETTIPBKCOLOR, mlngBackColor, 0&
    End If
    
    SendMessageLong m_lTTHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, mlngVisibleTime
    SendMessageLong m_lTTHwnd, TTM_SETDELAYTIME, TTDT_INITIAL, mlngDelayTime
    
End Function

Public Property Let Icon(ByVal vData As ttIconType)
    mtpeIcon = vData
    If m_lTTHwnd <> 0 And mstrTitle <> Empty And mtpeIcon <> TTNoIcon Then
        SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mtpeIcon), ByVal mstrTitle
    End If
End Property

Public Property Get Icon() As ttIconType
    Icon = mtpeIcon
End Property

Public Property Let ForeColor(ByVal vData As Long)
    mlngForeColor = vData
    If m_lTTHwnd <> 0 Then
        SendMessage m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mlngForeColor, 0&
    End If
End Property

Public Property Get ForeColor() As Long
    ForeColor = mlngForeColor
End Property

Public Property Let Title(ByVal vData As String)
    mstrTitle = vData
    If m_lTTHwnd <> 0 And mstrTitle <> Empty And mtpeIcon <> TTNoIcon Then
        SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mtpeIcon), ByVal mstrTitle
    End If
End Property

Public Property Get Title() As String
    Title = ti.lpStr
End Property

Public Property Let BackColor(ByVal vData As Long)
    mlngBackColor = vData
    If m_lTTHwnd <> 0 Then
        SendMessage m_lTTHwnd, TTM_SETTIPBKCOLOR, mlngBackColor, 0&
    End If
End Property

Public Property Get BackColor() As Long
    BackColor = mlngBackColor
End Property

Public Property Let TipText(ByVal vData As String)
    mstrTipText = vData
    ti.lpStr = vData
    If m_lTTHwnd <> 0 Then
        SendMessage m_lTTHwnd, TTM_UPDATETIPTEXTA, 0&, ti
    End If
End Property

Public Property Get TipText() As String
    TipText = mstrTipText
End Property

Private Sub Class_Initialize()
    InitCommonControls
    mlngDelayTime = 500
    mlngVisibleTime = 5000
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Sub Destroy()
    If m_lTTHwnd <> 0 Then
        DestroyWindow m_lTTHwnd
    End If
End Sub

Public Property Get VisibleTime() As Long
    VisibleTime = mlngVisibleTime
End Property

Public Property Let VisibleTime(ByVal lData As Long)
    mlngVisibleTime = lData
End Property

Public Property Get DelayTime() As Long
    DelayTime = mlngDelayTime
End Property

Public Property Let DelayTime(ByVal lData As Long)
    mlngDelayTime = lData
End Property
