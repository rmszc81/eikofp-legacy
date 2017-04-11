Attribute VB_Name = "bas_api"
Option Explicit

Public Const SHGFP_TYPE_CURRENT As Long = &H0
Public Const SHGFP_TYPE_DEFAULT As Long = &H1
Public Const MAX_LENGTH As Long = 260
Public Const S_OK  As Long = 0
Public Const S_FALSE As Long = 1

Public Const BIF_RETURNONLYFSDIRS As Long = 1
Public Const BIF_DONTGOBELOWDOMAIN As Long = 2
Public Const SW_SHOWNORMAL As Long = 1

Public Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Public Enum CSIDL_VALUES
    CSIDL_APPDATA = &H1A
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_DESKTOPDIRECTORY = &H10
End Enum

Public Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hWndOwner As Long, ByVal nFolder As Long, _
                                                                                     ByVal hToken As Long, ByVal dwReserved As Long, _
                                                                                     ByVal lpszPath As String) As Long

Public Declare Function lstrlenW Lib "kernel32" _
(ByVal lpString As Long) As Long

Public Declare Function SHBrowseForFolder Lib _
"shell32" (lpbi As BrowseInfo) As Long

Public Declare Function SHGetPathFromIDList Lib _
"shell32" (ByVal pidList As Long, ByVal lpBuffer _
As String) As Long

Public Declare Function lstrcat Lib "kernel32" _
Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
lpString2 As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function URLDownloadToFile Lib "urlmon.dll" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
                                                                                       ByVal szFileName As String, ByVal dwReserved As Long, _
                                                                                       ByVal lpfnCB As Long) As Long
                                                                                       
Public Declare Function IsUserAnAdmin Lib "shell32" () As Long

