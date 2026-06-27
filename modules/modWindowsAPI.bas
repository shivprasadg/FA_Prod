Attribute VB_Name = "modWindowsAPI"
Option Compare Database

Option Explicit


#If VBA7 Then ' Win API Declarations for 32 and 64 bit versions of Office 2010 and later

Public Declare PtrSafe Sub Sleep Lib "kernel32" _
    (ByVal dwMilliseconds As Long)

Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare PtrSafe Function SetForegroundWindow Lib "user32" _
    (ByVal hWnd As Long) As Long

Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare PtrSafe Function SendMessageByString Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare PtrSafe Sub keybd_event Lib "user32" _
    (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#Else

Public Declare Sub Sleep Lib "kernel32" _
    (ByVal dwMilliseconds As Long)

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hWnd As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Sub keybd_event Lib "user32" _
    (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#End If


Public Const BM_CLICK = &HF5
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE

Public Const VK_KEYDOWN = &H0
Public Const VK_KEYUP = &H2
Public Const VK_CONTROL = &H11
