Attribute VB_Name = "modSysTray"
Option Explicit

' Used to detect clicking on the TRAY icon
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205

' Used to control the TRAY icon
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

' Used as the ID of the call back message (TRAY ICON)
Public Const WM_MOUSEMOVE = &H200

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" ( _
                                ByVal dwMessage As Long, _
                                pnid As NOTIFYICONDATA) As Boolean

Public Declare Function DestroyIcon Lib "User32.dll" (ByVal hIcon As Long) As Long


' Used by Shell_NotifyIcon (TRAY ICON)
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public TrayIcon As NOTIFYICONDATA

