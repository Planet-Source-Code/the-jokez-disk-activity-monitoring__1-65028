Attribute VB_Name = "modAlwaysOnTop"
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
                                                    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
                                                    ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long

' Constantes de SetWindowPos :
Private Const HWND_TOPMOST      As Long = -1
Private Const HWND_NOTOPMOST    As Long = -2

Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOACTIVATE    As Long = &H10
'

Public Sub SetTop(Form As Form, _
                  ByVal Topmost As Boolean)

Dim hWndInsertAfter As Long

    If Topmost Then
        hWndInsertAfter = HWND_TOPMOST
    Else
        hWndInsertAfter = HWND_NOTOPMOST
    End If

    SetWindowPos Form.hWnd, hWndInsertAfter, 0, 0, 0, 0, _
                 SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-janv-14 01:31)  Decl: 14  Code: 19  Total: 33 Lines
':) CommentOnly: 2 (6,1%)  Commented: 0 (0%)  Empty: 8 (24,2%)  Max Logic Depth: 2
