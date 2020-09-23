Attribute VB_Name = "modRunAtStartUp"
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
                                ByVal hKey As Long, _
                                ByVal lpSubKey As String, _
                                ByVal ulOptions As Long, _
                                ByVal samDesired As Long, _
                                phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
                                ByVal hKey As Long, _
                                ByVal lpValueName As String, _
                                ByVal lpReserved As Long, _
                                lpType As Long, _
                                lpData As Any, _
                                lpcbData As Long) As Long
    ' Note : Si vous passez le paramètre lpData en String, faut le passer 'ByVal'
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" ( _
                                ByVal hKey As Long, _
                                ByVal lpSubKey As String, _
                                ByVal Reserved As Long, _
                                ByVal lpClass As String, _
                                ByVal dwOptions As Long, _
                                ByVal samDesired As Long, _
                                ByVal lpSecurityAttributes As Long, _
                                phkResult As Long, _
                                lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
                                ByVal hKey As Long, _
                                ByVal lpValueName As String, _
                                ByVal Reserved As Long, _
                                ByVal dwType As Long, _
                                lpData As Any, _
                                ByVal cbData As Long) As Long
    ' Note : Si vous passez le paramètre lpData en String, faut le passer 'ByVal'
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" ( _
                                ByVal hKey As Long, _
                                ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Const HKEY_LOCAL_MACHINE        As Long = &H80000002
Private Const SYNCHRONIZE               As Long = &H100000
Private Const READ_CONTROL              As Long = &H20000

Private Const STANDARD_RIGHTS_READ      As Long = (READ_CONTROL)
Private Const KEY_QUERY_VALUE           As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Private Const KEY_NOTIFY                As Long = &H10
Private Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const STANDARD_RIGHTS_WRITE     As Long = (READ_CONTROL)
Private Const KEY_SET_VALUE             As Long = &H2
Private Const KEY_CREATE_SUB_KEY        As Long = &H4
Private Const KEY_WRITE As Long = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Private Const REG_SZ                    As Long = 1
Private Const ERROR_SUCCESS             As Long = &H0&
'

' Renvoie True si l'application se lance avec la session Windows
Public Function WillRunAtStartup(ByVal app_name As String) As Boolean

    Dim hKey As Long
    Dim value_type As Long

    ' Regarde si la clé existe
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
                    "Software\Microsoft\Windows\CurrentVersion\Run", _
                    0, _
                    KEY_READ, _
                    hKey) = ERROR_SUCCESS Then
        WillRunAtStartup = (RegQueryValueEx(hKey, app_name, _
                                            ByVal 0&, _
                                            value_type, _
                                            ByVal 0&, _
                                            ByVal 0&) = ERROR_SUCCESS)
        ' Close the registry key handle.
        RegCloseKey hKey
    Else
        ' Can't find the key.
        WillRunAtStartup = False
    End If

End Function

'###################################################################################################
' Determine whether the program will run at startup.
' To run at startup, there should be a key in:
'   HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run
' named after the program's executable with value giving its path.

Public Sub SetRunAtStartup(ByVal app_name As String, _
                           ByVal app_path As String, _
                           Optional ByVal run_at_startup As Boolean = True, _
                           Optional ByVal FollowingParameter As String = "")
    ' Exemple :
    '    SetRunAtStartup App.EXEName, App.Path, True, "/debug"
    
    Dim hKey As Long
    Dim key_value As String

    On Error GoTo SetStartupError
    
    ' Open the key, creating it if it doesn't exist.
    If RegCreateKeyEx(HKEY_LOCAL_MACHINE, _
                      "Software\Microsoft\Windows\CurrentVersion\Run", _
                      ByVal 0&, ByVal 0&, ByVal 0&, _
                      KEY_WRITE, ByVal 0&, hKey, ByVal 0&) <> ERROR_SUCCESS Then
        Exit Sub
    End If

    ' See if we should run at startup.
    If run_at_startup Then
        ' Create the key.
        key_value = """" & app_path & "\" & app_name & ".exe"""
        If FollowingParameter <> "" Then key_value = key_value & " " & FollowingParameter
        key_value = key_value & vbNullChar
        Call RegSetValueEx(hKey, app_name, 0, REG_SZ, _
                           ByVal key_value, Len(key_value))
    Else
        ' Delete the value.
        Call RegDeleteValue(hKey, app_name)
    End If

    ' Close the key.
    RegCloseKey hKey
    Exit Sub

SetStartupError:
    'msgboxex Err.Number & " " & Err.Description
    Exit Sub
    
End Sub


