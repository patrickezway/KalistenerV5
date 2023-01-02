Attribute VB_Name = "MKA_Alerte"
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Public Const ABS_AUTOHIDE = &H1
Public Const ABS_ONTOP = &H2
Public Const ABM_GETSTATE = &H4
Public Const ABM_GETTASKBARPOS = &H5
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type APPBARDATA
    cbSize As Long
    hwnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
End Type
Public Type POINTAPI
    X As Long
    Y As Long
End Type
'Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
'Public Const SWP_NOSIZE = &H1
'Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const REG_SZ = 1
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_DYN_DATA = &H80000004
Dim lng As Long
Public va

' Fenetre au 1er PLAN
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_TOP = 0
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
                                                    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
                                                    ByVal cy As Long, ByVal wFlags As Long) As Long

Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    
    On Error Resume Next
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, Chr$(0))
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        End If
    End If
    
End Function

Function GetString(hKey As Long, strPath As String, strValue As String)
    
    On Error Resume Next
    
    Dim Ret
    
    RegOpenKey hKey, strPath, Ret
    GetString = RegQueryStringValue(Ret, strValue)
    RegCloseKey Ret

End Function

Public Function ex(valu As String, nom As String, texte As String)

    RegCreateKey HKEY_LOCAL_MACHINE, valu, lng
    Ret = GetString(HKEY_LOCAL_MACHINE, valu, nom)
    
    If Ret = "" Then
        RegCreateKey HKEY_LOCAL_MACHINE, valu, lng
        RegSetValueEx lng, nom, 0&, 1, texte, Len(texte) + 1
    End If

End Function

Public Function P_AuPremierPlan(hwnd As Long) As Long

  '  P_AuPremierPlan = SetWindowPos(hwnd, -1, 0, 0, 0, 0, FLAGS)
    
End Function
