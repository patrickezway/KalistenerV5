Attribute VB_Name = "KaliService"
Option Explicit

Public Function P_RecupPosteNom(ByVal v_num As Long, _
                                ByRef r_lib As String) As Integer

    Dim sql As String

    sql = "SELECT FT_Libelle FROM Poste, FctTrav" _
        & " WHERE PO_Num=" & v_num _
        & " AND FT_Num=PO_FTNum"
    If Odbc_RecupVal(sql, r_lib) = P_ERREUR Then
        P_RecupPosteNom = P_ERREUR
        Exit Function
    End If

    P_RecupPosteNom = P_OK

End Function

Public Function P_RecupSPLib(ByVal v_sp As String, _
                             ByRef r_lib As String) As Integer

    Dim sql As String, stype As String, s As String, lib As String, sp As String
    Dim n As Integer
    Dim num As Long

    If v_sp <> "" Then
        sp = STR_GetChamp(v_sp, "|", 0)
    Else
        sp = v_sp
    End If
    n = STR_GetNbchamp(sp, ";")
    If n = 0 Then
        r_lib = ""
        P_RecupSPLib = P_OK
        Exit Function
    End If
    If n = 1 Then
        s = STR_GetChamp(sp, ";", 0)
    Else
        s = STR_GetChamp(sp, ";", n - 1)
        If left$(s, 1) = "S" Then
            n = 1
        Else
            s = STR_GetChamp(sp, ";", n - 2)
        End If
    End If
    num = Mid$(s, 2)
    If P_RecupSrvNom(num, lib) = P_ERREUR Then
        P_RecupSPLib = P_ERREUR
        Exit Function
    End If
    r_lib = lib

    If n > 1 Then
        s = STR_GetChamp(sp, ";", n - 1)
        num = Mid$(s, 2)
        If P_RecupPosteNom(num, lib) = P_ERREUR Then
            P_RecupSPLib = P_ERREUR
            Exit Function
        End If
        r_lib = r_lib & " - " & lib
    End If

    P_RecupSPLib = P_OK

End Function

Public Function P_RecupSrvNom(ByVal v_num As Long, _
                              ByRef r_nom As String) As Integer

    Dim sql As String

    sql = "SELECT SRV_Nom FROM Service" _
        & " WHERE SRV_Num=" & v_num
    If Odbc_RecupVal(sql, r_nom) = P_ERREUR Then
        P_RecupSrvNom = P_ERREUR
        Exit Function
    End If

    P_RecupSrvNom = P_OK

End Function


