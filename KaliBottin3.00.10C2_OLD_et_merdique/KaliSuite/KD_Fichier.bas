Attribute VB_Name = "KD_Fichier"
Option Explicit

Public Function KF_CopierFichier(ByVal v_nomfich_src As String, _
                                 ByVal v_nomfich_dest As String) As Integer

    Dim liberr As String
    Dim cr As Integer
    
    cr = HTTP_Appel_copyfile(v_nomfich_src, v_nomfich_dest, liberr)
    If cr < 0 Then
        Call MsgBox("Impossible de copier " & v_nomfich_src & " dans " & v_nomfich_dest & vbCrLf & vbCrLf & liberr, vbInformation + vbOKOnly, "KF_CopierFichier (HTTP)")
        KF_CopierFichier = P_ERREUR
        Exit Function
    End If
    
End Function

Public Function KF_CreerRepertoire(ByVal v_nomrep As String) As Integer

    Dim liberr As String
    Dim cr As Integer
    
    cr = HTTP_Appel_creer_repertoire(v_nomrep, liberr)
    If cr = HTTP_OK Then
        KF_CreerRepertoire = P_OK
    Else
        KF_CreerRepertoire = P_ERREUR
    End If
    
End Function

Public Function KF_EffacerFichier(ByVal v_nomdoc As String, _
                                 ByVal v_bmesserr As Boolean) As Integer

    Dim liberr As String
    
    If HTTP_Appel_EffacerFichier(v_nomdoc, v_bmesserr, liberr) <> HTTP_OK Then
        KF_EffacerFichier = P_ERREUR
        Exit Function
    End If

End Function

Public Function KF_EffacerRepertoire(ByVal v_nomrep As String) As Integer

    Dim liberr As String
    Dim cr As Integer
    
    cr = HTTP_Appel_effacer_repertoire(v_nomrep, liberr)
    If cr = HTTP_OK Then
        KF_EffacerRepertoire = P_OK
    Else
        KF_EffacerRepertoire = P_ERREUR
    End If
    
End Function

Public Function KF_EstRepertoire(ByVal v_nomrep As String, _
                                 ByVal v_bmess As Boolean) As Boolean

    Dim liberr As String
    Dim cr As Integer
    
    cr = HTTP_Appel_est_repertoire(v_nomrep, v_bmess, liberr)
    If cr = HTTP_OK Then
        KF_EstRepertoire = True
    Else
        KF_EstRepertoire = False
    End If
    
End Function

Public Function KF_FichierExiste(ByVal v_nomfich As String) As Boolean

    Dim liberr As String
    Dim cr As Integer
    
    cr = HTTP_Appel_fichier_existe(v_nomfich, False, liberr)
    If cr = HTTP_OK Then
        KF_FichierExiste = True
    Else
        KF_FichierExiste = False
    End If
    
End Function

Public Function KF_GetFichier(ByVal v_nomfich_srv As String, _
                              ByVal v_nomfich_loc As String) As Integer

    Dim liberr As String
    
    Call FICH_EffacerFichier(v_nomfich_loc, False)
    
    If HTTP_Appel_GetFile(v_nomfich_srv, v_nomfich_loc, False, False, liberr) <> HTTP_OK Then
        Call MsgBox("Impossible de rapatrier " & v_nomfich_srv & " dans " & v_nomfich_loc & vbCrLf & "Erreur : " & liberr, vbInformation + vbOKOnly, "KF_GetFichier (HTTP)")
        KF_GetFichier = P_ERREUR
        Exit Function
    End If

    KF_GetFichier = P_OK

End Function

Public Function KF_PutFichier(ByVal v_nomfich_srv As String, _
                              ByVal v_nomfich_loc As String) As Integer
    
    Dim liberr As String
    
    If HTTP_Appel_PutFile(v_nomfich_srv, v_nomfich_loc, True, False, liberr) <> HTTP_OK Then
        KF_PutFichier = P_ERREUR
        Exit Function
    End If

    KF_PutFichier = P_OK

End Function

Public Function KF_RenommerFichier(ByVal v_nomsrc As String, _
                                   ByVal v_nomdest As String) As Integer

    Dim liberr As String
    Dim cr As Integer
    
    cr = HTTP_Appel_renamefile(v_nomsrc, v_nomdest, liberr)
    If cr < 0 Then
        Call MsgBox("Impossible de renommer " & v_nomsrc & " en " & v_nomdest & vbCrLf & vbCrLf & liberr, vbInformation + vbOKOnly, "KF_RenommerFichier (HTTP)")
        KF_RenommerFichier = P_ERREUR
        Exit Function
    End If
    
    KF_RenommerFichier = P_OK

End Function



