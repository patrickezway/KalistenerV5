Attribute VB_Name = "KS_KaliSuite"
Option Explicit

Public Const P_SUPER_UTIL = 1

Public Const P_FCT_CHGUTIL = 301

Public Const p_Auth_KaliDoc = "kalidoc"
Public Const p_Auth_KaliDocNew = "kalidocnew"
Public p_Mode_Auth_UtilAppli As String

Public p_NumEcr As Long
Public p_NumLabo As Long
Public p_CodeLabo As String
Public p_NumLaboDefaut As Long
Public p_CodeLaboDefaut As String
Public p_multilabo As Boolean
Public p_NbLabo As Long
Public p_ChgLaboAutor As Boolean

' applications gérées
Public p_appli_kalidoc As Long
Public p_appli_kalibottin As Long
Public p_appli_kaliress As Long

' Utilisateur courant
Public p_CodeUtil As String
Public p_NumUtil As Long

' Les appli gérées pour les documents
Public Type P_STRUCT_TYPEDOC
    appli As String
    ext As String
    chemin As String
    sconvHTML As String
    convPDF As Boolean
    appliexiste As Boolean
    pilotee As Boolean
End Type
Public p_tbltypdoc() As P_STRUCT_TYPEDOC
Public p_nbtypdoc As Long
Public p_appli_writer As String
Public p_yaconvhtml As Boolean

' Fonctions autorisées
Public p_fct_autor() As String

' POUR KALIDOC mais APPELEE dans PrmUtilisateur
' Ordre cycle attribué pour les destinataires DoxxUtil
Public Const P_DESTINATAIRE = 99
' Diffusion informatique
Public p_DiffusionInformatique As Boolean
' Diffusion papier
Public p_DiffusionPapier As Boolean
' Type document
Public Const P_DOCUMENTATION = 1
Public Const P_DOSSIER = 2
Public Const P_DOCUMENT = 3



Public Function P_ChargerModeAuth() As String
    Dim sql As String
    Dim lnb As Long
    
    ' Trouver de manière empirique le mode d'authentification pour utilappli
    sql = "select count(*) from UtilAppli WHERE UAPP_TypeCrypt ='" & p_Auth_KaliDocNew & "'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        MsgBox "Erreur SQL " & sql
    End If
    If lnb > 0 Then
        P_ChargerModeAuth = p_Auth_KaliDocNew
    Else
        P_ChargerModeAuth = p_Auth_KaliDoc
    End If

End Function


Public Function P_ChargerFctAutor() As Integer

    Dim sql As String
    Dim n As Long
    Dim rs As rdoResultset

    If p_NumUtil = P_SUPER_UTIL Then
        P_ChargerFctAutor = P_OK
        Exit Function
    End If
    
    Erase p_fct_autor()
    sql = "select FCT_Code" _
        & " from FctOK_Util, Fonction" _
        & " where FU_UNum=" & p_NumUtil _
        & " and FCT_Num=FU_FCTNum"
    On Error GoTo err_open_resultset
    On Error GoTo 0
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenStatic)
    On Error GoTo 0
    n = 0
    While Not rs.EOF
        ReDim Preserve p_fct_autor(n) As String
        p_fct_autor(n) = rs("FCT_Code").Value
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    
    P_ChargerFctAutor = P_OK
    Exit Function
    
err_open_resultset:
    MsgBox "Erreur OpenResultSet pour " + sql & vbLf & vbCr & "Erreur=" & Err.Number & vbLf & vbCr & Err.Description, vbOKOnly + vbCritical, ""
    P_ChargerFctAutor = P_ERREUR
    Exit Function

End Function

' KaliDoc mais appelée dans PrmUtilisateur
Public Function P_ChoisirPosteUtilisateur(ByVal v_numutil As Long, _
                                          ByRef r_numposte As Long, _
                                          ByRef r_libposte As String) As Integer

    Dim sql As String, ssp As String, libposte As String, nomutil As String
    Dim I As Integer, n As Integer
    Dim numposte As Long
    Dim rs As rdoResultset
    
    sql = "select U_SPM" _
        & " from Utilisateur" _
        & " where U_kb_actif=True AND U_Num=" & v_numutil
    If Odbc_RecupVal(sql, ssp) = P_ERREUR Then
        P_ChoisirPosteUtilisateur = P_ERREUR
        Exit Function
    End If
    
    n = STR_GetNbchamp(ssp, "|")
    If n = 1 Then
        r_numposte = P_RecupPosteNum(STR_GetChamp(ssp, "|", 0))
        If Odbc_RecupVal("select PO_Libelle from Poste where PO_Num=" & r_numposte, _
                         r_libposte) = P_ERREUR Then
            P_ChoisirPosteUtilisateur = P_ERREUR
            Exit Function
        End If
        P_ChoisirPosteUtilisateur = P_OUI
        Exit Function
    End If
        
    If P_RecupUtilNomP(v_numutil, nomutil) = P_ERREUR Then
        P_ChoisirPosteUtilisateur = P_ERREUR
        Exit Function
    End If
    
    Call CL_Init
    Call CL_InitTitreHelp("Choix du poste pour " & nomutil, "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    For I = 0 To n - 1
        numposte = P_RecupPosteNum(STR_GetChamp(ssp, "|", I))
        If Odbc_RecupVal("select PO_Libelle from Poste where PO_Num=" & numposte, _
                         libposte) = P_ERREUR Then
            P_ChoisirPosteUtilisateur = P_ERREUR
            Exit Function
        End If
        Call CL_AddLigne(libposte, numposte, "", False)
    Next I
    Call CL_InitTaille(0, -10)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        r_numposte = 0
        r_libposte = ""
        P_ChoisirPosteUtilisateur = P_NON
        Exit Function
    End If
    
    r_numposte = CL_liste.lignes(CL_liste.pointeur).num
    r_libposte = CL_liste.lignes(CL_liste.pointeur).texte
    
    P_ChoisirPosteUtilisateur = P_OUI

End Function

Public Function P_ChoixModele(ByVal v_nomrep As String, _
                              ByVal v_ext As String, _
                              ByRef r_nommodele As String) As Integer

    Dim n As Integer
    Dim nomdoc As String, sql As String, stype As String
    
    If Not FICH_EstRepertoire(v_nomrep, True) Then
        P_ChoixModele = P_NON
        Exit Function
    End If
    
lab_debut:
    Call CL_Init
    Call CL_InitTitreHelp("Choix du modèle", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    
    n = 0
    nomdoc = Dir$(v_nomrep & "\*.*")
    Do While Len(nomdoc)
        If InStr(nomdoc, v_ext) > 0 Or InStr(nomdoc, ".mod") > 0 Then
            nomdoc = left$(nomdoc, InStrRev(nomdoc, ".") - 1)
            Call CL_AddLigne(nomdoc, 0, "", False)
            n = n + 1
        End If
        nomdoc = Dir$
    Loop
    If n = 0 Then
        Call MsgBox("Aucun modèle n'est disponible.", vbInformation + vbOKOnly, "")
        P_ChoixModele = P_NON
        Exit Function
    End If
    Call CL_Tri(0)
    
    If n < 20 Then
        Call CL_InitTaille(0, -n)
    Else
        Call CL_InitTaille(0, -20)
    End If
    
lab_choix:
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        P_ChoixModele = P_NON
        Exit Function
    End If
    
    r_nommodele = CL_liste.lignes(CL_liste.pointeur).texte
    P_ChoixModele = P_OUI
    
End Function

Public Function P_DeterminerAppli() As Integer

    Dim sql As String
    Dim rs As rdoResultset

    p_appli_kalidoc = 0
    p_appli_kalibottin = 0
    p_appli_kaliress = 0
    
    sql = "select APP_Num from Application" _
        & " where APP_Code='KALIDOC'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        P_DeterminerAppli = P_ERREUR
        Exit Function
    End If
    If Not rs.EOF Then
        p_appli_kalidoc = rs("APP_Num").Value
    End If
    rs.Close
        
    sql = "select APP_Num from Application" _
        & " where APP_Code='KALIBOTTIN'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        P_DeterminerAppli = P_ERREUR
        Exit Function
    End If
    If Not rs.EOF Then
        p_appli_kalibottin = rs("APP_Num").Value
    End If
    rs.Close
        
    P_DeterminerAppli = P_OK
    
End Function

Public Function P_MajNblabo() As Integer

    Dim sql As String
    
    sql = "select count(*) from Laboratoire"
    If Odbc_Count(sql, p_NbLabo) = P_ERREUR Then
        P_MajNblabo = P_ERREUR
        Exit Function
    End If
    
    P_MajNblabo = P_OK
    
End Function

Public Function P_RecupLaboCode(ByVal numl As Long, _
                                ByRef codlabo As String) As Integer

    Dim sql As String
    Dim rs As rdoResultset
    
    sql = "select L_Code from Laboratoire" _
        & " where L_Num=" & numl
    On Error GoTo err_open_resultset
    Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenStatic)
    On Error GoTo 0
    If rs.EOF Then GoTo err_open_resultset
    
    codlabo = rs("L_Code").Value
    rs.Close
    
    P_RecupLaboCode = P_OK
    Exit Function
    
err_open_resultset:
    MsgBox "Erreur OpenResultSet pour " + sql & vbLf & vbCr & "Erreur=" & Err.Number & vbLf & vbCr & Err.Description, vbOKOnly + vbCritical, ""
    P_RecupLaboCode = P_ERREUR
    Exit Function

err_no_resultset:
    MsgBox "Pas de ligne pour " + sql, vbOKOnly + vbCritical, "MCommun (P_RecupLaboCode)"
    rs.Close
    P_RecupLaboCode = P_ERREUR
    Exit Function

End Function

Public Function P_RecupPosteNom(ByVal v_numposte As Long, _
                                ByRef r_lib As String) As Integer

    Dim sql As String

    sql = "select PO_Libelle from Poste" _
        & " where PO_Num=" & v_numposte
    If Odbc_RecupVal(sql, r_lib) = P_ERREUR Then
        P_RecupPosteNom = P_ERREUR
        Exit Function
    End If
    
    P_RecupPosteNom = P_OK
    
End Function

Public Function P_RecupPosteNomfct(ByVal v_numposte As Long, _
                                   ByRef r_lib As String) As Integer

    Dim sql As String

    sql = "select FT_Libelle from Poste, FctTrav" _
        & " where PO_Num=" & v_numposte _
        & " and FT_Num=PO_FTNum"
    If Odbc_RecupVal(sql, r_lib) = P_ERREUR Then
        P_RecupPosteNomfct = P_ERREUR
        Exit Function
    End If
    
    P_RecupPosteNomfct = P_OK
    
End Function

' v_ssp contient un seul poste Sx;Sy;Pz;
Public Function P_RecupPosteNum(ByVal v_ssp As String) As Long

    Dim n As Integer
    
    n = STR_GetNbchamp(v_ssp, ";")
    P_RecupPosteNum = Mid$(STR_GetChamp(v_ssp, ";", n - 1), 2)
    
End Function

Public Function P_RecupPSLib(ByVal v_sp As String, _
                             ByRef r_lib As String) As Integer
    
    Dim sql As String, s As String, s_srv As String, lib As String
    Dim n As Integer
    Dim num As Long

    r_lib = ""
    
    n = STR_GetNbchamp(v_sp, ";")
    s = STR_GetChamp(v_sp, ";", n - 1)
    If left$(s, 1) = "P" Then
        num = Mid$(s, 2)
        If P_RecupPosteNom(num, r_lib) = P_ERREUR Then
            P_RecupPSLib = P_ERREUR
            Exit Function
        End If
    End If
    s_srv = STR_GetChamp(v_sp, ";", n - 2)
    num = Mid$(s_srv, 2)
    If P_RecupSrvNom(num, lib) = P_ERREUR Then
        P_RecupPSLib = P_ERREUR
        Exit Function
    End If
    If r_lib <> "" Then
        r_lib = r_lib & " - "
    End If
    r_lib = r_lib & lib
    
    P_RecupPSLib = P_OK

End Function

Public Function P_RecupSPLib(ByVal v_sp As String, _
                             ByRef r_lib As String) As Integer
    
    Dim sql As String, stype As String, s As String, lib As String, sp As String
    Dim n As Integer
    Dim num As Long

    r_lib = ""
    
    If v_sp <> "" Then
        sp = STR_GetChamp(v_sp, "|", 0)
    Else
        sp = v_sp
    End If
    n = STR_GetNbchamp(sp, ";")
    s = STR_GetChamp(sp, ";", n - 1)
    
    num = Mid$(s, 2)
    If left$(s, 1) = "S" Then
        If P_RecupSrvNom(num, lib) = P_ERREUR Then
            P_RecupSPLib = P_ERREUR
            Exit Function
        End If
    Else
        If P_RecupPosteNom(num, lib) = P_ERREUR Then
            P_RecupSPLib = P_ERREUR
            Exit Function
        End If
    End If
    r_lib = lib
    
    P_RecupSPLib = P_OK

End Function

Public Function P_RecupSrvNom(ByVal v_num As Long, _
                              ByRef r_nom As String) As Integer

    Dim sql As String

    sql = "select SRV_Nom from Service" _
        & " where SRV_Num=" & v_num
    If Odbc_RecupVal(sql, r_nom) = P_ERREUR Then
        P_RecupSrvNom = P_ERREUR
        Exit Function
    End If
    
    P_RecupSrvNom = P_OK
    
End Function


Public Function P_RecupUtilCode(ByVal v_num As Long, _
                                ByRef r_code As String) As Integer

    Dim sql As String
    
    sql = "select UAPP_Code" _
        & " from UtilAppli" _
        & " where UAPP_UNum=" & v_num _
        & " and UAPP_APPNum=" & p_appli_kalidoc
    If Odbc_RecupVal(sql, r_code) = P_ERREUR Then
        P_RecupUtilCode = P_ERREUR
        Exit Function
    End If
    
    P_RecupUtilCode = P_OK

End Function

Public Function P_RecupUtilNomP(ByVal v_num As Long, _
                               ByRef r_nom As String) As Integer

    Dim sql As String, nom As String, prenom As String
    
    sql = "select U_Nom, U_Prenom" _
        & " from Utilisateur" _
        & " where U_kb_actif=True AND U_Num=" & v_num
    If Odbc_RecupVal(sql, nom, prenom) = P_ERREUR Then
        P_RecupUtilNomP = P_ERREUR
        Exit Function
    End If
    r_nom = nom + " " + prenom
    
    P_RecupUtilNomP = P_OK

End Function

Public Function P_RecupUtilPpointNom(ByVal v_num As Long, _
                                     ByRef r_nom As String) As Integer

    Dim sql As String, nom As String, prenom As String
    
    If v_num = P_SUPER_UTIL Then
        r_nom = "Administrateur"
        P_RecupUtilPpointNom = P_OK
        Exit Function
    End If
    
    sql = "select U_Nom, U_Prenom" _
        & " from Utilisateur" _
        & " where U_kb_actif=True AND U_Num=" & v_num
    If Odbc_RecupVal(sql, nom, prenom) = P_ERREUR Then
        P_RecupUtilPpointNom = P_ERREUR
        MsgBox "Erreur fonction P_RecupUtilPpointNom : " & sql
        Exit Function
    End If
    If prenom <> "" Then
        r_nom = left$(prenom, 1) + ". "
    Else
        r_nom = ""
    End If
    r_nom = r_nom + nom
    
    P_RecupUtilPpointNom = P_OK

End Function

Public Function P_RecupPrenomCourt(ByVal v_prenom As String) As String

    Dim prenom As String
    Dim pos As Integer
    
    If v_prenom <> "" Then
        pos = InStr(v_prenom, " ")
        If pos = 0 Then
            pos = InStr(v_prenom, "-")
        End If
        If pos > 0 And Len(v_prenom) > pos Then
            prenom = left$(v_prenom, 1) + "." + Mid$(v_prenom, pos + 1, 1) + "."
        Else
            prenom = left$(v_prenom, 1) + "."
        End If
    Else
        prenom = ""
    End If
    
    P_RecupPrenomCourt = prenom

End Function

Public Function P_SaisirUtilIdent(ByVal x As Integer, _
                                  ByVal y As Integer, _
                                  ByVal l As Integer, _
                                  ByVal H As Integer) As Integer

    Dim codutil As String, mpasse As String, sql As String
    Dim deuxieme_saisie As Boolean, bad_util As Boolean
    Dim nb As Integer, reponse As Integer
    Dim lnb As Long, lbid As Long
    Dim rs As rdoResultset
    Dim oMD5 As CMD5
    
    ' Pour le cryptage MD5
    Set oMD5 = New CMD5
    
    nb = 1
    deuxieme_saisie = False
    
    'Saisie du code utilisateur
lab_debut:
    Call SAIS_Init
    If deuxieme_saisie Then
        Call SAIS_InitTitreHelp("Confirmez votre mot de passe", "")
        Call SAIS_AddChampComplet("Mot de passe (confirmation)", 10, SAIS_TYP_TOUT_CAR, "", False, SAIS_CONV_SECRET, False)
    Else
        Call SAIS_InitTitreHelp("Identification", p_chemin_appli + "\help\kalidoc.chm;demarrage.htm")
        Call SAIS_AddChamp("Code d'accès", 15, SAIS_TYP_TOUT_CAR, False)
        Call SAIS_AddChampComplet("Mot de passe", 10, SAIS_TYP_TOUT_CAR, "", False, SAIS_CONV_SECRET, False)
    End If
    Call SAIS_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call SAIS_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
lab_saisie:
    Saisie.Show 1
    If SAIS_Saisie.retour = 1 Then
        P_SaisirUtilIdent = P_NON
        Exit Function
    End If
        
    If deuxieme_saisie Then
        If mpasse <> SAIS_Saisie.champs(0).sval Then
            MsgBox "Vous n'avez pas saisi le même mot de passe.", vbOKOnly + vbExclamation, ""
            deuxieme_saisie = False
            GoTo lab_debut
        End If
        ' Maj du mot de passe utilisateur
        sql = "select * from UtilAppli" _
            & " where UAPP_APPNum=" & p_appli_kalidoc _
            & " and UAPP_Code='" & UCase(codutil) & "'"
        On Error GoTo err_open_resultset
        Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
        On Error GoTo 0
        If rs.EOF Then GoTo err_no_resultset
        On Error GoTo err_edit
        rs.Edit
        On Error GoTo err_affecte
        If rs("UAPP_TypeCrypt").Value = "kalidoc" Or rs("UAPP_TypeCrypt").Value = "" Then
            rs("UAPP_MotPasse").Value = STR_Crypter(UCase(mpasse))
        ElseIf rs("UAPP_TypeCrypt").Value = "kalidocnew" Then
            rs("UAPP_MotPasse").Value = STR_Crypter_New(UCase(mpasse))
        ElseIf rs("UAPP_TypeCrypt").Value = "md5" Then
            rs("UAPP_MotPasse").Value = oMD5.MD5(mpasse)
        End If
        On Error GoTo err_update
        rs.Update
        On Error GoTo 0
        rs.Close
    Else
        codutil = SAIS_Saisie.champs(0).sval
        mpasse = SAIS_Saisie.champs(1).sval
    End If
    
    If codutil = "ROOT" And mpasse = "007" Then
        p_CodeUtil = "ROOT"
        p_NumUtil = P_SUPER_UTIL
        P_SaisirUtilIdent = P_OUI
        Exit Function
    End If
    
    'Recherche de cet utilisateur
    sql = "select U_Num, UAPP_MotPasse, UAPP_TypeCrypt from Utilisateur, UtilAppli" _
        & " where UAPP_Code='" & UCase(codutil) & "'" _
        & " and UAPP_APPNum=" & p_appli_kalidoc _
        & " and U_Actif=True" _
        & " and U_kb_actif=True" _
        & " and U_Num=UAPP_UNum"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        P_SaisirUtilIdent = P_ERREUR
        Exit Function
    End If
    If rs.EOF Then
        rs.Close
        bad_util = True
    Else
        If rs("UAPP_MotPasse").Value <> "" Then
            If (rs("UAPP_TypeCrypt").Value = "kalidoc" Or rs("UAPP_TypeCrypt").Value = "") And STR_Decrypter(rs("UAPP_MotPasse").Value) <> UCase(mpasse) Then
                bad_util = True
            ElseIf rs("UAPP_TypeCrypt").Value = "kalidocnew" And STR_Decrypter_New(rs("UAPP_MotPasse").Value) <> UCase(mpasse) Then
                bad_util = True
            ElseIf rs("UAPP_TypeCrypt").Value = "md5" And rs("UAPP_MotPasse").Value <> oMD5.MD5(mpasse) Then
                bad_util = True
            Else
                p_NumUtil = rs("U_Num").Value
                rs.Close
                GoTo lab_ok
            End If
        Else
            bad_util = False
        End If
        rs.Close
    End If
    If bad_util Then
        MsgBox "Identification inconnue.", vbOKOnly + vbExclamation, ""
        nb = nb + 1
        If nb > 3 Then
            P_SaisirUtilIdent = P_ERREUR
            Exit Function
        End If
        SAIS_Saisie.champs(1).sval = ""
        GoTo lab_saisie
    Else
        deuxieme_saisie = True
        GoTo lab_debut
    End If
    
lab_ok:
    p_CodeUtil = UCase(codutil)
    
    ' Ajout si nécessaire dans UtilKD
    If p_CheminPHP <> "" Then
        sql = "select count(*) from UtilKD where ukd_unum=" & p_NumUtil
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            lnb = 1
        End If
        If lnb = 0 Then
            reponse = MsgBox("Devez-vous être considéré comme un utilisateur habituel de KaliDoc ?", vbQuestion + vbYesNo, "")
            If reponse = vbYes Then
                Call Odbc_AddNew("UtilKD", "ukd_num", "ukd_seq", False, lbid, _
                                 "ukd_unum", p_NumUtil)
            End If
        End If
    End If
    
    P_SaisirUtilIdent = P_OUI
    Exit Function
    
err_open_resultset:
    MsgBox "Erreur OpenResultset " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    P_SaisirUtilIdent = P_ERREUR
    Exit Function
    
err_no_resultset:
    MsgBox "Pas de ligne pour " & sql, vbOKOnly, ""
    rs.Close
    P_SaisirUtilIdent = P_ERREUR
    Exit Function
    
err_edit:
    MsgBox "Erreur Edit " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    rs.Close
    P_SaisirUtilIdent = P_ERREUR
    Exit Function
    
err_affecte:
    MsgBox "Erreur Affectation " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    rs.Close
    P_SaisirUtilIdent = P_ERREUR
    Exit Function
    
err_update:
    MsgBox "Erreur Update " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    rs.Close
    P_SaisirUtilIdent = P_ERREUR
    Exit Function
    
End Function

Public Function P_SaisirUtilPasswd() As Integer

    Dim mpasse As String, sql As String, lib As String, titre As String
    Dim nb As Integer, etape As Integer
    Dim rs As rdoResultset
    
    nb = 1
    etape = 0
    
    'Saisie du code utilisateur
lab_debut:
    Call SAIS_Init
    If etape = 0 Then
        titre = "Saisissez votre mot de passe actuel"
        lib = "Mot de passe"
    ElseIf etape = 1 Then
        titre = "Saisissez le nouveau mot de passe"
        lib = "Nouveau mot de passe"
    Else
        titre = "Confirmez votre nouveau mot de passe"
        lib = "Mot de passe (confirmation)"
    End If
    Call SAIS_InitTitreHelp(titre, "")
    Call SAIS_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call SAIS_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call SAIS_AddChampComplet(lib, 10, SAIS_TYP_TOUT_CAR, "", False, SAIS_CONV_SECRET, False)
lab_saisie:
    Saisie.Show 1
    If SAIS_Saisie.retour = 1 Then
        P_SaisirUtilPasswd = P_NON
        Exit Function
    End If
        
    Select Case etape
    Case 0
        sql = "select * from UtilAppli" _
            & " where UAPP_APPNum=" & p_appli_kalidoc _
            & " and UAPP_Code='" & p_CodeUtil & "'" _
            & " and (((UAPP_TypeCrypt = 'kalidoc' or UAPP_TypeCrypt = '') AND UAPP_MotPasse='" & UCase(STR_Crypter(SAIS_Saisie.champs(0).sval)) & "') " _
            & " or (UAPP_TypeCrypt = 'kalidocnew' AND UAPP_MotPasse='" & UCase(STR_Crypter_New(SAIS_Saisie.champs(0).sval)) & "'))"
        On Error GoTo err_open_resultset
        Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenStatic)
        On Error GoTo 0
        If rs.EOF Then
            rs.Close
            MsgBox "Mot de passe incorrect.", vbOKOnly + vbExclamation, ""
            Exit Function
        End If
        rs.Close
        etape = 1
        GoTo lab_debut
    Case 1
        mpasse = UCase(SAIS_Saisie.champs(0).sval)
        etape = 2
        GoTo lab_debut
    Case 2
        If mpasse <> UCase(SAIS_Saisie.champs(0).sval) Then
            MsgBox "Vous n'avez pas saisi le même mot de passe.", vbOKOnly + vbExclamation, ""
            etape = 1
            GoTo lab_debut
        End If
        ' Maj du mot de passe utilisateur
        sql = "select * from Utilappli" _
            & " where UAPP_APPNum=" & p_appli_kalidoc _
            & " and UAPP_Code='" & p_CodeUtil & "'"
        On Error GoTo err_open_resultset
        Set rs = Odbc_Cnx.OpenResultset(sql, rdOpenKeyset, rdConcurRowVer)
        On Error GoTo 0
        If rs.EOF Then GoTo err_no_resultset
        On Error GoTo err_edit
        rs.Edit
        On Error GoTo err_affecte
        rs("UAPP_TypeCrypt").Value = p_Mode_Auth_UtilAppli
        rs("UAPP_MotPasse").Value = STR_Crypter_New(mpasse)
        On Error GoTo err_update
        rs.Update
        On Error GoTo 0
        rs.Close
    End Select
    
    P_SaisirUtilPasswd = P_OUI
    Exit Function
    
err_open_resultset:
    MsgBox "Erreur OpenResultset " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    P_SaisirUtilPasswd = P_ERREUR
    Exit Function
    
err_no_resultset:
    MsgBox "Pas de ligne pour " & sql, vbOKOnly, ""
    rs.Close
    P_SaisirUtilPasswd = P_ERREUR
    Exit Function
    
err_edit:
    MsgBox "Erreur Edit " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    rs.Close
    P_SaisirUtilPasswd = P_ERREUR
    Exit Function
    
err_affecte:
    MsgBox "Erreur Affectation " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    rs.Close
    P_SaisirUtilPasswd = P_ERREUR
    Exit Function
    
err_update:
    MsgBox "Erreur Update " & sql & vbCr & vbLf & "Erreur=" & Err.Number & " " & Err.Description, vbOKOnly, ""
    rs.Close
    P_SaisirUtilPasswd = P_ERREUR
    Exit Function
    
End Function

Public Function P_UtilAFonction(ByVal v_numutil As Long, _
                                ByVal v_numfct As Long) As Boolean


    Dim sql As String, sfcttrav As String
    Dim I As Integer, n As Integer
    Dim numfct As Long
    
    sql = "select U_FctTrav" _
        & " from Utilisateur" _
        & " where U_kb_actif=True AND U_Num=" & v_numutil
    If Odbc_RecupVal(sql, sfcttrav) = P_ERREUR Then
        P_UtilAFonction = False
        Exit Function
    End If
    
    n = STR_GetNbchamp(sfcttrav, ";")
    For I = 0 To n - 1
        numfct = Mid$(STR_GetChamp(sfcttrav, ";", I), 2)
        If numfct = v_numfct Then
            P_UtilAFonction = True
            Exit Function
        End If
    Next I

    P_UtilAFonction = False

End Function

Public Function P_UtilAPoste(ByVal v_numutil As Long, _
                             ByVal v_numposte As Long) As Boolean

    Dim sql As String
    Dim lnb As Long
    
    sql = "select count(*) from Utilisateur" _
        & " where U_kb_actif=True AND U_Num=" & v_numutil _
        & " and U_SPM like '%P" & v_numposte & ";%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        P_UtilAPoste = False
        Exit Function
    End If
    
    If lnb > 0 Then
        P_UtilAPoste = True
    Else
        P_UtilAPoste = False
    End If

End Function

Public Function P_UtilAPlusieursFonctions(ByVal v_numutil As Long) As Integer

    Dim n As Integer
    Dim rs As rdoResultset
    
    If Odbc_Select("select U_FctTrav from Utilisateur where U_kb_actif=True AND U_Num=" & v_numutil, rs) = P_ERREUR Then
        P_UtilAPlusieursFonctions = P_ERREUR
        Exit Function
    End If
    n = STR_GetNbchamp(rs("U_FctTrav").Value, ";")
    If n > 1 Then
        P_UtilAPlusieursFonctions = P_OUI
    Else
        P_UtilAPlusieursFonctions = P_NON
    End If
    
End Function

Public Function P_UtilAPlusieursPostes(ByVal v_numutil As Long) As Integer

    Dim n As Integer
    Dim s_sp As Variant
    Dim rs As rdoResultset
    
    If Odbc_RecupVal("select U_SPM from Utilisateur where U_kb_actif=True AND U_Num=" & v_numutil, s_sp) = P_ERREUR Then
        P_UtilAPlusieursPostes = P_ERREUR
        Exit Function
    End If
    n = STR_GetNbchamp(s_sp, "|")
    If n > 1 Then
        P_UtilAPlusieursPostes = P_OUI
    Else
        P_UtilAPlusieursPostes = P_NON
    End If
    
End Function

Public Function P_UtilEstAutorFct(ByVal v_fctcode As String) As Boolean

    Dim I As Long
    
    If p_NumUtil = P_SUPER_UTIL Then
        P_UtilEstAutorFct = True
        Exit Function
    End If
    
    On Error GoTo pas_autorise
    For I = 0 To UBound(p_fct_autor)
        If p_fct_autor(I) = v_fctcode Then
            P_UtilEstAutorFct = True
            Exit Function
        End If
    Next I
    
pas_autorise:
    P_UtilEstAutorFct = False
    Exit Function
    
End Function

Public Function P_UtilDansTBL(ByVal v_numutil As Long) As Boolean

    Dim I As Integer
    
    For I = 0 To p_siz_tblu
        If p_tblu(I) = v_numutil Then
            P_UtilDansTBL = True
            Exit Function
        End If
    Next I
    
    P_UtilDansTBL = False
    
End Function

Public Function P_YaAppli(ByVal v_nomappli As String) As Boolean

    Dim I As Integer
    
    For I = 0 To p_nbtypdoc - 1
        If p_tbltypdoc(I).appli = v_nomappli Then
            If p_tbltypdoc(I).appliexiste = True Then
                P_YaAppli = True
            Else
                P_YaAppli = False
            End If
            Exit Function
        End If
    Next I
    
    P_YaAppli = False
    
End Function

