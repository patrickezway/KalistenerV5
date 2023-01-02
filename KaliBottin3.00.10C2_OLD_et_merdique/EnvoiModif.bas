Attribute VB_Name = "MEnvoiModif"
Option Explicit

Private g_on_a_genere As Boolean
Private Function a_des_infoSuppl(ByVal v_unum As Long, _
                            ByVal v_app_infoSuppl As String) As Boolean
' chercher si on a des info suppl à envoyer pour cette personne
    Dim sql As String
    Dim lnb As Long, tis_num As Long
    Dim nbr_tis As Integer, I As Integer

    If Len(v_app_infoSuppl) = 0 Then GoTo lab_erreur

    nbr_tis = STR_GetNbchamp(v_app_infoSuppl, ";")
    ' parcourir les info suppl de l'application en cours
    For I = 0 To nbr_tis - 1
        tis_num = Mid$(STR_GetChamp(v_app_infoSuppl, ";", I), 2)
        sql = "SELECT COUNT(*) FROM InfoSupplEntite" & _
              " WHERE ISE_Type='U' AND ISE_TypeNum=" & v_unum & _
              " AND ISE_TisNum=" & tis_num
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            GoTo lab_erreur
        End If
        If lnb > 0 Then ' la personne est concernée !
            a_des_infoSuppl = True
            Exit Function
        End If
    Next I

lab_erreur:
    a_des_infoSuppl = False

End Function

Private Function CreerLesFichiers() As Integer
' ******************************************************************************************************
' 1° Parcourrir toutes les applications (sauf KaliDoc et KaliBottin)
' 2° Stocker les fichiers dans le repértoire approprié: pgd_chaminsmbk/kalibottin/mouvements/CODE_APPLI/
'       resume_env_AAAAMMJJ-gen.txt (U_Num#Mouvement#Description)
'       enclair_env_AAAAMMJJ-gen.txt
' 3° Créer 2 fichiers (résumé & détaillé) par application
' ******************************************************************************************************
    Dim sql As String, message As String, bilan As String, nomResponsable As String, _
        prenomResponsable As String, chemin As String, nomFichierEnClair As String, _
        nomFichierResume As String, listePostesValides As String, nomfich1 As String, nomfich2 As String
    Dim ecrire As Boolean, finfo As Boolean
    Dim cr As Integer, fd_resume As Integer, fd_enclair As Integer, reponse As Integer, _
        nbr_ecrit As Integer
    Dim numresp As Long
    Dim rs_appli As rdoResultset, rs_util As rdoResultset
    Dim frm As Form
    Dim nbtot_appli As Integer, nbtot_util As Integer
    Dim nb_util As Integer
    Dim nb_appli As Integer
    Dim bPoserQuestion As Boolean
    
    sql = "SELECT * FROM Application" _
        & " WHERE APP_Num<>" & p_appli_kalidoc & " AND APP_Num<>" & p_appli_kalibottin _
        & " AND APP_Num<>0" ' le zéro pour l'importation
    If Odbc_SelectV(sql, rs_appli) = P_ERREUR Then
        GoTo lab_erreur
    End If
    ' Parcourrir toutes les applications (sauf KaliDoc & KaliBottin)
    ' **************************************************************
    Menu.FrmFichiers.Visible = True
    nbtot_appli = rs_appli.RowCount
    Menu.PgBarAppli.Max = nbtot_appli
    nb_appli = 0
    While Not rs_appli.EOF
        nb_appli = nb_appli + 1
        Menu.PgBarAppli.Value = nb_appli
        nbr_ecrit = 0
        bPoserQuestion = True
        ' Le responsable de cette application doit il être informé ?
        ' Si NON => Passer à l'application suivante
        If Not rs_appli("APP_Informer").Value Then
            GoTo lab_appli_suivante
        End If
        If IsNull(rs_appli("app_lstresp")) Or rs_appli("app_lstresp") = "" Then
            Call MsgBox("Les personnes à informer pour l'application " & rs_appli("APP_Nom") & Chr(13) & Chr(10) & "ne sont pas renseignées", vbCritical + vbOKOnly)
            GoTo lab_appli_suivante
        End If
        ' Ici on commence à traiter l'application
        ' le responsable
        numresp = Mid$(STR_GetChamp(rs_appli("app_lstresp").Value, ";", 0), 2)
        If Odbc_RecupVal("SELECT U_Nom, U_Prenom FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & numresp, _
                          nomResponsable, prenomResponsable) = P_ERREUR Then
            GoTo lab_erreur
        End If
        ' créer le repértoire s'il n'existe pas déjà
        chemin = p_CheminKW & "/kalibottin/mouvements/" & rs_appli("APP_Code").Value
        If Not KF_EstRepertoire(chemin, False) Then
            Call KF_CreerRepertoire(chemin)
        End If
        ' ouverture des deux fichiers
        nomfich1 = p_chemin_appli + "/tmp/kb_enclair.txt"
        nomfich2 = p_chemin_appli + "/tmp/kb_resume.txt"
        Call FICH_EffacerFichier(nomfich1, False)
        Call FICH_EffacerFichier(nomfich2, False)
        If FICH_OuvrirFichier(nomfich1, FICH_ECRITURE, fd_enclair) = P_ERREUR Then GoTo lab_erreur
        If FICH_OuvrirFichier(nomfich2, FICH_ECRITURE, fd_resume) = P_ERREUR Then GoTo lab_erreur
        sql = "SELECT UM_UNum, UM_TypeMvt, U_SPM, U_Nom, U_Prenom" _
            & " FROM UtilMouvement, Utilisateur" _
            & " WHERE U_kb_actif=True AND UM_DateEnvoi IS NULL AND UM_UNum=U_Num" _
            & " GROUP BY UM_UNum, UM_TypeMvt, U_SPM, U_Nom, U_Prenom"
        If Odbc_Select(sql, rs_util) = P_ERREUR Then
            GoTo lab_erreur
        End If
        ' Ecriture du nom de l'application
        Print #fd_enclair, "" ' saut de ligne dans le fichier "en clair"
        Print #fd_enclair, "######################################################"
        Print #fd_enclair, vbTab & " APPLICATION :" & vbTab & rs_appli("APP_Code").Value
        Print #fd_enclair, "######################################################"
        Print #fd_enclair, "" ' saut de ligne dans le fichier "en clair"
        ' Parcourir les mouvements effectués sur les utilisateurs (groupés par utilisateur)
        ' **********************************************************************************
        nbtot_util = rs_util.RowCount
        Menu.PgBarUtil.Max = nbtot_util
        nb_util = 0
        While Not rs_util.EOF
            nb_util = nb_util + 1
            Menu.PgBarUtil.Value = nb_util
            ' Récupérer la liste des postes qui nous intéressent
            ' listePostesValides = "Px;Py;" (les postes à tester)
            listePostesValides = GetPostesValides(rs_util("UM_UNum").Value)
            ' Eviter les U_SPM vides
            If rs_util("U_SPM").Value <> "" Then
                ecrire = False
                ' l'utilisateur est-il concerné ? (passer la liste et non le u_spm)
                If UtilisateurConcerne(rs_appli("APP_Profil_Conc").Value, listePostesValides) Then
                    ecrire = True
                ElseIf Not UtilisateurConcerne(rs_appli("APP_Profil_NonConc").Value, listePostesValides) Then
                    ' l'utilisateur fait partie des personnes non renseignées (poser la question)
                    If Not bPoserQuestion Then
                        GoTo Lab_NextUtil
                    End If
                    reponse = MsgBox("Le poste de " & rs_util("U_Nom").Value & " " & rs_util("U_Prenom").Value _
                                & " (" & P_get_lib_srv_poste(P_get_num_srv_poste(rs_util("U_SPM").Value, P_POSTE), P_POSTE) _
                                & " - " & P_get_lib_srv_poste(P_get_num_srv_poste(rs_util("U_SPM").Value, P_SERVICE), P_SERVICE) & ")" _
                                & vbCrLf & vbCrLf & " n'a pas été renseigné comme étant concerné par l'application: " & rs_appli("APP_Nom").Value _
                                & "." & vbCrLf & vbCrLf & "Transmettre quand-même les modifications apportées à cette" _
                                & " personne au responsable" & vbCrLf & " de l'application " & rs_appli("APP_Nom").Value _
                                & " (" & nomResponsable & " " & prenomResponsable & ") ?", _
                                vbQuestion + vbYesNoCancel, "Manque d'informations")
                    ' Tester la réponse
                    If reponse = vbCancel Then
                        'GoTo lab_annuler_tout
                        bPoserQuestion = False
                    ElseIf reponse = vbYes Then
                        ecrire = True
                    End If
                End If
                If ecrire Then
                    ' a-t-on des information supplémentaires à communiquer ?
                    finfo = a_des_infoSuppl(rs_util("UM_UNum").Value, rs_appli("APP_InfoSuppl").Value & "")
                    nbr_ecrit = nbr_ecrit + GenererLesFichiers(rs_util("UM_UNum").Value, _
                                            rs_appli("APP_ZonePrev").Value, _
                                            fd_enclair, fd_resume, rs_appli("APP_Num").Value, finfo)
                End If
            End If
Lab_NextUtil:
            rs_util.MoveNext
        Wend
        rs_util.Close
        Close #fd_enclair
        Close #fd_resume
        If nbr_ecrit > 0 Then
            nomFichierEnClair = chemin & "/enclair_env_" & format(Date, "YYYYMMDD") & "-gen.txt"
            nomFichierResume = chemin & "/resume_env_" & format(Date, "YYYYMMDD") & "-gen.txt"
            Call KF_PutFichier(nomFichierEnClair, nomfich1)
            Call KF_PutFichier(nomFichierResume, nomfich2)
            g_on_a_genere = True
        End If
        Call FICH_EffacerFichier(nomfich1, False)
        Call FICH_EffacerFichier(nomfich2, False)
lab_appli_suivante:
        rs_appli.MoveNext
    Wend
    rs_appli.Close

    Menu.FrmFichiers.Visible = False
    
    CreerLesFichiers = P_OK
    Exit Function

lab_erreur:
    CreerLesFichiers = P_ERREUR
    Exit Function

lab_annuler_tout: ' même code de retour que ERREUR ?
' EFFACER les fichiers non finis !
    rs_util.Close
    rs_appli.Close
    Close #fd_enclair
    Close #fd_resume
    Call FICH_EffacerFichier(nomFichierResume, False)
    Call FICH_EffacerFichier(nomFichierEnClair, False)
    CreerLesFichiers = P_ERREUR

End Function

Private Function EnvoyerLesMessages(ByRef r_nbr_messages_non_envoyes As Integer, _
                                      ByRef r_bilan As String) As Integer
' **********************************************************************************
' Envoyer les e-mails pour les fichiers [nouvelement créés] et/ou
' [anciennement créés mais non envoyés] + un lien vers kb_visualiste.php?CODE_APPLI
' Renseigner le nombre de messages non envoyés et alimenter le bilan des opérations.
' **********************************************************************************
    Dim sql As String, message As String, liste_fichiers() As String, _
        str As String, chemin As String, nom_fich_avant As String, nom_fich_apres As String
    Dim nbr_fichiers As Long, I As Long, numresp As Long
    Dim rs As rdoResultset

    sql = "SELECT APP_Code, APP_lstResp, APP_Nom FROM Application" _
        & " WHERE APP_Num<>" & p_appli_kalidoc & " AND APP_Num<>" & p_appli_kalibottin
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    ' Parcourrir toutes les applications existantes (sauf KaliDoc et KaliBottin)
    While Not rs.EOF
        ' Parcourir le repertoire et chercher des fichiers ayant un sufixe "-gen.txt"
        chemin = p_CheminKW & "/kalibottin/mouvements/" & rs("APP_Code").Value
        nbr_fichiers = P_lister_fichiers(chemin, "*_env_*-gen.txt", "LISTE", liste_fichiers)
        If nbr_fichiers > 0 Then
            numresp = Mid$(STR_GetChamp(rs("app_lstresp").Value, ";", 0), 2)
            ' ------- l'envoi du mail --------------
            message = "Les caractéristiques de certaines personnes concernées par" _
                    & " l'application '" & rs("APP_Nom").Value & "' ont été modifiées." & vbCrLf & vbCrLf _
                    & "Cliquez sur le lien suivant pour accéder à la liste : " & vbCrLf & vbCrLf _
                    & p_CheminPHP & "/pident.php" _
                    & "?in=modifkb" & vbCrLf
            Call P_EnvoyerMessage(numresp, "", _
                                "Modifications apportées dans l'annuaire KaliBottin", _
                                message)
            For I = 0 To nbr_fichiers - 1
                ' renomme le fichier en omettant les "-gen"
                nom_fich_avant = liste_fichiers(I)
                nom_fich_apres = STR_GetChamp(nom_fich_avant, "-gen", 0) & STR_GetChamp(nom_fich_avant, "-gen", 1)
                If KF_RenommerFichier(chemin & "/" & nom_fich_avant, chemin & "/" & nom_fich_apres) = P_ERREUR Then
                    GoTo lab_erreur
                End If
            Next I
            g_on_a_genere = True
        End If
        rs.MoveNext
    Wend
    rs.Close

    EnvoyerLesMessages = P_OK
    Exit Function

lab_non_envoyes:
    EnvoyerLesMessages = P_NON
    Exit Function

lab_erreur:
    EnvoyerLesMessages = P_ERREUR

End Function

Public Function EnvoyerLesModifications() As Integer
' ********************************************************************************************
' Le point d'entrée pour créer et/ou envoyer les fichiers après la modifications des pesonnes
' Transmettre les modifications aux responsables des autres applications que KaliDoc:
' 1° Créer les fichiers "en clair" et "résumé" pour chaque application concernée.
' 2° Envoyer les mails aux responsables des applications + lien kb_visualiste.php?CODE_APPLI.
' 3° MAJ UtilMouvement.UM_DateEnvoi (avec la date de génération des fichiers).
' 4° Ajout des informations supplémentaires
' ********************************************************************************************
    Dim sql As String, bilan As String
    Dim nbr_messages_non_envoyes As Integer
    Dim lng As Long

    g_on_a_genere = False
' ************************************************ 1° *****************************************************
    If P_y_a_des_mouvements_a_envoyer = P_FICHIER_A_CREER Then
        If CreerLesFichiers() = P_ERREUR Then
            GoTo lab_erreur
        End If
' ************************************************ 3° *****************************************************
        ' maj UtilMouvement.UM_DateEnvoi
        If Odbc_UpdateP("UtilMouvement", "UM_Num", _
                       "WHERE UM_DateEnvoi IS NULL", lng, _
                       "UM_DateEnvoi", Date) = P_ERREUR Then
            GoTo lab_erreur
        End If
        P_y_a_des_mouvements_a_envoyer = P_MAILS_A_ENVOYER ' afin de REVENIR sur l'envoi des mails uniquement
    ' Else ' P_MOUVEMENTS_A_ENVOYER = P_MAILS_A_ENVOYER
        ' passer directement à l'envoi des emails
    End If
    Menu.FrmFichiers.Visible = False
' ************************************************ 2° *****************************************************
    ' on procèdera à envoyer des messages de toutes façons
    If EnvoyerLesMessages(nbr_messages_non_envoyes, bilan) = P_ERREUR Then
        'GoTo lab_erreur
    End If
' ************************************** MESSAGE DE CONFIRMATION ******************************************
    ' Le bilan des opérations
    If nbr_messages_non_envoyes = 0 Then
        Call MsgBox("Toutes les opérations se sont déroulées avec succès !", _
                    vbInformation + vbOKOnly, "Bilan des opérations")
    ElseIf nbr_messages_non_envoyes > 0 Then
        If nbr_messages_non_envoyes = 1 Then
            bilan = "Il reste un fichier non envoyé:" & bilan
        Else
            bilan = "Il reste " & nbr_messages_non_envoyes & " fichiers non envoyés:" _
                    & vbCrLf & bilan
        End If
        Call MsgBox(bilan, vbInformation + vbOKOnly, "Bilan des opérations")
        GoTo lab_erreur
    End If
' *********************************************************************************************************
    EnvoyerLesModifications = P_OK
    Exit Function

lab_erreur:
    ' Cacher l'indicateur d'envoi des modifs si on a généré au moins un fichier
    EnvoyerLesModifications = IIf(g_on_a_genere, P_OK, P_ERREUR)

End Function

Private Function GenererLesFichiers(ByVal v_um_unum As Long, ByVal v_app_zoneprev As String, _
                                    ByRef r_fd_enclair As Integer, ByRef r_fd_resume As Integer, _
                                    ByVal v_app_num As Long, ByVal v_infoSuppl As Boolean) As Integer
' *******************************************************************************************
' Ici on n'écrit que les lignes du fichier "resumé", pour ceux du fichier "en clair",
' on fait appel à la fonction remplir_lignes_enclair()
' *******************************************************************************************
    Dim sql As String, codeModifie As String, nom As String, prenom As String, _
        matricule As String, str As String, ligneModif As String
    Dim autre_utilisateur As Boolean
    Dim I As Integer, nbr_ecrit As Integer
    Dim rs As rdoResultset

    sql = "SELECT UM_UNum, UM_APPNum, UM_Date, UM_TypeMvt, UM_Commentaire FROM UtilMouvement" _
        & " WHERE UM_UNum=" & v_um_unum & " AND UM_DateEnvoi IS NULL"
    If Odbc_Select(sql, rs) = P_ERREUR Then
        Exit Function
    End If
    ' Parcourir les mouvements concernant cet utilisateur
    nbr_ecrit = 0
    autre_utilisateur = True
    ligneModif = ""
    While Not rs.EOF
        ' remplir la ligne "resumé"
        If rs("UM_TypeMvt").Value = "M" Then ' ************** MODIFICATION  **************
            codeModifie = STR_GetChamp(rs("UM_Commentaire").Value, "=", 0)
            If codeModifie <> "NOM" And codeModifie <> "PRENOM" And codeModifie <> "POSTE" Then
                ' tester si la modif de la COORDONNEÉ (principale) est à prendre en compte par l'appli
                ' --------------------------- COORDONNEÉ + MATRICULE -----------------------------------------------
                If PrendreEnCompteModif(codeModifie, v_app_zoneprev) Then
                    nbr_ecrit = nbr_ecrit + 1
                    ligneModif = IIf(ligneModif = "", _
                                      GetDescription("M", v_um_unum, rs("UM_Commentaire").Value), _
                                      ligneModif & "|" & GetDescription("M", v_um_unum, rs("UM_Commentaire").Value))
' **********************
                    Call RemplirLignesEnClair(r_fd_enclair, rs, codeModifie, v_app_zoneprev, _
                                              v_um_unum, autre_utilisateur, v_app_num, v_infoSuppl)
                    autre_utilisateur = False
' **********************
                End If
            Else ' --------------------------- NOM + PRENOM + POSTE ------------------------------------------------
                nbr_ecrit = nbr_ecrit + 1
                ligneModif = IIf(ligneModif = "", _
                                  GetDescription("M", v_um_unum, rs("UM_Commentaire").Value), _
                                  ligneModif & "|" & GetDescription("M", v_um_unum, rs("UM_Commentaire").Value))
' **********************
                Call RemplirLignesEnClair(r_fd_enclair, rs, codeModifie, v_app_zoneprev, _
                                          v_um_unum, autre_utilisateur, v_app_num, v_infoSuppl)
                autre_utilisateur = False
' **********************
            End If
        Else
            nbr_ecrit = nbr_ecrit + 1
            If rs("UM_TypeMvt").Value = "C" Then ' ************ CREATION *****************
                Print #r_fd_resume, v_um_unum & "#" & "C" & "#" & GetDescription("C", v_um_unum, "")
            ElseIf rs("UM_TypeMvt").Value = "I" Then ' ************ INACTIVE *****************
                Print #r_fd_resume, v_um_unum & "#" & "I" & "#"
            ElseIf rs("UM_TypeMvt").Value = "A" Then ' ************* ACTIVE ******************
                Print #r_fd_resume, v_um_unum & "#" & "A" & "#"
            End If
' **********************
            Call RemplirLignesEnClair(r_fd_enclair, rs, codeModifie, v_app_zoneprev, _
                                      v_um_unum, autre_utilisateur, v_app_num, v_infoSuppl)
            autre_utilisateur = False
' **********************
        End If
        ' remplir la ligne "en_clair"
        'Call RemplirLignesEnClair(r_fd_enclair, rs, codeModifie, v_app_zoneprev, v_um_unum, autre_utilisateur)
        'autre_utilisateur = False
        rs.MoveNext
    Wend
    rs.Close
    ' regrouper les modifications dans une seule ligne par personne dans le fichier "resumé"
    If ligneModif <> "" Then Print #r_fd_resume, v_um_unum & "#" & "M" & "#" & ligneModif

    GenererLesFichiers = nbr_ecrit

End Function

Private Function GetDescription(ByVal v_mode As String, ByVal v_u_num As Long, _
                                 ByVal v_commentaire As String) As String
' **********************************************************************************
' Formater le rs("UM_Commentaire").Value en une description pour le fichier "resumé"
' **********************************************************************************
    Dim nom As String, prenom As String, matricule As String
    Dim spm As String, code_modif As String, str As String
    Dim nbr As Integer, I As Integer

    If v_mode = "C" Then
        ' recherche les postes occupés par cette personne
        If Odbc_RecupVal("SELECT U_Nom, U_Prenom, U_Matricule, U_SPM FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & v_u_num, _
                         nom, prenom, matricule, spm) = P_ERREUR Then
            Exit Function
        End If
        nbr = STR_GetNbchamp(spm, "|")
        For I = 0 To nbr - 1
            str = str & P_get_num_srv_poste(STR_GetChamp(spm, "|", I), P_POSTE) & ";"
        Next I
        GetDescription = "NOM=" & nom & "|PRENOM=" & prenom & "|MATRICULE=" & matricule _
                        & "|POSTE=" & str
    Else ' v_mode = "M"
        code_modif = STR_GetChamp(v_commentaire, "=", 0)
        str = STR_GetChamp(v_commentaire, code_modif & "=", 1)
        If code_modif = "POSTE" Then
            GetDescription = "POSTE=" & STR_GetChamp(str, "|", 0) & "$" & STR_GetChamp(str, "|", 1)
        Else
            GetDescription = code_modif & "=" & str
        End If
    End If

End Function

Private Function GetPostesValides(ByVal v_u_num As Long) As String
' ****************************************************************
' Retourner la liste des postes qui sonts à tester s'ils sont
' à prendre en compte par l'applicatopn
' ****************************************************************

    Dim spm As String, spm_en_cours As String, lesAjouts As String
    Dim lesSuppressions As String, ajoutEnCours As String
    Dim debut As Boolean
    Dim nbr As Integer, nbr_spm_en_cours As Integer, I As Integer, _
        j As Integer, dernier_j As Integer, nbrAjouts As Integer, _
        nbrSuppressions As Integer
    Dim rs As rdoResultset

    debut = True
    lesAjouts = ""
    lesSuppressions = ""
    If Odbc_SelectV("SELECT UM_Commentaire FROM UtilMouvement WHERE UM_UNum=" & v_u_num _
                 & " AND UM_DateEnvoi IS NULL AND UM_TypeMvt='M' AND UM_Commentaire LIKE 'POSTE=%'" _
                 & " ORDER BY UM_Num", rs) = P_ERREUR Then
        Exit Function
    End If
    ' Pas de mouvements sur le poste => retourner la liste des postes de l'utilisateur
    If rs.EOF Then
        If Odbc_RecupVal("SELECT U_SPM FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & v_u_num, spm) = P_ERREUR Then
            Exit Function
        End If
        nbr = STR_GetNbchamp(spm, "|")
        For I = 0 To nbr - 1
            spm_en_cours = STR_GetChamp(spm, "|", I)
            nbr_spm_en_cours = STR_GetNbchamp(spm_en_cours, ";")
            GetPostesValides = GetPostesValides & STR_GetChamp(spm_en_cours, ";", nbr_spm_en_cours - 1) & ";"
        Next I
        Exit Function
    End If
    ' Parcourir les mouvements sur le poste de cette personne
    While Not rs.EOF
        If debut Then ' Les postes initiaux
            GetPostesValides = STR_GetChamp(STR_GetChamp(rs("UM_Commentaire").Value, "=", 1), "|", 0)
            debut = False
        End If
        ' Le cumul des opérations: ajouts et suppressions
        lesAjouts = lesAjouts & STR_GetChamp(STR_GetChamp(rs("UM_Commentaire").Value, "A:", 1), "S:", 0)
        lesSuppressions = lesSuppressions & STR_GetChamp(rs("UM_Commentaire").Value, "S:", 1)
        rs.MoveNext
    Wend
    rs.Close
    ' Générer la liste des postes intéressants
    dernier_j = 0
    nbrAjouts = STR_GetNbchamp(lesAjouts, ";")
    For I = 0 To nbrAjouts - 1
        ajoutEnCours = STR_GetChamp(lesAjouts, ";", I)
        nbrSuppressions = STR_GetNbchamp(lesSuppressions, ";")
        For j = dernier_j To nbrSuppressions - 1
            If ajoutEnCours = STR_GetChamp(lesSuppressions, ";", j) Then
                 ' Garder l'indice du dernier poste supprimé
                 ' Sachant que les postes
                dernier_j = j + 1
                GoTo lab_ajout_suivant
            End If
        Next j
        ' si on est là, c'est que le poste n'a pas été supprimé
        GetPostesValides = GetPostesValides & ajoutEnCours & ";"
lab_ajout_suivant:
    Next I

End Function

Private Function infoSuppl_a_envoyer(ByVal v_app_infoSuppl As String) As Boolean
' a-t-on des informations supplémentaires de personnes nouvellement crées
' à transmettre ?
    Dim sql As String
    Dim rs As rdoResultset
    Dim nbr_personnes As Long, tis_num As Long
    Dim I As Integer, nbr_tis As Integer

    If Len(v_app_infoSuppl) = 0 Then GoTo lab_erreur

    nbr_tis = STR_GetNbchamp(v_app_infoSuppl, ";")
    For I = 0 To nbr_tis
        tis_num = Mid$(STR_GetChamp(v_app_infoSuppl, ";", I), 2)
        sql = "SELECT COUNT(*) FROM Utilisateur, InfoSupplEntite, UtilMouvement" & _
              " WHERE U_kb_actif=True AND ISE_TypeNum=U_Num AND UM_UNum=U_Num AND ISE_Type='U'" & _
              " AND UM_TypeMvt='C' AND UM_DateEnvoi IS NULL" & _
              " AND ISE_TisNum=" & tis_num
        If Odbc_Count(sql, nbr_personnes) = P_ERREUR Then
            GoTo lab_erreur
        End If
        If nbr_personnes > 0 Then
            infoSuppl_a_envoyer = True
            Exit Function
        End If
    Next I

lab_erreur:
    infoSuppl_a_envoyer = False

End Function


Private Function PrendreEnCompteModif(ByVal v_zu_code As String, ByVal v_app_zoneprev As String) As Boolean
' ***********************************************************************************************
' Déterminer si on previent l'application en cours des modifications opérées sur cette coordonnée
' ***********************************************************************************************
    Dim I As Integer, nbr As Integer
    Dim zu_num As Long

    ' récupérer le ZU_Num pour cette modification
    If Odbc_RecupVal("SELECT ZU_Num FROM ZoneUtil WHERE ZU_Code='" & v_zu_code & "'", _
                     zu_num) = P_ERREUR Then
        GoTo lab_erreur
    End If
    nbr = STR_GetNbchamp(v_app_zoneprev, ";")
    For I = 0 To nbr - 1
        If zu_num = Mid$(STR_GetChamp(v_app_zoneprev, ";", I), 2) Then
            GoTo lab_ok
        End If
    Next I

lab_erreur:
    PrendreEnCompteModif = False
    Exit Function

lab_ok:
    PrendreEnCompteModif = True

End Function

Private Sub RemplirLignesEnClair(ByRef r_fd_enclair As Integer, ByVal v_rs As rdoResultset, _
                     ByVal v_code_modifie As String, v_app_zoneprev, ByVal v_um_unum As Long, _
                     ByVal v_autre_utilisateur As Boolean, ByVal v_app_num As Long, ByVal v_infoSuppl As Boolean)
' **********************************************************************************
' Ecrire la ligne en clair avec les infos adéquates:
'   date - application -   CREATION/ACTIVE/INACTIVE
'   date - application - MODIFICATION - valeur_avant - valeur_après
' Vérifier si les coordonnées sont à prendre en compte par cette application ou non
' **********************************************************************************
    Dim la_ligne As String, str As String, ajout As String, suppression As String, _
        nom As String, prenom As String, matricule As String, ligne_debut As String, _
        sql As String, app_infosuppl As String, kb_tisLibelle As String, _
        tis_num As String, ise_valeur As String
    Dim ecrire As Boolean
    Dim nbr As Integer, I As Integer, nbr_infosuppl As Integer
    Dim num_poste As Long, nbr_u_tis As Long
    Dim s_sp As Variant
    Dim rs As rdoResultset

    ' Écrire une seule fois la ligne identifiant la personne
    If v_autre_utilisateur Then
        If Odbc_RecupVal("SELECT U_Nom, U_Prenom, U_Matricule FROM Utilisateur" & _
                         " WHERE U_kb_actif=True AND U_Num=" & v_um_unum, _
                         nom, prenom, matricule) = P_ERREUR Then
            Exit Sub
        End If
        ' générer une ligne d'asterix de la même longueur que cette chaîne
        matricule = IIf(matricule = "", "non renseigné", matricule)
        For I = 0 To Len("Mouvements concernant " & nom & " " & prenom & " [ayant le MATRICULE: " & matricule & "]:") + 1
            str = str & "*"
        Next I
        ligne_debut = vbCrLf & str & vbCrLf & " Mouvements concernant " & nom & " " & prenom _
                      & " [ayant le MATRICULE: " & matricule & "]:" & vbCrLf & str & vbCrLf
    End If
    
    ecrire = True
    la_ligne = "Le " & format(v_rs("UM_Date").Value, "dddd d mmm yyyy") _
             & ", l'application [ " & P_get_nom_appli(v_rs("UM_APPNum").Value) & " ] a"
    Select Case v_rs("UM_TypeMvt").Value
        Case "C" ' ***************************** CRÉATION *************************************
            If v_autre_utilisateur And ecrire Then Print #r_fd_enclair, ligne_debut: ecrire = False
            Call Odbc_RecupVal("select u_spm from utilisateur where u_kb_actif=True and u_num=" & v_um_unum, s_sp)
            s_sp = STR_GetChamp(s_sp, "|", STR_GetNbchamp(s_sp, "|") - 1)
            num_poste = Mid$(STR_GetChamp(s_sp, ";", STR_GetNbchamp(s_sp, ";") - 1), 2)
            la_ligne = la_ligne & " créé le compte de cet utilisateur." & vbCrLf _
                       & vbTab & P_get_lib_srv_poste(num_poste, P_POSTE) _
                       & " - " & STR_GetChamp(P_get_service_du_poste(num_poste), "=", 1)
            Print #r_fd_enclair, la_ligne
        Case "A" ' ****************************** ACTIVE **************************************
            If v_autre_utilisateur And ecrire Then Print #r_fd_enclair, ligne_debut: ecrire = False
            la_ligne = la_ligne & " activé le compte de cette personne."
            Print #r_fd_enclair, la_ligne
        Case "I" ' ***************************** INACTIVE *************************************
            If v_autre_utilisateur And ecrire Then Print #r_fd_enclair, ligne_debut: ecrire = False
            la_ligne = la_ligne & " rendu le compte de cette personne inactif."
            Print #r_fd_enclair, la_ligne
        Case "M" ' *************************** MODIFICATION ***********************************
            If v_code_modifie = "NOM" Or v_code_modifie = "PRENOM" Or v_code_modifie = "POSTE" Then
            ' ---------------------------------------- NOM - PRENOM - POSTE ------------------------
                If v_autre_utilisateur And ecrire Then Print #r_fd_enclair, ligne_debut: ecrire = False
                la_ligne = la_ligne & " modifié le [ " & v_code_modifie & " ] de cette personne:"
                Print #r_fd_enclair, la_ligne
            ElseIf v_code_modifie = "MATRICULE" Then ' MATRICULE -----------------------------------
'                If PrendreEnCompteModif(v_code_modifie, v_app_zoneprev) Then
                    If v_autre_utilisateur And ecrire Then Print #r_fd_enclair, ligne_debut: ecrire = False
                    la_ligne = la_ligne & " modifié le [ " & v_code_modifie & " ] de cette personne:"
                    Print #r_fd_enclair, la_ligne
                    ' --------------- valeur avant
                    str = STR_GetChamp(STR_GetChamp(v_rs("UM_Commentaire").Value, "=", 1), ";", 0)
                    la_ligne = IIf(str = "", "   * AVANT: " & "[inexistant ou vide]", "   * AVANT: " & str)
                    Print #r_fd_enclair, la_ligne
                    ' --------------- valeur après
                    str = STR_GetChamp(STR_GetChamp(v_rs("UM_Commentaire").Value, "=", 1), ";", 1)
                    la_ligne = IIf(str = "", "   * APRÈS: " & "[vide ou supprimé]", "   * APRÈS: " & str)
                    Print #r_fd_enclair, la_ligne
'                End If
            Else ' ------------------------ changement(s) dans les COORDONNÉES ---------------------
                ' les coordonnées sont-ils à prendre en compte ?
'                If PrendreEnCompteModif(v_code_modifie, v_app_zoneprev) Then
                    If v_autre_utilisateur And ecrire Then Print #r_fd_enclair, ligne_debut: ecrire = False
                    la_ligne = la_ligne & " modifié le [ " & v_code_modifie & " ] de cette personne:"
                    Print #r_fd_enclair, la_ligne
                    ' --------------- valeur avant
                    str = STR_GetChamp(STR_GetChamp(v_rs("UM_Commentaire").Value, "=", 1), ";", 0)
                    la_ligne = IIf(str = "", "   * AVANT: " & "[inexistant ou vide]", "   * AVANT: " & str)
                    Print #r_fd_enclair, la_ligne
                    ' --------------- valeur après
                    str = STR_GetChamp(STR_GetChamp(v_rs("UM_Commentaire").Value, "=", 1), ";", 1)
                    la_ligne = IIf(str = "", "   * APRÈS: " & "[vide ou supprimé]", "   * APRÈS: " & str)
                    Print #r_fd_enclair, la_ligne
                    ' --------------- une précision
                    la_ligne = "(ces modifications concernent uniquement les coordonnées principales.)"
                    Print #r_fd_enclair, la_ligne
'                Else ' on n'écrit rien
'                    GoTo lab_rien_a_ecrire
'                End If
            End If
            If v_code_modifie = "POSTE" Then '      changement(s) dans le POSTE
                ' les postes avant les modifications
                If v_autre_utilisateur And ecrire Then Print #r_fd_enclair, ligne_debut: ecrire = False
                la_ligne = "   * LE(S) POSTE(S) OCCUPÉ(S) AVANT: "
                Print #r_fd_enclair, la_ligne
                la_ligne = "     ------------------------------ "
                Print #r_fd_enclair, la_ligne
                str = STR_GetChamp(STR_GetChamp(v_rs("UM_Commentaire").Value, "|", 0), "POSTE=", 1)
                nbr = STR_GetNbchamp(str, ";")
                For I = 0 To nbr - 1
                    num_poste = Mid$(STR_GetChamp(str, ";", I), 2)
                    la_ligne = I + 1 & ")" & vbTab & P_get_lib_srv_poste(num_poste, P_POSTE) _
                             & " (Service: " & STR_GetChamp(P_get_service_du_poste(num_poste), "=", 1) & ")"
                    Print #r_fd_enclair, la_ligne
                Next I
                Print #r_fd_enclair, "" ' saut de ligne
                ' les modifications effectuées sur les postes
                la_ligne = "   * LE(S) MOUVEMENT(S) EFFECTUÉ(S) SUR LES POSTES: "
                Print #r_fd_enclair, la_ligne
                la_ligne = "     ---------------------------------------------  "
                Print #r_fd_enclair, la_ligne
                str = STR_GetChamp(v_rs("UM_Commentaire").Value, "|", 1)
                ' LES AJOUTS
                ajout = STR_GetChamp(STR_GetChamp(str, "A:", 1), "S:", 0)
                nbr = STR_GetNbchamp(ajout, ";")
                For I = 0 To nbr - 1
                    num_poste = Mid$(STR_GetChamp(ajout, ";", I), 2)
                    la_ligne = " + Ajout du poste: " & P_get_lib_srv_poste(num_poste, P_POSTE) _
                             & " (Service: " & STR_GetChamp(P_get_service_du_poste(num_poste), "=", 1) & ")"
                    Print #r_fd_enclair, la_ligne
                Next I
                ' LES SUPPRESSIONS
                suppression = STR_GetChamp(str, "S:", 1)
                nbr = STR_GetNbchamp(suppression, ";")
                For I = 0 To nbr - 1
                    num_poste = Mid$(STR_GetChamp(suppression, ";", I), 2)
                    la_ligne = " - Suppression du poste: " & P_get_lib_srv_poste(num_poste, P_POSTE) _
                             & " (Service: " & STR_GetChamp(P_get_service_du_poste(num_poste), "=", 1) & ")"
                    Print #r_fd_enclair, la_ligne
                Next I
            ElseIf v_code_modifie = "NOM" Or v_code_modifie = "PRENOM" Then ' Or v_code_modifie = "MATRICULE" Then
                '                     changement(s) dans le NOM, PRENOM                           (ou le MATRICULE)
                If v_autre_utilisateur And ecrire Then Print #r_fd_enclair, ligne_debut: ecrire = False
                ' --------------- valeur avant
                str = STR_GetChamp(STR_GetChamp(v_rs("UM_Commentaire").Value, "=", 1), ";", 0)
                la_ligne = IIf(str = "", "   * AVANT: " & "[inexistant ou vide]", "   * AVANT: " & str)
                Print #r_fd_enclair, la_ligne
                ' --------------- valeur après
                str = STR_GetChamp(STR_GetChamp(v_rs("UM_Commentaire").Value, "=", 1), ";", 1)
                la_ligne = IIf(str = "", "   * APRÈS: " & "[vide ou supprimé]", "   * APRÈS: " & str)
                Print #r_fd_enclair, la_ligne
            Else ' Or v_code_modifie = "MATRICULE"
                '                     changement(s) dans le MATRICULE
                ' c'est déjà fait plus haut, ces lignes sont à supprimer :)
            End If
    End Select
    ' Uniquement pour les info supplementaire
    If v_infoSuppl Then
        If v_autre_utilisateur And ecrire Then Print #r_fd_enclair, ligne_debut
        la_ligne = ""
        ' chercher les infos suppl à renseigner
        sql = "SELECT APP_InfoSuppl FROM Application WHERE APP_Num=" & v_app_num
        If Odbc_RecupVal(sql, app_infosuppl) = P_ERREUR Then
            Exit Sub
        End If
        ' parcourir les info suppl de cette appli
        nbr_infosuppl = STR_GetNbchamp(app_infosuppl, ";")
        For I = 0 To nbr_infosuppl - 1
            tis_num = Mid$(STR_GetChamp(app_infosuppl, ";", I), 2)
            ' chercher si on a cette infoSuppl renseignée pour cette personne
            sql = "SELECT COUNT(*) FROM InfoSupplEntite" & _
                  " WHERE ISE_TisNum=" & tis_num & " AND ISE_Type='U' AND ISE_TypeNum=" & v_um_unum
            If Odbc_Count(sql, nbr_u_tis) = P_ERREUR Then
                Exit Sub
            End If
            If nbr_u_tis > 0 Then ' on a trouvé une infosuppl
                ' recuperer les libelle+valeur
                sql = "SELECT KB_TisLibelle, ISE_Valeur FROM KB_TypeInfoSuppl, InfoSupplEntite" & _
                      " WHERE KB_TisNum=ISE_TisNum AND KB_TisNum=" & tis_num & _
                      " AND ISE_Type='U' AND ISE_TypeNum=" & v_um_unum
                If Odbc_RecupVal(sql, kb_tisLibelle, ise_valeur) = P_ERREUR Then
                    Exit Sub
                End If
                ' ecrire la ligne
                la_ligne = la_ligne & vbTab & UCase(kb_tisLibelle) & " = " & ise_valeur & vbCrLf
            End If
        Next I
        Print #r_fd_enclair, la_ligne
    End If

'lab_rien_a_ecrire: ' lorsque la modif est sur des coordonnées qui ne sont pas à renseigner

End Sub

Private Function UtilisateurConcerne(ByVal v_profil As String, ByVal v_ListePostes As String) As Boolean
' ****************************************************************************
' Déterminer si la liste v_ListePostes est prise en compte par le v_profil
' par exemple: v_ListePoste = "P20;P117;P68;" & v_profil="T" ou "" ou autre...
' OU
' le compte de la personne vient d'être créé donc envoie d'infosuppl
' ****************************************************************************
    Dim profil_en_cours As String, str As String, sql As String
    Dim nbr_profil As Integer, nbr_postes As Integer, I As Integer, j As Integer, nbr As Integer
    Dim num_poste_en_cours As Long, num_srv_en_cours As Long, num_srv_pere As Long
    Dim s_srv_pere As String, snum As String
    Dim rs As rdoResultset

    If v_profil = "T" Then GoTo lab_succes ' tout le monde est concerné
    If v_profil = "" Then GoTo lab_erreur ' personne n'est concerné
    nbr_profil = STR_GetNbchamp(v_profil, "|")
    nbr_postes = STR_GetNbchamp(v_ListePostes, ";")
    ' Parcourrir les postes de l'utilisateur
    For I = 0 To nbr_postes - 1
        num_poste_en_cours = Mid$(STR_GetChamp(v_ListePostes, ";", I), 2)
        For j = 0 To nbr_profil - 1
            profil_en_cours = STR_GetChamp(v_profil, "|", j)
            nbr = STR_GetNbchamp(profil_en_cours, ";")
            str = STR_GetChamp(profil_en_cours, ";", nbr - 1)
            If Mid$(str, 1, 1) = "P" Then
                If Mid$(str, 2) = num_poste_en_cours Then ' le poste est concerné
                    GoTo lab_succes
                End If
            ElseIf Mid$(str, 1, 1) = "S" Then
                num_srv_en_cours = Mid$(str, 2)
' *************************************************************************************************************
                ' Est-ce que le poste concerne ce service ?
                s_srv_pere = STR_GetChamp(P_get_service_du_poste(num_poste_en_cours), "=", 0)
                If s_srv_pere <> "" Then
                    num_srv_pere = STR_GetChamp(P_get_service_du_poste(num_poste_en_cours), "=", 0)
                    If num_srv_pere = num_srv_en_cours Then
                        GoTo lab_succes ' le poste est concerné
                    Else ' rechercher dans les service parents
                        Do Until num_srv_pere = 0
                            snum = P_get_num_srv_pere(num_srv_pere)
                            If snum <> "" Then
                                num_srv_pere = STR_GetChamp(snum, "=", 0)
                                If num_srv_pere = num_srv_en_cours Then
                                    GoTo lab_succes
                                End If
                            Else
                                num_srv_pere = 0
                            End If
                        Loop
                    End If
                End If
' *************************************************************************************************************
            End If
        Next j
    Next I

lab_erreur:
    UtilisateurConcerne = False
    Exit Function

lab_succes:
    UtilisateurConcerne = True

End Function

