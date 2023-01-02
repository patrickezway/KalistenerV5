Attribute VB_Name = "MKaliBottin"
Option Explicit

' Cycles documentation
Public p_cycle_relecture As Long
Public p_cycle_verifliens As Long
Public p_cycle_consultable As Long
Public p_cycle_diffusion As Long

Public Type SCYCLEDOCS
    etape As String
    acteur As String
    action As String
    ordre_si_refus As Long
    informer_si_refus As Boolean
    modifiable As Boolean
End Type
Public p_scycledocs() As SCYCLEDOCS

Public Type PGB_TIS_POS
    prmgenb_tis_num As Long
    prmgenb_tis_pos As Integer
    prmgenb_tis_long As Integer
    prmgenb_tis_lien As Long    ' si lié à une vraie coordonnée
    prmgenb_tis_lib As String
End Type

Public LISTE_TIS_POS() As PGB_TIS_POS

Public p_nbr_lstInfoSuppl As Integer

Public p_multilabo As Boolean
Public p_chemin_appli As String
Public p_nomini As String

Public g_fd1 As Integer
Public g_fd2 As Integer
Public g_fd3 As Integer
Public g_nomfichHTML As String

Public p_CheminPHP As String
Public p_CheminKW As String
Public p_sversconf As String

Public p_ouvrir_log As Boolean
Public p_traitement_background As Boolean
Public p_traitement_préimport_seul As Boolean
Public p_traitement_background_semiauto As Boolean
Public p_mess_fait_background As String
Public p_background_synchro_auto As Boolean
Public p_corps_background As String
Public p_mess_pasfait_background As String

Public bFaireRemove As Boolean

Public p_type_messagerie1 As Integer
Public p_type_messagerie2 As Integer

Public p_tblu() As Long
Public p_tblu_sel() As Long
Public p_siz_tblu As Long
Public p_siz_tblu_sel As Long

' Le fichier d'importation
Public p_programme_preimport_exe As String
Public p_programme_preimport_lock As String
Public p_programme_preimport_log As String
Public p_nom_fichier_importation_local As String
Public p_nom_fichier_importation As String
Public p_est_sur_serveur As Boolean
Public p_type_fichier As String
Public p_separateur As String
Public p_nom_fichier_ini_kalidoc As String              'RV05042011

' Les differentes extension possibles pour les fichiers d'importation
Public Const P_EXTENSIONS_FICHIERS_IMPORTATION = "*.txt;*.csv"
' Positions des champs dans le fichier d'importation
Public p_pos_matricule As Integer
Public p_pos_titre As Integer
Public p_pos_nom As Integer
Public p_pos_njf As Integer
Public p_pos_prenom As Integer
Public p_pos_civilite As Integer
Public p_pos_code_section As Integer
Public p_pos_lib_section As Integer
Public p_pos_code_emploi As Integer
Public p_pos_lib_emploi As Integer
Public p_format_code As String
Public p_mdp As String
' longueur des champs dans le fichier d'importation
Public p_long_matricule As Integer
Public p_long_titre As Integer
Public p_long_nom As Integer
Public p_long_njf As Integer
Public p_long_prenom As Integer
Public p_long_civilite As Integer
Public p_long_code_section As Integer
Public p_long_lib_section As Integer
Public p_long_code_emploi As Integer
Public p_long_lib_emploi As Integer

Public Type p_tb_car_traités
    ASC_in As String
    CAR_out As String
End Type
Public p_tbl_car_traités() As p_tb_car_traités
Public p_bool_p_tbl_car_traités_chargé As Boolean

Public Const P_POSTE = 1 ' !!! ne pas toucher à ce chiffre
Public Const P_SERVICE = 2 ' !!! ne pas toucher à ce chiffre

' Gestion des mouvements
Public P_y_a_des_mouvements_a_envoyer As Integer ' --------------------------------
Public Const P_FICHIER_A_CREER = 1 ' !!! laisser des nombre entier positifs
Public Const P_MAILS_A_ENVOYER = 2 ' !!! laisser des nombre entier positifs

'**********************************************************************
Private Const MAX_PATH = 260
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const ERROR_NO_MORE_FILES = 18&

Public Const P_PERSONNE_SELECTIONNEE_NON = -186

Public Function P_ChangerCar(ByVal v_str As String, ByRef r_tbcaractere_nontraite()) As String

    Dim s As String, msg As String
    Dim I As Integer, code_car As Integer
    Dim j As Integer
    Dim chemin_serv As String, chemin_loc As String
    Dim fd As Integer
    Dim bOK As Boolean
    
    If Not p_bool_p_tbl_car_traités_chargé Then
        chemin_serv = p_CheminKW & "/includes/Specifique_Site/kaliBottinCar.txt"
        If Not KF_FichierExiste(chemin_serv) Then
            Call Menu.FaitFichierCar("")
        End If
        chemin_loc = p_chemin_appli & "/kaliBottinCar.txt"
        Call FICH_EffacerFichier(chemin_loc, False)
        Call KF_GetFichier(chemin_serv, chemin_loc)
        FICH_OuvrirFichier chemin_loc, FICH_LECTURE, fd
        I = 0
        While Not EOF(fd)
            Line Input #fd, s
            If Mid(s, 1, 1) <> "#" Then
                ReDim Preserve p_tbl_car_traités(I)
                p_tbl_car_traités(I).ASC_in = STR_GetChamp(s, "=", 0)
                p_tbl_car_traités(I).CAR_out = STR_GetChamp(s, "=", 1)
                I = I + 1
            End If
        Wend
        p_bool_p_tbl_car_traités_chargé = True
        Close #fd
    End If
    s = ""
    For I = 1 To Len(v_str)
        code_car = Asc(Mid$(v_str, I, 1))
        If code_car > 127 Then
            bOK = False
            For j = 0 To UBound(p_tbl_car_traités)
                If p_tbl_car_traités(j).ASC_in = code_car Then
                    s = s & p_tbl_car_traités(j).CAR_out
                    bOK = True
                    Exit For
                End If
            Next j
            If Not bOK Then
                s = s & Mid$(v_str, I, 1)
                Call caractere_non_traite(code_car, v_str, r_tbcaractere_nontraite)
            End If
        Else
            s = s & Mid$(v_str, I, 1)
        End If
    Next I
    
    P_ChangerCar = s

End Function

Private Sub caractere_non_traite(ByVal v_num As Integer, _
                                 ByVal v_str As String, _
                                 ByRef r_tbcaractere_nontraite())
    
    Dim iBoucle As Integer, I As Integer
    Dim bChaineTrouvee As Boolean
    
    bChaineTrouvee = False
    iBoucle = 0
    On Error Resume Next
    iBoucle = UBound(r_tbcaractere_nontraite)
    If iBoucle = 0 Then
        ReDim Preserve r_tbcaractere_nontraite(1)
        r_tbcaractere_nontraite(1) = v_num & " (" & Chr(v_num) & ") dans " & v_str
        Exit Sub
    Else
        I = 1
        Do
            If STR_GetChamp(r_tbcaractere_nontraite(I), " ", 0) = v_num Then
                bChaineTrouvee = True
                Exit Do
            End If
            I = I + 1
        Loop Until I > iBoucle
    End If
    
    If Not bChaineTrouvee Then
       ReDim Preserve r_tbcaractere_nontraite(UBound(r_tbcaractere_nontraite) + 1)
       r_tbcaractere_nontraite(UBound(r_tbcaractere_nontraite)) = v_num & " dans " & v_str
    End If
    
End Sub


Private Function charger_prmgen() As Integer

    Dim sql As String
    Dim rs As rdoResultset
    
    p_appli_kalidoc = 0
    p_appli_kalibottin = 0

    sql = "SELECT APP_Num FROM Application WHERE APP_Code='KALI"
    If Odbc_RecupVal(sql & "DOC'", p_appli_kalidoc) = P_ERREUR Then
        charger_prmgen = P_ERREUR
        Exit Function
    End If
    If Odbc_RecupVal(sql & "BOTTIN'", p_appli_kalibottin) = P_ERREUR Then
        charger_prmgen = P_ERREUR
        Exit Function
    End If
    
    sql = "select * from PrmGen_http"
    If Odbc_Select(sql, rs) = P_ERREUR Then
        charger_prmgen = P_ERREUR
        Exit Function
    End If
    p_AdrServeur = rs("PG_Serveur").Value
    p_CheminPHP = rs("PG_CheminPHP").Value
    p_sversconf = rs("PG_sversconf").Value
    
    p_CheminKW = rs("PG_CheminKW").Value
    p_HTTP_CheminDepot = rs("PG_HTTP_CheminDepot").Value
    p_HTTP_MaxParFichier = rs("PG_HTTP_MaxParFichier").Value
    If p_HTTP_MaxParFichier = 0 Then
        p_HTTP_MaxParFichier = 1024000
    End If
    p_HTTP_MaxParPaquet = rs("PG_HTTP_MaxParPaquet").Value
    If p_HTTP_MaxParPaquet = 0 Then
        p_HTTP_MaxParPaquet = 1024
    End If
    rs.Close
    
    p_smtp_adrsrv = ""
    sql = "select pg_adrsmtp, pg_adrwebmaster from prmgen_http"
    If Odbc_RecupVal(sql, p_smtp_adrsrv, p_smtp_webmaster) = P_ERREUR Then
        charger_prmgen = P_ERREUR
        Exit Function
    End If
    
    p_Mode_Auth_UtilAppli = P_ChargerModeAuth()
    
    charger_prmgen = P_OK
    
End Function

Private Function formater_prenom(ByVal v_prenom As String) As String
' ******************************************************************************
' Mettre la 1° lettre (avec/sans séparateur) en majiscule, le reste en miniscule
' ******************************************************************************
    Dim sous_str As String
    Dim nbr As Integer, I As Integer

    nbr = STR_GetNbchamp(v_prenom, " ")
    If nbr > 1 Then
        For I = 0 To nbr - 1
            sous_str = UCase$(Mid$(STR_GetChamp(v_prenom, " ", I), 1, 1)) & LCase$(Mid$(STR_GetChamp(v_prenom, " ", I), 2))
            If I = 0 Then
                formater_prenom = sous_str
            Else
                formater_prenom = formater_prenom & " " & sous_str
            End If
        Next I
        Exit Function
    End If
    nbr = STR_GetNbchamp(v_prenom, "-")
    If nbr > 1 Then
        For I = 0 To nbr - 1
            sous_str = UCase$(Mid$(STR_GetChamp(v_prenom, "-", I), 1, 1)) & LCase$(Mid$(STR_GetChamp(v_prenom, "-", I), 2))
            If I = 0 Then
                formater_prenom = sous_str
            Else
                formater_prenom = formater_prenom & "-" & sous_str
            End If
        Next I
        Exit Function
    End If

    formater_prenom = UCase$(Mid$(v_prenom, 1, 1)) & LCase$(Mid$(v_prenom, 2))

End Function

Public Function P_RenommerFichierImportation() As Integer
' ********************************
' Renomme le fichier d'importation
' ********************************
    Dim sql As String, liste_pos As String, chemin As String, lst_info_suppl As String
    Dim I As Integer, n As Integer, nb As Integer
    Dim num As Long, lnb As Long
    Dim rs As rdoResultset
    Dim frm As Form

    If SYS_GetIni("SERVEUR", "FICHIER_RENOMMER", p_nomini) <> "OUI" Then
        Exit Function
    End If
    
    chemin = ""
    
    ' Récupérer le chemin du fichier d'importation
    sql = "SELECT PGB_Chemin, PGB_fichsurserveur FROM PrmGenB"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        P_RenommerFichierImportation = P_ERREUR
        Exit Function
    End If
    If rs.EOF Then
        rs.Close
        Call MsgBox("Veuillez d'abord indiquer les paramètres du fichier d'importation (Divers - Paramétrage général)", _
                    vbOKOnly + vbInformation, "")
        P_RenommerFichierImportation = P_NON
        Exit Function
    End If
    
    chemin = rs("PGB_Chemin").Value
    p_est_sur_serveur = rs("PGB_fichsurserveur").Value
    rs.Close
    If p_est_sur_serveur Then
        chemin = Replace(chemin, "\", "/")
        If KF_FichierExiste(chemin) Then
            p_nom_fichier_importation = chemin
        ElseIf p_traitement_background Then
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> Le fichier '" & chemin & "' n'existe pas (serveur)"
            If p_traitement_background_semiauto Then
                MsgBox p_mess_fait_background
                End
            End If
        Else
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> Le fichier '" & chemin & "' n'existe pas (serveur)"
        End If
    Else
        If FICH_FichierExiste(chemin) Then
            p_nom_fichier_importation = chemin
        ElseIf p_traitement_background Then
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> Le fichier '" & chemin & "' n'existe pas (local)"
            If p_traitement_background_semiauto Then
                MsgBox p_mess_fait_background
                End
            End If
        Else
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> Le fichier '" & chemin & "' n'existe pas (local)"
        End If
    End If
    
    If p_nom_fichier_importation = "" Then
        Call MsgBox("Fichier '" & chemin & "' non trouvé " & IIf(p_est_sur_serveur, "(serveur)", "(local)"), vbInformation + vbOKOnly, "")
        ' Récupérer le nom du fichier d'importation
        Set frm = Com_ChoixFichier
        p_nom_fichier_importation = Com_ChoixFichier.AppelFrm("Choix du fichier d'importation", _
                                                              "", _
                                                              "c:", _
                                                              P_EXTENSIONS_FICHIERS_IMPORTATION, _
                                                              False)
        Set frm = Nothing
        p_est_sur_serveur = False
    End If
    
    Call KF_RenommerFichier(p_nom_fichier_importation, p_nom_fichier_importation & "_" & format(Date, "YYYYMMDD"))

End Function

Public Function P_InitFichierImportation() As Integer
' ***********************************************************************
' Récupérer le nom du fichier d'importation et le caractère de séparation
' ***********************************************************************
    Dim sql As String, liste_pos As String, chemin As String, lst_info_suppl As String
    Dim I As Integer, n As Integer, nb As Integer
    Dim num As Long, lnb As Long
    Dim rs As rdoResultset
    Dim type_fichier As String
    Dim frm As Form
    Dim s As String
    Dim lib As String

    chemin = ""
    
    ' Récupérer le chemin du fichier d'importation
    sql = "SELECT PGB_Chemin, PGB_FichType, PGB_fichsurserveur FROM PrmGenB"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        P_InitFichierImportation = P_ERREUR
        Exit Function
    End If
    If rs.EOF Then
        rs.Close
        Call MsgBox("Veuillez d'abord indiquer les paramètres du fichier d'importation (Divers - Paramétrage général)", _
                    vbOKOnly + vbInformation, "")
        P_InitFichierImportation = P_NON
        Exit Function
    End If
    
    chemin = rs("PGB_Chemin").Value
    p_est_sur_serveur = rs("PGB_fichsurserveur").Value
    p_type_fichier = rs("PGB_FichType").Value
    rs.Close
    ' Voir si on outrepasse ce chemin
    If p_nom_fichier_importation_local <> "" Then
        If Not KF_FichierExiste(p_nom_fichier_importation_local) Then
            MsgBox p_nom_fichier_importation_local & " n'existe pas (voir le .ini)"
        Else
            p_nom_fichier_importation = p_nom_fichier_importation_local
            p_est_sur_serveur = False
        End If
    End If
    If p_est_sur_serveur Then
        chemin = Replace(chemin, "\", "/")
        If KF_FichierExiste(chemin) Then
            p_nom_fichier_importation = chemin
        ElseIf p_traitement_background Then
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> Le fichier '" & chemin & "' n'existe pas (serveur)"
            If p_traitement_background_semiauto Then
                MsgBox p_mess_fait_background
                End
            Else
                End
            End If
        Else
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> Le fichier '" & chemin & "' n'existe pas (serveur)"
        End If
    Else
        If FICH_FichierExiste(chemin) Then
            p_nom_fichier_importation = chemin
        ElseIf p_traitement_background Then
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> Le fichier '" & chemin & "' n'existe pas (local)"
            If p_traitement_background_semiauto Then
                MsgBox p_mess_fait_background
                End
            Else
                End
            End If
        Else
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> Le fichier '" & chemin & "' n'existe pas (local)"
        End If
    End If
    
    If p_nom_fichier_importation = "" Then
        Call MsgBox("Fichier '" & chemin & "' non trouvé " & IIf(p_est_sur_serveur, "(serveur)", "(local)"), vbInformation + vbOKOnly, "")
        ' Récupérer le nom du fichier d'importation
        Set frm = Com_ChoixFichier
        p_nom_fichier_importation = Com_ChoixFichier.AppelFrm("Choix du fichier d'importation", _
                                                              "", _
                                                              "c:", _
                                                              P_EXTENSIONS_FICHIERS_IMPORTATION, _
                                                              False)
        Set frm = Nothing
        p_est_sur_serveur = False
    End If
    
    If p_nom_fichier_importation = "" Then
        P_InitFichierImportation = P_NON
        Exit Function
    End If

    sql = "SELECT * FROM PrmGenB"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        P_InitFichierImportation = P_ERREUR
        Exit Function
    End If
    If rs.EOF Then
        rs.Close
        P_InitFichierImportation = P_NON
        Exit Function
    End If
        
    liste_pos = rs("PGB_FichLstPos").Value
    p_separateur = rs("PGB_FichSep").Value
    p_format_code = rs("PGB_code").Value
    p_mdp = rs("PGB_mdp").Value
    type_fichier = rs("PGB_FichType").Value
    If p_separateur = "TAB" Then
        p_separateur = Chr(9)
    End If
    lst_info_suppl = rs("PGB_LstPosInfoAutre").Value
    rs.Close
    
    ' Exemple du champ PGB_FichLstPos de la table PrmGenB
' NOM=3;PRENOM=4;MATRICULE=1;CODE_SECTION=5;LIB_SECTION=6;CODE_FONCTION=7;LIB_FONCTION=8;
    If type_fichier = 1 Then
        p_pos_nom = STR_GetChamp(STR_GetChamp(liste_pos, ";", 0), "=", 1)
        p_pos_prenom = STR_GetChamp(STR_GetChamp(liste_pos, ";", 1), "=", 1)
        p_pos_matricule = STR_GetChamp(STR_GetChamp(liste_pos, ";", 2), "=", 1)
        p_pos_code_section = STR_GetChamp(STR_GetChamp(liste_pos, ";", 3), "=", 1)
        p_pos_lib_section = STR_GetChamp(STR_GetChamp(liste_pos, ";", 4), "=", 1)
        p_pos_code_emploi = STR_GetChamp(STR_GetChamp(liste_pos, ";", 5), "=", 1)
        p_pos_lib_emploi = STR_GetChamp(STR_GetChamp(liste_pos, ";", 6), "=", 1)
        nb = STR_GetNbchamp(liste_pos, ";")
        If nb > 7 Then
            p_pos_civilite = STR_GetChamp(STR_GetChamp(liste_pos, ";", 7), "=", 1)
        Else
            p_pos_civilite = -1
        End If
        If nb > 8 Then
            p_pos_njf = STR_GetChamp(STR_GetChamp(liste_pos, ";", 8), "=", 1)
        Else
            p_pos_njf = -1
        End If
    Else
        ' positionnel
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 0), "=", 1), ":", 0)
        If s <> "" Then p_pos_nom = s
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 1), "=", 1), ":", 0)
        If s <> "" Then p_pos_prenom = s
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 2), "=", 1), ":", 0)
        If s <> "" Then p_pos_matricule = s
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 3), "=", 1), ":", 0)
        If s <> "" Then p_pos_code_section = s
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 4), "=", 1), ":", 0)
        If s <> "" Then p_pos_lib_section = s
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 5), "=", 1), ":", 0)
        If s <> "" Then p_pos_code_emploi = s
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 6), "=", 1), ":", 0)
        If s <> "" Then p_pos_lib_emploi = s
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 7), "=", 1), ":", 0)
        If s <> "" Then
            p_pos_civilite = s
        Else
            p_pos_civilite = -1
        End If
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 8), "=", 1), ":", 0)
        If s <> "" Then
            p_pos_njf = s
        Else
            p_pos_njf = -1
        End If
        '
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 0), "=", 1), ":", 1)
        If s <> "" Then p_long_nom = s
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 1), "=", 1), ":", 1)
        If s <> "" Then p_long_prenom = s
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 2), "=", 1), ":", 1)
        If s <> "" Then p_long_matricule = s
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 3), "=", 1), ":", 1)
        If s <> "" Then p_long_code_section = s
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 4), "=", 1), ":", 1)
        If s <> "" Then p_long_lib_section = s
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 5), "=", 1), ":", 1)
        If s <> "" Then p_long_code_emploi = s
        s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 6), "=", 1), ":", 1)
        If s <> "" Then p_long_lib_emploi = s
        If p_pos_civilite <> -1 Then
            s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 7), "=", 1), ":", 1)
            If s <> "" Then
                p_long_civilite = s
            Else
                p_pos_civilite = -1
                p_long_civilite = -1
            End If
        Else
            p_pos_civilite = -1
            p_long_civilite = -1
        End If
        If p_pos_njf <> -1 Then
            s = STR_GetChamp(STR_GetChamp(STR_GetChamp(liste_pos, ";", 8), "=", 1), ":", 1)
            If s <> "" Then
                p_long_njf = s
            Else
                p_pos_njf = -1
                p_long_njf = -1
            End If
        Else
            p_pos_njf = -1
            p_long_njf = -1
        End If
    End If
    ' les type d'informations supplémentaires
    n = STR_GetNbchamp(lst_info_suppl, "|")
    p_nbr_lstInfoSuppl = 0
    For I = 0 To n - 1
        num = Mid$(STR_GetChamp(STR_GetChamp(lst_info_suppl, "|", I), ";", 0), 2)
        sql = "select count(*) from KB_TypeInfoSuppl where KB_TisNum=" & num
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            P_InitFichierImportation = P_ERREUR
            Exit Function
        End If
        If lnb > 0 Then
            ReDim Preserve LISTE_TIS_POS(p_nbr_lstInfoSuppl)
            LISTE_TIS_POS(p_nbr_lstInfoSuppl).prmgenb_tis_num = num
            LISTE_TIS_POS(p_nbr_lstInfoSuppl).prmgenb_tis_lien = VoirSiAlimente(num)
            If type_fichier = 1 Then
                LISTE_TIS_POS(p_nbr_lstInfoSuppl).prmgenb_tis_pos = STR_GetChamp(STR_GetChamp(lst_info_suppl, "|", I), ";", 1) - 1
            Else
                LISTE_TIS_POS(p_nbr_lstInfoSuppl).prmgenb_tis_pos = STR_GetChamp(STR_GetChamp(lst_info_suppl, "|", I), ";", 1)
            End If
            If type_fichier = 2 Then
                s = STR_GetChamp(STR_GetChamp(lst_info_suppl, "|", I), ";", 2)
                If IsNumeric(s) And val(s) > 0 Then
                    LISTE_TIS_POS(p_nbr_lstInfoSuppl).prmgenb_tis_long = s
                End If
            End If
            sql = "SELECT KB_TISLibelle From KB_TypeInfoSuppl WHERE KB_TisNum=" & num
            Call Odbc_RecupVal(sql, lib)
            LISTE_TIS_POS(p_nbr_lstInfoSuppl).prmgenb_tis_lib = lib
            p_nbr_lstInfoSuppl = p_nbr_lstInfoSuppl + 1
        End If
    Next I

    P_InitFichierImportation = P_OUI

End Function


Function VoirSiAlimente(v_num)
    Dim sql As String, rs As rdoResultset
    
    sql = "SELECT ZU_Num FROM zoneutil WHERE ZU_Alimente=" & v_num
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        VoirSiAlimente = 0
    ElseIf rs.EOF Then
        VoirSiAlimente = 0
    Else
        VoirSiAlimente = rs("ZU_Num")
    End If
End Function

Public Function P_get_fcttrav(ByVal v_spm As String) As String
' ************************************************************
' Construire la chaine U_FctTrav sous la forme "Fx;Fy;..."
' ************************************************************
    Dim sql As String, spm_en_cours As String
    Dim nbr_postes As Integer, I As Integer, nbr As Integer
    Dim num_poste As Long

    P_get_fcttrav = ""
    nbr_postes = STR_GetNbchamp(v_spm, "|")
    For I = 0 To nbr_postes - 1
        spm_en_cours = STR_GetChamp(v_spm, "|", I)
        nbr = STR_GetNbchamp(spm_en_cours, ";")
        sql = "SELECT PO_FTNum FROM Poste WHERE PO_Num=" & Mid$(STR_GetChamp(spm_en_cours, ";", nbr - 1), 2)
        Call Odbc_RecupVal(sql, num_poste)
        If InStr(P_get_fcttrav, "F" & num_poste & ";") = 0 Then
            P_get_fcttrav = P_get_fcttrav & "F" & num_poste & ";"
        End If
    Next I

End Function

Public Function P_get_num_fct(ByVal v_num_poste As Long) As Long
' *************************************************************
' Retourne le numero de la fonction du poste v_num_poste
' *************************************************************
    If Odbc_RecupVal("SELECT PO_FtNum FROM Poste" _
                   & " WHERE PO_Num = " & v_num_poste, P_get_num_fct) = P_ERREUR Then
        Exit Function
    End If

End Function

Public Function P_get_num_srv_pere(ByVal v_num_srv As Long) As String
' *************************************************************
' Retourne le numero et le libellé du service pere d'un service
' sous la forme "num_srv=LIB_SRV"
' *************************************************************
    Dim num_srv As Long

    If Odbc_RecupVal("SELECT SRV_NumPere FROM Service WHERE SRV_Num=" & v_num_srv, num_srv) = P_ERREUR Then
        P_get_num_srv_pere = ""
        Exit Function
    End If

    If num_srv = 0 Then
        P_get_num_srv_pere = num_srv & "=" & "" ' pas de nom pour la racine ?
    Else
        P_get_num_srv_pere = num_srv & "=" & P_get_lib_srv_poste(num_srv, P_SERVICE)
    End If

End Function

Public Function P_get_service_du_poste(ByVal v_num_poste As Long) As String
' *****************************************************************************
' Retourne le numero et le libellé du service dans le quel est rataché ce poste
' sous la forme "num_srv=LIB_SRV"
' *****************************************************************************
    Dim num_srv As Long
    Dim sql As String, rs As rdoResultset
    
    sql = "SELECT PO_SRVNum FROM Poste WHERE PO_Num=" & v_num_poste
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        P_get_service_du_poste = ""
    ElseIf rs.EOF Then
        P_get_service_du_poste = ""
    ElseIf Odbc_RecupVal("SELECT PO_SRVNum FROM Poste WHERE PO_Num=" & v_num_poste, num_srv) = P_ERREUR Then
        P_get_service_du_poste = ""
    Else
        P_get_service_du_poste = num_srv & "=" & P_get_lib_srv_poste(num_srv, P_SERVICE)
    End If
    rs.Close
End Function

Public Function P_FctRecupNiveau(ByVal v_srvnum As Integer) As String
    Dim sql As String, rs As rdoResultset
    
    sql = "select * from service where SRV_Num=" & v_srvnum
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        P_FctRecupNiveau = ""
        Exit Function
    Else
        If rs("srv_NivsNum") = 0 Then
            P_FctRecupNiveau = ""
        Else
            sql = "select Nivs_Nom from niveau_structure where Nivs_Num=" & rs("srv_NivsNum")
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                P_FctRecupNiveau = ""
            ElseIf rs.EOF Then
                P_FctRecupNiveau = ""
            Else
                P_FctRecupNiveau = rs("Nivs_Nom")
            End If
        End If
    End If
End Function

Public Function P_get_lib_srv_poste(ByVal v_num As Long, _
                                    ByVal v_P_POSTE_ou_P_SERVICE As Integer) As String
' **********************************************************
' Retourne le libellé du SERVICE/POSTE determiné par son NUM
' **********************************************************
    Dim sql As String
    Dim rs As rdoResultset

    P_get_lib_srv_poste = ""
    If v_P_POSTE_ou_P_SERVICE = P_SERVICE Then
        sql = "SELECT SRV_Nom FROM Service WHERE SRV_Num=" & v_num
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Exit Function
        End If
        If Not rs.EOF Then
            P_get_lib_srv_poste = P_FctRecupNiveau(v_num) & " " & rs("SRV_Nom").Value
        End If
        rs.Close
    Else ' v_P_POSTE_ou_P_SERVICE = P_POSTE
        sql = "SELECT FT_Libelle FROM FctTrav, Poste" & _
              " WHERE FT_Num=PO_FTNum AND PO_Num=" & v_num
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Exit Function
        End If
        If Not rs.EOF Then
            P_get_lib_srv_poste = rs("FT_Libelle").Value
        End If
        rs.Close
    End If

End Function

Public Function P_get_nom_appli(ByVal v_numappli As Long) As String
' ****************************************************************
' Retourner le NOM de l'application determinées par son v_numappli
' ****************************************************************
    If v_numappli = 0 Then
        P_get_nom_appli = "IMPORT"
    Else
        If Odbc_RecupVal("SELECT APP_Code FROM Application WHERE APP_Num=" & v_numappli, _
                         P_get_nom_appli) = P_ERREUR Then
            P_get_nom_appli = "[ INTROUVABLE ]"
        End If
    End If

End Function

Public Function P_get_nom_piece(ByVal v_num_piece As Long) As String
' ************************************
' Rtourner le APP_Num de l'application
' ************************************
    If Odbc_RecupVal("SELECT PC_Nom FROM Piece WHERE PC_Num=" & v_num_piece, P_get_nom_piece) = P_ERREUR Then
    End If

End Function

Public Function P_get_num_appli(ByVal v_nomappli As String) As Long
' ************************************
' Rtourner le APP_Num de l'application
' ************************************
    If Odbc_RecupVal("SELECT APP_Num FROM Application WHERE APP_Nom='" & v_nomappli & "'", P_get_num_appli) = P_ERREUR Then
    End If

End Function

Public Function P_get_num_srv_poste(ByVal v_spm As String, ByVal v_P_POSTE_ou_P_SERVICE As Integer) As Long
' *****************************************************************************************
' Retourner le NUM_SRV ou NUM_POSTE en n'examinant que la première suite "Sx;Sy;Pz;|"
' *****************************************************************************************
    Dim le_premier_spm As String

    If v_spm = "" Then
        P_get_num_srv_poste = 0
        Exit Function
    End If
    If InStr(v_spm, "|") = 0 Then
        'Call MsgBox("P_get_num_srv_poste : v_spm mal formatté " & v_spm)
        v_spm = v_spm & "|"
    End If
    le_premier_spm = STR_GetChamp(v_spm, "|", 0)
    ' Le v_P_POSTE_ou_P_SERVICE a comme valeurs : 1 pour POSTE ou 2 pour SERVICE
    P_get_num_srv_poste = Mid$(STR_GetChamp(le_premier_spm, ";", STR_GetNbchamp(le_premier_spm, ";") - v_P_POSTE_ou_P_SERVICE), 2)

End Function

Public Function P_get_poste_modif(ByVal v_old_spm As String, ByVal v_new_spm As String) As String
' ******************************************************************************************************
' Appelée depuis ImportationAnnuaire.get_commentaire_modif() & PrmPersonne.remplir_utilmouvement().
' Renseigner la partie 'POSTE' du champs 'commentaire' de la table UtilMouvement.
' Retourne une chaine qui designe: les valeurs avant|le(s) ajout(s);le(s) suppression(s)
' sous la forme: "POSTE=Pa;Pb;..;Pn;Po;|A:Pb;Pc;..;S:Pn;"
' où la première partie designe les postes occupés avant, A: pour les Ajouts et S: pour les Suppressions
' ******************************************************************************************************
    Dim spm_en_cours As String, resultat_avant As String, poste_en_cours As String
    Dim A_existe As Boolean, S_existe As Boolean
    Dim nbr As Integer, nbr2 As Integer, I As Integer, j As Integer

    A_existe = False
    S_existe = False
    ' ---------------- POSTES OCCUPES AVANT -----------------
    nbr = STR_GetNbchamp(v_old_spm, "|")
    ' parcourrir les champs séparés par des "|"
    For I = 0 To nbr - 1
        spm_en_cours = STR_GetChamp(v_old_spm, "|", I)
        resultat_avant = resultat_avant & "P" & P_get_num_srv_poste(spm_en_cours, P_POSTE) & ";"
    Next I
    ' -------------------------------------------------------
    P_get_poste_modif = "POSTE=" & resultat_avant & "|"
    ' -------------- MOUVEMENTS SUR LES POSTES --------------
    nbr = STR_GetNbchamp(v_new_spm, "|")
    nbr2 = STR_GetNbchamp(resultat_avant, ";")

    ' LES AJOUTS:
    ' parcourrir les champs du nouveau SPM séparés par des "|"
    For I = 0 To nbr - 1
        spm_en_cours = STR_GetChamp(v_new_spm, "|", I)
        poste_en_cours = "P" & P_get_num_srv_poste(spm_en_cours, P_POSTE)
        ' parcourrir les champs du "resulat_avant" séparés par des ";"
        For j = 0 To nbr2 - 1
            If poste_en_cours = STR_GetChamp(resultat_avant, ";", j) Then
                GoTo lab_i_suivant
            End If
        Next j
        P_get_poste_modif = P_get_poste_modif & IIf(A_existe, "", "A:") & poste_en_cours & ";"
        A_existe = True
lab_i_suivant:
    Next I

    ' LES SUPPRESSIONS:
    ' parcourrir les champs du "resulat_avant" séparés par des ";"
    For j = 0 To nbr2 - 1
        poste_en_cours = STR_GetChamp(resultat_avant, ";", j)
        ' parcourrir les champs du nouveau SPM séparés par des "|"
        For I = 0 To nbr - 1
            spm_en_cours = STR_GetChamp(v_new_spm, "|", I)
            If poste_en_cours = "P" & P_get_num_srv_poste(spm_en_cours, P_POSTE) Then
                GoTo lab_j_suivant
            End If
        Next I
        P_get_poste_modif = P_get_poste_modif & IIf(S_existe, "", "S:") & poste_en_cours & ";"
        S_existe = True
lab_j_suivant:
    Next j

End Function

Private Function init_param_debug() As Integer

    Dim stype_bdd As String, nom_bdd As String, nomini As String
    Dim ask_enreg As Boolean
    Dim reponse As Integer
    Dim s_derutil As String
    
    p_multilabo = True

    p_chemin_appli = "c:\kalidoc"

    'p_nomini = InputBox("Chemin du .ini : ", , "c:\kalidoc\kalidoc.ini")
    
    ' p_nomini = "c:\kalidoc\kalidoc_demo_dev.ini"
    ' chercher les .ini
    nomini = SYS_GetIni("FICHIER_INI", "DERNIER", p_chemin_appli & "\dernier_ini_ouvert.txt")
    p_nomini = Choisir_Ini(p_nomini, p_chemin_appli)
    
    If p_nomini = "" Then
        init_param_debug = P_ERREUR
        Exit Function
    End If

    ask_enreg = False
    
    ' Dernier utilisateur qui a ouvert
    If nomini = p_nomini Then
        s_derutil = SYS_GetIni("FICHIER_INI", "UTILISATEUR", p_chemin_appli & "\dernier_ini_ouvert.txt")
        If s_derutil <> "" Then
            p_NumUtil = s_derutil
        End If
    End If
    
    ' Type de base
    stype_bdd = SYS_GetIni("BASE", "TYPE", p_nomini)
    If stype_bdd = "" Then
lab_sais_typb:
        stype_bdd = InputBox("Type de base (PG, MDB) : ", , "MDB")
        If stype_bdd = "" Then
            init_param_debug = P_ERREUR
            Exit Function
        End If
        If stype_bdd <> "MDB" And stype_bdd <> "PG" Then
            GoTo lab_sais_typb
        End If
        ask_enreg = True
    End If
    ' Nom base
    nom_bdd = SYS_GetIni("BASE", "NOM", p_nomini)
    If nom_bdd = "" Then
        nom_bdd = InputBox("Nom de la base : ", , "c:\kalidoc\kalidoc.mdb")
        If nom_bdd = "" Then
            init_param_debug = P_ERREUR
            Exit Function
        End If
        ask_enreg = True
    End If
    
    ' Enregistrement des infos base
    If ask_enreg Then
        reponse = MsgBox("Voulez-vous enregistrer les informations saisies ?", vbQuestion + vbYesNo, "")
        If reponse = vbYes Then
            Call SYS_PutIni("BASE", "TYPE", stype_bdd, p_chemin_appli & "\kalibottin.ini")
            Call SYS_PutIni("BASE", "NOM", nom_bdd, p_chemin_appli & "\kalibottin.ini")
        End If
    End If

    ' Connexion à la base
    If Odbc_Init(stype_bdd, nom_bdd) = P_ERREUR Then
        init_param_debug = P_ERREUR
        Exit Function
    End If

    init_param_debug = P_OK

End Function


Private Function Choisir_Ini(ByVal v_p_nomini As String, ByVal v_p_chemin_appli As String) As String
    Dim s As String
    Dim n As Integer
    Dim nomfich As String
    
lab_choix:
    'Call FRM_ResizeForm(Me, 0, 0)
    
    Call CL_Init
    Call CL_InitTaille(0, -15)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("Ouvrir le ini", "", 0, 0, 1500)
    Call CL_AddBouton("ODBC", "", 0, 0, 1500)
    Call CL_AddBouton("Créer à partir de", "", 0, 0, 1800)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    s = Dir$(v_p_chemin_appli & "\kalibot*.ini")
    n = 0
    While s <> ""
        If UCase(v_p_chemin_appli & "\" & s) = UCase(v_p_nomini) Then
            Call CL_AddLigne(s, n, "", True)
        Else
            Call CL_AddLigne(s, n, "", False)
        End If
        n = n + 1
        s = Dir$()
    Wend
    Call CL_AffiSelFirst
    
    ChoixListe.Show 1
    ' Quitter
    If CL_liste.retour = 4 Then
        Choisir_Ini = P_NON
        Exit Function
    End If
    
    nomfich = v_p_chemin_appli & "\" & CL_liste.lignes(CL_liste.pointeur).texte
    
    ' Ouvrir
    If CL_liste.retour = 1 Then
        If FICH_FichierExiste(nomfich) Then
            Call SYS_StartProcess(nomfich)
        Else
            MsgBox "Le fichier '" & nomfich & "' n'a pas été trouvé."
        End If
        GoTo lab_choix
    End If
    
    ' ODBC
    If CL_liste.retour = 2 Then
        Call SYS_StartProcess("c:\windows\system32\control", "odbccp32.cpl")
        GoTo lab_choix
    End If
        
    ' Nouveau à partir de
    If CL_liste.retour = 3 Then
        If FICH_FichierExiste(nomfich) Then
            Choisir_Ini = nomfich
        Else
            MsgBox nomfich & " n'existe pas"
        End If
encore:
        s = InputBox("Nom du nouveau .ini", "Créer un .ini", "kalidoc_nouveau.ini")
        If s <> "" Then
            If FICH_FichierExiste(UCase(v_p_chemin_appli & "\" & s)) Then
                MsgBox s & " existe déjà"
                GoTo encore
            Else
                If FICH_CopierFichier(nomfich, v_p_chemin_appli & "\" & s) = P_ERREUR Then
                    MsgBox "Erreur pour copier " & nomfich & " vers " & s
                Else
                    nomfich = v_p_chemin_appli & "\" & s
                    If FICH_FichierExiste(nomfich) Then
                        Call SYS_StartProcess(nomfich)
                    End If
                End If
            End If
        End If
    End If
    
    If FICH_FichierExiste(nomfich) Then
        Choisir_Ini = nomfich
        '
        Call SYS_PutIni("FICHIER_INI", "DERNIER", nomfich, v_p_chemin_appli & "\dernier_ini_ouvert.txt")
        '
    Else
        MsgBox nomfich & " n'existe pas"
    End If
End Function

Private Function init_param_exe(ByVal v_scmd As String) As Integer

    Dim s As String, stype_bdd As String, nom_bdd As String
    Dim saction As String, snumdos As String, tbldoscli() As String, snumcli As String
    Dim snumdosp As String, titredos As String, lstresp As String
    Dim etat As Boolean
    Dim nbprm As Integer, n As Integer, I As Integer
    Dim frm As Form

    nbprm = STR_GetNbchamp(v_scmd, ";")
    If nbprm < 4 Then
        Call MsgBox("Usage : KaliDoc <Chemin application>;<Type BDD>;<Nom BDD>;<MULT>;[NOM INI]" & vbCr & vbLf _
                    & "cmd:" & v_scmd & vbCr & vbLf _
                    & "L'application ne peut pas être lancée.", vbInformation + vbOKOnly)
        init_param_exe = P_ERREUR
        Exit Function
    End If

    ' Chemin appli
    p_chemin_appli = STR_GetChamp(v_scmd, ";", 0)
    If p_chemin_appli = "" Then
        Call MsgBox("<Chemin application> est vide." & vbCr & vbLf _
                    & "cmd:" & v_scmd & vbCr & vbLf _
                    & "Usage : KaliDoc <Chemin application>;<Type BDD>;<Nom BDD>;<MULT>", vbInformation + vbOKOnly)
        init_param_exe = P_ERREUR
        Exit Function
    End If
    ' Type de base
    stype_bdd = STR_GetChamp(v_scmd, ";", 1)
    If stype_bdd <> "PG" And stype_bdd <> "MDB" Then
        Call MsgBox("Type de Base incorrect : " & stype_bdd, vbInformation + vbOKOnly)
        init_param_exe = P_ERREUR
        Exit Function
    End If
    ' Nom de la base
    nom_bdd = STR_GetChamp(v_scmd, ";", 2)
    If nom_bdd = "" Then
        Call MsgBox("Pas de nom de base.", vbInformation + vbOKOnly)
        init_param_exe = P_ERREUR
        Exit Function
    End If
    ' Multi
    s = STR_GetChamp(v_scmd, ";", 3)
    If s = "1" Then
        p_multilabo = True
    Else
        p_multilabo = False
    End If

    p_nomini = ""
    If nbprm > 4 Then
        p_nomini = STR_GetChamp(v_scmd, ";", 4)
    End If
    If p_nomini = "" Then
        p_nomini = "kalibottin.ini"
    End If
    p_nomini = p_chemin_appli & "\" & p_nomini

    ' Connexion à la base
    If Odbc_Init(stype_bdd, nom_bdd) = P_ERREUR Then
        init_param_exe = P_ERREUR
        Exit Function
    End If

    init_param_exe = P_OK

End Function

Public Function P_InsertIntoUtilmouvement(ByVal v_numutil As Long, ByVal v_typemvt As String, _
                            ByVal v_commentaire As String, ByVal v_num_appli As Long) As Integer
' *************************************************************************
' Ajouter un nouvel enregistrement dans la table UtilMouvement
' Appelée depuis:
' 1. ImportaionAnnuaire{remplir_utilmouvement() & creer_cette_personne() _
                      & associer_cette_personne() & destruction()}
' 2. PrmPersonne{remplir_utilmouvement() & enregistrer_coordonnees()}
' *************************************************************************
    Dim lng As Long
    
    If Odbc_AddNew("UtilMouvement", "UM_Num", "UM_Seq", False, lng, _
                   "UM_AppNum", v_num_appli, _
                   "UM_Date", Date, _
                   "UM_TypeMvt", v_typemvt, _
                   "UM_Commentaire", v_commentaire, _
                   "UM_UNum", v_numutil) = P_ERREUR Then
        P_InsertIntoUtilmouvement = P_ERREUR
        Exit Function
    End If

    P_InsertIntoUtilmouvement = P_OK

End Function

Sub Main()

    Dim scmd As String, sql As String
    Dim s As String

    If App.PrevInstance Then
        Call MsgBox("KaliBottin a déjà été lancé.", vbInformation + vbOKOnly, "")
        End
    End If

    p_NumUtil = 0

    ' Param de l'application
    scmd = Command$

    ' Mode DEBUG
    'scmd = "c:\kalidoc;PG;VM_KALI_CHMULH;;kalibottin_chmul.ini;AUTO"
    'scmd = "c:\kalidoc;PG;VM_KALI_GAP;;Kalibottin_GAP.ini;AUTO"
    'scmd = "c:\kalidoc;PG;VM_KALI_GAP;;Kalibottin_GAP.ini;PREIMPORT"
    'scmd = "c:\kalidoc;PG;VM_KALI_GAP;;Kalibottin_GAP.ini;AUTO"

    If InStr(scmd, "DEBUG") > 0 Then
        If init_param_debug() = P_ERREUR Then
            End
        End If
    Else
        If init_param_exe(scmd) = P_ERREUR Then
            End
        End If
    End If

    If InStr(scmd, "AUTO") > 0 Then
        p_traitement_background = True
    ElseIf InStr(scmd, "PREIMPORT") > 0 Then
        p_traitement_background = True
        p_traitement_préimport_seul = True
    End If
    
    If charger_prmgen() = P_ERREUR Then
        End
    End If
    ' fichier si on outrepasse en local
    p_nom_fichier_importation_local = SYS_GetIni("FICHIER", "LOCAL", p_nomini)
    
    ' fichier à executer en pré-import
    p_programme_preimport_exe = SYS_GetIni("PREIMPORT", "FICHIER_EXE", p_nomini)
    p_programme_preimport_log = SYS_GetIni("PREIMPORT", "FICHIER_LOG", p_nomini)
    p_programme_preimport_lock = SYS_GetIni("PREIMPORT", "FICHIER_LOCK", p_nomini)
    
    ' fichier ini pour lancer kalidoc avec lance.exe
    p_nom_fichier_ini_kalidoc = SYS_GetIni("KALIDOC", "INI", p_nomini)
    If p_nom_fichier_ini_kalidoc = "" Then
        s = "Définir la variable [KALIDOC] INI= dans " & p_nomini
        If p_traitement_background Then
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> " & s
            If p_traitement_background_semiauto Then
                MsgBox p_mess_fait_background
            End If
        Else
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> " & s
            Call MsgBox(s)
        End If
        'p_nom_fichier_ini_kalidoc = "KaliDoc.Ini"
    ElseIf Not FICH_FichierExiste(p_chemin_appli & "\" & p_nom_fichier_ini_kalidoc) Then
        s = "Le fichier " & p_nom_fichier_ini_kalidoc & " défini par la variable [KALIDOC] INI= dans " & p_nomini & " n'existe pas"
        If p_traitement_background Then
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> " & s
            Call ImportationAnnuaire.Traitement_Background(s)
        Else
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> " & s
            Call MsgBox(s)
            End
        End If
    End If
    
    sql = "SELECT L_Num FROM Laboratoire ORDER BY L_Num"
    If Odbc_RecupVal(sql, p_NumLaboDefaut) = P_ERREUR Then
        End
    End If

    If p_traitement_background Then
        Call Menu.mnuImportation_Click
    Else
        ' Menu
        Menu.Show 0
    End If
End Sub

Public Function P_lire_valeur(v_type_fichier, v_ligne_lue, v_separateur, v_position, v_longueur, v_libelle)
    Dim s As String
    
    If v_type_fichier = 2 Then
        If v_position >= 0 And v_longueur > 0 Then
            If IsNumeric(v_position) And IsNumeric(v_longueur) Then
                s = Trim$(Mid(v_ligne_lue, v_position, v_longueur))
            Else
                MsgBox "Incorrect : position=" & v_position & " longueur=" & v_longueur & " pour " & v_libelle
                End
            End If
        Else
            MsgBox "Incorrect : position=" & v_position & " longueur=" & v_longueur & " pour " & v_libelle
            End
        End If
    Else
        s = Trim$(STR_GetChamp(v_ligne_lue, v_separateur, v_position))
    End If
    P_lire_valeur = s
End Function

Public Function P_lister_fichiers(ByVal v_nomrep As String, _
                                 ByVal v_nomfich As String, _
                                 ByVal v_mode As String, _
                                 ByRef r_tblfich() As String) As Integer

    Dim liberr As String, path As String, nomdoc As String, listeFich As String
    Dim ext As String
    Dim cr As Integer, nbfich As Integer, I As Integer
    
    If Not KF_EstRepertoire(v_nomrep, False) Then
        P_lister_fichiers = 0
        Exit Function
    End If
    
    path = v_nomrep
    If v_nomfich <> "" Then
        path = path & "/" & v_nomfich
    End If
    cr = HTTP_Appel_ListeFichiers(path, liberr)
    If cr = HTTP_LISTEFICH_DOSINEX Then
        Call MsgBox("Le répertoire '" & v_nomrep & "' est inexistant.", vbOKOnly + vbInformation, "")
        P_lister_fichiers = 0
        Exit Function
    ElseIf cr = HTTP_LISTEFICH_DOSINACC Then
        Call MsgBox("Le répertoire '" & v_nomrep & "' est inaccessible.", vbOKOnly + vbInformation, "")
        P_lister_fichiers = 0
        Exit Function
    ElseIf cr = HTTP_LISTEFICH_ERREUR Then
        Call MsgBox("Erreur lors du listage du répertoire '" & v_nomrep & "' " & liberr, vbOKOnly + vbInformation, "")
        P_lister_fichiers = 0
        Exit Function
    Else
        nbfich = STR_GetChamp(liberr, "|", 1)
        If v_mode = "NOMBRE" Then
            P_lister_fichiers = nbfich
            Exit Function
        End If
        ' v_mode = LISTE
        If nbfich > 0 Then
            ReDim r_tblfich(nbfich - 1) As String
            listeFich = STR_GetChamp(liberr, "|", 2)
            For I = 0 To nbfich
                nomdoc = STR_GetChamp(listeFich, ";", I)
                If nomdoc = "!FIN!" Then
                    Exit For
                Else
                    r_tblfich(I) = nomdoc
                End If
            Next I
        End If
        P_lister_fichiers = nbfich
    End If

End Function

Public Function P_mouvements_a_envoyer() As Boolean
' *********************************************************************
' Ne pas générer plus q'une fois les fichiers le jours même.
' Déterminer s'il y a des mouvement à envoyer aux responsables:
' UM_DateEnvoi IS NNULL - ou - il reste des fichiers non encore envoyés
' *********************************************************************
    Dim sql As String, tbl_fich() As String, chemin As String
    Dim nbr_mvt_a_envoyer As Long
    Dim rs As rdoResultset

    ' ******************************************************************************************
    ' INTERDIRE LA DOUBLE GÉNÉRATION DE FICHIERS LE JOURS MÊME
    sql = "SELECT APP_Code FROM Application" _
        & " WHERE APP_Num<>" & p_appli_kalidoc _
        & " AND APP_Num<>" & p_appli_kalibottin
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF
        chemin = p_CheminKW & "/kalibottin/mouvements/" & rs("APP_Code").Value
        nbr_mvt_a_envoyer = P_lister_fichiers(chemin, _
                                             "*" & format(Date, "YYYYMMDD") & ".txt", _
                                             "NOMBRE", tbl_fich())
        If nbr_mvt_a_envoyer > 0 Then
            rs.Close
            GoTo lab_email
        End If
        rs.MoveNext
    Wend
    rs.Close
    ' ******************************************************************************************
    ' reste-t-il des fichier à générer ?
    sql = "SELECT COUNT(*) FROM UtilMouvement WHERE UM_DateEnvoi IS NULL"
    If Odbc_Count(sql, nbr_mvt_a_envoyer) = P_ERREUR Then
        Exit Function
    End If
    If nbr_mvt_a_envoyer > 0 Then
        P_y_a_des_mouvements_a_envoyer = P_FICHIER_A_CREER
        GoTo lab_ok
    End If
    ' ******************************************************************************************
    ' reste-t-il des e-mails à envoyer ?
lab_email:
    sql = "SELECT APP_Code FROM Application" _
        & " WHERE APP_Num<>" & p_appli_kalidoc & " AND APP_Num<>" & p_appli_kalibottin
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF
        chemin = p_CheminKW & "/kalibottin/mouvements/" & rs("APP_Code").Value
        nbr_mvt_a_envoyer = P_lister_fichiers(chemin, _
                                             "*_env_*-gen.txt", _
                                             "NOMBRE", tbl_fich())
        If nbr_mvt_a_envoyer > 0 Then
            rs.Close
            P_y_a_des_mouvements_a_envoyer = P_MAILS_A_ENVOYER
            GoTo lab_ok
        End If
        rs.MoveNext
    Wend
    rs.Close
    ' ******************************************************************************************

lab_erreur:
    P_mouvements_a_envoyer = False
    Exit Function

lab_ok:
    P_mouvements_a_envoyer = True

End Function

Public Function P_EnvoyerMail(ByVal v_nomsrc As String, _
                                ByVal v_adrsrc As String, _
                                ByVal v_nomdest As String, _
                                ByVal v_adrdest As String, _
                                ByVal v_sujet As String, _
                                ByVal v_message As Variant, _
                                ByVal v_nomfich As String) As Integer

    Dim cr As Integer
    Dim frm As Form
    
    Set frm = FMail_SMTP
    cr = FMail_SMTP.EnvoiMessage(v_nomsrc, v_adrsrc, v_nomdest, v_adrdest, v_sujet, v_message, v_nomfich)
    Set frm = Nothing
    
    P_EnvoyerMail = cr
    
End Function

Private Function convertir_kalimail(ByVal v_txt As Variant) As Variant

    Dim url As String
    Dim pos As Integer, pos_blanc As Integer, pos_cr As Integer
    Dim s As Variant
    
'   preg_replace( "/\[url\](.+?)\[\/url\]/si", "<a href=\"$1\" target=\"km_url\">$1</a>", $v_txt );
    s = v_txt
    pos = InStr(v_txt, "http://")
    If pos > 0 Then
        pos_blanc = InStr(pos, v_txt, " ")
        pos_cr = InStr(pos, v_txt, vbCr)
        If pos_blanc > 0 Or pos_cr > 0 Then
            If pos_blanc > 0 And (pos_blanc < pos_cr Or pos_cr = 0) Then
                url = Mid$(v_txt, pos, pos_blanc - pos)
            Else
                url = Mid$(v_txt, pos, pos_cr - pos)
            End If
            s = left$(v_txt, pos - 1) & "<a href=""" & url & """ target=""km_url"">" & url & "</a>"
        End If
    End If
    s = Replace(s, vbCrLf, "<br/>")
    
    convertir_kalimail = s

End Function

Public Sub P_EnvoyerMessage(ByVal v_numutil As Long, _
                            ByVal v_adrdest As String, _
                            ByVal v_stitre As String, _
                            ByVal v_scomm As String, _
                            Optional v_nomfich As Variant)

    Dim asrc As String, adest As String, nomfich As String, sql As String
    Dim nomdest As String, nomsrc As String
    Dim cr As Integer
    Dim numzone As Long, numkm As Long, lbid As Long
    Dim rs As rdoResultset
    Dim frm As Form

    cr = P_ERREUR
    
    ' N° de la zone Adrmail
    If Odbc_RecupVal("SELECT ZU_Num FROM ZoneUtil WHERE ZU_Code='ADRMAIL'", _
                      numzone) = P_ERREUR Then
        GoTo lab_kalimail
    End If

    ' Destinataire
    If v_numutil = 0 Then
        adest = v_adrdest
        asrc = SYS_GetIni("MAILEXPED", "MAILEXPED", p_nomini)
        If adest <> "" And asrc <> "" Then
            GoTo LabSuite
        Else
            GoTo lab_kalimail
        End If
    End If
    sql = "SELECT UC_Valeur FROM UtilCoordonnee" _
        & " WHERE UC_Type='U'" _
        & " AND UC_TypeNum=" & v_numutil _
        & " AND UC_ZUNum=" & numzone
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_kalimail
    End If
    If rs.EOF Then ' pas d'envoi s'il n'y a pas d'adresse mail
        rs.Close
        GoTo lab_kalimail
    End If
    adest = rs("UC_Valeur").Value
    rs.Close
    Call P_RecupUtilPpointNom(v_numutil, nomdest)
    
    ' Emetteur
    sql = "SELECT UC_Valeur FROM UtilCoordonnee" _
        & " WHERE UC_Type='U'" _
        & " AND UC_TypeNum=" & p_NumUtil _
        & " AND UC_ZUNum=" & numzone
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_kalimail
    End If
    If rs.EOF Then
        rs.Close
        GoTo lab_kalimail
    End If
    asrc = rs("UC_Valeur").Value
    rs.Close
    Call P_RecupUtilPpointNom(p_NumUtil, nomsrc)

LabSuite:
    ' Envoi du message
    If v_numutil = 0 Then Exit Sub
    If p_smtp_adrsrv <> "" Then
        'Pour que les messages arrivent à Kalitech ...
        If p_smtp_webmaster = "kalitech@kalitech.fr" Then
            nomsrc = "Utilisateur" & p_NumUtil
            asrc = "kalitech@kalitech.fr"
            nomdest = "Utilisateur" & v_numutil
            adest = "kalitech@kalitech.fr"
        End If
        If IsMissing(v_nomfich) Then
            nomfich = ""
        Else
            nomfich = v_nomfich
        End If
        cr = P_EnvoyerMail(nomsrc, asrc, nomdest, adest, "KaliBottin - " & v_stitre, v_scomm, nomfich)
    End If
    
lab_kalimail:
    ' Pas de mail envoyé -> on envoie kalimail
    If cr = P_ERREUR Then
        v_scomm = convertir_kalimail(v_scomm)
        ' Enregistrer le sujet et le corps du KaliMail
        If Odbc_AddNew("kalimail", "km_num", "km_seq", True, numkm, _
                        "km_sujet", "KaliBottin - " & v_stitre, _
                        "km_corps", v_scomm, _
                        "km_typelien", 0, _
                        "km_liblien", "", _
                        "km_urllien", "") = P_ERREUR Then
            Exit Sub
        End If
        ' Enregistrer le destinataire
        Call Odbc_AddNew("kalimaildetail", "kmd_num", "kmd_seq", False, lbid, _
                       "kmd_kmnum", numkm, _
                       "kmd_expnum", p_NumUtil, _
                       "kmd_destnum", v_numutil, _
                       "kmd_dateenvoi", format(Date, "yyyy-mm-dd") & " " & format(Time, "hh:mm:ss"), _
                       "kmd_prioritaire", False, _
                       "kmd_niveau", 1, _
                       "kmd_suppexp", 0, _
                       "kmd_suppdest", 0, _
                       "kmd_numpere", 0, _
                       "kmd_accusereception", False, _
                       "kmd_kmdsnum", 0)
    End If

End Sub

Public Function P_RemplacerResponsableAppli(ByVal v_u_num As Long, ByVal v_nom As String, _
                        ByVal v_prenom As String, ByVal v_matricule As String) As Integer
' ****************************************************************************
' Remplacer la personne par une autre pour la responsabilité d'une application
' ****************************************************************************
    Dim sql As String, nom_appli As String, choix As String, _
        new_nom As String, new_prenom As String, new_matricule As String
    Dim lstresp As String, s As String
    Dim reponse As Integer, n As Integer, I As Integer
    Dim new_u_num As Long
    Dim rs As rdoResultset
    Dim frm As Form


    sql = "SELECT APP_Num, APP_Nom, app_lstresp FROM Application WHERE APP_lstResp like '%U" & v_u_num & ";%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF ' la liste des applications ayant comme responsable la personne à désactiver/supprimer
        nom_appli = rs("APP_Nom").Value
        s = v_nom & " " & v_prenom & " est responsable de l'application suivante: " & nom_appli _
                       & "." & vbCrLf & vbCrLf & "Vous devez choisir une autre personne pour la remplacer."
        If True Or p_traitement_background Then
            p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> " & s
            GoTo lab_fin
        End If
        reponse = MsgBox(s & " Voulez-vous la faire remplacer et continuer à la désactiver ?", _
                       vbExclamation + vbYesNo, "Attention")
        If reponse = vbNo Then
            rs.Close
            GoTo lab_erreur
        End If
lab_debut_liste_personnes:
        Set frm = ChoixUtilisateur
        choix = ChoixUtilisateur.AppelFrm("Choisir le responsable de: " & nom_appli, _
                                        "", _
                                        False, _
                                        False, _
                                        "")
        Set frm = Nothing
        If choix = "" Then
            reponse = MsgBox("Vous n'avez pas choisi de personne pour remplacer " & v_nom & " " & v_prenom _
                            & vbCrLf & " à la responsabilité de l'application " & nom_appli & "." _
                            & vbCrLf & vbCrLf & "Voulez-vous abandonner cette opération ?", _
                            vbCritical + vbYesNo, "Attention")
            If reponse = vbYes Then
                GoTo lab_erreur
            End If
            rs.Close
            GoTo lab_debut_liste_personnes
        End If
        new_u_num = p_tblu_sel(0)
        If new_u_num = v_u_num Then
            If MsgBox("Vous avez choisi la même personne ou il ne reste que cette personne." & vbCrLf & vbCrLf _
                    & "Voulez-vous abandonner cette opération ?", _
                    vbQuestion + vbYesNo, "Attention") = vbYes Then
                GoTo lab_erreur
            End If
            GoTo lab_debut_liste_personnes
        End If
        sql = "SELECT U_Nom, U_Prenom, U_Matricule" _
            & " FROM Utilisateur" _
            & " WHERE U_kb_actif=True AND U_Actif=True AND U_Num=" & new_u_num
        If Odbc_RecupVal(sql, new_nom, new_prenom, new_matricule) = P_ERREUR Then
            GoTo lab_erreur
        End If
        If MsgBox("Etes-vous sûr de vouloir remplacer:" & vbCrLf & vbCrLf _
                & vbTab & v_nom & " " & v_prenom & " (ayant le MATRICULE: " & v_matricule & ") ?" _
                & vbCrLf & vbCrLf & "par:" & vbCrLf & vbTab & new_nom & " " & new_prenom _
                & " (ayant le MATRICULE: " & new_matricule & ")" & vbCrLf & vbCrLf _
                & "à la responsabilité de " & nom_appli & " ?", _
                vbQuestion + vbYesNo, "") = vbNo Then
            GoTo lab_debut_liste_personnes
        End If
        
        lstresp = ""
        n = STR_GetNbchamp(rs("app_lstresp").Value, ";")
        For I = 0 To n - 1
            s = STR_GetChamp(rs("app_lstresp").Value, ";", I)
            If s = "U" & v_u_num Then
                s = "U" & new_u_num
            End If
            If InStr(lstresp, s & ";") = 0 Then
                lstresp = lstresp & s & ";"
            End If
        Next I
        If Odbc_Update("Application", _
                       "APP_Num", _
                       "where APP_Num=" & rs("APP_Num").Value, _
                       "APP_lstResp", lstresp) = P_ERREUR Then
            GoTo lab_erreur
        End If
        rs.MoveNext
    Wend
    rs.Close

lab_fin:
    P_RemplacerResponsableAppli = P_OK
    Exit Function

lab_erreur:
    P_RemplacerResponsableAppli = P_ERREUR

End Function

Public Function P_SelectionnerPersonne(ByVal v_titre As String, ByRef r_unum As Long, _
                            ByRef r_nom As String, ByRef r_prenom As String) As Integer
' ***************************************************************************
' Retourne le code P_PERSONNE_SELECTIONNEE_NON si personne n'est selectionnée
' ***************************************************************************
    Dim sql As String, choix As String
    Dim frm As Form

    Set frm = ChoixUtilisateur
    choix = ChoixUtilisateur.AppelFrm(v_titre, _
                                    "", _
                                    False, _
                                    False, _
                                    "")
    Set frm = Nothing
    If choix = "" Then
        If MsgBox(vbTab & "Vous n'avez selectionné aucune personne !" & vbTab & vbCrLf & vbCrLf _
                  & vbTab & "Etes-vous sûr de vouloir annuler cette opération ?" & vbTab, _
                  vbExclamation + vbYesNo, "Attention") = vbYes Then
            P_SelectionnerPersonne = P_ERREUR
            Exit Function
        Else
            P_SelectionnerPersonne = P_PERSONNE_SELECTIONNEE_NON
            Exit Function
        End If
    End If
    sql = "SELECT U_Nom, U_Prenom FROM Utilisateur WHERE U_kb_actif=True AND U_Actif=True AND U_Num=" & p_tblu_sel(0)
    If Odbc_RecupVal(sql, r_nom, r_prenom) = P_ERREUR Then
        r_unum = p_tblu_sel(0)
        P_SelectionnerPersonne = P_ERREUR
        Exit Function
    End If

    r_unum = p_tblu_sel(0)
    P_SelectionnerPersonne = P_OK

End Function

Public Function P_ValiderCode(ByVal v_numutil As Long, ByVal v_code As String, ByRef r_message) As Boolean
' *********************************************************************
' Verifier que le code a inserer/maj est unique POUR CHAQUE APPLICATION
' *********************************************************************
    Dim nom As String, prenom As String
    Dim numutil As Long
    Dim rs_code As rdoResultset

    If v_code = "" Then
        P_ValiderCode = P_OUI
        Exit Function
    End If
    
    ' limiter la longueur du code
    If Len(v_code) > 0 And Len(v_code) > 15 Then
        Call MsgBox("Le code " & v_code & " dépasse 15 caractères." & vbCrLf & vbCrLf & "Veuillez en saisir un autre.", vbInformation + vbOKOnly, "")
        P_ValiderCode = False
        Exit Function
    End If

    If Odbc_SelectV("SELECT UAPP_CODE, uapp_unum FROM UtilAppli" _
                 & " WHERE UAPP_APPNum = " & p_appli_kalidoc _
                 & " AND UAPP_Code = " & Odbc_String(UCase(v_code)) _
                 & " AND UAPP_UNum <> " & v_numutil, rs_code) = P_ERREUR Then
        P_ValiderCode = False
        Exit Function
    End If
    If Not rs_code.EOF Then
        numutil = rs_code("uapp_unum").Value
        rs_code.Close
        If Odbc_RecupVal("select u_nom, u_prenom from utilisateur where u_num=" & numutil, nom, prenom) = P_OK Then
            If p_traitement_background Then
                p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> Le code " & v_code & " est déjà attribué à " & nom & " " & prenom
            Else
                p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> Le code " & v_code & " est déjà attribué à " & nom & " " & prenom
                Call MsgBox("Le code " & v_code & " est déjà attribué à " & nom & " " & prenom & "." & vbCrLf & vbCrLf & "Veuillez en saisir un autre.", vbInformation + vbOKOnly, "")
            End If
            r_message = "==> Le code " & v_code & " est déjà attribué à " & nom & " " & prenom
        End If
        P_ValiderCode = False
        Exit Function
    Else
        rs_code.Close
    End If
    
    P_ValiderCode = True
    Exit Function

End Function

