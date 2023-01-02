Attribute VB_Name = "Module1"
Option Explicit
Global p_nomini As String
Global p_traitement_background As String

Global p_nom_fichier_agent As String
Global p_nom_fichier_structure As String
Global p_nom_fichier_grade As String
Global p_nom_fichier_OUT As String
Global p_nom_fichier_LOG As String
Global p_nom_fichier_LOCK As String

Global pos_code_UF As String
Global pos_libelle_UF As String
Global pos_code_grade As String
Global pos_libelle_grade As String

Global fd_agent As Integer
Global fd_structure As Integer
Global fd_grade As Integer
Global fd_OUT As Integer
Global fd_LOG As Integer
Global fd_LOCK As Integer

Global p_type_base As String
Global p_nom_base As String
Global p_NumUtil As Long
Global p_chemin_appli As String

Global NOM As String
Global PRENOM As String
Global MATRICULE As String
Global CODE_SECTION As String
Global LIB_SECTION As String
Global CODE_FONCTION As String
Global LIB_FONCTION As String
Global LIB_CIVILITE As String
Global NJF As String


Sub Main()
    Dim ligne_lu As String, ligne_OUT As String, Message As String
    Dim pos_nom As Integer, lon_nom As Integer
    Dim pos_prenom As Integer, lon_prenom As Integer
    Dim pos_matricule As Integer, lon_matricule As Integer
    Dim pos_present As Integer, lon_present As Integer
    Dim pos_code_section As Integer, lon_code_section As Integer
    Dim pos_code_fonction As Integer, lon_code_fonction As Integer
    Dim pos_civilite As Integer, lon_civilite As Integer
    Dim pos_NJF As Integer, lon_NJF As Integer
    Dim nbligIN As Integer
    Dim yaErreur As Boolean
    Dim liberr As String
    Dim i As Integer
    Dim iUF As Integer
    Dim nbUF As Integer
    Dim Spos As String, Slon As String
    Dim s As String, IN_MATRICULE As String, IN_NOM As String, IN_PRENOM As String
    Dim IN_CODE_SECTION As String, IN_LIB_SECTION As String, IN_CODE_FONCTION As String, IN_LIB_FONCTION As String
    Dim IN_PRESENT As String
    Dim IN_CIVILITE As String, IN_NJF As String
    Dim sql As String, rs As rdoResultset
    Dim deb As String, lon As String
    Dim pos_NBUFSEC As String, lon_NBUFSEC As String
    Dim nb_UfSec As String, str_UfSec As String
    Dim TbParamUFSEC()
    Dim TbDatafUFSEC()

    ' Voir si les fichiers existent
    p_chemin_appli = "c:\kalidoc"
    p_nomini = Command$
p_nomini = "c:\kalidoc\Kalibottin_GAP.ini"
    If Not FICH_FichierExiste(p_nomini) Then
        MsgBox p_nomini & " n'existe pas"
        End
    End If
    p_nom_fichier_agent = SYS_GetIni("PREIMPORT_AGENTS", "Fichier", p_nomini)

    ' Connexion à la base
    p_type_base = SYS_GetIni("BASE", "TYPE", p_nomini)
    p_nom_base = SYS_GetIni("BASE", "NOM", p_nomini)
    If Odbc_Init(p_type_base, p_nom_base) = P_ERREUR Then
        End
    End If
    ' position des champs de sortie
    sql = "SELECT PGB_FichLstPos FROM PrmGenB"
    If Odbc_Select(sql, rs) = P_ERREUR Then
        MsgBox "Erreur sur " & sql
        End
    Else
        s = rs("PGB_FichLstPos")
        NOM = STR_GetChamp(STR_GetChamp(s, ";", 0), "=", 1)
        PRENOM = STR_GetChamp(STR_GetChamp(s, ";", 1), "=", 1)
        MATRICULE = STR_GetChamp(STR_GetChamp(s, ";", 2), "=", 1)
        CODE_SECTION = STR_GetChamp(STR_GetChamp(s, ";", 3), "=", 1)
        LIB_SECTION = STR_GetChamp(STR_GetChamp(s, ";", 4), "=", 1)
        CODE_FONCTION = STR_GetChamp(STR_GetChamp(s, ";", 5), "=", 1)
        LIB_FONCTION = STR_GetChamp(STR_GetChamp(s, ";", 6), "=", 1)
        LIB_CIVILITE = STR_GetChamp(STR_GetChamp(s, ";", 7), "=", 1)
        NJF = STR_GetChamp(STR_GetChamp(s, ";", 8), "=", 1)
    End If
    p_NumUtil = 0
    ' vérification des fichiers
    If HTTP_Appel_fichier_existe(p_nom_fichier_agent, False, liberr) = HTTP_OK Then
    End If
    If FICH_FichierExiste(p_nom_fichier_agent) Then
        If FICH_OuvrirFichier(p_nom_fichier_agent, FICH_LECTURE, fd_agent) = P_ERREUR Then
            MsgBox "Impossible d'ouvrir " & p_nom_fichier_agent
            End
        Else
            ' fichier de structure
            p_nom_fichier_structure = SYS_GetIni("PREIMPORT_STRUCTURE", "Fichier", p_nomini)
            If FICH_OuvrirFichier(p_nom_fichier_structure, FICH_LECTURE, fd_structure) = P_ERREUR Then
                MsgBox "Impossible d'ouvrir " & p_nom_fichier_structure
                Close #fd_agent
                End
            Else
                pos_code_UF = SYS_GetIni("PREIMPORT_STRUCTURE", "code_UF", p_nomini)
                pos_libelle_UF = SYS_GetIni("PREIMPORT_STRUCTURE", "libelle_UF", p_nomini)
                Close #fd_structure
            End If
            ' fichier des grades
            p_nom_fichier_grade = SYS_GetIni("PREIMPORT_GRADE", "Fichier", p_nomini)
            If FICH_OuvrirFichier(p_nom_fichier_grade, FICH_LECTURE, fd_structure) = P_ERREUR Then
                MsgBox "Impossible d'ouvrir " & p_nom_fichier_grade
                Close #fd_grade
                End
            Else
                pos_code_grade = SYS_GetIni("PREIMPORT_GRADE", "code_grade", p_nomini)
                pos_libelle_grade = SYS_GetIni("PREIMPORT_GRADE", "libelle_grade", p_nomini)
                Close #fd_grade
            End If
            ' fichier de sortie
            p_nom_fichier_OUT = SYS_GetIni("PREIMPORT_SORTIE", "Fichier", p_nomini)
            If FICH_FichierExiste(p_nom_fichier_OUT) Then
                Call FICH_EffacerFichier(p_nom_fichier_OUT, True)
            End If
            If FICH_OuvrirFichier(p_nom_fichier_OUT, FICH_ECRITURE, fd_OUT) = P_ERREUR Then
                MsgBox "Impossible d'ouvrir " & p_nom_fichier_OUT
                Close #fd_agent
                End
            End If
            ' fichier de LOG
            p_nom_fichier_LOG = SYS_GetIni("PREIMPORT", "FICHIER_LOG", p_nomini)
            If FICH_FichierExiste(p_nom_fichier_LOG) Then
                Call FICH_EffacerFichier(p_nom_fichier_LOG, True)
            End If
            If FICH_OuvrirFichier(p_nom_fichier_LOG, FICH_ECRITURE, fd_LOG) = P_ERREUR Then
                MsgBox "Impossible d'ouvrir " & p_nom_fichier_LOG
                Close #fd_OUT
                Close #fd_agent
                End
            End If
            ' fichier de LOCK
            p_nom_fichier_LOCK = SYS_GetIni("PREIMPORT", "FICHIER_LOCK", p_nomini)
            If Not FICH_FichierExiste(p_nom_fichier_LOCK) Then
                MsgBox "Le fichier " & p_nom_fichier_LOCK & " n'est pas présent"
                'Close #fd_OUT
                'Close #fd_agent
                'End
            End If
            ' lecture des positions des données
            s = SYS_GetIni("PREIMPORT_AGENTS", "nom", p_nomini)
            pos_nom = STR_GetChamp(s, ":", 0)
            lon_nom = STR_GetChamp(s, ":", 1)
            
            s = SYS_GetIni("PREIMPORT_AGENTS", "prenom", p_nomini)
            pos_prenom = STR_GetChamp(s, ":", 0)
            lon_prenom = STR_GetChamp(s, ":", 1)
            
            s = SYS_GetIni("PREIMPORT_AGENTS", "matricule", p_nomini)
            pos_matricule = STR_GetChamp(s, ":", 0)
            lon_matricule = STR_GetChamp(s, ":", 1)
            
            s = SYS_GetIni("PREIMPORT_AGENTS", "code_section", p_nomini)
            pos_code_section = STR_GetChamp(s, ":", 0)
            lon_code_section = STR_GetChamp(s, ":", 1)
            
            s = SYS_GetIni("PREIMPORT_AGENTS", "code_fonction", p_nomini)
            pos_code_fonction = STR_GetChamp(s, ":", 0)
            lon_code_fonction = STR_GetChamp(s, ":", 1)
            
            s = SYS_GetIni("PREIMPORT_AGENTS", "civilite", p_nomini)
            pos_civilite = STR_GetChamp(s, ":", 0)
            lon_civilite = STR_GetChamp(s, ":", 1)
            
            s = SYS_GetIni("PREIMPORT_AGENTS", "NJF", p_nomini)
            pos_NJF = STR_GetChamp(s, ":", 0)
            lon_NJF = STR_GetChamp(s, ":", 1)
            
            s = SYS_GetIni("PREIMPORT_AGENTS", "PRESENT", p_nomini)
            pos_present = STR_GetChamp(s, ":", 0)
            lon_present = STR_GetChamp(s, ":", 1)
            
            ' UF Secondaires
            s = SYS_GetIni("PREIMPORT_AGENTS", "NBUFSEC", p_nomini)
            pos_NBUFSEC = STR_GetChamp(s, ":", 0)
            lon_NBUFSEC = STR_GetChamp(s, ":", 1)
            nbUF = 0
            For i = 1 To 10
                s = SYS_GetIni("PREIMPORT_AGENTS", "UFSEC" & i, p_nomini)
                If s <> "" Then
                    Spos = STR_GetChamp(s, ":", 0)
                    Slon = STR_GetChamp(s, ":", 1)
                    ReDim Preserve TbParamUFSEC(2, nbUF)
                    TbParamUFSEC(1, nbUF) = Spos
                    TbParamUFSEC(2, nbUF) = Slon
                    nbUF = nbUF + 1
                End If
            Next i

            ' lire le fichier
            Ajouter_Erreur ("Debut le " & Date & " à " & Time)
            While Not EOF(fd_agent)
                Line Input #fd_agent, ligne_lu
                nbligIN = nbligIN + 1
                yaErreur = False
                Ajouter_Erreur ("****************  Ligne num. " & nbligIN)
                ' IN_nom
                IN_NOM = Mid(ligne_lu, pos_nom, lon_nom)
                ' IN_prenom
                IN_PRENOM = Mid(ligne_lu, pos_prenom, lon_prenom)
                ' IN_matricule
                IN_MATRICULE = Mid(ligne_lu, pos_matricule, lon_matricule)
                ' IN_CODE_SECTION
                IN_CODE_SECTION = Mid(ligne_lu, pos_code_section, lon_code_section)
                ' IN_CODE_FONCTION
                IN_CODE_FONCTION = Mid(ligne_lu, pos_code_fonction, lon_code_fonction)
                ' IN_Civilité
                IN_CIVILITE = Mid(ligne_lu, pos_civilite, lon_civilite)
                ' IN_NJF
                IN_NJF = Mid(ligne_lu, pos_NJF, lon_NJF)
                ' IN_PRESENT
                IN_PRESENT = Trim(Mid(ligne_lu, pos_present, lon_present))
                
                If IN_PRESENT = "" Then
                    ' Chercher le grade
                    IN_LIB_FONCTION = Chercher_Grade(IN_CODE_FONCTION, Message)
                    If IN_LIB_FONCTION = "ERREUR" Then
                        Ajouter_Erreur ("ERREUR Chercher le grade : " & Message)
                        IN_LIB_FONCTION = "ERREUR : " & Message
                        yaErreur = True
                    End If
                    
                    ' Chercher l'UF
                    IN_LIB_SECTION = Chercher_UF(IN_CODE_SECTION, Message)
                    If IN_LIB_SECTION = "ERREUR" Then
                        Ajouter_Erreur ("ERREUR Chercher l'UF : " & Message)
                        IN_LIB_SECTION = "ERREUR : " & Message
                        yaErreur = True
                    End If
                    
                    ' constituer la ligne de sortie
                    If Not yaErreur Then
                        ligne_OUT = Space(1500)
                        Mid(ligne_OUT, STR_GetChamp(NOM, ":", 0), STR_GetChamp(NOM, ":", 1)) = IN_NOM
                        Mid(ligne_OUT, STR_GetChamp(PRENOM, ":", 0), STR_GetChamp(PRENOM, ":", 1)) = IN_PRENOM
                        Mid(ligne_OUT, STR_GetChamp(MATRICULE, ":", 0), STR_GetChamp(MATRICULE, ":", 1)) = IN_MATRICULE
                        Mid(ligne_OUT, STR_GetChamp(CODE_SECTION, ":", 0), STR_GetChamp(CODE_SECTION, ":", 1)) = IN_CODE_SECTION
                        Mid(ligne_OUT, STR_GetChamp(LIB_SECTION, ":", 0), STR_GetChamp(LIB_SECTION, ":", 1)) = IN_LIB_SECTION
                        Mid(ligne_OUT, STR_GetChamp(CODE_FONCTION, ":", 0), STR_GetChamp(CODE_FONCTION, ":", 1)) = IN_CODE_FONCTION
                        Mid(ligne_OUT, STR_GetChamp(LIB_FONCTION, ":", 0), STR_GetChamp(LIB_FONCTION, ":", 1)) = IN_LIB_FONCTION
                        Mid(ligne_OUT, STR_GetChamp(LIB_CIVILITE, ":", 0), STR_GetChamp(LIB_CIVILITE, ":", 1)) = IN_CIVILITE
                        Mid(ligne_OUT, STR_GetChamp(NJF, ":", 0), STR_GetChamp(NJF, ":", 1)) = IN_NJF
                        ' ecrire la sortie
                        Print #fd_OUT, ligne_OUT
                    End If
                    'Voir si UF secondaires
                    nb_UfSec = Trim(Mid(ligne_lu, pos_NBUFSEC, lon_NBUFSEC))
                    If nb_UfSec > 0 Then
                        For i = 0 To UBound(TbParamUFSEC, 2)
                            str_UfSec = Mid(ligne_lu, TbParamUFSEC(1, i), TbParamUFSEC(2, i))
                            If Trim(str_UfSec) <> "" Then
                                ' Chercher l'UF
                                IN_LIB_SECTION = Chercher_UF(str_UfSec, Message)
                                If IN_LIB_SECTION = "ERREUR" Then
                                    Ajouter_Erreur ("ERREUR UF secondaires n° " & i & " : " & Message)
                                    IN_LIB_SECTION = "ERREUR : " & Message
                                    yaErreur = True
                                End If
                                
                                ' constituer la ligne de sortie
                                If Not yaErreur Then
                                    ligne_OUT = Space(1500)
                                    Mid(ligne_OUT, STR_GetChamp(NOM, ":", 0), STR_GetChamp(NOM, ":", 1)) = IN_NOM
                                    Mid(ligne_OUT, STR_GetChamp(PRENOM, ":", 0), STR_GetChamp(PRENOM, ":", 1)) = IN_PRENOM
                                    Mid(ligne_OUT, STR_GetChamp(MATRICULE, ":", 0), STR_GetChamp(MATRICULE, ":", 1)) = IN_MATRICULE
                                    Mid(ligne_OUT, STR_GetChamp(CODE_SECTION, ":", 0), STR_GetChamp(CODE_SECTION, ":", 1)) = str_UfSec
                                    Mid(ligne_OUT, STR_GetChamp(LIB_SECTION, ":", 0), STR_GetChamp(LIB_SECTION, ":", 1)) = IN_LIB_SECTION
                                    Mid(ligne_OUT, STR_GetChamp(CODE_FONCTION, ":", 0), STR_GetChamp(CODE_FONCTION, ":", 1)) = IN_CODE_FONCTION
                                    Mid(ligne_OUT, STR_GetChamp(LIB_FONCTION, ":", 0), STR_GetChamp(LIB_FONCTION, ":", 1)) = IN_LIB_FONCTION
                                    Mid(ligne_OUT, STR_GetChamp(LIB_CIVILITE, ":", 0), STR_GetChamp(LIB_CIVILITE, ":", 1)) = IN_CIVILITE
                                    Mid(ligne_OUT, STR_GetChamp(NJF, ":", 0), STR_GetChamp(NJF, ":", 1)) = IN_NJF
                                    ' ecrire la sortie
                                    Print #fd_OUT, ligne_OUT
                                End If
                            End If
                        Next i
                    End If
                Else
                    Ajouter_Erreur ("****************  IN_PRESENT= " & IN_PRESENT & " pos_present=" & pos_present & " " & lon_present)
                End If
            Wend
            Ajouter_Erreur ("Fin le " & Date & " à " & Time)
            Close #fd_agent
            Close #fd_OUT
            Close #fd_LOG
            ' Effacer le fichier de lock
            Call FICH_EffacerFichier(p_nom_fichier_LOCK, False)
        End If
    End If

    End
End Sub

Private Function Ajouter_Erreur(v_Message)
    'MsgBox "Ajouter_Erreur(" & v_Message & ")"
    Print #fd_LOG, v_Message
End Function

Private Function Chercher_Grade(v_code_fonction, ByRef r_Message)
    Dim ligne_lu As String
    Dim nb As Integer
    Dim fini As Boolean
    
    If FICH_OuvrirFichier(p_nom_fichier_grade, FICH_LECTURE, fd_grade) = P_ERREUR Then
        r_Message = "Impossible d'ouvrir " & p_nom_fichier_grade
        Chercher_Grade = "ERREUR"
    Else
        pos_code_grade = SYS_GetIni("PREIMPORT_GRADE", "code_grade", p_nomini)
        pos_libelle_grade = SYS_GetIni("PREIMPORT_GRADE", "libelle_grade", p_nomini)
        fini = False
        ' lire le fichier
        While Not fini
            If EOF(fd_grade) Then
                fini = True
            Else
                Line Input #fd_grade, ligne_lu
                If Mid(ligne_lu, STR_GetChamp(pos_code_grade, ":", 0), STR_GetChamp(pos_code_grade, ":", 1)) = v_code_fonction Then
                    nb = nb + 1
                    fini = True
                    Chercher_Grade = Mid(ligne_lu, STR_GetChamp(pos_libelle_grade, ":", 0), STR_GetChamp(pos_libelle_grade, ":", 1))
                End If
            End If
        Wend
        Close #fd_grade
        If nb = 0 Then
            r_Message = "pas de grade : " & v_code_fonction
            Chercher_Grade = "ERREUR"
        ElseIf nb > 1 Then
            r_Message = "plusieurs grade : " & v_code_fonction
            Chercher_Grade = "ERREUR"
        End If
    End If
End Function

Private Function Chercher_UF(v_code_section, ByRef r_Message)
    Dim ligne_lu As String
    Dim nb As Integer
    Dim fini As Boolean
    
    If FICH_OuvrirFichier(p_nom_fichier_structure, FICH_LECTURE, fd_structure) = P_ERREUR Then
        r_Message = "Impossible d'ouvrir " & p_nom_fichier_structure
        Chercher_UF = "ERREUR"
    Else
        pos_code_UF = SYS_GetIni("PREIMPORT_STRUCTURE", "code_UF", p_nomini)
        pos_libelle_UF = SYS_GetIni("PREIMPORT_STRUCTURE", "libelle_UF", p_nomini)
        fini = False
        ' lire le fichier
        While Not fini
            If EOF(fd_structure) Then
                fini = True
            Else
                Line Input #fd_structure, ligne_lu
                If Mid(ligne_lu, STR_GetChamp(pos_code_UF, ":", 0), STR_GetChamp(pos_code_UF, ":", 1)) = v_code_section Then
                    nb = nb + 1
                    fini = True
                    Chercher_UF = Mid(ligne_lu, STR_GetChamp(pos_libelle_UF, ":", 0), STR_GetChamp(pos_libelle_UF, ":", 1))
                End If
            End If
        Wend
        Close #fd_structure
        If nb = 0 Then
            r_Message = "pas d'UF : " & v_code_section
            Chercher_UF = "ERREUR"
        ElseIf nb > 1 Then
            r_Message = "plusieurs UF : " & v_code_section
            Chercher_UF = "ERREUR"
        End If
    End If
End Function

