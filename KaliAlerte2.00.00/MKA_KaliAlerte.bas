Attribute VB_Name = "MKA_KaliAlerte"
Public Const P_SUPER_UTIL = 1

Public p_stype_bdd As String
Public p_nom_bdd As String

Public p_numUtil As Long
Public p_nbaction As Integer
Public p_newaction As Boolean
Public p_nbdiff As Integer
Public p_newdiff As Boolean
Public p_nbinfo As Integer
Public p_newinfo As Boolean
Public p_nbdem As Integer
Public p_newdem As Boolean

Public p_salerte As String
Public p_cheminphp As String
Public p_nomini As String
Public p_chemin_appli As String
Public p_nomappli As String
Public p_slstaction As String
Public p_Result As Long

Private Function init_param_debug() As Integer
    
    p_chemin_appli = "c:\kalidoc"
    
    p_nomini = InputBox("Chemin du .ini : ", , "c:\kalidoc\kalidoc.ini")
    'p_nomini = InputBox("Chemin du .ini : ", , "c:\kalidoc\KaliDoc_VM_CHSTBRIEUC.ini")
    'p_nomini = InputBox("Chemin du .ini : ", , "c:\kalidoc\KaliDoc_CHMUL.ini")
    If p_nomini = "" Then
        init_param_debug = P_ERREUR
        Exit Function
    End If
    
    init_param_debug = P_OK
    
End Function

Public Function saisi_nomini() As Integer

    Call SAIS_Init
    Call SAIS_InitTitreHelp("Fichier .ini", "")
    Call SAIS_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call SAIS_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    If p_nomini <> "" Then
        Call SAIS_AddChamp("Chemin du fichier", 100, SAIS_TYP_TOUT_CAR, False, p_nomini)
    Else
        Call SAIS_AddChamp("Chemin du fichier", 100, SAIS_TYP_TOUT_CAR, False, "c:\kalidoc\kalidoc.ini")
    End If
    
    KS_Saisie.Show 1
    If SAIS_Saisie.retour = 1 Then
        saisi_nomini = P_NON
        Exit Function
    End If
    
    If SAIS_Saisie.champs(0).sval <> GetSetting(p_nomappli, "Section", "NomIni") Then
        p_nomini = SAIS_Saisie.champs(0).sval

        'Ajout dans la base de registre
        SaveSetting p_nomappli, "Section", "NomIni", p_nomini
        
        On Error Resume Next
        DeleteSetting p_nomappli, "Section", NumUtil
        DeleteSetting p_nomappli, "Section", Alerte
        On Error GoTo 0
        
        saisi_nomini = P_OUI
        
    Else
        saisi_nomini = P_NON
    
    End If
    

End Function

Private Sub Main()
    
    Dim scmd As String
    
    ' Param de l'application
    scmd = Command$
    
    ' Mode DEBUG
    'MsgBox "scmd=" & scmd
    p_nomappli = App.EXEName

    'scmd = "c:\kalidoc\kalialerte.ini"

    If scmd = "DEBUG" Then
        direct = False
        If init_param_debug() = P_ERREUR Then
            End
        End If
    Else
        p_chemin_appli = App.Path
        'MsgBox "p_chemin_appli=" & p_chemin_appli
lab_init:
        'If scmd <> "" Then
        '    p_nomini = scmd
        '    pos = InStrRev(p_nomini, "\")
        '    If pos > 0 Then
        '        s = Mid$(scmd, pos + 1)
        '        pos = InStrRev(s, ".")
        '        If pos > 0 Then
        '            p_nomappli = Left$(s, pos - 1)
        '        End If
        '    End If
        'End If
        p_nomini = scmd
        If p_nomini = "" Then
            MsgBox "Le nom du fichier Ini est vide"
            End
        ElseIf Not FICH_FichierExiste(p_nomini) Then
            MsgBox "Le fichier " & p_nomini & " est introuvable"
            If p_nomappli = "" Then
                p_nomini = GetSetting(App.EXEName, "Section", "NomIni")
                'MsgBox "p_nomini=" & p_nomini
                If p_nomini = "" Then
                    If saisi_nomini() = P_NON Then
                        End
                    End If
                    scmd = p_nomini
                    GoTo lab_init
                Else
                    p_nomappli = App.EXEName
                End If
            End If
        End If
    End If
    
    'MsgBox "p_nomappli=" & p_nomappli
    ask_enreg = False

    ' Type de base
    p_stype_bdd = SYS_GetIni("BASE", "TYPE", p_nomini)
    If p_stype_bdd = "" Then
lab_sais_typb:
        p_stype_bdd = InputBox("Type de base (PG, MDB) : ", , "PG")
        If p_stype_bdd = "" Then
            Exit Sub
        End If
        If p_stype_bdd <> "MDB" And p_stype_bdd <> "PG" Then
            GoTo lab_sais_typb
        End If
        ask_enreg = True
    End If
    ' Nom base
    p_nom_bdd = SYS_GetIni("BASE", "NOM", p_nomini)
    If p_nom_bdd = "" Then
        p_nom_bdd = InputBox("Nom de la base : ", , "kalidoc")
        If p_nom_bdd = "" Then
            Exit Sub
        End If
        ask_enreg = True
    End If
    ' Enregistrement des infos base
    If ask_enreg Then
        reponse = MsgBox("Voulez-vous enregistrer les informations saisies ?", vbQuestion + vbYesNo, "")
        If reponse = vbYes Then
            Call SYS_PutIni("BASE", "TYPE", p_stype_bdd, p_chemin_appli & "\kalidoc.ini")
            Call SYS_PutIni("BASE", "NOM", p_nom_bdd, p_chemin_appli & "\kalidoc.ini")
        End If
    End If

    ' Connexion à la base
    If Odbc_Init(p_stype_bdd, p_nom_bdd, False) = P_ERREUR Then
        Exit Sub
    End If
    
    ' Initialisation du chemin pour le serveur
    Call Odbc_RecupVal("SELECT PG_CheminPHP FROM PRMGen_http", p_cheminphp)
    
    KA_PrmAlerte.Show
    
End Sub
