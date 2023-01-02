VERSION 5.00
Begin VB.Form ChoixPrmPers 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Choix d'une personne"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2865
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6945
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   2
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rechercher par "
         ForeColor       =   &H00800080&
         Height          =   1095
         Left            =   450
         TabIndex        =   11
         Top             =   1620
         Width           =   6045
         Begin VB.CommandButton cmd 
            Caption         =   "&Application"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   2
            Left            =   2640
            TabIndex        =   4
            ToolTipText     =   "Lister les applications"
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   5
            Left            =   510
            Picture         =   "ChoixPrmPers.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Rechercher par fonction"
            Top             =   360
            Width           =   660
         End
         Begin VB.CommandButton cmd 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   4
            Left            =   4860
            Picture         =   "ChoixPrmPers.frx":0463
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Rechercher par service"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   750
         End
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   2040
         MaxLength       =   40
         TabIndex        =   1
         Top             =   510
         Width           =   2655
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Liste complète"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   3
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   390
         Width           =   1020
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Matricule"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   825
      Left            =   0
      TabIndex        =   0
      Top             =   2730
      Width           =   6945
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   5970
         Picture         =   "ChoixPrmPers.frx":09F2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Quitter"
         Top             =   220
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   0
         Left            =   480
         Picture         =   "ChoixPrmPers.frx":0FAB
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Sélectionner"
         Top             =   220
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "ChoixPrmPers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index sur les objets cmd
Private Const CMD_OK = 0
Private Const CMD_CHOIX_APPLICATION = 2
Private Const CMD_FERMER = 1
Private Const CMD_CHOIX_UTIL = 3
Private Const CMD_CHOIX_FONCTION = 5
Private Const CMD_CHOIX_SERVICE = 4

Private Const TXT_NOM = 0
Private Const TXT_MATRICULE = 1

Private g_crutil_autor As Boolean

Private g_mode_saisie As Boolean
Private g_txt_nom_avant As String
Private g_txt_matricule_avant As String
Private g_form_active As Boolean

Private Function build_sql_ft(ByVal v_schoix As String, _
                               ByVal v_fordre_norm As Boolean) As String

    Dim sDest As String, sql As String
    Dim s As String, s2 As String, slst As String
    Dim n As Integer, I As Integer
    Dim rs As rdoResultset

    sDest = ""
    If left$(v_schoix, 1) = "F" Then
        sDest = sDest + v_schoix
    End If
    sql = "SELECT * FROM Utilisateur, UtilAppli" _
            & " WHERE U_kb_actif=True AND U_Num>" & P_SUPER_UTIL _
            & " AND UAPP_UNum=U_Num"
    If sDest <> "" Then
        n = STR_GetNbchamp(sDest, "|")
        For I = 0 To n - 1
            s = STR_GetChamp(sDest, "|", I)
            If I = 0 Then
                sql = sql & " AND ("
            Else
                sql = sql & " OR"
            End If
            Select Case left$(s, 1)
            Case "F"
                sql = sql & " U_FctTrav LIKE '%" & s & ";%'"
            Case "S"
                sql = sql & " U_SPM LIKE '%" & s & "%'"
            Case "U"
                sql = sql & " U_Num=" & Mid$(s, 2)
            End Select
        Next I
        sql = sql & ")"
    End If
    If v_fordre_norm Then
        sql = sql & " ORDER BY U_Nom, U_Prenom"
    Else
        sql = sql & " ORDER BY U_Nom DESC, U_Prenom DESC"
    End If

    build_sql_ft = sql

End Function

Private Function build_titre(ByVal v_schoix As String) As String

    Dim stitre As String, libfct As String, nomgrp As String

    stitre = "Liste des personnes"
    If left$(v_schoix, 1) = "F" Then
        Call Odbc_RecupVal("SELECT FT_Libelle FROM FctTrav WHERE FT_Num=" & Mid$(v_schoix, 2), _
                           libfct)
        stitre = stitre + " 'Fonction:" + libfct + " '"
    End If
    build_titre = stitre

End Function

Private Sub choisir_application()

    Dim sql As String, postes_con As String, poste_en_cours As String, _
        str As String, chaine_en_cours As String
    Dim str_concat As Boolean
    Dim n As Integer, I As Integer, nbr_poste_conc As Integer
    Dim rs As rdoResultset

    Call CL_Init
    n = 0
    sql = "SELECT * FROM Application ORDER BY App_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    While Not rs.EOF
        Call CL_AddLigne(rs("App_Nom").Value, rs("App_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    If n = 0 Then
        Exit Sub
    End If

    Call CL_InitTitreHelp("Les applications existantes", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -15)
lab_afficher: ' pour réafficher en cas ou il n'y a pas de postes concernés
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then ' QUITTER
        Exit Sub
    End If
    ' on a choisi une application
    sql = "SELECT APP_Profil_Conc FROM Application WHERE App_Num=" & CL_liste.lignes(CL_liste.pointeur).num
    If Odbc_RecupVal(sql, postes_con) = P_ERREUR Then
        Exit Sub
    End If
    nbr_poste_conc = STR_GetNbchamp(postes_con, "|")
    If nbr_poste_conc = 0 Then
        Call MsgBox("Aucune personne n'a les critères indiqués.", vbInformation + vbOKOnly, "")
        GoTo lab_afficher
    End If
    str_concat = False
    For I = 0 To nbr_poste_conc - 1
        chaine_en_cours = STR_GetChamp(postes_con, "|", I)
        poste_en_cours = STR_GetChamp(chaine_en_cours, ";", STR_GetNbchamp(chaine_en_cours, ";") - 1)
        If Mid$(chaine_en_cours, 1, 1) = "S" Then     ' un service
            If str_concat Then
                str = str & " OR U_SPM LIKE '%" & poste_en_cours & ";%' "
            Else
                 str = " AND ( U_SPM LIKE '%" & poste_en_cours & ";%' "
                 str_concat = True
            End If
        ElseIf Mid$(chaine_en_cours, 1, 1) = "P" Then ' un poste
            If str_concat Then
                str = str & " OR U_SPM LIKE '%" & poste_en_cours & ";%' "
            Else
                 str = " AND ( U_SPM LIKE '%" & poste_en_cours & ";%' "
                 str_concat = True
            End If
        Else ' "L" => afficher tout le monde
            str = ""
        End If
    Next I
    ' recherche des personnes
    sql = "SELECT * FROM Utilisateur, UtilAppli" _
        & " WHERE U_kb_actif=True AND U_Num>" & P_SUPER_UTIL & " AND UAPP_UNum=U_Num" _
        & IIf(str = "", "", str & ")")

    Call choisir_utilisateur("A" & sql)

End Sub

Private Sub choisir_fonction()

    Dim sql As String
    Dim n As Integer, I As Integer
    Dim rs As rdoResultset

    Call CL_Init
    n = 0
    sql = "SELECT * FROM FctTrav ORDER BY FT_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    While Not rs.EOF
        Call CL_AddLigne(rs("FT_Libelle").Value, rs("FT_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close
    If n = 0 Then
        Exit Sub
    End If

    Call CL_InitTitreHelp("Fonctions du personnel", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -15)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Exit Sub
    End If

    Call choisir_utilisateur("F" & CL_liste.lignes(CL_liste.pointeur).num)

End Sub

Private Sub choisir_service()

    Dim sret As String, s_serv As String, sprm As String, nomserv As String
    Dim encore As Boolean
    Dim n As Integer
    Dim numutil As Long
    Dim frm As Form

    Call CL_Init
    Set frm = PrmService
    sret = PrmService.AppelFrm("Choix d'un service", "S", False, "", "SL", False)
    Set frm = Nothing
    If sret = "" Then
        Exit Sub
    End If

    n = STR_GetNbchamp(sret, ";")
    s_serv = STR_GetChamp(sret, ";", n - 1)
    If left$(s_serv, 1) = "L" Then
        Call P_RecupLaboCode(Mid$(s_serv, 2), nomserv)
    Else
        Call P_RecupSrvNom(Mid$(s_serv, 2), nomserv)
    End If
    encore = True
    Do
        Set frm = PrmService
        sret = PrmService.AppelFrm("Personnes rattachées au service : " & nomserv, "P", False, "", s_serv, True)
        Set frm = Nothing
        If sret = "" Then
            encore = False
        Else
            Set frm = PrmPersonne
            numutil = STR_GetChamp(sret, "|", 0)
            If numutil = 0 Then
                sprm = "POSTE=" & Mid$(STR_GetChamp(sret, "|", 1), 2)
            Else
                sprm = ""
            End If
            Call PrmPersonne.AppelFrm(numutil, sprm)
            Set frm = Nothing
        End If
    Loop Until encore = False

    txt(TXT_NOM).Text = ""
    txt(TXT_NOM).SetFocus

End Sub

Private Sub choisir_utilisateur(ByVal v_schoix As String)

    Dim sql As String, stext As String, s As String, lib As String
    Dim stitre As String, libfct As String, sprm As String
    Dim n As Integer, lig_crt As Integer
    Dim numutil As Long, numappli As Long
    Dim rs As rdoResultset
    Dim frm As Form

Lab_Debut:
    If left$(v_schoix, 1) = "U" Then
        numutil = Mid$(STR_GetChamp(v_schoix, "|", 0), 2)
        sprm = STR_GetChamp(v_schoix, "|", 1)
        GoTo lab_prm
    ElseIf left$(v_schoix, 1) = "A" Then
        sql = Mid$(v_schoix, 2)
        GoTo lab_appli ' sql est déjà construite, passer directement à l'affichage
    End If

    sql = build_sql_ft(v_schoix, True)
lab_appli:
    Call FRM_ResizeForm(Me, 0, 0)
    Call CL_Init

    'Choix de l'utilisateur
    n = 0
    lig_crt = 0
    If g_crutil_autor And v_schoix = "0" Then
        Call CL_AddLigne("<Nouveau>", 0, "", False)
        n = 1
    End If
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    While Not rs.EOF
        stext = rs("U_Matricule").Value & vbTab _
                & IIf(rs("U_Actif").Value, "[   ]", "[ Inactive ]") & " " _
                & IIf(rs("U_ExterneFich").Value, "[Externe au fichier]", "[   ]") & vbTab _
                & rs("U_Nom").Value & " " & rs("U_Prenom").Value & vbTab
        If rs("U_SPM").Value <> "" Then
            's = STR_GetChamp(rs("U_SPM").Value, "|", 0)
            'If P_RecupSPLib(s, lib) = P_ERREUR Then
            '    Exit Sub
            'End If
            'stext = stext & vbTab & lib
            stext = stext & vbTab & P_get_lib_srv_poste(P_get_num_srv_poste(rs("U_SPM").Value, P_POSTE), P_POSTE) _
                    & " - " & P_get_lib_srv_poste(P_get_num_srv_poste(rs("U_SPM").Value, P_SERVICE), P_SERVICE)
        End If
        Call CL_AddLigne(stext, _
                         rs("U_Num").Value, _
                         "", _
                         False)
        If numutil = rs("U_Num").Value Then
            lig_crt = n
        End If
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close

    If n = 0 Then
        Call MsgBox("Aucune personne n'a les critères indiqués.", vbInformation + vbOKOnly, "")
        GoTo lab_init
    End If
    stitre = build_titre(v_schoix)
    Call CL_InitTitreHelp(stitre, _
                          p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_f_utilisateur.htm")
    Call CL_InitTaille(0, -20)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnimprimer.gif", vbKeyI, vbKeyF3, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    If lig_crt > 0 Then
        Call CL_InitPointeur(lig_crt)
    End If
    ChoixListe.Show 1
    ' Quitter
    If CL_liste.retour = 2 Then
        GoTo lab_init
    End If
    ' Imprimer
    If CL_liste.retour = 1 Then
        Call imprimer(v_schoix)
        GoTo Lab_Debut
    End If
    numutil = CL_liste.lignes(CL_liste.pointeur).num

lab_prm:
    Set frm = PrmPersonne
    Call PrmPersonne.AppelFrm(numutil, sprm)
    Set frm = Nothing
    If left$(v_schoix, 1) = "U" Then
        GoTo lab_init
    Else
        GoTo Lab_Debut
    End If

lab_init:
    Call FRM_ResizeForm(Me, Me.width, Me.Height)
    txt(TXT_NOM).Text = ""
    txt(TXT_MATRICULE).Text = ""
    On Error Resume Next
    txt(TXT_NOM).SetFocus
    On Error GoTo 0
End Sub

Private Sub imprimer(ByVal v_schoix As String)

    Dim modele As String, sql As String, nomdata As String, s As String
    Dim code As String, lib As String, s_sp As String, sadr As String
    Dim libfct As String, codlabo As String, stitre As String
    Dim fl_fax As Boolean, fl_bid As Boolean, trouve As Boolean
    Dim n As Integer, ilabo As Integer, fd As Integer, I As Integer, ns As Integer
    Dim np As Integer, j As Integer, ndecal As Integer, n2 As Integer
    Dim m As Integer
    Dim numl As Long, numf As Long, numutil As Long, numzone As Long
    Dim tbl_s() As Long, tbl_p() As Long, tbl_ps() As Long
    Dim nums As Long, nump As Long
    Dim tbl_numfct() As Long, tbl_numfctp() As Long, numpere As Long, tbl_fcttr() As Long
    Dim ligne As Variant
    Dim rs As rdoResultset, rs2 As rdoResultset

'    If P_ChoixModele(p_chemin_modele + "\Personne", _
'                     modele) <> P_OUI Then
'        Exit Sub
'    End If

    ' Choix de l'imprimante
    If PR_ChoixImp(False, False, fl_fax, fl_bid) = False Then Exit Sub
        
    nomdata = p_chemin_appli & "\tmp\" & p_CodeUtil & format(Time, "hhmmss") & ".txt"
'    If FICH_CopierFichier(p_chemin_modele + "\Personne\entete.txt", _
'                          nomdata) = P_ERREUR Then
'        Exit Sub
'    End If
    If FICH_OuvrirFichier(nomdata, FICH_ECRITURE, fd) = P_ERREUR Then
        Exit Sub
    End If

    ' N° de la zone Adrmail
    If Odbc_RecupVal("SELECT ZU_Num FROM ZoneUtil WHERE ZU_Code='ADRMAIL'", _
                      numzone) = P_ERREUR Then
        GoTo lab_err
    End If

    ' Date du jour
    ligne = format(Date, "dd/mm/yyyy") & ";" & vbCr & vbLf
    ' Titre
    stitre = build_titre(v_schoix)
    ligne = ligne & stitre & ";" & vbCrLf
    Print #fd, ligne
    sql = build_sql_ft(v_schoix, False)
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_err
    End If
    While Not rs.EOF
        ' Nom
        ligne = rs("U_Nom").Value & " " & rs("U_Prenom").Value & "|"
        ' Code
        ligne = ligne & rs("UAPP_Code").Value & "|"
        ' Actif
        If rs("U_Actif").Value Then
            s = "Actif"
        Else
            s = "Non actif"
        End If
        ligne = ligne & s & "|"
        ' Adr msg
        sql = "SELECT UC_Valeur FROM UtilCoordonnee" _
            & " WHERE UC_Type='U'" _
            & " AND UC_TypeNum=" & rs("U_Num").Value _
            & " AND UC_ZUNum=" & numzone
        If Odbc_SelectV(sql, rs2) = P_ERREUR Then
            GoTo lab_err
        End If
        If rs2.EOF Then
            sadr = ""
        Else
            sadr = rs2("UC_Valeur").Value
        End If
        rs2.Close
        ligne = ligne & sadr & " |"
        ' Catégorie prof
        If rs("U_CATPNum").Value > 0 Then
            If Odbc_RecupVal("SELECT CATP_Nom FROM CategorieProf WHERE CATP_Num=" & rs("U_CATPNum").Value, _
                             s) = P_ERREUR Then
                s = " "
            End If
        Else
            s = " "
        End If
        ligne = ligne & s & "|"
        ' Sites
        s = ""
        n = STR_GetNbchamp(rs("U_Labo").Value, ";")
        For I = 0 To n - 1
            numl = CLng(Mid$(STR_GetChamp(rs("U_Labo").Value, ";", I), 2))
'            If P_RecupLaboCode(numl, code) = P_ERREUR Then
'                code = ""
'            End If
            If s <> "" Then s = s + " - "
            s = s & code
            If numl = rs("U_LNumPrinc").Value Then
                s = s & " (Princ) "
            End If
        Next I
        ligne = ligne & s & "|"
        ' Fonctions
        s = ""
        n = STR_GetNbchamp(rs("U_FctTrav").Value, ";")
        For I = 0 To n - 1
            numf = CLng(Mid$(STR_GetChamp(rs("U_FctTrav").Value, ";", I), 2))
            If Odbc_RecupVal("SELECT FT_Libelle FROM FctTrav WHERE FT_Num=" & numf, _
                             lib) = P_ERREUR Then
                lib = "???"
            End If
            If s <> "" Then s = s + " - "
            s = s & lib
        Next I
        ligne = ligne & s & " |"
        ' Services
        s_sp = rs("U_SPM").Value
        ns = -1
        np = -1
        n = STR_GetNbchamp(s_sp, "|")
        For I = 0 To n - 1
            s = STR_GetChamp(s_sp, "|", I)
            n2 = STR_GetNbchamp(s, ";")
            nums = CLng(Mid$(STR_GetChamp(s, ";", n2 - 2), 2))
            nump = CLng(Mid$(STR_GetChamp(s, ";", n2 - 1), 2))
            trouve = False
            For j = 0 To ns
                If tbl_s(j) = nums Then
                    trouve = True
                    Exit For
                End If
            Next j
            If Not trouve Then
                ns = ns + 1
                ReDim Preserve tbl_s(ns) As Long
                tbl_s(ns) = nums
            End If
            np = np + 1
            ReDim Preserve tbl_p(np) As Long
            tbl_p(np) = nump
            ReDim Preserve tbl_ps(np) As Long
            tbl_ps(np) = nums
        Next I
        s = ""
        For I = 0 To ns
            If P_RecupSrvNom(tbl_s(I), lib) = P_ERREUR Then
                lib = "???"
            End If
            s = s & lib & "##"
            For j = 0 To np
                If tbl_ps(j) = tbl_s(I) Then
                    If Odbc_RecupVal("SELECT FT_Libelle, L_Code" _
                                    & " FROM Poste, FctTrav, Laboratoire" _
                                    & " WHERE PO_Num=" & tbl_p(j) & " AND L_Num=PO_LNum AND FT_Num=PO_FTNum", _
                                     lib, _
                                     codlabo) = P_ERREUR Then
                        lib = "???"
                    End If
                    If p_NbLabo > 1 Then
                        lib = lib + " - " + codlabo
                    End If
                    s = s & "   " & lib & "##"
                End If
            Next j
        Next I
        ligne = ligne & s & " |"
        ' Fct Autor
        sql = "SELECT FCT_Num, FCT_NumPere, FCT_Libelle" _
            & " FROM FctOK_Util, Fonction" _
            & " WHERE FU_UNum=" & rs("U_Num").Value _
            & " AND FCT_Num=FU_FCTNum" _
            & " ORDER BY FCT_Num"
        s = ""
        n = -1
        If Odbc_SelectV(sql, rs2) <> P_ERREUR Then
            While Not rs2.EOF
                n = n + 1
                ReDim Preserve tbl_numfct(n) As Long
                tbl_numfct(n) = rs2("FCT_Num").Value
                ReDim Preserve tbl_numfctp(n) As Long
                tbl_numfctp(n) = rs2("FCT_NumPere").Value
                rs2.MoveNext
            Wend
            rs2.Close
        End If
        m = -1
        For I = 0 To n
            If Odbc_RecupVal("SELECT FCT_Libelle FROM Fonction WHERE FCT_Num=" & tbl_numfct(I), _
                             lib) = P_ERREUR Then
                lib = "???"
            End If
            If tbl_numfctp(I) > 0 Then
                For j = 0 To m
                    If tbl_numfctp(I) = tbl_fcttr(j) Then
                        ndecal = j + 1
                        m = j
                        Exit For
                    End If
                Next j
            Else
                m = -1
                ndecal = 0
            End If
            m = m + 1
            ReDim Preserve tbl_fcttr(m) As Long
            tbl_fcttr(m) = tbl_numfct(I)
            If ndecal > 0 Then lib = String$(ndecal * 3, " ") + lib
            s = s & lib & "##"
        Next I
        ligne = ligne & s & " "
        rs.MoveNext
        If Not rs.EOF Then
            ligne = ligne & "%" & vbCr & vbLf
        Else
            ligne = ligne & "#;"
        End If
        Print #fd, ligne
    Wend
    Close #fd

'    Call Word_Fusionner(p_chemin_modele + "\" + "Personne\" & modele, _
'                        "", _
'                        nomdata, _
'                        False, _
'                        "", _
'                        True, _
'                        "", _
'                        False, _
'                        WORD_IMPRESSION, _
'                        1, _
'                        WORD_DEB_CROBJ, _
'                        WORD_FIN_RAZOBJ)
'    Call PR_RestoreImp
    Exit Sub

lab_err:
    On Error Resume Next
    Printer.KillDoc
    Call PR_RestoreImp
    On Error GoTo 0

End Sub

Private Sub initialiser()

    g_mode_saisie = True

    txt(TXT_NOM).SetFocus

End Sub

Private Sub quitter()

    Unload Me

End Sub

Private Function recup_fcttrav(ByVal v_sfct As String, _
                               ByRef r_lib As String) As Integer

    Dim s As String
    Dim num As Long

    s = STR_GetChamp(v_sfct, ";", 0)
    If s <> "" Then
        num = CLng(Mid$(s, 2))
        If Odbc_RecupVal("SELECT FT_Libelle FROM FctTrav WHERE FT_Num=" & num, _
                         r_lib) = P_ERREUR Then
            recup_fcttrav = P_ERREUR
            Exit Function
        End If
    Else
        r_lib = ""
    End If

    recup_fcttrav = P_OK

End Function

Private Function recup_fsp(ByVal v_numutil As Long, _
                           ByRef r_nomutil As String, _
                           ByRef r_libfct As String, _
                           ByVal v_cndActif As String, _
                           ByRef r_libsp As String) As Integer

    Dim sfct As String, prenom As String
    Dim s_sp As Variant
    Dim rs As rdoResultset

    If Odbc_RecupVal("SELECT U_Nom, U_Prenom, U_FctTrav, U_SPM FROM Utilisateur Where " & v_cndActif & " U_Num=" & v_numutil, _
                     r_nomutil, _
                     prenom, _
                     sfct, _
                     s_sp) = P_ERREUR Then
        recup_fsp = P_ERREUR
        Exit Function
    End If

    r_nomutil = r_nomutil + " " + prenom
    If recup_fcttrav(sfct, r_libfct) = P_ERREUR Then
        recup_fsp = P_ERREUR
        Exit Function
    End If

    s_sp = STR_GetChamp(s_sp, "|", 0)
    If P_RecupSPLib(s_sp, r_libsp) = P_ERREUR Then
        recup_fsp = P_ERREUR
        Exit Function
    End If

    recup_fsp = P_OK

End Function

Private Function verif_util(ByVal v_mode As Integer) As Integer

    Dim sql As String, nomutil As String, libfct As String, libspm As String, _
        MATRICULE As String, schoix As String, lib_srv_poste As String
    Dim est_ok As Boolean
    Dim n As Integer, I As Integer, j As Integer, nbchp_u As Integer, nbchp_p As Integer, _
        reponse As Integer
    Dim numutil As Long, numlabo As Long, lnb As Long
    Dim rs As rdoResultset
    Dim cndavecinactifs As String

    cndavecinactifs = "U_kb_actif=True AND "
Début:
    If v_mode = TXT_NOM Then
        nomutil = txt(TXT_NOM).Text
        ' Recher du nom exact
        sql = "SELECT U_Num, U_Nom, U_Matricule, U_Prenom, U_Actif, U_KB_Actif, U_ExterneFich, U_SPM" _
            & " FROM Utilisateur" _
            & " WHERE " & cndavecinactifs & " U_Num<>" & P_SUPER_UTIL _
            & " AND " & Odbc_upper & "(U_Nom)=" & Odbc_String(UCase(nomutil))
    Else ' v_mode = TXT_MATRICULE
        MATRICULE = txt(TXT_MATRICULE).Text
        ' Recherche sur le matricule
        sql = "SELECT U_Num, U_Nom, U_Matricule, U_Prenom, U_Actif, U_KB_Actif, U_ExterneFich, U_SPM" _
            & " FROM Utilisateur" _
            & " WHERE " & cndavecinactifs & " U_Num<>" & P_SUPER_UTIL _
            & " AND " & Odbc_upper & "(U_Matricule)=" & Odbc_String(UCase(MATRICULE))
    End If
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        verif_util = P_ERREUR
        Exit Function
    End If
    ' Il y en a
    If Not rs.EOF Then
        GoTo lab_affiche
    End If
    rs.Close
    If v_mode = TXT_NOM Then
        If Right$(nomutil, 1) <> "*" Then
            nomutil = nomutil + "*"
        End If
        ' Recherche NOM commencant par 'nomutil'
        sql = "SELECT U_Num, U_Nom, U_Prenom, U_Matricule, U_KB_Actif, U_Actif, U_ExterneFich, U_Labo, U_SPM" _
            & " FROM Utilisateur" _
            & " WHERE " & cndavecinactifs & " U_Num>1" _
            & " AND " & Odbc_upper() & "(U_Nom) LIKE " & Odbc_String(UCase(nomutil)) _
            & " ORDER BY U_Nom"
    Else ' v_mode = TXT_MATRICULE Then
        If Right$(MATRICULE, 1) <> "*" Then
            MATRICULE = MATRICULE + "*"
        End If
        ' Recherche MATRICULE commencant par 'matricule'
        sql = "SELECT U_Num, U_Nom, U_Prenom, U_Matricule, U_KB_Actif, U_Actif, U_ExterneFich, U_Labo, U_SPM" _
            & " FROM Utilisateur" _
            & " WHERE " & cndavecinactifs & " U_Num>1" _
            & " AND " & Odbc_upper() & "(U_Matricule) LIKE " & Odbc_String(UCase(MATRICULE)) _
            & " ORDER BY U_Nom"
    End If
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        verif_util = P_ERREUR
        Exit Function
    End If
    ' Il n'y en a pas -> proposition de création
    If rs.EOF Then
        rs.Close
        If cndavecinactifs = " true and " Then
            If v_mode = TXT_MATRICULE Then
                Call MsgBox("Aucune personne dont le matricule commence par '" & txt(TXT_MATRICULE).Text & "' n'a été trouvée." _
                          & vbCrLf & "Veuillez renouveler votre recherche.", _
                            vbExclamation + vbOKOnly, _
                            "")
                verif_util = P_NON
                Exit Function
            End If
            reponse = MsgBox("Aucune personne dont le nom commence par '" _
                    & txt(TXT_NOM).Text & "' n'a été trouvée." & vbCrLf & "Voulez-vous la créer ?", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "")
            If reponse = vbNo Then
                verif_util = P_NON
                Exit Function
            End If
            If Right$(txt(TXT_NOM).Text, 1) = "*" Then
                schoix = "U0|NOM=" & UCase(Mid$(txt(TXT_NOM).Text, 1, Len(txt(TXT_NOM).Text) - 1))
            ElseIf left$(txt(TXT_NOM).Text, 1) = "*" Then
                schoix = "U0|NOM=" & UCase(Mid$(txt(TXT_NOM).Text, 2))
            Else
                schoix = "U0|NOM=" & UCase(txt(TXT_NOM).Text)
            End If
            GoTo lab_fin
        Else
            cndavecinactifs = " true and "
            GoTo Début
        End If
    End If
    

    ' Affichage de la liste trouvée
lab_affiche:
    Call CL_Init
    Call CL_InitTitreHelp("Personnes ayant le critère recherché", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("Voir les inactifs de KaliBottin", "", 0, 0, 2000)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_AddLigne("<Nouveau>", 0, "", False)
    n = 1
    While Not rs.EOF
        If recup_fsp(rs("U_Num").Value, nomutil, libfct, cndavecinactifs, libspm) = P_ERREUR Then
            verif_util = P_ERREUR
            Exit Function
        End If
        lib_srv_poste = P_get_lib_srv_poste(P_get_num_srv_poste(rs("U_SPM").Value, P_POSTE), P_POSTE) _
                      & " - " & P_get_lib_srv_poste(P_get_num_srv_poste(rs("U_SPM").Value, P_SERVICE), P_SERVICE)
        Call CL_AddLigne(rs("U_Matricule").Value & vbTab & IIf(rs("U_Actif").Value, "[   ]", "[ Inactive KD]") & vbTab & IIf(rs("U_KB_Actif").Value, "[   ]", "[ Inactive KB]") _
                        & " " & IIf(rs("U_ExterneFich").Value, "[Externe au fichier]", "[   ]") & vbTab & " " _
                        & rs("U_Nom").Value & " " & rs("U_Prenom").Value & vbTab _
                        & lib_srv_poste, _
                         rs("U_Num").Value, _
                         "", _
                         False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close

    If n = 0 Then
        Call MsgBox("Aucune personne n'a été trouvé avec les critères désirés.", vbInformation + vbOKOnly, "")
        verif_util = P_NON
        Exit Function
    End If

    Call CL_InitTaille(0, -20)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        cndavecinactifs = ""
        GoTo Début
    End If
    If CL_liste.retour = 2 Then
        verif_util = P_NON
        Exit Function
    End If
    schoix = "U" & CL_liste.lignes(CL_liste.pointeur).num
    If CL_liste.lignes(CL_liste.pointeur).num = 0 Then
        schoix = schoix & "|NOM=" & UCase(txt(TXT_NOM).Text)
    End If

lab_fin:
    Call choisir_utilisateur(schoix)

    verif_util = P_OUI

End Function

Private Sub cmd_Click(Index As Integer)

    Select Case Index
    Case CMD_CHOIX_UTIL
        Call choisir_utilisateur("0")
    Case CMD_CHOIX_FONCTION
        Call choisir_fonction
    Case CMD_CHOIX_SERVICE
        Call choisir_service
    Case CMD_CHOIX_APPLICATION
        Call choisir_application
    Case CMD_FERMER
        Call quitter
    End Select

End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Index = CMD_FERMER Then g_mode_saisie = False

End Sub

Private Sub Form_Activate()

    If g_form_active Then Exit Sub

    g_form_active = True
    Call initialiser

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyH And Shift = vbAltMask Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_f_utilisateur.htm")
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Call quitter
    End If

End Sub

Private Sub Form_Load()

    g_mode_saisie = False
    g_form_active = False

End Sub

Private Sub txt_GotFocus(Index As Integer)

    g_txt_nom_avant = txt(TXT_NOM).Text
    g_txt_matricule_avant = txt(TXT_MATRICULE).Text

End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        If Index = TXT_NOM Then
            Call choisir_utilisateur("0")
        End If
    End If
    
End Sub

Private Sub txt_LostFocus(Index As Integer)

    If g_mode_saisie Then
        If g_txt_nom_avant <> txt(TXT_NOM).Text Then
            If Index = TXT_NOM Then
                If verif_util(TXT_NOM) <> P_OUI Then
                    txt(TXT_NOM).Text = ""
                    txt(TXT_NOM).SetFocus
                    txt(TXT_MATRICULE).Text = ""
                    Exit Sub
                End If
            End If
        ElseIf g_txt_matricule_avant <> txt(TXT_MATRICULE).Text Then
            If Index = TXT_MATRICULE Then
                If verif_util(TXT_MATRICULE) <> P_OUI Then
                        txt(TXT_NOM).SetFocus
                        txt(TXT_NOM).Text = ""
                        txt(TXT_MATRICULE).Text = ""
                    Exit Sub
                End If
            End If
        End If
    End If

End Sub
