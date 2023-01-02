VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm Menu 
   BackColor       =   &H8000000C&
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10560
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pct 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00016BC5&
      Height          =   11500
      Left            =   0
      ScaleHeight     =   11445
      ScaleWidth      =   10500
      TabIndex        =   1
      Top             =   0
      Width           =   10560
      Begin VB.Frame FrmFichiers 
         Caption         =   "Génération des fichiers"
         Height          =   1335
         Left            =   4320
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   4575
         Begin ComctlLib.ProgressBar PgBarAppli 
            Height          =   375
            Left            =   1320
            TabIndex        =   6
            Top             =   360
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   327682
            Appearance      =   1
         End
         Begin ComctlLib.ProgressBar PgBarUtil 
            Height          =   375
            Left            =   1320
            TabIndex        =   8
            Top             =   840
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   327682
            Appearance      =   1
         End
         Begin VB.Label Label2 
            Caption         =   "Destinataires"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Applications"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.CommandButton cmd 
         Height          =   735
         Index           =   2
         Left            =   11040
         Picture         =   "Menu.frx":05FA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Accès à KaliWeb"
         Top             =   2880
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmd 
         Height          =   495
         Index           =   1
         Left            =   11160
         Picture         =   "Menu.frx":11C4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Vous avez des mouvements à envoyer..."
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmd 
         Height          =   495
         Index           =   0
         Left            =   11160
         Picture         =   "Menu.frx":1F1A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Personnes modifiées dans KaliDoc"
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Timer TimerCodope 
         Left            =   4320
         Top             =   3360
      End
      Begin VB.Label LabPreImport 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   13935
      End
      Begin VB.Image pctfond 
         Height          =   9000
         Left            =   1320
         Picture         =   "Menu.frx":2C70
         Top             =   2880
         Width           =   12000
      End
      Begin ComctlLib.ImageList imglst 
         Left            =   2430
         Top             =   1230
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   71
         ImageHeight     =   71
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   11
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":99A3
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":A1B5
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":A9C7
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":B1D9
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":B9EB
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":C1FD
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":CA0F
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":D221
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":DA33
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":E07D
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Menu.frx":E6C7
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   10560
      TabIndex        =   0
      Top             =   11505
      Width           =   10560
   End
   Begin VB.Menu mnuAnnuaire 
      Caption         =   "&Annuaire       "
      Begin VB.Menu mnuVerification 
         Caption         =   "&Vue globale p/r à un fichier d'importation"
      End
      Begin VB.Menu mnuImportation 
         Caption         =   "I&mportation depuis un fichier"
      End
      Begin VB.Menu mnuImportationAuto 
         Caption         =   "Importation en &Automatique"
      End
   End
   Begin VB.Menu mnuPrm 
      Caption         =   "&Dictionnaires          "
      Begin VB.Menu mnuPrmTypeCoordonne 
         Caption         =   "&Types de coordonnées"
      End
      Begin VB.Menu mnuTypeInfoSuppl 
         Caption         =   "Typ&es d'informations supplémentaires"
      End
      Begin VB.Menu mnuPrmAppli 
         Caption         =   "&Applications"
      End
      Begin VB.Menu mnuPrmFonction 
         Caption         =   "&Fonctions du personnel"
      End
      Begin VB.Menu mnuPrmService 
         Caption         =   "&Services "
      End
      Begin VB.Menu mnuPrmPersonne 
         Caption         =   "&Personnes"
      End
      Begin VB.Menu mnuPrmAssociations 
         Caption         =   "Ass&ociations"
      End
      Begin VB.Menu mnuPrmAssociationsInit 
         Caption         =   "&Initialiser les associations à partir du fichier"
      End
   End
   Begin VB.Menu mnuDivers 
      Caption         =   "Di&vers          "
      Begin VB.Menu MnuDivSession 
         Caption         =   "Nouvelle &Session"
      End
      Begin VB.Menu mnuPrmGen 
         Caption         =   "Paramétrage &général"
      End
      Begin VB.Menu mnuCopierFichier 
         Caption         =   "... Copier le fichier vers le Serveur"
      End
      Begin VB.Menu mnuOuvrirFichier 
         Caption         =   "... Ouvrir le fichier sur mon poste"
      End
      Begin VB.Menu mnuFichierCar 
         Caption         =   "... Définir le fichier des caractères traités"
      End
      Begin VB.Menu mnuActualiser 
         Caption         =   "Réactualiser les modifications"
      End
      Begin VB.Menu mnuVoirLog 
         Caption         =   "Voir le journal d'activité"
      End
      Begin VB.Menu mnuDéfinirMailDest 
         Caption         =   "Définir les destinataire du journal d'activité"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
   Begin VB.Menu mnuAide 
      Caption         =   "        ?"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const g_version = "- 3.00.10C"
Private g_stitre As String
Private g_prems_connex As Boolean

' Indice des boutons
Private Const INDIC_PERSONNE_MODIFIEE = 0
Private Const INDIC_MODIFICATIONS_A_ENVOYER = 1
Private Const INDIC_CMD_KALIWEB = 2

' Image acces kaliweb
Private Const IMG_ACCES_KALIWEB = 11


' Une minute
Private Const TIMER_INTERVAL = 65535

Private g_timeout As Integer

Private g_form_active As Boolean

Private Function init_user(ByVal v_fdemuser As Boolean, _
                           ByVal v_fdemdocs As Boolean) As Integer

    Dim sql As String
    Dim cr As Integer, I As Integer, n As Integer
    Dim sav_numdocs As Long, tbl_docs() As Long
    Dim rs As rdoResultset

    If v_fdemuser Then
        cr = P_SaisirUtilIdent(Me.left, Me.Top, Me.width, Me.Height)
        If cr <> P_OUI Then
            init_user = cr
            Exit Function
        End If
    Else
        If p_NumUtil = 1 Then
            p_CodeUtil = "ROOT"
            init_user = P_OUI
            Exit Function
        End If
        If Odbc_RecupVal("SELECT UAPP_Code FROM UtilAppli WHERE UAPP_UNum=" & p_NumUtil, _
                         p_CodeUtil) = P_ERREUR Then
            init_user = P_ERREUR
            Exit Function
        End If
    End If

'    If charger_fct_autor() = P_ERREUR Then
'        init_user = P_ERREUR
'        Exit Function
'    End If

    init_user = P_OUI

End Function
Private Function lancer_kaliweb() As Integer
    
    Dim numutil As String, url As String

    ' Cryptage du N° d'utilisateur
    numutil = STR_CrypterNombre(format(p_NumUtil, "#0000000"))
    
    url = p_CheminPHP & "/pident.php?in=kalibottin&V_util=" & numutil
    If p_sversconf <> "" Then
        url = url & "&s_vers_conf=" & p_sversconf
    End If
    
    ' Chargement de la page
    Shell "C:\Program Files\Internet Explorer\iexplore.exe " & url, vbMaximizedFocus
    
End Function

Private Sub initialiser()

    Dim fdem As Boolean
    Dim cr_docs As Integer
    Dim taille As Long
    Dim scmd As String

    g_prems_connex = True

'    pctfond.left = 500
'    pctfond.Top = 500
'    pctfond.Height = Me.Height - 2000
'    pctfond.width = Me.width - 1000
'    pctfond.Stretch = True
    pctfond.left = (Me.width - pctfond.width) / 2
    pctfond.Top = (Me.Height - pctfond.Height - 1000) / 2
    pctfond.Visible = True

    
    ' Identification + documentation
    If p_NumUtil = 0 Then
        fdem = True
    Else
        ' Kalidoc est appelé via une autre application -> l'identification
        ' de l'utilisateur est connue
        fdem = False
    End If
    If init_user(fdem, False) <> P_OUI Then
        If Command$ = "DEBUG" Then
            p_NumUtil = 0
            fdem = True
            Call MnuDivSession_Click
        Else
            End
        End If
    End If

    ' utilisateur en Automatique si DEBUG
    scmd = Command$
    If scmd = "DEBUG" Then
        Call SYS_PutIni("FICHIER_INI", "UTILISATEUR", p_NumUtil, p_chemin_appli & "\dernier_ini_ouvert.txt")
    End If
    
    DoEvents
    Me.MousePointer = 11

    ' Vérification nom écran <-> laboratoire
'    If init_ecran() <> P_OUI Then
'        End
'    End If

    ' Vérification si des mouvements sont à envoyer (UM_DateEnvoi)
    If P_mouvements_a_envoyer() Then
        cmd(INDIC_MODIFICATIONS_A_ENVOYER).left = Screen.width - cmd(INDIC_MODIFICATIONS_A_ENVOYER).width
        cmd(INDIC_MODIFICATIONS_A_ENVOYER).Visible = True
    End If
    
    ' Image kaliweb pour rediriger vers l'annuaire
    'taille = 800
    'cmd(INDIC_CMD_KALIWEB).width = taille
    'cmd(INDIC_CMD_KALIWEB).Height = taille
    'cmd(INDIC_CMD_KALIWEB).left = Screen.width - taille
    'cmd(INDIC_CMD_KALIWEB).Top = 6 * taille
    'cmd(INDIC_CMD_KALIWEB).ZOrder 0
    cmd(INDIC_CMD_KALIWEB).left = Screen.width - cmd(INDIC_CMD_KALIWEB).width
    cmd(INDIC_CMD_KALIWEB).Picture = imglst.ListImages(IMG_ACCES_KALIWEB).Picture
    cmd(INDIC_CMD_KALIWEB).Visible = True
    
    If g_timeout > 0 Then
        TimerCodope.Enabled = True
        TimerCodope.Interval = TIMER_INTERVAL
        TimerCodope.tag = 0
    End If

    Me.MousePointer = 0

    g_prems_connex = False

End Sub

Public Sub maj_titre(ByVal v_stitre As String)

    Dim s As String

    g_stitre = v_stitre
    s = "KaliBottin " _
        & g_version _
        & "        " & v_stitre _
        & "        " & p_CodeUtil _
        & "      " & p_CodeLabo _
        & "     Copyright (C) 2003 KALITECH             Tél : 01 69 41 97 54"
    Me.Caption = s

End Sub

Private Sub quitter(ByVal v_mode As Boolean)

    If MsgBox("Êtes-vous sûr de vouloir quitter KaliBottin ?", _
               vbQuestion + vbYesNo + vbDefaultButton2, "") = vbNo Then Exit Sub

    Odbc_Cnx.Close
    Unload Me
    End

End Sub


Private Sub reactive_timercodope()

    If g_timeout > 0 Then
        TimerCodope.Enabled = False
        TimerCodope.Interval = TIMER_INTERVAL
        TimerCodope.tag = 0
        TimerCodope.Enabled = True
    End If

End Sub

Private Sub saisir_ini()

    Dim nom_ini As String, s As String

    nom_ini = p_chemin_appli + "\kalibottin.ini"

    Call SAIS_Init
    Call SAIS_InitTitreHelp("Paramètres du .ini", "")
    s = SYS_GetIni("DOC", "WORD", nom_ini)
    Call SAIS_AddChampComplet("Application Word", 80, SAIS_TYP_TOUT_CAR, "", False, 0, False, s)
    s = SYS_GetIni("DOC", "MODELE", nom_ini)
    Call SAIS_AddChampComplet("Chemin des modèles", 80, SAIS_TYP_TOUT_CAR, "", False, 0, False, s)
    s = SYS_GetIni("DOC", "ARCHIVE", nom_ini)
    Call SAIS_AddChampComplet("Chemin des archives", 80, SAIS_TYP_TOUT_CAR, "", False, 0, False, s)
    s = SYS_GetIni("BASE", "TYPE", nom_ini)
    Call SAIS_AddChampComplet("Type de base", 3, SAIS_TYP_LETTRE, "", False, 0, False, s)
    s = SYS_GetIni("BASE", "NOM", nom_ini)
    Call SAIS_AddChampComplet("Nom de la base", 50, SAIS_TYP_TOUT_CAR, "", False, 0, False, s)
    s = SYS_GetIni("SERVEUR", "CHEMIN", nom_ini)
    Call SAIS_AddChampComplet("Chemin inst.exe", 80, SAIS_TYP_TOUT_CAR, "", False, 0, False, s)

    Call SAIS_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call SAIS_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)

    Saisie.Show 1
    If SAIS_Saisie.retour = 1 Then
        Exit Sub
    End If

    Call SYS_PutIni("DOC", "WORD", SAIS_Saisie.champs(0).sval, nom_ini)
    Call SYS_PutIni("DOC", "MODELE", SAIS_Saisie.champs(1).sval, nom_ini)
    Call SYS_PutIni("DOC", "ARCHIVE", SAIS_Saisie.champs(2).sval, nom_ini)
    Call SYS_PutIni("BASE", "TYPE", SAIS_Saisie.champs(3).sval, nom_ini)
    Call SYS_PutIni("BASE", "NOM", SAIS_Saisie.champs(4).sval, nom_ini)
    Call SYS_PutIni("SERVEUR", "CHEMIN", SAIS_Saisie.champs(5).sval, nom_ini)

End Sub

Private Sub cmd_Click(Index As Integer)
' **************************************************
' 1° Afficher les personnes d modifiées dans KaliDoc
' 2° Générer les fichiers des mouvements
' **************************************************
    Select Case Index
        Case INDIC_MODIFICATIONS_A_ENVOYER
            If cmd(INDIC_PERSONNE_MODIFIEE).Visible Then
            Call MsgBox("Vous devez tout d'abord prendre en compte les modifications" & vbCrLf _
                      & "faites dans KaliDoc afin de continuer." & vbCrLf & vbCrLf _
                      & "Une fois fini, vous pouvez relancer cette opération.", _
                      vbInformation + vbOKOnly, "Attention")
            Else
                If EnvoyerLesModifications() = P_OK Then
                    cmd(INDIC_MODIFICATIONS_A_ENVOYER).Visible = False
                End If
            End If
        Case INDIC_CMD_KALIWEB
            Call lancer_kaliweb
    End Select

End Sub

Private Sub MDIForm_Activate()

    If g_form_active Then
        If Me.tag = "1" Then
            If init_user(True, True) <> P_OUI Then
                End
            End If
            Me.tag = ""
        End If
        If g_timeout > 0 Then
            TimerCodope.Enabled = True
            TimerCodope.Interval = TIMER_INTERVAL
            TimerCodope.tag = 0
        End If
        Exit Sub
    End If
    pct.Height = Screen.Height
    g_form_active = True

    Call initialiser

End Sub

Private Sub MDIForm_Click()

    Call reactive_timercodope

End Sub

Private Sub MDIForm_Deactivate()

    TimerCodope.Enabled = False

End Sub

Private Sub MDIForm_Load()

    g_form_active = False

    Me.Caption = "KaliBottin " & g_version & Space(80) & "     Copyright (C) 2003 KALITECH"

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Call quitter(False)
        Exit Sub
    End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
' **************************************************************
' Pour ne pas quitter quand-t-on clique sur la croix de 'Fermer'
' Laisser Quitter() mettre fin à l'application
' **************************************************************
    Cancel = True

End Sub

Private Sub mnuActualiser_Click()

    ' Vérification si des mouvements sont à envoyer (UM_DateEnvoi)
    If P_mouvements_a_envoyer() Then
        cmd(INDIC_MODIFICATIONS_A_ENVOYER).left = Screen.width - cmd(INDIC_MODIFICATIONS_A_ENVOYER).width
        cmd(INDIC_MODIFICATIONS_A_ENVOYER).Visible = True
    End If

End Sub

Private Sub mnuAide_Click()

    Dim nomfich As String, s As String
    'Call MsgBox("L'aide de KaliBottin n'est pas encore disponible.", vbInformation + vbOKOnly, "")
    'Exit Sub
    
    nomfich = p_chemin_appli + "\help\kalibottin.pdf"
    If FICH_FichierExiste(nomfich) Then
        s = nomfich
        Call SYS_StartProcess(s)
    Else
        Call MsgBox("L'aide de KaliBottin n'est pas disponible sur votre poste." & Chr(13) & Chr(10) & p_chemin_appli + "\help\kalibottin.pdf", vbInformation + vbOKOnly, "")
    End If
    'Call HtmlHelp(0, p_chemin_appli + "\help\kalibottin.pdf", HH_DISPLAY_TOPIC, "generalites.htm")

End Sub

Private Sub mnuCopierFichier_Click()
    Dim rs As rdoResultset
    Dim liberr As String
    Dim chemin As String
    Dim fichier_importation As String
    
    ' Récupérer le chemin du fichier d'importation
    If Odbc_SelectV("SELECT PGB_Chemin, PGB_fichsurserveur FROM PrmGenB", rs) = P_ERREUR Then
        Exit Sub
    End If
    fichier_importation = ""
    If Not rs.EOF Then
        fichier_importation = rs("PGB_Chemin").Value
        p_est_sur_serveur = rs("PGB_fichsurserveur").Value
        rs.Close
    End If
    If fichier_importation <> "" Then
        chemin = p_chemin_appli & "\tmp\personnel_kalibottin.txt"
        chemin = Com_ChoixFichier.AppelFrm("Choix du fichier d'importation", "", p_chemin_appli & "\tmp", P_EXTENSIONS_FICHIERS_IMPORTATION, False)
        If HTTP_Appel_fichier_existe(fichier_importation, False, liberr) = HTTP_OK Then
            If p_est_sur_serveur Then
                If FICH_FichierExiste(chemin) Then
                    If KF_PutFichier(fichier_importation, chemin) = P_ERREUR Then
                        Exit Sub
                    Else
                        MsgBox "Transfert effectué"
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub mnuDéfinirMailDest_Click()
    'MsgBox "ici"
    ' Lancement KD pour maj des diffusions
    's = p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";KaliDoc;" & p_nom_fichier_ini_kalidoc & ";CONNEXION=" & p_NumUtil & ";ACTIVER_PERS=" & numutil & "[WAIT];[KBAUTO]"
    'Call SYS_ExecShell(s, True, True)
End Sub

Private Sub MnuDivSession_Click()
    Dim scmd As String
    
    If init_user(True, True) = P_ERREUR Then
        End
    End If
    
    ' utilisateur en Automatique si DEBUG
    scmd = Command$
    If scmd = "DEBUG" Then
        Call SYS_PutIni("FICHIER_INI", "UTILISATEUR", p_NumUtil, p_chemin_appli & "\dernier_ini_ouvert.txt")
    End If
    
    SendKeys "%{V}"
    

End Sub

Public Function FaitFichierCar(v_trait)
    Dim s As String
    Dim chemin_serv As String, chemin_loc As String
    Dim fd As Integer
    Dim fso As FileSystemObject
        
    chemin_serv = p_CheminKW & "/includes/Specifique_Site/kaliBottinCar.txt"
    chemin_loc = p_chemin_appli & "/kaliBottinCar.txt"
    Call FICH_EffacerFichier(chemin_loc, False)
    If Not KF_FichierExiste(chemin_serv) Then
        FICH_OuvrirFichier chemin_loc, FICH_ECRITURE, fd
        Print #fd, "#******************************"
        Print #fd, "#Fichier des caractères traités"
        Print #fd, "#******************************"
        
        s = "203=E"
        Print #fd, s
        s = "176=°"
        Print #fd, s
        s = "220=u"
        Print #fd, s
        s = "133=à"
        Print #fd, s
        s = "224=à"
        Print #fd, s
        s = "226=â"
        Print #fd, s
        s = "199=" & UCase("ç")
        Print #fd, s
        s = "231=ç"
        Print #fd, s
        s = "130=é"
        Print #fd, s
        s = "233=é"
        Print #fd, s
        s = "234=é"
        Print #fd, s
        s = "201=" & UCase("é")
        Print #fd, s
        s = "200=" & UCase("è")
        Print #fd, s
        s = "235=ë"
        Print #fd, s
        s = "144=E"
        Print #fd, s
        s = "207=" & UCase("ï")
        Print #fd, s
        s = "239=ï"
        Print #fd, s
        s = "206=" & UCase("î")
        Print #fd, s
        s = "238=" & UCase("î")
        Print #fd, s
        s = "212=" & UCase("ô")
        Print #fd, s
        s = "246=" & UCase("ô")
        Print #fd, s
        s = "244=" & UCase("ô")
        Print #fd, s
        Close #fd
        Call KF_PutFichier(chemin_serv, chemin_loc)
    End If
    If v_trait = "Ouvrir" Then
        ' ouvrir avec notepad ou autre
        s = SYS_GetIni("DOC", "EDITEUR_JS", p_nomini)
        If s <> "" Then
            Set fso = CreateObject("Scripting.FileSystemObject")
            If Not fso.FileExists(s) Then
                MsgBox "La constante DOC->EDITEUR_JS = " & s & " : fichier introuvable"
                s = "C:\WINDOWS\NOTEPAD.exe"
            End If
        Else
            MsgBox "La constante DOC->EDITEUR_JS n'est pas renseignée dans config.php"
            s = "C:\WINDOWS\NOTEPAD.exe"
            If Not FICH_FichierExiste(s) Then
                MsgBox "La constante DOC->EDITEUR_JS = " & s & " : fichier introuvable"
            End If
        End If
        s = s & " " & chemin_loc
        Call KF_GetFichier(chemin_serv, chemin_loc)
        Call SYS_ExecShell(s, True, True)
        Call KF_PutFichier(chemin_serv, chemin_loc)
        p_bool_p_tbl_car_traités_chargé = False
    End If
End Function
Private Sub mnuFichierCar_Click()
    
    Call FaitFichierCar("Ouvrir")

End Sub

' Menu Annuaire

Public Sub mnuImportation_Click()

    Dim importer As Boolean
    Dim cr As Integer
    Dim frm As Form
    Dim nomfich As String
    Dim fd As Integer
    Dim s As String
    Dim I As Integer
    
    nomfich = p_chemin_appli & "\rapport"
    If FICH_EstFichierOuRep(nomfich) = FICH_REP Then
    Else
        MkDir nomfich
    End If
Lab1:
    nomfich = p_chemin_appli & "\rapport\rapport_kalibottin_fait_" & p_NumUtil & ".txt"
    Call FICH_EffacerFichier(nomfich, False)
    g_fd1 = 11
    On Error GoTo LabClose1
    Open nomfich For Append As g_fd1
    GoTo Lab2
LabClose1:
    Close g_fd1
    Resume
Lab2:
    nomfich = p_chemin_appli & "\rapport\rapport_kalibottin_pasfait_" & p_NumUtil & ".txt"
    Call FICH_EffacerFichier(nomfich, False)
    g_fd2 = 12
    On Error GoTo LabClose2
    Open nomfich For Append As g_fd2
    GoTo Lab3
LabClose2:
    Close g_fd2
    Resume
Lab3:
    nomfich = p_chemin_appli & "\rapport\rapport_kalibottin_assoc_" & p_NumUtil & ".txt"
    Call FICH_EffacerFichier(nomfich, False)
    g_fd3 = 13
    On Error GoTo LabClose3
    Open nomfich For Append As g_fd3
    GoTo LabSuite
LabClose3:
    Close g_fd3
    Resume

LabSuite:

    On Error GoTo 0
    Call Gerer_PreImport
    If p_traitement_préimport_seul Then ' on ne fait que le pré-import
        End
    End If
    
    If P_InitFichierImportation() = P_OUI Then ' on selectionne le fichier d'import.
lab_reverifier:
        If p_mess_fait_background <> "" Then
            Print #g_fd1, p_mess_fait_background
        End If
        If p_mess_pasfait_background <> "" Then
            Print #g_fd2, p_mess_pasfait_background
        End If
        Set frm = VerificationAnnuaire
        ' Vérification avant importation
        importer = VerificationAnnuaire.AppelFrm(1)
        Set frm = Nothing
        DoEvents
        If importer Then
            ' L'importation effective
            Set frm = ImportationAnnuaire
            cr = ImportationAnnuaire.AppelFrm
            If cr = P_ERREUR Then
                Set frm = Nothing
                GoTo lab_reverifier
            End If
            Set frm = Nothing
            ' Ouvrir le log
            If p_ouvrir_log Then
                Call mnuVoirLog_Click
            End If
        End If
    End If

lab_fin:
    SendKeys "%{A}{DOWN}"

End Sub

Private Function Gerer_PreImport()
    Dim fd_Lock As Integer
    
    ' Voir si un programme de pré_import est à executer
    If p_programme_preimport_exe <> "" Then
        Menu.LabPreImport.Visible = True
        Menu.LabPreImport.Caption = "Execution de : " & p_programme_preimport_exe & " " & p_nomini
        Menu.LabPreImport.Refresh
        If Not FICH_FichierExiste(p_programme_preimport_exe) Then
            MsgBox p_programme_preimport_exe & " introuvable"
        Else
            If FICH_FichierExiste(p_programme_preimport_log) Then
                Call FICH_EffacerFichier(p_programme_preimport_log, True)
            End If
            If FICH_FichierExiste(p_programme_preimport_lock) Then
                Call FICH_EffacerFichier(p_programme_preimport_lock, True)
            End If
            ' Créer le fichier de lock
            If FICH_OuvrirFichier(p_programme_preimport_lock, FICH_ECRITURE, fd_Lock) = P_ERREUR Then
                MsgBox "Impossible d'ouvrir " & p_programme_preimport_lock
                End
            Else
                Print #fd_Lock, "Date " & Date
                Print #fd_Lock, "Fichier : " & p_programme_preimport_lock
                Close #fd_Lock
            End If
            
            ' lancer l'exe
            Call SYS_ExecShell(p_programme_preimport_exe & " " & p_nomini, True, True)

            ' Effacer le fichier de lock
            Call FICH_EffacerFichier(p_programme_preimport_lock, False)
        End If
        Menu.LabPreImport.Visible = False
    End If

End Function

Private Sub mnuImportationAuto_Click()
    
    p_traitement_background = True
    p_traitement_background_semiauto = True
    Call Menu.mnuImportation_Click

End Sub

Private Sub mnuOuvrirFichier_Click()
    Dim rs As rdoResultset
    Dim chemin As String
    Dim fichier_importation As String
    Dim pos As Integer, sext As String
    Dim liberr As String
    
    ' Récupérer le chemin du fichier d'importation
    If Odbc_SelectV("SELECT PGB_Chemin, PGB_fichsurserveur FROM PrmGenB", rs) = P_ERREUR Then
        Exit Sub
    End If
    fichier_importation = ""
    If Not rs.EOF Then
        fichier_importation = rs("PGB_Chemin").Value
        p_est_sur_serveur = rs("PGB_fichsurserveur").Value
        rs.Close
    End If
    If fichier_importation <> "" Then
        If HTTP_Appel_fichier_existe(fichier_importation, False, liberr) = HTTP_OK Then
            If p_est_sur_serveur Then
                pos = InStrRev(fichier_importation, ".")
                sext = Mid$(fichier_importation, pos)
                chemin = p_chemin_appli & "\tmp\personnel_kalibottin.txt"
                If HTTP_Appel_GetFile(fichier_importation, chemin, False, False, liberr) = P_ERREUR Then
                    MsgBox liberr
                    Exit Sub
                End If
                Call SYS_StartProcess(chemin)
            End If
        End If
    End If
End Sub

' Menu Dictionnaire

Private Sub mnuPrmAppli_Click()

    PrmApplication.Show (1)

    SendKeys "%{D}"

End Sub

Private Sub mnuPrmAssociations_Click()
    Call ImportationAnnuaire.Liste_Assoc("", "", "", "")
End Sub

Private Sub mnuPrmAssociationsInit_Click()
    Dim frm As Form
    Dim cret As String
    
    ' à partir du fichier
    Set frm = AnalyseFichier
    cret = AnalyseFichier.AppelFrm

End Sub

Private Sub mnuPrmFonction_Click()

    Dim frm As Form

'    Set frm = PrmFonction
    Call PrmFonction.AppelFrm("F", -1)
'    Set frm = Nothing

    SendKeys "%{D}"

End Sub

Private Sub mnuPrmService_Click()

    Dim sret As String, sprm As String
    Dim encore As Boolean
    Dim numlabo As Long, numutil As Long
    Dim frm As Form

    numlabo = p_NumLabo

    encore = True
    Do
        Set frm = PrmService
        sret = PrmService.AppelFrm("Paramétrage des services - postes", "M", False, "", "", True)
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
    
    p_NumLabo = numlabo

    SendKeys "%{D}"

End Sub

Private Sub mnuPrmPersonne_Click()

    ChoixPrmPers.Show 1

    SendKeys "%{D}"

End Sub

' MENU Divers

Private Sub mnuPrmGen_Click()

    Call PrmGeneral.AppelFrm

    SendKeys "%{V}"

End Sub

Private Sub mnuPrmTypeCoordonne_Click()
    
    PrmTypeCoordonnees.AppelFrm (-1)

    SendKeys "%{D}"

End Sub

' MENU Quitter
Private Sub mnuQuitter_Click()

    Call quitter(False)

End Sub

Private Sub mnuTypeInfoSuppl_Click()
    
    Call PrmTypeInfoSuppl.AppelFrm(-1, -1)

    SendKeys "%{D}"

End Sub

' Menu Annuaire

Private Sub mnuVerification_Click()

    Dim frm As Form

    If P_InitFichierImportation() = P_OUI Then
        Set frm = VerificationAnnuaire
        Call VerificationAnnuaire.AppelFrm(0)
        Set frm = Nothing
    End If

    SendKeys "%{A}"

End Sub

Private Sub mnuVoirLog_Click()
    Dim nomfich As String
    Dim s As String
    
    If g_nomfichHTML = "" Then
        MsgBox "Aucun journal d'activité récent"
        nomfich = p_chemin_appli & "\rapport\rapport_kalibottin_" & p_NumUtil & "_" & Replace(Date, "/", "-") & "_" & Time & ".html"
        ' La recherche du fichier image
        nomfich = Com_ChoixFichier.AppelFrm("Selectionnez le fichier de traces", "", p_chemin_appli & "\rapport", _
                         "*.html", False)
    Else
        nomfich = g_nomfichHTML
    End If
    If nomfich <> "" And FICH_FichierExiste(nomfich) Then
        s = nomfich
        Call SYS_StartProcess(s)
    End If
End Sub

Private Sub TimerCodope_Timer()

    If TimerCodope.tag = g_timeout - 1 Then
        If Not FRM_EstEnCours(Me) Then
'            Debug.Print "Pas en cours"
            Me.tag = 1
            Exit Sub
        End If
        Me.tag = ""
        If init_user(True, True) <> P_OUI Then
            End
        End If
        Call reactive_timercodope
    Else
        TimerCodope.tag = TimerCodope.tag + 1
        TimerCodope.Enabled = False
        TimerCodope.Interval = TIMER_INTERVAL
        TimerCodope.Enabled = True
    End If

End Sub
