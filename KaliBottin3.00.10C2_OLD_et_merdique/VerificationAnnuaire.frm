VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form VerificationAnnuaire 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rechercher une personne dans le dictionnaire"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   735
      Index           =   2
      Left            =   2400
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   9375
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   0
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   2
         Left            =   8520
         Picture         =   "VerificationAnnuaire.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Rechercher"
         Top             =   180
         Width           =   550
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
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   495
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
         Index           =   5
         Left            =   4320
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Associer"
      Enabled         =   0   'False
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
      Index           =   1
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Associer les MATRICULEs des deux personnes selectionnées."
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cette CAPTION va être changée dans le code !"
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
      Height          =   7695
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12015
      Begin VB.Frame frmPatience 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Chargement en cours ..."
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
         Height          =   1935
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Visible         =   0   'False
         Width           =   11535
         Begin ComctlLib.ProgressBar pgb 
            Height          =   495
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   873
            _Version        =   327682
            Appearance      =   1
            Max             =   1000
         End
         Begin ComctlLib.ProgressBar pgb2 
            Height          =   495
            Left            =   5160
            TabIndex        =   19
            Top             =   1200
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   873
            _Version        =   327682
            Appearance      =   1
            Max             =   1000
         End
         Begin VB.Label LbGauge2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   20
            Top             =   1200
            Width           =   4815
         End
      End
      Begin VB.ComboBox cmbox 
         Height          =   315
         ItemData        =   "VerificationAnnuaire.frx":056D
         Left            =   10080
         List            =   "VerificationAnnuaire.frx":056F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Selectionner un matricule dans la liste afin d'y accéder dans le tableau suivant."
         Top             =   1250
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   2475
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   5040
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4366
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   8454143
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   2475
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   1680
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4366
         _Version        =   393216
         Rows            =   1
         Cols            =   12
         FixedCols       =   0
         BackColorFixed  =   8454143
         AllowUserResizing=   1
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Matricules REDONDANTS"
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
         Index           =   3
         Left            =   7800
         TabIndex        =   12
         ToolTipText     =   "Selectionner un matricule dans la liste afin d'y accéder dans le tableau suivant."
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Personnes du fichier d'importation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   4710
         Visible         =   0   'False
         Width           =   10095
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Personnes de KaliBottin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   1350
         Visible         =   0   'False
         Width           =   7335
      End
      Begin ComctlLib.ImageList imglst 
         Left            =   11160
         Top             =   4320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   16777215
         ImageWidth      =   58
         ImageHeight     =   10
         MaskColor       =   -2147483638
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   4
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "VerificationAnnuaire.frx":0571
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "VerificationAnnuaire.frx":0937
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "VerificationAnnuaire.frx":0CFD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "VerificationAnnuaire.frx":10C3
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frm 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   7560
      Width           =   12030
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Associer en au&to"
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
         Index           =   3
         Left            =   450
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1725
      End
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
         Index           =   0
         Left            =   11355
         Picture         =   "VerificationAnnuaire.frx":1489
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Quitter l'application en cours"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
   Begin VB.Menu mnuContextuel 
      Caption         =   "menu contextuel"
      Visible         =   0   'False
      Begin VB.Menu mnuAcceder 
         Caption         =   "Accéder"
      End
      Begin VB.Menu mnuAssocier 
         Caption         =   "Associer"
      End
   End
End
Attribute VB_Name = "VerificationAnnuaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Index des boutons CMD
Private Const CMD_QUITTER = 0
Private Const CMD_ASSOCIER = 1
Private Const CMD_RECHERCHER = 2
Private Const CMD_ASSOCIER_AUTO = 3

' Index des Grid
Private Const GRD_HAUT = 0 ' 9 colonnes
Private Const GRD_BAS = 1 ' 7 colonnes

' Index des TextBox
Private Const TXT_NOM = 0
Private Const TXT_MATRICULE = 1

' Index des FRAMES
Private Const FRM_PRINCIPALE = 0 ' ne sert à rien
Private Const FRM_QUITTER = 1 ' ne sert à rien
Private Const FRM_RECHERCHER = 2

' Index des colonnes
' GRD_HAUT
Private Const GRDC_HAUT_U_NUM = 0
Private Const GRDC_HAUT_IMPORTE = 1
Private Const GRDC_HAUT_MATRICULE = 2
Private Const GRDC_HAUT_NOM = 3
Private Const GRDC_HAUT_PRENOM = 4
Private Const GRDC_HAUT_CODE_SECTION = 5
Private Const GRDC_HAUT_LIB_SECTION = 6
Private Const GRDC_HAUT_CODE_EMPLOI = 7
Private Const GRDC_HAUT_LIB_EMPLOI = 8
Private Const GRDC_HAUT_ENCORE_EMPLOI = 9
Private Const GRDC_HAUT_MODIFIER = 10
Private Const GRDC_HAUT_VERT_ROUGE = 11 ' l'état de l'importation: V, R  ou X pour mi-vert mi-rouge (doublons)
Private Const GRDC_HAUT_PLUSIEURS_POSTES = 12
' GRD_BAS
Private Const GRDC_BAS_U_NUM = 0
Private Const GRDC_BAS_IMPORTE = 1
Private Const GRDC_BAS_MATRICULE = 2
Private Const GRDC_BAS_NOM = 3
Private Const GRDC_BAS_PRENOM = 4
Private Const GRDC_BAS_CODE_SECTION = 5
Private Const GRDC_BAS_LIB_SECTION = 6
Private Const GRDC_BAS_CODE_EMPLOI = 7
Private Const GRDC_BAS_LIB_EMPLOI = 8
Private Const GRDC_BAS_VERT_ROUGE = 9   ' l'état de l'importation: V, R  ou X pour mi-vert mi-rouge (doublons)

' Index des ImageList
Private Const IMG_SUITE = 1
Private Const IMG_NON_IMPORTE = 2
Private Const IMG_IMPORTE = 3
Private Const IMG_MATRICULE_REDONDANT = 4

Private tbcaractere_nontraite()

' Index des positions des images
Private Const POS_GAUCHE = 1
Private Const POS_CENTRE = 4
Private Const POS_DROITE = 7

' Constantes pour les ToolTipTexts et le ComboBox
Private Const IMPORTE = "V"
Private Const NON_IMPORTE = "R"
Private Const MATRICULE_REDONDANT = "X" ' mi-vert mi-rouge

' Index des lbl
Private Const LBL_GRD_HAUT = 1
Private Const LBL_GRD_BAS = 2
Private Const LBL_COMBOBOX = 3
Private Const LBL_NOM = 4 ' ne sert à rien
Private Const LBL_MATRICULE = 5 ' ne sert à rien

' Index des modes de traitement
Private Const MODE_VERIFICATION = 0
Private Const MODE_VERIF_AVANT_IMPORT = 1

' Index des couleurs du tri
Private Const COLOR_PAS_DE_TRI = &HFFFFFF '&H00FFFFFF& BLANC
Private Const COLOR_DU_TRI = &HE0E0E0    ' &H00E0E0E0& GRIS CLAIR

' Mode de traitement du fichier:
Private g_mode_traitement As Integer

' Coordonnées des cellules selectionnées pour chaque grid
Private g_grd_haut_row As Integer
Private g_grd_haut_col As Integer
Private g_grd_bas_row As Integer
Private g_grd_bas_col As Integer

' Sur quelle colonne le tri a été fait
Dim g_col_tri_haut As Integer
Dim g_col_tri_bas As Integer
Private g_sens_tri_haut As Integer
Private g_sens_tri_bas As Integer
Private g_ancienne_key As Integer
Private g_form_active As Boolean
Private g_grd_haut_caption As String
Private g_nbr_sans_matricule As Integer
' Nombre des premiers caractères du PRENOM à comparer
Private Const g_nbr_car_prenom = 1
Private g_redim_grd_haut As Boolean
Private g_tout_est_associe As Boolean

Private Sub ajouter_ligne_bas(ByVal v_ligne_lu As String)

    Dim matricule_en_cours As String
    Dim matricule_ancien As String
    Dim etat_importe As Boolean

    With grd(GRD_BAS)
        '''matricule_en_cours = STR_GetChamp(v_ligne_lu, p_separateur, p_pos_matricule)
        matricule_en_cours = P_lire_valeur(p_type_fichier, v_ligne_lu, p_separateur, p_pos_matricule, p_long_matricule, "matricule")
        If .Rows > 1 Then
            matricule_ancien = .TextMatrix(.Rows - 2, GRDC_BAS_MATRICULE)
            If matricule_en_cours = matricule_ancien Then
                GoTo Prochain
            End If
        End If
        .AddItem ""
        .TextMatrix(.Rows - 1, GRDC_BAS_MATRICULE) = matricule_en_cours
        .Row = .Rows - 1
        .col = GRDC_BAS_IMPORTE
        etat_importe = False
        If matricule_existe(matricule_en_cours, etat_importe) Then
            If etat_importe Then
                Set .CellPicture = imglst.ListImages(IMG_IMPORTE).Picture
                .TextMatrix(.Rows - 1, GRDC_BAS_VERT_ROUGE) = IMPORTE
            Else
                Set .CellPicture = imglst.ListImages(IMG_NON_IMPORTE).Picture
                .TextMatrix(.Rows - 1, GRDC_BAS_VERT_ROUGE) = NON_IMPORTE
            End If
        Else
            Set .CellPicture = imglst.ListImages(IMG_NON_IMPORTE).Picture
            .TextMatrix(.Rows - 1, GRDC_BAS_VERT_ROUGE) = NON_IMPORTE
        End If
        .TextMatrix(.Rows - 1, GRDC_BAS_NOM) = Trim$(UCase$(P_ChangerCar(P_lire_valeur(p_type_fichier, v_ligne_lu, p_separateur, p_pos_nom, p_long_nom, "nom"), tbcaractere_nontraite)))
        .TextMatrix(.Rows - 1, GRDC_BAS_PRENOM) = Trim$(formater_prenom(P_lire_valeur(p_type_fichier, v_ligne_lu, p_separateur, p_pos_prenom, p_long_prenom, "prénom")))
        .TextMatrix(.Rows - 1, GRDC_BAS_CODE_SECTION) = Trim$(P_lire_valeur(p_type_fichier, v_ligne_lu, p_separateur, p_pos_code_section, p_long_code_section, "code section"))
        .TextMatrix(.Rows - 1, GRDC_BAS_LIB_SECTION) = Trim$(P_ChangerCar(P_lire_valeur(p_type_fichier, v_ligne_lu, p_separateur, p_pos_lib_section, p_long_lib_section, "libellé section"), tbcaractere_nontraite))
        .TextMatrix(.Rows - 1, GRDC_BAS_CODE_EMPLOI) = Trim$(P_lire_valeur(p_type_fichier, v_ligne_lu, p_separateur, p_pos_code_emploi, p_long_code_emploi, "code emploi"))
        .TextMatrix(.Rows - 1, GRDC_BAS_LIB_EMPLOI) = Trim$(P_ChangerCar(P_lire_valeur(p_type_fichier, v_ligne_lu, p_separateur, p_pos_lib_emploi, p_long_lib_emploi, "libellé emploi"), tbcaractere_nontraite))
Prochain:
    End With

End Sub

Private Sub ajouter_ligne_haut(ByVal v_numutil As Long, _
                               ByVal v_import As String, _
                               ByVal v_matricule As String, _
                               ByVal v_nom As String, _
                               ByVal v_prenom As String, _
                               ByVal v_spm As Variant, _
                               ByVal v_numposte As Long)

    Dim lig As Integer

    With grd(GRD_HAUT)
        .AddItem ""
        lig = .Rows - 1

        .TextMatrix(lig, GRDC_HAUT_U_NUM) = v_numutil

        .Row = lig
        .col = GRDC_HAUT_IMPORTE
        Select Case v_import
            Case IMPORTE
                Set .CellPicture = imglst.ListImages(IMG_IMPORTE).Picture
            Case NON_IMPORTE
                Set .CellPicture = imglst.ListImages(IMG_NON_IMPORTE).Picture
        End Select
        
        .TextMatrix(lig, GRDC_HAUT_VERT_ROUGE) = v_import
        .TextMatrix(lig, GRDC_HAUT_MATRICULE) = v_matricule
        .TextMatrix(lig, GRDC_HAUT_NOM) = v_nom
        .TextMatrix(lig, GRDC_HAUT_PRENOM) = formater_prenom(v_prenom)

        ' Remplir des 4 colonnes concernant le SERVICE et le POSTE
        Call remplir_srv_fct(v_numutil, v_spm, v_numposte, lig)

        .TextMatrix(lig, GRDC_HAUT_MODIFIER) = "->"
        .col = GRDC_HAUT_MODIFIER
        .CellBackColor = P_VERT
        .CellFontBold = True

    End With
    
End Sub

Private Function ajouter_synchro(ByVal v_row_haut As Integer, _
                                 ByVal v_row_bas As Integer) As Integer

    Dim sql As String, mon_service As String, mon_poste As String, section As String
    Dim emploi As String
    Dim I As Integer, nbr As Integer
    Dim lnb As Long, mon_num_service As Long, mon_num_poste As Long
    Dim rs As rdoResultset, rs2 As rdoResultset

    With grd(GRD_HAUT)
        If .TextMatrix(v_row_haut, GRDC_HAUT_PLUSIEURS_POSTES) Then
        ' La personne est affectée à plus d'un poste
            Call CL_Init
            sql = "SELECT  * FROM Utilisateur WHERE U_kb_actif=True AND U_Actif=TRUE AND U_Num=" & .TextMatrix(v_row_haut, GRDC_HAUT_U_NUM)
            If Odbc_Select(sql, rs) = P_ERREUR Then
                GoTo lab_erreur
            End If
            For I = 0 To STR_GetNbchamp(rs("U_SPM").Value, "|") - 1
                ' Récupérer le nombre de serices + le poste, au mois nbr=2
                nbr = STR_GetNbchamp(STR_GetChamp(rs("U_SPM").Value, "|", I), ";")
                mon_num_service = Mid$(STR_GetChamp(STR_GetChamp(rs("U_SPM").Value, "|", I), ";", nbr - 2), 2)
                mon_num_poste = Mid$(STR_GetChamp(STR_GetChamp(rs("U_SPM").Value, "|", I), ";", nbr - 1), 2)
                ' Le SRVICE
                sql = "SELECT SRV_Nom FROM Service WHERE SRV_Num=" & mon_num_service
                If Odbc_RecupVal(sql, mon_service) = P_ERREUR Then
                    GoTo lab_erreur
                End If
                ' Le POSTE
                sql = "SELECT FT_Libelle FROM FctTrav, Poste" _
                    & " WHERE FT_Num=PO_FTNum AND PO_Num=" & mon_num_poste
                If Odbc_RecupVal(sql, mon_poste) = P_ERREUR Then
                    GoTo lab_erreur
                End If
                'Call CL_AddLigne(mon_poste  & ":" & mon_service ,
                Call CL_AddLigne(mon_poste & vbTab & mon_service, _
                                 mon_num_poste, _
                                 mon_num_service, _
                                 False)
            Next I
            rs.Close
            Call CL_InitTaille(0, -15)
            Call CL_InitTitreHelp("Les postes associés à " & .TextMatrix(v_row_haut, GRDC_HAUT_NOM) & " " _
                                 & .TextMatrix(v_row_haut, GRDC_HAUT_PRENOM), "")
            Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
            Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
            ' Afficher la liste de tous les postes
            ChoixListe.Show 1
            ' Tester le choix:
            ' Quitter
            If CL_liste.retour = 1 Then
                ' Remettre les lignes en mode selection
                Call mettre_lignes_selectionnees(v_row_haut, v_row_bas)
                ' Activer le bouton CMD_ASSOCIER
                cmd(CMD_ASSOCIER).Enabled = True
                GoTo lab_erreur
            End If
            ' Enregistrer
            .TextMatrix(v_row_haut, GRDC_HAUT_CODE_SECTION) = CL_liste.lignes(CL_liste.pointeur).tag
            .TextMatrix(v_row_haut, GRDC_HAUT_LIB_SECTION) = STR_GetChamp(CL_liste.lignes(CL_liste.pointeur).texte, vbTab, 1)
            .TextMatrix(v_row_haut, GRDC_HAUT_CODE_EMPLOI) = CL_liste.lignes(CL_liste.pointeur).num
            .TextMatrix(v_row_haut, GRDC_HAUT_LIB_EMPLOI) = STR_GetChamp(CL_liste.lignes(CL_liste.pointeur).texte, vbTab, 0)
        End If

        ' Ne pas insérer si le triplé SECTION-EMPLOI-SPNUM existe dans Synchro
        section = grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_CODE_SECTION)
        emploi = grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_CODE_EMPLOI)
        sql = "SELECT COUNT(*) FROM Synchro WHERE SYNC_SPNum=" & .TextMatrix(v_row_haut, GRDC_HAUT_CODE_EMPLOI) _
            & " AND SYNC_Section=" & Odbc_String(section) _
            & " AND SYNC_Emploi=" & Odbc_String(emploi)
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            GoTo lab_erreur
        End If
        If lnb = 0 Then
            If Odbc_AddNew("Synchro", "SYNC_Num", "SYNC_Seq", False, 99, _
                            "SYNC_Section", section, _
                            "SYNC_Emploi", emploi, _
                            "SYNC_SPNum", .TextMatrix(v_row_haut, GRDC_HAUT_CODE_EMPLOI), _
                            "Sync_auto", False) Then
            End If
        End If
    End With

    ajouter_synchro = P_OK
    Exit Function

lab_erreur:
    ajouter_synchro = P_ERREUR

End Function

Public Function AppelFrm(ByVal v_mode_traitement As Integer) As Boolean
' Appelée depuis mnuVerification ou mnuImportation

    Call FRM_ResizeForm(Me, Me.width, Me.Height)

    g_mode_traitement = v_mode_traitement

    Me.Show 1

    If g_mode_traitement = MODE_VERIF_AVANT_IMPORT Then
        If g_tout_est_associe Then
            cmd(CMD_QUITTER).Visible = False
            AppelFrm = True ' => Passer à l'importation
        Else
            AppelFrm = False
        End If
    End If

End Function

Private Sub apres_remplir_grid()
' quelques opérations après le remplissage des grids

    Dim message As String
    Dim I As Integer
    
    ' Affichage des grids
    frmPatience.Visible = False
    Me.MousePointer = 0
    lbl(LBL_GRD_HAUT).Visible = True
    lbl(LBL_GRD_BAS).Visible = True
    ' S'il n'y a pas de ScrollBar => redimensionner les grids
    If grd(GRD_HAUT).Rows <= 10 Then
        With grd(GRD_HAUT)
            .ColWidth(GRDC_HAUT_LIB_SECTION) = .ColWidth(GRDC_HAUT_LIB_SECTION) + 150
            .ColWidth(GRDC_HAUT_LIB_EMPLOI) = .ColWidth(GRDC_HAUT_LIB_EMPLOI) + 105
            g_redim_grd_haut = False
        End With
    End If
    If grd(GRD_BAS).Rows <= 10 Then
        With grd(GRD_BAS)
            .ColWidth(GRDC_BAS_LIB_SECTION) = .ColWidth(GRDC_BAS_LIB_SECTION) + 105
            .ColWidth(GRDC_BAS_LIB_EMPLOI) = .ColWidth(GRDC_BAS_LIB_EMPLOI) + 150
        End With
    End If
    frm(FRM_RECHERCHER).Visible = True
    grd(GRD_HAUT).Visible = True
    grd(GRD_BAS).Visible = True

    ' Le MsgBox si necessaire
    If cmbox.Visible Then
        lbl(LBL_COMBOBOX).Visible = True
        With cmbox
            If .ListCount = 1 Then
                message = "Il y a un matricule redondant: "
            Else ' .ListCount > 1
                message = "Il y a " & .ListCount & " matricules redondants: "
            End If
            For I = 0 To .ListCount - 1
                message = message & vbCrLf & "- " & .List(I)
            Next I
        End With
        Call MsgBox(message, vbExclamation + vbOKOnly, "Matricule:")
    End If

End Sub

Private Sub associer(ByVal v_row_haut As Integer, ByVal v_row_bas As Integer)
'****************************************************************************
' ENTRÉE: les lignes à associer
' SORTIE: associer les deux lignes selectionnées après quelques vérifications
'****************************************************************************
    Dim sql As String, mon_service As String, mon_poste As String
    Dim modifier_matricule As Boolean
    Dim I As Integer, j As Integer, reponse As Integer, nbr As Integer, p As Integer
    Dim modifier_nom As Integer, modifier_prenom As Integer, cr As Integer
    Dim mon_num_service As Long, mon_num_poste As Long, lnb As Long
    Dim mess As Variant
    Dim rs As rdoResultset, rs2 As rdoResultset

    With grd(GRD_HAUT)
        modifier_nom = 0
        modifier_prenom = 0
        modifier_matricule = False
        ' Confirmer la modification s'il y en a
        If .TextMatrix(v_row_haut, GRDC_HAUT_MATRICULE) <> "" Then
            If .TextMatrix(v_row_haut, GRDC_HAUT_MATRICULE) <> grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_MATRICULE) Then
                reponse = MsgBox("Voulez-vous remplacer le matricule " & .TextMatrix(v_row_haut, GRDC_HAUT_MATRICULE) _
                                 & " par " & grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_MATRICULE), _
                                 vbYesNo + vbDefaultButton2 + vbQuestion, "Association")
                If reponse = vbNo Then
                    Exit Sub
                End If
                modifier_matricule = True
            End If
        Else
            modifier_matricule = True
            g_nbr_sans_matricule = g_nbr_sans_matricule - 1
        End If
        ' Confirmation si les PRENOMS diffèrent
        If .TextMatrix(v_row_haut, GRDC_HAUT_NOM) = grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_NOM) Then
            If left$(.TextMatrix(v_row_haut, GRDC_HAUT_PRENOM), g_nbr_car_prenom) _
            <> left$(grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_PRENOM), g_nbr_car_prenom) Then
                mess = "La personne que vous voulez associer, " & .TextMatrix(v_row_haut, GRDC_HAUT_NOM) _
                     & ", a un prénom différent à celui du fichier d'importation:" _
                     & vbCrLf & vbCrLf & "Dictionnaire : " & .TextMatrix(v_row_haut, GRDC_HAUT_PRENOM) _
                     & vbCrLf & "Nouveau PRENOM : " & grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_PRENOM) _
                     & vbCrLf & vbCrLf & "Confirmez-vous l'association ?"
                cr = question_asso(mess)
                If cr = 0 Then
                    Exit Sub
                End If
                modifier_prenom = cr
            End If
        ' Confirmation si les NOMS different
        Else ' .TextMatrix(v_row_haut, GRDC_HAUT_NOM) <> grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_NOM)
            mess = "Le NOM de cette personne est différent de celui du fichier d'importation" _
                     & vbCrLf & vbCrLf & "Dictionnaire : " & .TextMatrix(v_row_haut, GRDC_HAUT_NOM) _
                            & " " & .TextMatrix(v_row_haut, GRDC_HAUT_PRENOM) _
                     & vbCrLf & "Nouveau NOM et PRENOM : " & grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_NOM) _
                            & " " & grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_PRENOM) _
                     & vbCrLf & vbCrLf & "Confirmez-vous l'association ?"
            cr = question_asso(mess)
            If cr = 0 Then
                Exit Sub
            End If
            modifier_nom = cr
            If left$(.TextMatrix(v_row_haut, GRDC_HAUT_PRENOM), g_nbr_car_prenom) _
            <> left$(grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_PRENOM), g_nbr_car_prenom) Then
                modifier_prenom = cr
            End If
        End If
        ' MAJ de la table Synchro
        If ajouter_synchro(v_row_haut, v_row_bas) = P_ERREUR Then
            Exit Sub
        End If
        ' Mettre à jour la table [Utilisateur]
        If Odbc_Update("Utilisateur", "U_Num", _
                       "WHERE u_num=" & .TextMatrix(v_row_haut, GRDC_HAUT_U_NUM), _
                       "U_Importe", True) = P_ERREUR Then
            Exit Sub
        End If
        If modifier_matricule Then
            If Odbc_Update("Utilisateur", "U_Num", _
                           "WHERE u_num=" & .TextMatrix(v_row_haut, GRDC_HAUT_U_NUM), _
                           "U_Matricule", grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_MATRICULE)) = P_ERREUR Then
                Exit Sub
            End If
            .TextMatrix(v_row_haut, GRDC_HAUT_MATRICULE) = grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_MATRICULE)
        End If
        If modifier_nom = 1 Then
            If Odbc_Update("Utilisateur", "U_Num", _
                           "WHERE u_num=" & .TextMatrix(v_row_haut, GRDC_HAUT_U_NUM), _
                           "U_Nom", grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_NOM)) = P_ERREUR Then
                Exit Sub
            End If
            .TextMatrix(v_row_haut, GRDC_HAUT_NOM) = grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_NOM)
        ElseIf modifier_nom = 2 Then
            If Odbc_Update("Utilisateur", "U_Num", _
                           "WHERE u_num=" & .TextMatrix(v_row_haut, GRDC_HAUT_U_NUM), _
                           "U_NomJunon", grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_NOM)) = P_ERREUR Then
                Exit Sub
            End If
        End If
        If modifier_prenom = 1 Then
            If Odbc_Update("Utilisateur", "U_Num", _
                           "WHERE u_num=" & .TextMatrix(v_row_haut, GRDC_HAUT_U_NUM), _
                           "U_Prenom", grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_PRENOM)) = P_ERREUR Then
                Exit Sub
            End If
            .TextMatrix(v_row_haut, GRDC_HAUT_PRENOM) = grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_PRENOM)
        ElseIf modifier_prenom = 2 Then
            If Odbc_Update("Utilisateur", "U_Num", _
                           "WHERE u_num=" & .TextMatrix(v_row_haut, GRDC_HAUT_U_NUM), _
                           "U_PrenomJunon", grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_PRENOM)) = P_ERREUR Then
                Exit Sub
            End If
        End If
        ' Alimenter la table [UtilMouvement]
        If modifier_matricule Or modifier_nom = 1 Or modifier_prenom = 1 Then
            If remplir_utilmouvement(modifier_nom, modifier_prenom, modifier_matricule, v_row_haut, v_row_bas) = P_ERREUR Then
                Exit Sub
            End If
        End If
        ' Les pastilles deviennent VERTES
        .col = GRDC_HAUT_IMPORTE
        Set .CellPicture = imglst.ListImages(IMG_IMPORTE).Picture
        .TextMatrix(v_row_haut, GRDC_HAUT_VERT_ROUGE) = IMPORTE
        grd(GRD_BAS).col = GRDC_BAS_IMPORTE
        Set grd(GRD_BAS).CellPicture = imglst.ListImages(IMG_IMPORTE).Picture
        grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_VERT_ROUGE) = IMPORTE
        ' Remettre les couleurs des pastilles
'        sql = "SELECT * FROM Utilisateur WHERE U_kb_actif=True AND U_Actif=TRUE AND U_Importe=TRUE AND U_Num<>" _
'              & .TextMatrix(v_row_haut, GRDC_HAUT_U_NUM) & " ORDER BY U_Matricule"
'        If Odbc_SelectV(sql, rs) = P_ERREUR Then
'            Exit Sub
'        End If
'        While Not rs.EOF
'            For i = 1 To .Rows - 1
'                If .TextMatrix(i, GRDC_HAUT_U_NUM) = rs("U_Num").Value _
'                        And .TextMatrix(i, GRDC_HAUT_VERT_ROUGE) = MATRICULE_REDONDANT Then
'                    .TextMatrix(i, GRDC_HAUT_VERT_ROUGE) = IMPORTE
'                    .Row = i
'                    .col = GRDC_HAUT_IMPORTE
'                    Set .CellPicture = imglst.ListImages(IMG_IMPORTE).Picture
'                End If
'            Next i
'            rs.MoveNext
'        Wend
'        rs.Close
        ' Mettre à jour le COMBOBOX
'        cmbox.Clear
'        Call remplir_cbo_matricules_double
        ' Remettre les lignes en mode selection
        Call mettre_lignes_selectionnees(v_row_haut, v_row_bas)
        ' Désactiver le bouton CMD_ASSOCIER
        cmd(CMD_ASSOCIER).Enabled = False

        ' mettre à jour le label g_nbr_sans_matricule
        If g_nbr_sans_matricule = 0 Then
            lbl(LBL_GRD_HAUT).Caption = g_grd_haut_caption
        Else
            lbl(LBL_GRD_HAUT).Caption = g_grd_haut_caption & " (" & g_nbr_sans_matricule & " sans matricule)"
        End If

        If g_nbr_sans_matricule = 0 Then
            If g_mode_traitement = MODE_VERIF_AVANT_IMPORT Then
                Call tout_est_associe
                If g_tout_est_associe Then
                    Call MsgBox("Il ne reste plus de personne à associer." & vbCrLf & vbCrLf _
                              & "L'importation va être lancée.", vbInformation + vbOKOnly, _
                                "Fin de la vérification")
                    Call quitter(False)
                End If
            End If
        End If
    End With

End Sub

Private Sub associer_auto()

    Dim lig As Integer
    
    frmPatience.Visible = True
    Me.pgb2.Max = grd(GRD_HAUT).Rows - 1
    Me.pgb2.Value = 0
    For lig = 1 To grd(GRD_HAUT).Rows - 1
        Call chercher(GRD_HAUT, lig, vbKeySpace)
        Me.pgb2.Value = Me.pgb2.Value + 1
        If cmd(CMD_ASSOCIER).Enabled Then
            Call associer(grd(GRD_HAUT).Row, grd(GRD_BAS).Row)
            If g_tout_est_associe Then
                Exit Sub
            End If
            If lig > 1 Then
                grd(GRD_HAUT).TopRow = lig - 1
            End If
        End If
    Next lig
    frmPatience.Visible = False
    
End Sub

Private Sub chercher(ByVal v_index_grd As Integer, ByVal v_row As Integer, ByVal v_keycode As Integer)
' Bare d'espace: chercher le nom dans l'autre grid
' Les touches de 'A' à 'Z': accéder à la ligne contenant le nom commançant par cette lettre

    Dim I As Integer, ligne_tempo As Integer, col_nom As Integer

    ' Désactiver le bouton ASSOCIER
    cmd(CMD_ASSOCIER).Enabled = False
    ' Déselectionner les lignes lors de la recherche
    Call deselectionner_lignes(v_index_grd)
    ' le numéro de la colonne du NOM à rechercher et/ou "synchroniser" avec l'autre grid
    col_nom = IIf(v_index_grd = GRD_HAUT, GRDC_HAUT_NOM, GRDC_BAS_NOM)

    If v_keycode <> vbKeySpace And v_keycode <> vbKeyReturn Then ' une touche ALPHABETIQUE ********
        Call recherche_alphanumerique(v_index_grd, v_row, v_keycode)
    Else ' ********************* la touche d'ESPACE, ENTER ou un Double-Click *********************
        ligne_tempo = -1
        ' Récupérer le caractère de la recherche pour les autres touches
        g_ancienne_key = Asc(Mid$(grd(v_index_grd).TextMatrix(v_row, col_nom), 1, 1))
        With grd(GRD_HAUT)
            '                                *************************************************************************
            If v_index_grd = GRD_HAUT Then ' ********************* Double-clique sur le GRD_HAUT *********************
            '                                *************************************************************************
                ' Le MATRICULE est vide dans le GRD_HAUT
                If .TextMatrix(v_row, GRDC_HAUT_MATRICULE) = "" Then
                    For I = 1 To grd(GRD_BAS).Rows - 1
                        ' Si le même NOM dans les deux grids
                        If UCase(grd(GRD_BAS).TextMatrix(I, GRDC_BAS_NOM)) = UCase(.TextMatrix(v_row, GRDC_HAUT_NOM)) Then
                            ligne_tempo = I ' pour se positionner même si on ne trouve pas de PRENOM
                            ' Rechercher le(s) g_nbr_car_prenom premier(s) caractère(s) du PRENOM
                            If UCase(Mid$(grd(GRD_BAS).TextMatrix(I, GRDC_BAS_PRENOM), 1, g_nbr_car_prenom)) _
                            <> UCase(Mid$(.TextMatrix(v_row, GRDC_HAUT_PRENOM), 1, g_nbr_car_prenom)) Then
                                GoTo i_suivant_haut
                            End If
                            ' selection des lignes
                            Call mettre_lignes_selectionnees(v_row, I)
                            ' Associer ssi GRD_BAS.img = R
                            If grd(GRD_BAS).TextMatrix(I, GRDC_BAS_VERT_ROUGE) = NON_IMPORTE Then
                                If .TextMatrix(v_row, GRDC_HAUT_VERT_ROUGE) = MATRICULE_REDONDANT Then
                                    cmd(CMD_ASSOCIER).Enabled = False
                                Else
                                    cmd(CMD_ASSOCIER).Enabled = True
                                End If
                                Exit Sub
                            End If
                            cmd(CMD_ASSOCIER).Enabled = True
                            'Exit Sub
                        End If ' même NOM
i_suivant_haut:
                    Next I
                    ' Pas de résultat pour le NOM, rechercher un nom à peu près similaire
                    If ligne_tempo = -1 Then
                        Call recherche_approximative(GRD_BAS, v_row)
                    ' Se positionner même si on n'a que le NOM
                    ElseIf ligne_tempo <> -1 Then ' => ligne_tempo<>0
                        Call mettre_lignes_selectionnees(v_row, ligne_tempo)
                        cmd(CMD_ASSOCIER).Enabled = False
                    End If
                Else ' GRD_HAUT.MATRICULE<>'' => recherche sur MATRICULE uniquement
                    For I = 1 To grd(GRD_BAS).Rows - 1
                        If .TextMatrix(v_row, GRDC_HAUT_MATRICULE) = grd(GRD_BAS).TextMatrix(I, GRDC_BAS_MATRICULE) Then
                            ' selection des lignes
                            Call mettre_lignes_selectionnees(v_row, I)
                            If grd(GRD_BAS).TextMatrix(I, GRDC_BAS_VERT_ROUGE) = NON_IMPORTE Then
                                If .TextMatrix(v_row, GRDC_HAUT_VERT_ROUGE) <> MATRICULE_REDONDANT Then
                                    cmd(CMD_ASSOCIER).Enabled = True
                                Else
                                    cmd(CMD_ASSOCIER).Enabled = False
                                End If
                                Exit Sub
                            Else ' img = V ou DOUBLONS
                                ' Associer ssi img = R
                            End If
                            Exit Sub
                        ElseIf .TextMatrix(v_row, GRDC_HAUT_NOM) = grd(GRD_BAS).TextMatrix(I, GRDC_BAS_NOM) Then
                            ligne_tempo = I
                        End If
                    Next I
                    ' Pas de résultat pour le MATRICULE ni le NOM, rechercher un nom à peu près similaire
                    If ligne_tempo = -1 Then
                        Call recherche_approximative(GRD_BAS, v_row)
                    ' Se positionner même si on n'a pas trouvé de MATRICULEs mais le même NOM
                    ElseIf ligne_tempo <> -1 Then
                        ' Trier le grid du BAS sur le NOM
                        grd(GRD_BAS).col = GRDC_BAS_NOM
                        grd(GRD_BAS).Sort = 1
                        Call mettre_lignes_selectionnees(v_row, ligne_tempo)
                        grd(GRD_BAS).TopRow = ligne_tempo
                        cmd(CMD_ASSOCIER).Enabled = False
                    End If
                End If
            '      **********************************************************************************************
            Else ' ******************************** Double-clique sur le GRD_BAS ********************************
            '      **********************************************************************************************
                ' parcourir les GRD_HAUT
                For I = 1 To .Rows - 1
                    ' Le MATRICULE est vide dans le GRD_HAUT
                    If .TextMatrix(I, GRDC_HAUT_MATRICULE) = "" Then
                        ' Si le même NOM dans les deux grids
                        If UCase(grd(GRD_BAS).TextMatrix(v_row, GRDC_BAS_NOM)) = UCase(.TextMatrix(I, GRDC_HAUT_NOM)) Then
                            ligne_tempo = I ' pour se positionner même si on ne trouve pas de PRENOM
                            ' Rechercher le(s) g_nbr_car_prenom premier(s) caractère(s) du PRENOM
                            If UCase(Mid$(grd(GRD_BAS).TextMatrix(v_row, GRDC_BAS_PRENOM), 1, g_nbr_car_prenom)) _
                            <> UCase(Mid$(.TextMatrix(I, GRDC_HAUT_PRENOM), 1, g_nbr_car_prenom)) Then
                                GoTo i_suivant_bas
                            End If
                            ' selection des lignes
                            Call mettre_lignes_selectionnees(I, v_row)
                            ' Associer ssi GRD_BAS.img = R
                            If grd(GRD_BAS).TextMatrix(v_row, GRDC_BAS_VERT_ROUGE) = NON_IMPORTE Then
                                If .TextMatrix(I, GRDC_HAUT_VERT_ROUGE) = MATRICULE_REDONDANT Then
                                    cmd(CMD_ASSOCIER).Enabled = False
                                Else
                                    cmd(CMD_ASSOCIER).Enabled = True
                                End If
                                Exit Sub
                            End If
                            cmd(CMD_ASSOCIER).Enabled = True
                            Exit Sub
                        End If ' même NOM dans les deux grids
                    Else ' GRD_HAUT.MATRICULE<>'' => recherche sur MATRICULE uniquement
                        If .TextMatrix(I, GRDC_HAUT_MATRICULE) = grd(GRD_BAS).TextMatrix(v_row, GRDC_BAS_MATRICULE) Then
                            ' selection des lignes
                            Call mettre_lignes_selectionnees(I, v_row)
                            If grd(GRD_BAS).TextMatrix(v_row, GRDC_BAS_VERT_ROUGE) = NON_IMPORTE Then
                                If .TextMatrix(I, GRDC_HAUT_VERT_ROUGE) <> MATRICULE_REDONDANT Then
                                    cmd(CMD_ASSOCIER).Enabled = True
                                Else
                                    cmd(CMD_ASSOCIER).Enabled = False
                                End If
                                Exit Sub
                            End If
                            Exit Sub
                        ElseIf .TextMatrix(I, GRDC_HAUT_NOM) = grd(GRD_BAS).TextMatrix(v_row, GRDC_BAS_NOM) Then
                            ligne_tempo = I
                        End If
                    End If ' MATRICULE est vide
i_suivant_bas:
                Next I
                ' Pas de résultat pour le MATRICULE ni le NOM, rechercher un nom à peu près similaire
                If ligne_tempo = -1 Then
                    Call recherche_approximative(GRD_HAUT, v_row)
                ' Se positionner même si on n'a que le NOM
                ElseIf ligne_tempo <> -1 Then
                    ' Trier le grid du HAUT sur le NOM
                    .col = GRDC_HAUT_NOM
                    .Sort = 1
                    Call mettre_lignes_selectionnees(ligne_tempo, v_row)
                    cmd(CMD_ASSOCIER).Enabled = False
                End If
            End If ' Choix du grid
        End With
    End If

End Sub

Private Sub colorer_colonne_triee(ByVal v_index_grd As Integer, ByVal v_old_col As Integer, ByVal v_new_col As Integer)
' ***********************************************
' Distinguer la colonne triée des autres colonnes
' ***********************************************
    Dim I As Integer

    With grd(v_index_grd)
        For I = 1 To .Rows - 1
            .Row = I
            ' Décolorer
            .col = v_old_col
            .CellBackColor = COLOR_PAS_DE_TRI
            ' Colorer
            '.col = v_new_col
            '.CellBackColor = COLOR_DU_TRI
        Next I
        For I = 1 To .Rows - 1
            .Row = I
            ' Décolorer
            '.col = v_old_col
            '.CellBackColor = COLOR_PAS_DE_TRI
            ' Colorer
            .col = v_new_col
            .CellBackColor = COLOR_DU_TRI
        Next I
        .col = 0
        .Row = 0
    End With

End Sub

Private Sub deselectionner_lignes(ByVal v_index_grd As Integer)

    Dim autre_grd As Integer

    If v_index_grd = GRD_HAUT Then
        autre_grd = GRD_BAS
    Else
        autre_grd = GRD_HAUT
    End If
    With grd(autre_grd)
        .col = 0
    End With

End Sub

Private Function formater_prenom(ByVal v_prenom As String) As String
' Mettre la 1° lettre (avec/sans séparateur) en majiscule, et le reste en miniscule

    Dim sous_str As String
    Dim nbr As Integer, I As Integer

    nbr = STR_GetNbchamp(v_prenom, " ")
    If nbr > 1 Then
        For I = 0 To nbr - 1
            If LCase$(STR_GetChamp(v_prenom, " ", I)) = "de" _
                    Or LCase$(STR_GetChamp(v_prenom, " ", I)) = "du" _
                    Or LCase$(STR_GetChamp(v_prenom, " ", I)) = "des" _
                    Or LCase$(STR_GetChamp(v_prenom, " ", I)) = "la" _
                    Or LCase$(STR_GetChamp(v_prenom, " ", I)) = "le" Then
                If I = 0 Then
                    formater_prenom = STR_GetChamp(v_prenom, " ", I)
                Else
                    formater_prenom = formater_prenom & " " & STR_GetChamp(v_prenom, " ", I)
                End If
            Else
                sous_str = UCase$(Mid$(STR_GetChamp(v_prenom, " ", I), 1, 1)) & LCase$(Mid$(STR_GetChamp(v_prenom, " ", I), 2))
                If I = 0 Then
                    formater_prenom = sous_str
                Else
                    formater_prenom = formater_prenom & " " & sous_str
                End If
            End If
        Next I
        Exit Function
    End If
    nbr = STR_GetNbchamp(v_prenom, "-")
    If nbr > 1 Then
        For I = 0 To nbr - 1
            If LCase$(STR_GetChamp(v_prenom, "-", I)) = "de" _
                    Or LCase$(STR_GetChamp(v_prenom, "-", I)) = "du" _
                    Or LCase$(STR_GetChamp(v_prenom, "-", I)) = "des" _
                    Or LCase$(STR_GetChamp(v_prenom, "-", I)) = "la" _
                    Or LCase$(STR_GetChamp(v_prenom, "-", I)) = "le" Then
                If I = 0 Then
                    formater_prenom = STR_GetChamp(v_prenom, " ", I)
                Else
                    formater_prenom = formater_prenom & "-" & STR_GetChamp(v_prenom, " ", I)
                End If
            Else
                sous_str = UCase$(Mid$(STR_GetChamp(v_prenom, "-", I), 1, 1)) & LCase$(Mid$(STR_GetChamp(v_prenom, "-", I), 2))
                If I = 0 Then
                    formater_prenom = sous_str
                Else
                    formater_prenom = formater_prenom & "-" & sous_str
                End If
            End If
        Next I
        Exit Function
    End If

    formater_prenom = UCase$(Mid$(v_prenom, 1, 1)) & LCase$(Mid$(v_prenom, 2))

End Function

Private Sub initialiser()

    Dim reponse As String
    Dim I As Integer
    Dim lng As Long
    Dim nomfich As String

    cmd(CMD_ASSOCIER).left = (Me.width / 4) * 3 - (cmd(CMD_QUITTER).width / 2)
    frm(FRM_PRINCIPALE).Caption = "Vérification avec le fichier d'importation: "
    cmd(CMD_QUITTER).left = (Me.width / 2) - (cmd(CMD_QUITTER).width / 2)
    cmd(CMD_QUITTER).Visible = True

    frm(FRM_PRINCIPALE).Caption = frm(FRM_PRINCIPALE).Caption & " " & p_nom_fichier_importation & " " & IIf(p_est_sur_serveur, "(serveur)", "(local)")
    frmPatience.left = (Me.width / 2) - (frmPatience.width / 2)
    frmPatience.Top = (Me.Height / 2) - (frmPatience.Height / 2)
    frmPatience.Visible = True
    g_nbr_sans_matricule = 0
    g_tout_est_associe = False
    Me.MousePointer = 11

    g_col_tri_haut = GRDC_HAUT_MATRICULE ' l'état du tri
    g_sens_tri_haut = 1

    With grd(GRD_HAUT) ' GRID DU HAUT -------------------------------------------------
        .FormatString = "u_num||MATRIC.|NOM|PRENOM| code|SERVICE| CODE->|POSTE|||importe_mais_rouge|plusieurs_potes?"
        .ScrollTrack = True
        .ColWidth(GRDC_HAUT_U_NUM) = 0
        .ColWidth(GRDC_HAUT_IMPORTE) = 240
        .ColWidth(GRDC_HAUT_MATRICULE) = 1000 ' 1400
        .ColWidth(GRDC_HAUT_NOM) = 2080
        .ColWidth(GRDC_HAUT_PRENOM) = 1500
        .ColWidth(GRDC_HAUT_CODE_SECTION) = 0
        .ColWidth(GRDC_HAUT_LIB_SECTION) = 2910
        .ColWidth(GRDC_HAUT_CODE_EMPLOI) = 0
        .ColWidth(GRDC_HAUT_LIB_EMPLOI) = 2910 ' 2640
        .ColWidth(GRDC_HAUT_ENCORE_EMPLOI) = 250
        .ColWidth(GRDC_HAUT_MODIFIER) = 250
        .ColWidth(GRDC_HAUT_VERT_ROUGE) = 0
        .ColWidth(GRDC_HAUT_PLUSIEURS_POSTES) = 0
        .ColAlignment(GRDC_HAUT_MATRICULE) = flexAlignLeftCenter
        .ColAlignment(GRDC_HAUT_NOM) = flexAlignLeftCenter
        .ColAlignment(GRDC_HAUT_PRENOM) = flexAlignLeftCenter
        .ColAlignment(GRDC_HAUT_LIB_SECTION) = flexAlignLeftCenter
        .ColAlignment(GRDC_HAUT_LIB_EMPLOI) = flexAlignLeftCenter
        .SelectionMode = flexSelectionByRow
        For I = 1 To .Cols - 1
            .Row = 0
            .col = I
            .CellFontBold = True
        Next I
        g_grd_haut_caption = lbl(LBL_GRD_HAUT).Caption
    End With
    With grd(GRD_BAS) ' GRID DU BAS ---------------------------------------------------
        .FormatString = "||MATRIC.|NOM|PRENOM| CODE->|SECTION| CODE->|EMPLOI|importe_mais_rouge"
        .ScrollTrack = True
        .ColWidth(GRDC_BAS_U_NUM) = 0
        .ColWidth(GRDC_BAS_IMPORTE) = 240
        .ColWidth(GRDC_BAS_MATRICULE) = 1000 ' 1400
        .ColWidth(GRDC_BAS_NOM) = 2080 '1896
        .ColWidth(GRDC_BAS_PRENOM) = 1500 '1900
        .ColWidth(GRDC_BAS_CODE_SECTION) = 800 '1000
        .ColWidth(GRDC_BAS_LIB_SECTION) = 2370 '2705
        .ColWidth(GRDC_BAS_CODE_EMPLOI) = 800 '1000
        .ColWidth(GRDC_BAS_LIB_EMPLOI) = 2370 '2720
        .ColWidth(GRDC_BAS_VERT_ROUGE) = 0
        .ColAlignment(GRDC_BAS_MATRICULE) = flexAlignLeftCenter
        .ColAlignment(GRDC_BAS_NOM) = flexAlignLeftCenter
        .ColAlignment(GRDC_BAS_PRENOM) = flexAlignLeftCenter
        .ColAlignment(GRDC_BAS_CODE_SECTION) = flexAlignLeftCenter
        .ColAlignment(GRDC_BAS_LIB_SECTION) = flexAlignLeftCenter
        .ColAlignment(GRDC_BAS_CODE_EMPLOI) = flexAlignLeftCenter
        .ColAlignment(GRDC_BAS_LIB_EMPLOI) = flexAlignLeftCenter
        .SelectionMode = flexSelectionByRow
        .Row = 0
        .col = GRDC_BAS_MATRICULE
        For I = 1 To .Cols - 1
            .Row = 0
            .col = I
            .CellFontBold = True
        Next I
    End With

    Me.Refresh
    
    ' MAJ la table des Utilisateurs pour mettre IMPORTE à FALSE et KB_ACTIF à TRUE pour tous ceux qui n'ont pas de MATRICULE
    Call Odbc_Cnx.Execute("UPDATE Utilisateur SET U_Importe=FALSE, U_kb_actif=True WHERE U_Matricule='' AND U_Importe=True")
    g_redim_grd_haut = True
    If g_mode_traitement = MODE_VERIFICATION Then
        Call remplir_grid
    Else ' g_mode_traitement = MODE_VERIF_AVANT_IMPORT
        Call verif_avant_import
    End If
    cmd(CMD_ASSOCIER).Visible = True

    On Error Resume Next
    grd(GRD_HAUT).SetFocus
    On Error GoTo 0
    
End Sub

Private Sub maj_ligne(ByVal v_row As Integer, ByVal v_u_num As Integer)
' Cette procedure est appelée lorsqu'on clique sur la cellule GRDC_MODIFER
' Elle met-à-jour la ligne selectionnée + la ligne correspondante dans GRD_BAS

    Dim sql As String, ancien_matricule As String, ancienne_couleur As String
    Dim I As Integer, nbr_spm As Integer, nbr As Integer, j As Integer, p As Integer
    Dim rs As rdoResultset

    With grd(GRD_HAUT)
        ancien_matricule = .TextMatrix(v_row, GRDC_HAUT_MATRICULE)
        ancienne_couleur = .TextMatrix(v_row, GRDC_HAUT_VERT_ROUGE)

        If Odbc_SelectV("SELECT * FROM Utilisateur" _
                      & " Where U_kb_actif=True AND U_ExterneFich = False" _
                      & " And U_Actif = True" _
                      & " And U_Num = " & v_u_num, rs) = P_ERREUR Then
            Exit Sub
        End If
        If rs.EOF Then ' personne supprimée et/ou inactive
            If ancien_matricule = "" Then g_nbr_sans_matricule = g_nbr_sans_matricule - 1
            If .Rows = 2 Then
                .Rows = 1
            Else
                .RemoveItem v_row
                For I = 1 To grd(GRD_BAS).Rows - 1
                    grd(GRD_BAS).Row = I
                    grd(GRD_BAS).col = GRDC_BAS_IMPORTE
                    If grd(GRD_BAS).TextMatrix(I, GRDC_BAS_MATRICULE) = ancien_matricule _
                            And grd(GRD_BAS).TextMatrix(I, GRDC_BAS_VERT_ROUGE) = IMPORTE Then
                            'And grd(GRD_BAS).CellPicture = imglst.ListImages(IMG_IMPORTE).Picture Then
                        grd(GRD_BAS).col = GRDC_BAS_IMPORTE
                        Set grd(GRD_BAS).CellPicture = imglst.ListImages(IMG_NON_IMPORTE).Picture
                        grd(GRD_BAS).TextMatrix(I, GRDC_BAS_VERT_ROUGE) = NON_IMPORTE
                    End If
                Next I
            End If
            rs.Close
        Else ' le personne est toujours active / existe dans le dictionnaire
            rs.Close
            sql = "SELECT * FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & .TextMatrix(v_row, GRDC_HAUT_U_NUM)
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                Exit Sub
            End If
            .Row = v_row
            .col = GRDC_HAUT_IMPORTE
            ' Des modifications importatntes => non importée
            If Not rs("U_Importe").Value Then
                Set .CellPicture = imglst.ListImages(IMG_NON_IMPORTE).Picture
                If ancien_matricule <> "" Then
                    For I = 1 To grd(GRD_BAS).Rows - 1
                        If grd(GRD_BAS).TextMatrix(I, GRDC_BAS_MATRICULE) = ancien_matricule Then
                            grd(GRD_BAS).Row = I
                            grd(GRD_BAS).col = GRDC_BAS_IMPORTE
                            Set grd(GRD_BAS).CellPicture = imglst.ListImages(IMG_NON_IMPORTE).Picture
                            grd(GRD_BAS).TextMatrix(I, GRDC_BAS_VERT_ROUGE) = NON_IMPORTE
                        End If
                    Next I
                End If
            End If
            .TextMatrix(v_row, GRDC_HAUT_MATRICULE) = rs("U_Matricule").Value
            ' calculer le g_nbr_sans_matricule
            If ancien_matricule = "" Then
                If rs("U_Matricule").Value <> "" Then
                    g_nbr_sans_matricule = g_nbr_sans_matricule - 1
                End If
            ElseIf rs("U_Matricule").Value = "" Then
                g_nbr_sans_matricule = g_nbr_sans_matricule + 1
            End If
            .TextMatrix(v_row, GRDC_HAUT_NOM) = UCase$(rs("U_Nom").Value)
            .TextMatrix(v_row, GRDC_HAUT_PRENOM) = formater_prenom(rs("U_Prenom").Value)
            ' Remplir des 4 colonnes concernant le SERVICE et le POTE
            Call remplir_srv_fct(rs("U_Num").Value, rs("U_SPM").Value, rs("U_Po_Princ").Value, v_row)
            ' Remettre la bonne couleur des pastilles
            ' si le matricule a changé
'            If .TextMatrix(v_row, GRDC_HAUT_MATRICULE) <> ancien_matricule Then
'                .Row = v_row
'                .col = GRDC_HAUT_IMPORTE
'                Set .CellPicture = imglst.ListImages(IMG_NON_IMPORTE).Picture
'                .TextMatrix(i, GRDC_HAUT_VERT_ROUGE) = NON_IMPORTE
'                For i = 1 To grd(GRD_BAS).Rows - 1
'                    grd(GRD_BAS).Row = i
'                    grd(GRD_BAS).col = GRDC_BAS_IMPORTE
'                    If grd(GRD_BAS).TextMatrix(i, GRDC_BAS_MATRICULE) = ancien_matricule _
'                            And grd(GRD_BAS).TextMatrix(i, GRDC_BAS_VERT_ROUGE) = IMPORTE Then
'                            'And grd(GRD_BAS).CellPicture = imglst.ListImages(IMG_IMPORTE).Picture Then
'                        grd(GRD_BAS).col = GRDC_BAS_IMPORTE
'                        Set grd(GRD_BAS).CellPicture = imglst.ListImages(IMG_NON_IMPORTE).Picture
'                        grd(GRD_BAS).TextMatrix(i, GRDC_BAS_VERT_ROUGE) = NON_IMPORTE
'                    End If
'                Next i
'            End If
            Call remplir_srv_fct(rs("U_Num").Value, rs("U_SPM").Value, rs("U_Po_Princ").Value, v_row)
            rs.Close
            ' Eviter la MAJ dans le GRD_BAS si on n'a pas modifié: MATRICULE et/ou U_ACTIF et/ou U_ExterneFich
            If utilisateur_modifie(v_row) Then
                For I = 1 To grd(GRD_BAS).Rows - 1
                    grd(GRD_BAS).Row = I
                    grd(GRD_BAS).col = GRDC_BAS_IMPORTE
                    If grd(GRD_BAS).TextMatrix(I, GRDC_BAS_MATRICULE) = ancien_matricule _
                            And grd(GRD_BAS).TextMatrix(I, GRDC_BAS_VERT_ROUGE) = IMPORTE Then
                            'And grd(GRD_BAS).CellPicture = imglst.ListImages(IMG_IMPORTE).Picture Then
                        grd(GRD_BAS).col = GRDC_BAS_IMPORTE
                        Set grd(GRD_BAS).CellPicture = imglst.ListImages(IMG_NON_IMPORTE).Picture
                        grd(GRD_BAS).TextMatrix(I, GRDC_BAS_VERT_ROUGE) = NON_IMPORTE
                    End If
                Next I
            End If
        End If
        ' Compter le nombre de personnes restantes & Remplir le ComboBox si necessaire
        'g_nbr_sans_matricule = 0 ' =*=*=**=*
        cmbox.Clear
        Call remplir_cbo_matricules_double
        ' compter la dernière ligne
        'If .TextMatrix(.Rows - 1, GRDC_HAUT_MATRICULE) = "" Then g_nbr_sans_matricule = g_nbr_sans_matricule + 1 ' =*=*=*=*
        ' Le GRD_BAS
        For I = 0 To cmbox.ListCount - 1
            For j = 1 To grd(GRD_BAS).Rows - 1
                grd(GRD_BAS).Row = j
                grd(GRD_BAS).col = GRDC_BAS_IMPORTE
                If grd(GRD_BAS).TextMatrix(j, GRDC_BAS_MATRICULE) = cmbox.List(I) _
                        And grd(GRD_BAS).TextMatrix(j, GRDC_BAS_VERT_ROUGE) = IMPORTE Then
                    grd(GRD_BAS).col = GRDC_BAS_IMPORTE
                    Set grd(GRD_BAS).CellPicture = imglst.ListImages(IMG_NON_IMPORTE).Picture
                    grd(GRD_BAS).TextMatrix(j, GRDC_BAS_VERT_ROUGE) = NON_IMPORTE
                End If
            Next j
        Next I
        lbl(LBL_GRD_HAUT).Caption = g_grd_haut_caption & " (" & g_nbr_sans_matricule & " sans matricule)"
        ' Mettre la ligne en mode selection
        If .Rows - 1 > 1 Then
            .col = 0
            If v_row = .Rows Then ' on a supprimé la ligne => on selectionne celle du dessus
                .Row = v_row - 1
                .RowSel = v_row - 1
            Else
                .Row = v_row
                .RowSel = v_row
            End If
            .ColSel = .Cols - 1
            If v_row > 4 Then
                .TopRow = v_row - 4
            End If
        ElseIf .Rows - 1 = 1 Then
            .col = 0
            .Row = 1
            .RowSel = 1
            .ColSel = .Cols - 1
        ElseIf .Rows - 1 = 0 Then
            .col = 0
            .Enabled = False
            If g_mode_traitement = MODE_VERIF_AVANT_IMPORT Then
                Call MsgBox("Il ne reste plus de personne à associer." & vbCrLf & vbCrLf _
                          & "Vous pouvez maintenant procéder à l'importation.", vbQuestion + vbOKOnly, _
                            "Fin de la vérification")
                Call quitter(False)
            End If
        End If

    ' S'il n'y a pas de ScrollBar => redimensionner les grids
    If grd(GRD_HAUT).Rows <= 10 Then
        If g_redim_grd_haut Then
            .ColWidth(GRDC_HAUT_LIB_SECTION) = .ColWidth(GRDC_HAUT_LIB_SECTION) + 150
            .ColWidth(GRDC_HAUT_LIB_EMPLOI) = .ColWidth(GRDC_HAUT_LIB_EMPLOI) + 105
            g_redim_grd_haut = False
        End If
    End If
    End With

End Sub

Private Sub maj_utilis_apre_recher(ByVal v_u_num As Long)
' *******************************************************************
' Mettre à jour une ligne après une modification dans le dictionnaire
' *******************************************************************
    Dim I As Integer
    Dim actif As Boolean
    Dim rs As rdoResultset

    With grd(GRD_HAUT)
        If .Rows - 1 = 0 Then Exit Sub
        For I = 1 To .Rows - 1
            If .TextMatrix(I, GRDC_HAUT_U_NUM) = v_u_num Then
                Call maj_ligne(I, v_u_num)
                Exit Sub
            End If
        Next I
        ' Si on est là, c'est que la ligne n'existe pas dans le grid
        If Odbc_SelectV("SELECT * FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & v_u_num, rs) = P_ERREUR Then
            Exit Sub
        End If
        If Not rs.EOF Then
            If rs("U_Actif").Value Then
                .AddItem ""
                .Row = .Rows - 1
                .col = GRDC_HAUT_IMPORTE
                Set .CellPicture = imglst.ListImages(IMG_NON_IMPORTE).Picture
                .TextMatrix(.Rows - 1, GRDC_HAUT_VERT_ROUGE) = IIf(rs("U_Importe").Value, IMPORTE, NON_IMPORTE)
                .TextMatrix(.Rows - 1, GRDC_HAUT_U_NUM) = v_u_num
                .TextMatrix(.Rows - 1, GRDC_HAUT_MATRICULE) = rs("U_Matricule").Value
                .TextMatrix(.Rows - 1, GRDC_HAUT_NOM) = rs("U_Nom").Value
                .TextMatrix(.Rows - 1, GRDC_HAUT_PRENOM) = formater_prenom(rs("U_Prenom").Value)
                ' Remplir des 4 colonnes concernant le SERVICE et le POSTE
                Call remplir_srv_fct(v_u_num, rs("U_SPM").Value, rs("U_Po_Princ").Value, .Rows - 1)
                ' La dernière colonne
                .TextMatrix(.Rows - 1, GRDC_HAUT_MODIFIER) = " Accéder"
                .col = GRDC_HAUT_MODIFIER
                .CellBackColor = P_VERT
                .CellFontBold = True
                ' Le nombre de personne sans matricule
                If rs("U_Matricule").Value = "" Then
                    g_nbr_sans_matricule = g_nbr_sans_matricule + 1
                End If
            End If
            rs.Close
        End If
    End With

End Sub

Private Function matricule_existe(ByVal v_matricule As String, ByRef r_importe As Boolean) As Boolean
' ENTREE: le V_MATRICULE a chrcher dans GRD_HAUT, le boolean à modifier si le V_MATRICULE existe
' SORTIE: le matricule existe _et_ l'état de ce matricule (importé ou non)

    Dim I As Integer

    With grd(GRD_HAUT)
        For I = 1 To .Rows - 1
            If .TextMatrix(I, GRDC_HAUT_MATRICULE) = v_matricule Then
                If .TextMatrix(I, GRDC_HAUT_VERT_ROUGE) = IMPORTE Then
                    r_importe = True
                Else
                    r_importe = False
                End If
                matricule_existe = True
                Exit Function
            End If
        Next I
    End With

    r_importe = False
    matricule_existe = False

End Function

Private Sub mettre_lignes_selectionnees(ByVal v_row_haut As Integer, ByVal v_row_bas As Integer)
'**************************************
' Remettre les lignes en mode selection
'**************************************
    With grd(GRD_HAUT)
        .col = 0
        .Row = v_row_haut
        .RowSel = v_row_haut
        .ColSel = .Cols - 1
        If Not .RowIsVisible(v_row_haut) Or (v_row_haut > .TopRow + 8) Then ' la ligne du haut n'est pas visible
            If v_row_haut > 10 Then
                .TopRow = v_row_haut - 8
            Else
                .TopRow = 10 - v_row_haut + 1
            End If
        ' Else ' la ligne du haut est visible
        End If
    End With
    With grd(GRD_BAS)
        .col = 0
        .Row = v_row_bas
        .RowSel = v_row_bas
        .ColSel = .Cols - 1
        If Not .RowIsVisible(v_row_bas) Or (v_row_haut > .TopRow + 8) Then ' la ligne d'en bas n'est pas visible
            If v_row_bas > 10 Then
                .TopRow = v_row_bas - 8
            Else
                .TopRow = 10 - v_row_bas + 1
            End If
        ' Else ' la ligne d'en bas est visible
        End If
    End With

End Sub

Private Function question_asso(ByVal v_mess As Variant) As Integer

    Dim tbl_libelle(2) As String, tbl_tooltip(2) As String
    Dim cr As Integer
    Dim frm As Form
    
    tbl_libelle(0) = "Non"
    tbl_tooltip(0) = ""
    tbl_libelle(1) = "Oui et mettre à jour le dictionnaire"
    tbl_tooltip(1) = ""
    tbl_libelle(2) = "Oui sans mettre à jour le dictionnaire"
    tbl_tooltip(2) = ""
    Set frm = Com_Message
    cr = Com_Message.AppelFrm(v_mess, _
                              "", _
                              tbl_libelle(), _
                              tbl_tooltip())
    Set frm = Nothing
    
    question_asso = cr
    
End Function

Private Sub quitter(ByVal v_bforce As Boolean)

    Dim reponse As Integer
    Dim s As String
    
    If Not p_traitement_background And v_bforce Then
        g_tout_est_associe = False
        Unload Me
        Exit Sub
    End If

    ' Vérification avant "Unload Me" pour pouvoir tester si tout le monde est associé ou non
    If g_mode_traitement = MODE_VERIF_AVANT_IMPORT Then
        Call tout_est_associe
        If Not g_tout_est_associe Then
            s = "La vérification n'est pas encore terminée." _
                           & " Vous n'avez certainement pas encore géré au moins un des cas suivants:" _
                           & vbCrLf & vbCrLf & "* Il reste des personne sans matricule," & vbCrLf _
                           & "* Il reste des matricules en double," & vbCrLf _
                           & "* Il reste des personnes non associées avec le fichier d'importation." & vbCrLf & vbCrLf
            If True Or p_traitement_background Then
                p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> " & s
                p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "L 'importation ne peut pas être effectuée."
            End If
            reponse = MsgBox(s & "Afin de finaliser cette procédure, voulez-vous revenir sur la vérification ?", _
                           vbQuestion + vbYesNo, "Attention:")
            If reponse = vbYes Then
                Exit Sub
            End If
            Call MsgBox("L'importation ne peut pas être effectuée.", vbInformation + vbOKOnly, "")
        End If
    End If

    Unload Me

End Sub

Private Sub recherche_alphanumerique(ByVal v_index_grd As Integer, ByVal v_row As Integer, ByVal v_keycode As Integer)
' ***********************************************************************************
' Rechercher une pesronne en se basant sur le premier caractère dans la colonne triée
' ***********************************************************************************
    Dim I As Integer, col_u_num As Integer, col_tri As Integer

    ' le numéro de la colonne triée du grid
    col_tri = IIf(v_index_grd = GRD_HAUT, g_col_tri_haut, g_col_tri_bas)
    ' le numéro de la colonne du GRDC_U_NUM à rechercher dans le grid
    col_u_num = IIf(v_index_grd = GRD_HAUT, GRDC_HAUT_U_NUM, GRDC_BAS_U_NUM)

    If v_keycode <> g_ancienne_key Then ' une nouvelle touche ?
        g_ancienne_key = v_keycode
    End If
    With grd(v_index_grd)
        .SelectionMode = flexSelectionByRow
        If v_row < .Rows - 1 Then
            For I = v_row + 1 To .Rows - 1
                If Mid$(.TextMatrix(I, col_tri), 1, 1) = Chr(v_keycode) Then
                    .Row = I
                    If I > 4 Then
                        .TopRow = I - 4
                    End If
                    .Row = I
                    .col = col_u_num
                    .RowSel = I
                    .ColSel = .Cols - 1
                    If Not .RowIsVisible(I) Then .TopRow = I
                    .SetFocus
                    Exit Sub
                End If
            Next I
            For I = 1 To v_row ' On recommence depuis le début jusqu'à la ligne en cours
                If Mid$(.TextMatrix(I, col_tri), 1, 1) = Chr(v_keycode) Then
                    .Row = I
                    If I > 4 Then
                        .TopRow = I - 4
                    End If
                    .col = col_u_num
                    .RowSel = I
                    .ColSel = .Cols - 1
                    If Not .RowIsVisible(I) Then .TopRow = I
                    .SetFocus
                    Exit Sub
                End If
            Next I
            ' On a pas trouvé la lettre recherchée => on déselectionne la ligne en cours
            .RowSel = .Row
            .ColSel = col_u_num
        Else ' v_row = .Rows - 1
            For I = 1 To v_row ' On recommence depuis le début jusqu'à la ligne en cours
                If Mid$(.TextMatrix(I, col_tri), 1, 1) = Chr(v_keycode) Then
                    .Row = I
                    If I > 4 Then
                        .TopRow = I - 4
                    End If
                    .col = col_u_num
                    .RowSel = I
                    .ColSel = .Cols - 1
                    If Not .RowIsVisible(I) Then .TopRow = I
                    .SetFocus
                    Exit Sub
                End If
            Next I
        End If
    End With

End Sub

Private Sub recherche_approximative(ByVal v_index_grid As Integer, ByVal v_row As Integer)
' ************************************************************************************
' Positionner la ligne du grid sur un nom approximativemant similaire au nom recherché
' v_index_grid = le grid sur le quelle on doit faire la recherche et v_row = la ligne
' ************************************************************************************
    Dim nom_recherche As String, nom_trouve As String, str As String
    Dim I As Integer, j As Integer, le_plus_long As Integer, la_ligne As Integer

    la_ligne = -1
    le_plus_long = 0
    nom_recherche = grd(IIf(v_index_grid = GRD_HAUT, GRD_BAS, GRD_HAUT)).TextMatrix(v_row, IIf(v_index_grid = GRD_HAUT, GRDC_BAS_NOM, GRDC_HAUT_NOM))
    ' parcourrir le grid dans lequel on doit effectuer la recherche
    With grd(v_index_grid)
        For I = 1 To .Rows - 1
            nom_trouve = .TextMatrix(I, IIf(v_index_grid = GRD_HAUT, GRDC_HAUT_NOM, GRDC_BAS_NOM))
            str = ""
            ' Parcourrir tous les caractères du nom recherché
            For j = 1 To Len(nom_recherche)
                str = str & Mid$(nom_recherche, j, 1)
                If str = Mid$(nom_trouve, 1, j) Then
                    ' récupérer la longueur du nom trouvé et la ligne correspondante
                    If le_plus_long < j Then
                        le_plus_long = j
                        la_ligne = I
                    End If
                End If
            Next j
        Next I
    End With
    ' Mettre les lignes en surbrillance (si on a trouvé une similitude)
    If la_ligne <> -1 Then
        Call mettre_lignes_selectionnees(IIf(v_index_grid = GRD_HAUT, la_ligne, v_row), IIf(v_index_grid = GRD_HAUT, v_row, la_ligne))
    End If

End Sub

Private Sub rechercher_personne()
' ********************************************
' Rechercher une personne dans le dictionnaire
' Le NOM est avantagé sur le MATRICULE
' ********************************************
    Dim sql As String, text_nom As String, text_matricule As String, str As String
    Dim reafficher As Boolean
    Dim nbr_affiche As Integer
    Dim rs As rdoResultset
    Dim frm As Form

    reafficher = False
lab_afficher:
    nbr_affiche = 0
    text_nom = txt(TXT_NOM).Text
    text_matricule = txt(TXT_MATRICULE).Text
    sql = "SELECT * FROM Utilisateur"
    If text_nom <> "" Then
        If text_matricule <> "" Then ' on a un NOM et un MATRICULE
            sql = sql & " WHERE U_kb_actif=True AND U_Nom LIKE " & Odbc_upper() & "(" & Odbc_String(text_nom) & ")" _
                      & " AND U_Matricule LIKE '%" & text_matricule & "%'ORDER BY U_Nom"
        Else ' que le NOM
            sql = sql & " WHERE U_kb_actif=True AND U_Nom LIKE " & Odbc_upper() & "(" & Odbc_String(text_nom & "%") & ") ORDER BY U_Nom"
        End If
    Else ' que le MATRICULE
        sql = sql & " WHERE U_kb_actif=True AND U_Matricule LIKE " & Odbc_String(text_matricule & "%") & " ORDER BY U_Nom"
    End If
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    Call CL_Init
    If rs.EOF Then ' pas de resultat
        ' renouveler la recherche sur le NOM
        If txt(TXT_NOM).Text <> "" And txt(TXT_MATRICULE).Text <> "" Then
            txt(TXT_MATRICULE).Text = ""
            GoTo lab_afficher
        End If
        If Not reafficher Then ' éviter le message si on a modifié la personne
            Call MsgBox("Aucune personne comportant " & IIf(text_nom <> "", IIf(text_matricule <> "", _
                        "ce nom ni ce matricule", "ce nom"), IIf(text_matricule <> "", " ce matricule", "")) _
                      & " n'a été trouvée dans le dictionnaire !" & vbCrLf & vbCrLf _
                      & "Veuillez modifier vos critères de recherche.", vbExclamation + vbOKOnly, "")
        End If
        rs.Close
        Exit Sub
    Else ' on a une liste
        While Not rs.EOF
            str = ""
            str = IIf(rs("U_Actif").Value, "[]", "[Inactive]")
            str = str & " " & IIf(rs("U_ExterneFich").Value, "[Externe au fichier]", "[]")
            Call CL_AddLigne(str & vbTab & UCase$(rs("U_Nom").Value) & " " & formater_prenom(rs("U_Prenom").Value) _
                            & vbTab & rs("U_Matricule").Value, _
                             rs("U_Num").Value, _
                             rs("U_Actif").Value, _
                             False)
            nbr_affiche = nbr_affiche + 1
            rs.MoveNext
        Wend
        rs.Close
    End If
    Call CL_InitTaille(0, -15)
    Call CL_InitTitreHelp("Le résultat de la recherche sur le" & IIf(text_nom <> "", IIf(text_matricule <> "", _
                          " NOM (" & text_nom & ") et le MATRICULE (" & text_matricule & ")", _
                          " NOM (" & text_nom & ")"), IIf(text_matricule <> "", _
                          " MATRICULE (" & text_matricule & ")", "")), "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1
    ' Tester le choix
    If CL_liste.retour = 1 Then ' QUITTER
        Exit Sub
    Else                        ' OK
        Set frm = PrmPersonne
        If Not PrmPersonne.AppelFrm(CL_liste.lignes(CL_liste.pointeur).num, "") Then
            Set frm = Nothing
            GoTo lab_afficher
        Else
            Set frm = Nothing
            Call maj_utilis_apre_recher(CL_liste.lignes(CL_liste.pointeur).num)
            reafficher = True
            If nbr_affiche > 1 Then GoTo lab_afficher
        End If
    End If

End Sub

Private Sub remplir_cbo_matricules_double()

    Dim I As Integer, j As Integer, p As Integer, nbr As Integer

    pgb2.Max = pgb2.Max + grd(GRD_HAUT).Rows - 2
    pgb.Value = 0
    Me.LbGauge2.Caption = "Vérification des matricules en double"
    Me.Refresh
    
    With grd(GRD_HAUT)
        nbr = 0
        For I = 1 To .Rows - 2
            pgb2.Value = pgb2.Value + 1
            If pgb.Value = pgb.Max Then pgb.Value = 0
            pgb.Value = pgb.Value + 1
            If .TextMatrix(I, GRDC_HAUT_MATRICULE) <> "" Then
                For j = I + 1 To .Rows - 1
                    If .TextMatrix(I, GRDC_HAUT_MATRICULE) = .TextMatrix(j, GRDC_HAUT_MATRICULE) Then
                        ' Etat de la pastille
                        .Row = j
                        .col = GRDC_HAUT_IMPORTE
                        Set .CellPicture = imglst.ListImages(IMG_MATRICULE_REDONDANT).Picture
                        .TextMatrix(.Row, GRDC_HAUT_VERT_ROUGE) = MATRICULE_REDONDANT
                        .Row = I
                        Set .CellPicture = imglst.ListImages(IMG_MATRICULE_REDONDANT).Picture
                        .TextMatrix(.Row, GRDC_HAUT_VERT_ROUGE) = MATRICULE_REDONDANT
                        ' Mettre à jour la table Utilisateur, afin de rester cohérant lors de la recharge du grid
                        If I <> j Then ' éviter de le faire deux fois!
                            Call Odbc_Cnx.Execute("UPDATE Utilisateur SET U_Importe=False, U_kb_actif=True WHERE U_Num=" _
                                                & .TextMatrix(I, GRDC_HAUT_U_NUM))
                            Call Odbc_Cnx.Execute("UPDATE Utilisateur SET U_Importe=False, U_kb_actif=True WHERE U_Num=" _
                                                & .TextMatrix(j, GRDC_HAUT_U_NUM))
                        End If
                        ' Le ComboBox
                        For p = 0 To cmbox.ListCount - 1
                            If .TextMatrix(I, GRDC_HAUT_MATRICULE) = cmbox.List(p) Then
                                GoTo combobox_suivant
                            End If
                        Next p
                        cmbox.AddItem Item:=.TextMatrix(I, GRDC_HAUT_MATRICULE), Index:=nbr
                        cmbox.ItemData(nbr) = .TextMatrix(I, GRDC_HAUT_U_NUM)
                        nbr = nbr + 1
combobox_suivant:
                    End If
                Next j
            End If
        Next I
        If cmbox.ListCount > 0 Then
'            lbl(LBL_COMBOBOX).Caption = "Liste des MATRICULES redondants"
'            lbl(LBL_COMBOBOX).Enabled = True
        Else ' plus de matricule en double
            ' enlever les pastilles à deux couleurs s'il y en a =*=*=*=
' ARRET : VERIFIER L'UTILITE
            pgb2.Max = pgb2.Max + grd(GRD_HAUT).Rows - 2
            Me.LbGauge2.Caption = "Vérification des matricules en double"
            Me.Refresh
            For I = 1 To .Rows - 1 ' =*=*=*=
                If .TextMatrix(I, GRDC_HAUT_VERT_ROUGE) = MATRICULE_REDONDANT Then ' =*=*=*=
                    .Row = I ' =*=*=*=
                    .col = GRDC_HAUT_IMPORTE ' =*=*=*=
                    Set .CellPicture = imglst.ListImages(IMG_NON_IMPORTE).Picture ' =*=*=*=
                    .TextMatrix(I, GRDC_HAUT_VERT_ROUGE) = NON_IMPORTE ' =*=*=*=
                End If ' =*=*=*=
            Next I ' =*=*=*=
            cmbox.Visible = False
            lbl(LBL_COMBOBOX).Visible = False
        End If
    End With

End Sub

Private Sub remplir_grid()

    Dim nomfich As String, sext As String
    Dim fd As Integer, pos As Integer
    Dim rs As rdoResultset
    Dim ligne_lu As Variant

    ' ******************************** Remplissage du grid du HAUT ********************************
    ' g_mode_traitement = MODE_VERIFICATION
    If Odbc_SelectV("SELECT * FROM Utilisateur" _
                  & " WHERE U_kb_actif=True AND U_Actif=TRUE AND U_ExterneFich=FALSE" _
                  & " ORDER BY U_Nom, U_Prenom", rs) = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
        pgb2.Value = 0
        pgb2.Max = rs.RowCount
        Me.LbGauge2.Caption = "Initialisation / KaliBottin"
        Me.Refresh
    End If
    
    With grd(GRD_HAUT)
        While Not rs.EOF
            If pgb.Value = pgb.Max Then
                pgb.Value = 0
            End If
            pgb.Value = pgb.Value + 1
            pgb2.Value = pgb2.Value + 1
            Call ajouter_ligne_haut(rs("U_Num").Value, _
                                    IIf(rs("U_Importe").Value, IMPORTE, NON_IMPORTE), _
                                    rs("U_Matricule").Value, _
                                    rs("U_Nom").Value, _
                                    rs("U_Prenom").Value, _
                                    rs("U_SPM").Value, _
                                    rs("U_Po_Princ").Value)
            If rs("U_Matricule").Value = "" Then
                g_nbr_sans_matricule = g_nbr_sans_matricule + 1
            End If
            rs.MoveNext
        Wend
        rs.Close
        ' Trier le grid du HAUT (déjà trié sur le U_Nom et le , U_Prenom)
        g_sens_tri_haut = 1
        g_col_tri_haut = GRDC_HAUT_NOM
        Call colorer_colonne_triee(GRD_HAUT, 0, GRDC_HAUT_NOM)

        ' Remplissage du ComboBox si necessaire
        Call remplir_cbo_matricules_double
        If cmbox.ListCount > 0 Then
            cmbox.Visible = True
        End If
        
        If .Rows > 1 Then ' il est possible d'avoir un grid vide (si la table utilisateur est vide par ex.)
            .Row = 1
            .TopRow = 1
        Else
            .col = 0
        End If

        lbl(LBL_GRD_HAUT).Caption = g_grd_haut_caption & " (" & g_nbr_sans_matricule & " sans matricule)"
    End With

    ' ******************************** Remplissage du grid du BAS ********************************
    ' Ouverture du fichier en lecture seule
    If p_est_sur_serveur Then
        pos = InStrRev(p_nom_fichier_importation, ".")
        sext = Mid$(p_nom_fichier_importation, pos)
        nomfich = p_chemin_appli & "\tmp\personnel_" & format(Date, "hhmmss") & sext
        If KF_GetFichier(p_nom_fichier_importation, nomfich) = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
    Else
        nomfich = p_nom_fichier_importation
    End If
    
    If FICH_OuvrirFichier(nomfich, FICH_LECTURE, fd) = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    Me.LbGauge2.Caption = "Initialisation / fichier d'importation"
    Me.Refresh
    While Not EOF(fd)
        Line Input #fd, ligne_lu
        pgb2.Max = pgb2.Max + 1
    Wend
    Close #fd
    
    If FICH_OuvrirFichier(nomfich, FICH_LECTURE, fd) = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If

    ' Commencer de la ligne indiquée (1°, 2°, ...)
    ' For i = 0 To g_ligne_debut_fichier -1
        'Line Input #fd, ligne_lu
    ' Next i

    With grd(GRD_BAS)
        While Not EOF(fd)
            If pgb.Value = pgb.Max Then
                pgb.Value = 0
            End If
            pgb.Value = pgb.Value + 1
            pgb2.Value = pgb2.Value + 1
            Line Input #fd, ligne_lu
            If ligne_lu <> "" Then
                Call ajouter_ligne_bas(ligne_lu)
            End If
        Wend
        ' Trier le grid du BAS
        .col = GRDC_BAS_NOM
        g_col_tri_bas = GRDC_BAS_NOM
        g_sens_tri_bas = 1
        .Sort = 1
        Call colorer_colonne_triee(GRD_BAS, 0, GRDC_BAS_NOM)
        If .Rows > 1 Then ' il est possible d'avoir un grid vide (si le fichier est vide)
            .Row = 1
            .TopRow = 1
        Else
            .col = 0
        End If
    End With ' **************************************************************************************

    ' Fermeture du fichier
    Close #fd
    If p_est_sur_serveur Then
        Call FICH_EffacerFichier(nomfich, False)
    End If
    
    Call apres_remplir_grid

End Sub

Private Sub remplir_srv_fct(ByVal v_unum As Long, _
                            ByVal v_uspm As String, _
                            ByVal v_unumposte As Long, _
                            ByVal v_row As Integer)
' ENTREE: le num utilisateur + le chemin de son poste + la ligne en question
' SORTIE: ajouter les colonnes SERVICE et POSTE dans le GRD_HAUT + l'image s'il y a plusieurs postes

    Dim sql As String, mon_service As String, mon_poste As String, my_spm As String
    Dim I As Integer, nbr As Integer, nbr_spm As Integer
    Dim mon_num_service As Long, mon_num_poste As Long
    Dim rs As rdoResultset

    mon_num_poste = v_unumposte
    
    ' Le SRVICE
    sql = "SELECT po_srvnum, SRV_Nom FROM Poste, Service" _
        & " WHERE po_num=" & mon_num_poste _
        & " and SRV_Num=po_srvnum"
    If Odbc_RecupVal(sql, mon_num_service, mon_service) = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    
    ' La Fonction
    sql = "SELECT FT_Libelle FROM FctTrav, Poste" _
        & " WHERE FT_Num=PO_FTNum AND PO_Num=" & mon_num_poste
    If Odbc_RecupVal(sql, mon_poste) = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    
    ' Le nombre de postes
    nbr_spm = STR_GetNbchamp(v_uspm, "|")
        
    With grd(GRD_HAUT)
        .TextMatrix(v_row, GRDC_HAUT_CODE_SECTION) = mon_num_service
        .TextMatrix(v_row, GRDC_HAUT_LIB_SECTION) = mon_service
        .TextMatrix(v_row, GRDC_HAUT_CODE_EMPLOI) = mon_num_poste
        .TextMatrix(v_row, GRDC_HAUT_LIB_EMPLOI) = mon_poste
        ' Déterminer l'image s'il y a plusieurs postes
        .Row = v_row
        .col = GRDC_HAUT_ENCORE_EMPLOI
        .CellPictureAlignment = POS_CENTRE
        If nbr_spm > 1 Then ' plusieurs postes pour cette personne
            Set .CellPicture = imglst.ListImages(IMG_SUITE).Picture
            .TextMatrix(v_row, GRDC_HAUT_PLUSIEURS_POSTES) = True
        Else
            Set .CellPicture = LoadPicture("")
             .TextMatrix(v_row, GRDC_HAUT_PLUSIEURS_POSTES) = False
        End If
        .TextMatrix(v_row, GRDC_HAUT_PLUSIEURS_POSTES) = False
    End With

End Sub

Private Sub OLD_remplir_srv_fct(ByVal v_unum As Long, _
                            ByVal v_uspm As String, _
                            ByVal v_row As Integer)
' ENTREE: le num utilisateur + le chemin de son poste + la ligne en question
' SORTIE: ajouter les colonnes SERVICE et POSTE dans le GRD_HAUT + l'image s'il y a plusieurs postes

    Dim sql As String, mon_service As String, mon_poste As String, my_spm As String
    Dim I As Integer, nbr As Integer, nbr_spm As Integer
    Dim mon_num_service As Long, mon_num_poste As Long
    Dim rs As rdoResultset

    ' Le nombre de postes
    nbr_spm = STR_GetNbchamp(v_uspm, "|")
    If nbr_spm = 1 Then ' ----------------- UN SEUL POSTE ------------------
        my_spm = STR_GetChamp(v_uspm, "|", 0)
        
        ' Récupérer le nombre (services + le poste), au mois nbr=2
        nbr = STR_GetNbchamp(my_spm, ";")
        mon_num_service = Mid$(STR_GetChamp(my_spm, ";", nbr - 2), 2)
        mon_num_poste = Mid$(STR_GetChamp(my_spm, ";", nbr - 1), 2)

        ' Le SRVICE
        sql = "SELECT SRV_Nom FROM Service WHERE SRV_Num=" & mon_num_service
        If Odbc_RecupVal(sql, mon_service) = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
        ' Le POSTE
        sql = "SELECT FT_Libelle FROM FctTrav, Poste" _
            & " WHERE FT_Num=PO_FTNum AND PO_Num=" & mon_num_poste
        If Odbc_RecupVal(sql, mon_poste) = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If

        With grd(GRD_HAUT)
            .TextMatrix(v_row, GRDC_HAUT_CODE_SECTION) = mon_num_service
            .TextMatrix(v_row, GRDC_HAUT_LIB_SECTION) = mon_service
            .TextMatrix(v_row, GRDC_HAUT_CODE_EMPLOI) = mon_num_poste
            .TextMatrix(v_row, GRDC_HAUT_LIB_EMPLOI) = mon_poste
            .Row = v_row
            .col = GRDC_HAUT_ENCORE_EMPLOI
            .CellPictureAlignment = POS_CENTRE
            Set .CellPicture = LoadPicture("")
            .TextMatrix(v_row, GRDC_HAUT_PLUSIEURS_POSTES) = False
        End With
    Else ' ------------------- PLUSIEURS POSTES ------------------
        ' Parcourir tous les postes, s'arrêter au premier poste conïcidant avec la table Synchro
        For I = 0 To nbr_spm - 1
            my_spm = STR_GetChamp(v_uspm, "|", I)
            ' Récupérer le nombre (services + le poste), au mois nbr=2
            nbr = STR_GetNbchamp(my_spm, ";")
            mon_num_poste = Mid$(STR_GetChamp(my_spm, ";", nbr - 1), 2)
            sql = "SELECT SYNC_SPNum FROM Synchro" ' WHERE SYNC_SPNum=" & mon_num_poste
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                Call quitter(True)
                Exit Sub
            End If

            While Not rs.EOF
                If mon_num_poste = rs("SYNC_SPNum").Value Then
                    rs.Close
                    mon_num_service = Mid$(STR_GetChamp(my_spm, ";", nbr - 2), 2)
                    ' Le SRVICE
                    sql = "SELECT SRV_Nom FROM Service WHERE SRV_Num=" & mon_num_service
                    If Odbc_RecupVal(sql, mon_service) = P_ERREUR Then
                        Call quitter(True)
                        Exit Sub
                    End If
                    ' Le POSTE
                    sql = "SELECT FT_Libelle FROM FctTrav, Poste" _
                        & " WHERE FT_Num=PO_FTNum AND PO_Num=" & mon_num_poste
                    If Odbc_RecupVal(sql, mon_poste) = P_ERREUR Then
                        Call quitter(True)
                        Exit Sub
                    End If

                    With grd(GRD_HAUT)
                        .TextMatrix(v_row, GRDC_HAUT_CODE_SECTION) = mon_num_service
                        .TextMatrix(v_row, GRDC_HAUT_LIB_SECTION) = mon_service
                        .TextMatrix(v_row, GRDC_HAUT_CODE_EMPLOI) = mon_num_poste
                        .TextMatrix(v_row, GRDC_HAUT_LIB_EMPLOI) = mon_poste
                    End With
                    GoTo lab_suite_grd
                End If
                rs.MoveNext
            Wend
            rs.Close
        Next I

        ' On n'a pas trouvé dans la table Synchro, on affiche le dernier poste trouvé
        mon_num_service = Mid$(STR_GetChamp(my_spm, ";", nbr - 2), 2)
        ' Le SRVICE
        sql = "SELECT SRV_Nom FROM Service WHERE SRV_Num=" & mon_num_service
        If Odbc_RecupVal(sql, mon_service) = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
        ' Le POSTE
        sql = "SELECT FT_Libelle FROM FctTrav, Poste" _
            & " WHERE FT_Num=PO_FTNum AND PO_Num=" & mon_num_poste
        If Odbc_RecupVal(sql, mon_poste) = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
lab_suite_grd:
        With grd(GRD_HAUT)
            .TextMatrix(v_row, GRDC_HAUT_CODE_SECTION) = mon_num_service
            .TextMatrix(v_row, GRDC_HAUT_LIB_SECTION) = mon_service
            .TextMatrix(v_row, GRDC_HAUT_CODE_EMPLOI) = mon_num_poste
            .TextMatrix(v_row, GRDC_HAUT_LIB_EMPLOI) = mon_poste
            ' Déterminer l'image s'il y a plusieurs postes
            If nbr_spm > 1 Then ' plusieurs postes pour cette personne
                .Row = v_row
                .col = GRDC_HAUT_ENCORE_EMPLOI
                .CellPictureAlignment = POS_CENTRE
                Set .CellPicture = imglst.ListImages(IMG_SUITE).Picture
                .TextMatrix(v_row, GRDC_HAUT_PLUSIEURS_POSTES) = True
            Else
                .TextMatrix(v_row, GRDC_HAUT_PLUSIEURS_POSTES) = False
            End If
        End With
    End If

End Sub

Private Function remplir_utilmouvement(ByVal v_nom As Integer, ByVal v_prenom As Integer, _
                                       ByVal v_matricule As Boolean, _
                                       ByVal v_row_haut As Integer, ByVal v_row_bas As Integer) As Integer
'****************************************************************************************************
' Insérer dans [UtilMouvement] les lignes des changements opérés sur les coordonnées d'un utilisateur
'****************************************************************************************************
    If v_nom = 1 Then ' **************** CHANGEMENT DU NOM ****************
        If P_InsertIntoUtilmouvement(grd(GRD_HAUT).TextMatrix(v_row_haut, GRDC_HAUT_U_NUM), _
                    "M", "NOM=" & grd(GRD_HAUT).TextMatrix(v_row_haut, GRDC_HAUT_NOM) & ";" _
                                & grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_NOM) & ";", _
                    0) = P_ERREUR Then
            GoTo lab_erreur
        End If
    End If
    If v_prenom = 1 Then ' ************* CHANGEMENT DU PRENOM *************
        If P_InsertIntoUtilmouvement(grd(GRD_HAUT).TextMatrix(v_row_haut, GRDC_HAUT_U_NUM), _
                    "M", "PRENOM=" & grd(GRD_HAUT).TextMatrix(v_row_haut, GRDC_HAUT_PRENOM) & ";" _
                                   & grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_PRENOM) & ";", _
                    0) = P_ERREUR Then
            GoTo lab_erreur
        End If
    End If
    If v_matricule Then ' ********** CHANGEMENT DE MATRICULE **********
        If P_InsertIntoUtilmouvement(grd(GRD_HAUT).TextMatrix(v_row_haut, GRDC_HAUT_U_NUM), _
                    "M", "MATRICULE=" & grd(GRD_HAUT).TextMatrix(v_row_haut, GRDC_HAUT_MATRICULE) & ";" _
                                      & grd(GRD_BAS).TextMatrix(v_row_bas, GRDC_BAS_MATRICULE) & ";", _
                    0) = P_ERREUR Then
            GoTo lab_erreur
        End If
    End If

    remplir_utilmouvement = P_OK
    Exit Function

lab_erreur:
    remplir_utilmouvement = P_ERREUR

End Function

Private Sub tout_est_associe()
' Fonction appelée depuis Quitter(False) lorsque le g_mode_traitement = MODE_VERIF_AVANT_IMPORT
' Déterminer si le GRD_HAUT est OK et on peut passer à l'importation

    Dim I As Integer

    With grd(GRD_HAUT)
        For I = 1 To .Rows - 1
            ' Il suffit d'une autre couleur que le VERT pour ne pas passer à l'importation
            If .TextMatrix(I, GRDC_HAUT_VERT_ROUGE) <> IMPORTE Then
                g_tout_est_associe = False
                Exit Sub
            End If
        Next I
    End With

    g_tout_est_associe = True

End Sub

Private Function utilisateur_modifie(ByVal v_row As Integer) As Boolean
' Appelée uniquement depuis maj_ligne() pour déterminer si les coordonnées importatntes
' de la personne ont été modifiées: SUPPRESSION (INNACTIVITEE)

    Dim sql As String, msg As String
    Dim I As Integer
    Dim rs As rdoResultset

    sql = "SELECT * FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & grd(GRD_HAUT).TextMatrix(v_row, GRDC_HAUT_U_NUM)
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Call quitter(True)
        Exit Function
    End If
    If rs.EOF Then ' On a supprimé cette personne de la base !!!!
        utilisateur_modifie = True
        Exit Function
    End If
    With grd(GRD_HAUT)
        If Not rs("U_Actif").Value Or rs("U_ExterneFich").Value Then ' supprimée ou inactive
            utilisateur_modifie = True
            Exit Function
        End If
    End With

    ' Pas de modifications importantes
    utilisateur_modifie = False

End Function

Private Sub verif_avant_import()
' GRD_HAUT => Matricule='' & U_Importe='F' & Doublons
' GRD_BAS  => Fichier.Matricule !€ KaliBottin.Utilisateur

    Dim sql As String, message As String, matricule_en_cours As String, msg As String
    Dim nomfich As String, sext As String
    Dim etat_importe As Boolean
    Dim nbr_spm As Integer, I As Integer, fd As Integer, nbr As Integer, j As Integer, p As Integer
    Dim pos As Integer
    Dim rs As rdoResultset, rs2 As rdoResultset
    Dim ligne_lu As Variant
    Dim iBoucle As Integer
    
    ' g_mode_traitement = MODE_VERIF_AVANT_IMPORT
    ' Remplir le GRD_HAUT avec les SANS MATRICULE & U_Importe=FALSE & les doublons
    With grd(GRD_HAUT) ' ****************************** GRID_DU_HAUT **********************************
        ' Les Actifs à gérer dans le fichier d'importation SANS MATRICULE ou dont U_Importe=FALSE
        ' se baser sur le nom Junon s'il existe
'        sql = "SELECT * FROM Utilisateur " _
'            & " WHERE U_kb_actif=True AND U_Actif=TRUE AND U_ExterneFich=FALSE " _
'            & " AND (U_Matricule='' OR U_Importe=FALSE) " _
'            & " ORDER BY U_Nom"
        sql = "SELECT * FROM Utilisateur " _
            & " WHERE U_kb_actif=True AND U_Actif=TRUE AND U_ExterneFich=FALSE " _
            & " AND U_Matricule=''" _
            & " ORDER BY U_Nom"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
        
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            pgb2.Value = 0
            pgb2.Max = rs.RowCount
        End If
        
        While Not rs.EOF
            If pgb.Value = pgb.Max Then
                pgb.Value = 0
            End If
            pgb.Value = pgb.Value + 1
            pgb2.Value = pgb2.Value + 1
            Call ajouter_ligne_haut(rs("U_Num").Value, _
                                    NON_IMPORTE, _
                                    rs("U_Matricule").Value, _
                                    rs("U_Nom").Value, _
                                    rs("U_Prenom").Value, _
                                    rs("U_SPM").Value, _
                                    rs("U_Po_Princ").Value)
            If rs("U_Matricule").Value = "" Then
                g_nbr_sans_matricule = g_nbr_sans_matricule + 1
            End If
            rs.MoveNext
        Wend
        rs.Close

        ' Les personnes MATRICULES en DOUBLE
        sql = "SELECT U_Matricule, COUNT(*) AS NBR_DOUBLE" _
            & " FROM Utilisateur" _
            & " WHERE U_kb_actif=True AND U_Matricule<>'' AND U_Importe=TRUE" _
            & " GROUP BY U_Matricule  HAVING COUNT(U_Matricule) > 1"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
        
        
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            pgb2.Max = pgb2.Max + rs.RowCount
        End If
        
        While Not rs.EOF
            sql = "SELECT * FROM Utilisateur WHERE U_kb_actif=True AND U_Matricule='" & rs("U_Matricule").Value & "'"
            If Odbc_SelectV(sql, rs2) = P_ERREUR Then
                Call quitter(True)
                Exit Sub
            End If
            While Not rs2.EOF
                If pgb.Value = pgb.Max Then
                    pgb.Value = 0
                End If
                pgb.Value = pgb.Value + 1
                Call ajouter_ligne_haut(rs2("U_Num").Value, _
                                        IIf(rs2("U_Importe").Value, IMPORTE, NON_IMPORTE), _
                                        rs2("U_Matricule").Value, _
                                        rs2("U_Nom").Value, _
                                        rs2("U_Prenom").Value, _
                                        rs2("U_SPM").Value, _
                                        rs2("U_Po_Princ").Value)
                rs2.MoveNext
            Wend
            rs2.Close
            rs.MoveNext
        Wend
        rs.Close

        ' Trier le grid du HAUT, quel que soit le g_mode_traitement, le tri se porte sur le nom
        .col = GRDC_HAUT_NOM
        g_sens_tri_haut = 1
        g_col_tri_haut = GRDC_HAUT_NOM
        Call colorer_colonne_triee(GRD_HAUT, 0, GRDC_HAUT_NOM)

        ' Remplissage du ComboBox matricules en double si nécessaire
        Call remplir_cbo_matricules_double

        ' La vérification est terminée s'il ne reste plus de personne à associer
        If .Rows = 1 Then
            Call quitter(False)
            Exit Sub
        End If
        If .Rows > 1 Then
            .Row = 1
            .TopRow = 1
        Else
            .col = 0
        End If
        lbl(LBL_GRD_HAUT).Caption = g_grd_haut_caption & " (" & g_nbr_sans_matricule & " sans matricule)"
    End With

    If p_est_sur_serveur Then
        pos = InStrRev(p_nom_fichier_importation, ".")
        sext = Mid$(p_nom_fichier_importation, pos)
        nomfich = p_chemin_appli & "\tmp\personnel_" & format(Date, "hhmmss") & sext
        If KF_GetFichier(p_nom_fichier_importation, nomfich) = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
    Else
        nomfich = p_nom_fichier_importation
    End If
    
    ' Remplir le GRD_BAS avec les lignes du fichier :
    '   * qui ont un MATRICULE qui n'existe pas dans la table Utilisateur
    '   * les cas des doublons
    With grd(GRD_BAS) ' ****************************** GRID_DU_BAS **********************************
        ' Ouverture du fichier en lecture seule
'            For j = 0 To cmbox.ListCount - 1
'                If cmbox.List(j) = rs("U_Matricule").Value Then
                ' Ne pas exclure les doublons
'                    GoTo rs_suivant
'                End If
'            Next j

        If FICH_OuvrirFichier(nomfich, FICH_LECTURE, fd) = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If

        ' Commencer de la ligne indiquée (1°, 2°, ...)
        ' For i = 0 To g_ligne_debut_fichier -1
            'Line Input #fd, ligne_lu
        ' Next i
        While Not EOF(fd)
            Line Input #fd, ligne_lu
            pgb2.Max = pgb2.Max + 1
        Wend
        Close #fd
        
        If FICH_OuvrirFichier(nomfich, FICH_LECTURE, fd) = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
        While Not EOF(fd)
            Line Input #fd, ligne_lu
            If ligne_lu <> "" Then
                If pgb.Value = pgb.Max Then
                    pgb.Value = 0
                End If
                pgb.Value = pgb.Value + 1
                pgb2.Value = pgb2.Value + 1
                matricule_en_cours = P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_matricule, p_long_matricule, "matricule")
                sql = "SELECT U_Matricule FROM Utilisateur WHERE U_kb_actif=True AND U_Matricule='" & matricule_en_cours & "'" _
                    & " AND U_Importe=TRUE ORDER BY U_Matricule"
                If Odbc_SelectV(sql, rs) = P_ERREUR Then
                    Call quitter(True)
                    Exit Sub
                End If
                If Not rs.EOF Then ' le matricule_en_cours existe dans la table Utilisateur
                    rs.Close
                    ' Fait partie des matricules à ne pas afficher
                    GoTo ligne_suivante
                Else               ' le matricule_en_cours n'existe pas dans la table Utilisateur
                    .AddItem ""
                    .TextMatrix(.Rows - 1, GRDC_BAS_MATRICULE) = matricule_en_cours
                    .Row = .Rows - 1
                    .col = GRDC_BAS_IMPORTE
                    etat_importe = False
                    Set .CellPicture = imglst.ListImages(IMG_NON_IMPORTE).Picture
                    .TextMatrix(.Rows - 1, GRDC_BAS_VERT_ROUGE) = NON_IMPORTE
                    .TextMatrix(.Rows - 1, GRDC_BAS_NOM) = Trim$(UCase$(P_ChangerCar(P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_nom, p_long_nom, "nom"), tbcaractere_nontraite)))
                    .TextMatrix(.Rows - 1, GRDC_BAS_PRENOM) = Trim$(formater_prenom(P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_prenom, p_long_prenom, "prénom")))
                    .TextMatrix(.Rows - 1, GRDC_BAS_CODE_SECTION) = Trim$(P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_code_section, p_long_code_section, "code section"))
                    .TextMatrix(.Rows - 1, GRDC_BAS_LIB_SECTION) = Trim$(P_ChangerCar(P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_lib_section, p_long_lib_section, "libelle section"), tbcaractere_nontraite))
                    .TextMatrix(.Rows - 1, GRDC_BAS_CODE_EMPLOI) = Trim$(P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_code_emploi, p_long_code_emploi, "code emploi"))
                    .TextMatrix(.Rows - 1, GRDC_BAS_LIB_EMPLOI) = Trim$(P_ChangerCar(P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_lib_emploi, p_long_lib_emploi, "libelle emploi"), tbcaractere_nontraite))
                End If
            End If
ligne_suivante:
        Wend ' fin de lecture du fichier
        ' Fermeture du fichier
        Close #fd
        If p_est_sur_serveur Then
            Call FICH_EffacerFichier(nomfich, False)
        End If
        ' Trier le grid du BAS
        .col = GRDC_BAS_NOM
        g_col_tri_bas = GRDC_BAS_NOM
        g_sens_tri_bas = 1
        .Sort = 1
        Call colorer_colonne_triee(GRD_BAS, 0, GRDC_BAS_NOM)
        If .Rows > 1 Then
            .Row = 1
            .TopRow = 1
        Else
            .col = 0
        End If
    End With ' --------------------------------------------------------------------------------------

    ' Affichage des grids
    frmPatience.Visible = False
    Me.MousePointer = 0
    ' S'il n'y a pas de ScrollBar => redimensionner les grids
    With grd(GRD_HAUT)
        If .Rows <= 10 Then
            .ColWidth(GRDC_HAUT_LIB_SECTION) = .ColWidth(GRDC_HAUT_LIB_SECTION) + 150
            .ColWidth(GRDC_HAUT_LIB_EMPLOI) = .ColWidth(GRDC_HAUT_LIB_EMPLOI) + 105
            g_redim_grd_haut = False
        ElseIf .Rows > 1 Then
            Call MsgBox("La vérification n'est pas encore terminée." & vbCrLf _
                      & "Avant de pouvoir commencer l'importation, vous devez continuer à traiter les personnes non associées.", _
                      vbExclamation + vbOKOnly, "Attention")
        End If
    End With
    With grd(GRD_BAS)
        If .Rows <= 10 Then
            .ColWidth(GRDC_BAS_LIB_SECTION) = .ColWidth(GRDC_BAS_LIB_SECTION) + 105
            .ColWidth(GRDC_BAS_LIB_EMPLOI) = .ColWidth(GRDC_BAS_LIB_EMPLOI) + 150
        End If
    End With

    g_grd_haut_caption = "Liste des personnes à associer"
    lbl(LBL_GRD_HAUT).left = 360
    lbl(LBL_GRD_HAUT).Caption = g_grd_haut_caption & " (" & g_nbr_sans_matricule & " sans matricule)"
    lbl(LBL_GRD_HAUT).Visible = True
    lbl(LBL_GRD_BAS).Caption = "Liste des personnes du fichier dont le matricule n'existe pas dans KaliBottin"
    lbl(LBL_GRD_BAS).Visible = True
    frm(FRM_RECHERCHER).Visible = True
    grd(GRD_HAUT).Visible = True
    grd(GRD_BAS).Visible = True

    If grd(GRD_BAS).Rows = 1 Then ' le GRD_BAS est vide
        Call MsgBox("Il n'y a pas de données à afficher dans le tableau du bas." & vbCrLf & vbCrLf & "Veuillez vérifier les paramètres généraux.", _
                   vbExclamation + vbOKOnly, "Le tableau de vérification est vide:")
        grd(GRD_BAS).Enabled = False
    Else
        iBoucle = 0
        On Error Resume Next
        iBoucle = UBound(tbcaractere_nontraite)
        If iBoucle > 0 Then
            msg = ""
            For I = 1 To UBound(tbcaractere_nontraite)
               msg = msg & " - " & tbcaractere_nontraite(I) & vbCrLf
            Next I
            MsgBox "Certains caractères n'ont pas été traités : " & vbCrLf _
                   & msg & vbCrLf & "Merci d'en informer KaliTech"
        End If
    End If
    ' Le MsgBox si necessaire
    With cmbox
        If .ListCount > 0 Then
            .Visible = True
            If .ListCount = 1 Then
                message = "Il y a un matricule redondant: "
            Else ' .ListCount > 1
                message = "Il existe : " & .ListCount & " matricules redondants: "
            End If
            For I = 0 To .ListCount - 1
                message = message & vbCrLf & " * " & .List(I)
            Next I
            Call MsgBox(message, vbExclamation + vbOKOnly, "Matricule:")
        End If
    End With

End Sub

Private Sub cmbox_Click()

    Dim I As Integer, j As Integer

    With grd(GRD_HAUT)
        For I = 1 To .Rows - 1
            If .TextMatrix(I, GRDC_HAUT_U_NUM) = cmbox.ItemData(cmbox.ListIndex) Then
                ' Supprimer les autres flèches
                .Row = 0
                For j = 1 To .Cols - 1
                    .col = j
                    Set .CellPicture = LoadPicture("")
                Next j
                ' Trier le grid du HAUT selon les matricules afin de reperer les doublons
                .col = GRDC_HAUT_MATRICULE
                .Row = 0
                .Sort = 1
                g_sens_tri_haut = 1
                g_col_tri_haut = GRDC_HAUT_MATRICULE
                ' Acceder à la personne en question (matricule boublon)
                .col = GRDC_HAUT_U_NUM
                .Row = I
                .ColSel = GRDC_HAUT_MODIFIER
                .RowSel = I
                If I > 4 Then
                    .TopRow = I - 4
                End If
            End If
        Next I
        .SetFocus ' sur le grid du HAUT
    End With

End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case CMD_QUITTER
            Call quitter(False)
        Case CMD_ASSOCIER
            Call associer(grd(GRD_HAUT).Row, grd(GRD_BAS).Row)
        Case CMD_ASSOCIER_AUTO
            Call associer_auto
        Case CMD_RECHERCHER
            Call rechercher_personne
            txt(TXT_NOM).Text = ""
            txt(TXT_MATRICULE).Text = ""
    End Select

End Sub

Private Sub Form_Activate()

    If g_form_active Then Exit Sub

    g_form_active = True

    Call initialiser

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Call quitter(False)
    End If
    
End Sub

Private Sub Form_Load()

    g_form_active = False

End Sub

Private Sub grd_Click(Index As Integer)

    Dim mouse_row As Integer, mouse_col As Integer
    Dim frm As Form
    Dim B_retour As Boolean

    Select Case Index
        Case GRD_HAUT ' ******************* HAUT *******************
            With grd(GRD_HAUT)
                If .Rows - 1 = 0 Then Exit Sub
                mouse_row = .MouseRow
                mouse_col = .MouseCol
                g_grd_haut_row = mouse_row
                g_grd_haut_col = mouse_col
                ' Modification d'une personne
                If mouse_row <> 0 And mouse_col = GRDC_HAUT_MODIFIER Then
                    Set frm = PrmPersonne
                    B_retour = PrmPersonne.AppelFrm(.TextMatrix(mouse_row, GRDC_HAUT_U_NUM), "")
                    'Call PrmPersonne.AppelFrm(.TextMatrix(mouse_row, GRDC_HAUT_U_NUM), "")
                    If B_retour Then
                        Set frm = Nothing
                        ' Mettre à jour la ligne du grid(GRD_HAUT) après cet modification
                        Call maj_ligne(mouse_row, .TextMatrix(mouse_row, GRDC_HAUT_U_NUM))
                    End If
                End If
            End With
        Case GRD_BAS ' ******************* BAS *******************
            With grd(GRD_BAS)
                If .Rows - 1 = 0 Then Exit Sub
                mouse_row = .MouseRow
                mouse_col = .MouseCol
                g_grd_bas_row = mouse_row
                g_grd_bas_col = mouse_col
            End With
    End Select

End Sub

Private Sub grd_DblClick(Index As Integer)

    Dim mouse_row As Integer, mouse_col As Integer, I As Integer, col_modifier As Integer, col_importe As Integer

    If Index = GRD_HAUT Then
    Else ' Index = GRD_BAS
    End If

    With grd(Index)
        If Index = GRD_HAUT Then ' ************************ GRID DU HAUT ************************
            If .Rows - 1 = 0 Then Exit Sub
            mouse_row = .MouseRow
            mouse_col = .MouseCol
            ' Trier selon la colonne choisie dans la ligne fixe uniquement
'            If mouse_row = 0 And mouse_col <> GRDC_HAUT_MODIFIER And mouse_col <> GRDC_HAUT_IMPORTE _
                                And mouse_col <> GRDC_HAUT_ENCORE_EMPLOI Then
             If mouse_row = 0 And mouse_col <> GRDC_HAUT_MODIFIER And mouse_col <> GRDC_HAUT_ENCORE_EMPLOI Then
                .Row = 0
                If mouse_col = GRDC_HAUT_IMPORTE Then
                    .col = GRDC_HAUT_VERT_ROUGE
                    If g_col_tri_haut = mouse_col Then
                        If g_sens_tri_haut = 1 Then
                            .Sort = 2
                            g_sens_tri_haut = 2
                        Else
                            .Sort = 1
                            g_sens_tri_haut = 1
                        End If
                    Else ' tri par défaut
                        .Sort = 1
                        g_sens_tri_haut = 1
                    End If
                Else
                    .col = mouse_col
                    If g_col_tri_haut = .col Then
                        If g_sens_tri_haut = 1 Then
                            .Sort = 2
                            g_sens_tri_haut = 2
                        Else
                            .Sort = 1
                            g_sens_tri_haut = 1
                        End If
                    Else ' tri par défaut
                        .Sort = 1
                        g_sens_tri_haut = 1
                    End If
                    If .Rows > 1 Then
                        .TopRow = 1
                        .col = 0
                        .Row = 1
                    End If
                End If
                ' *************************************************************
                Call colorer_colonne_triee(GRD_HAUT, g_col_tri_haut, mouse_col)
                ' *************************************************************
                g_col_tri_haut = mouse_col
            ElseIf mouse_row <> 0 Then ' n'importe où sur le GRD_HAT sauf la ligne fixe ou la colonne GRDC_MODIFIER
                Call chercher(Index, mouse_row, vbKeySpace)
            End If
        ElseIf Index = GRD_BAS Then ' ************************ GRID DU BAS ************************
            If .Rows - 1 = 0 Then Exit Sub
            mouse_row = .MouseRow
            mouse_col = .MouseCol
            ' Trier selon la colonne choisie dans la ligne fixe uniquement
'            If mouse_row = 0 And mouse_col <> GRDC_BAS_IMPORTE Then
             If mouse_row = 0 Then
                ' Trier selon la colonne choisie
                .Row = 0
                If mouse_col = GRDC_BAS_IMPORTE Then
                    .col = GRDC_BAS_VERT_ROUGE
                    If g_col_tri_bas = mouse_col Then
                        If g_sens_tri_bas = 1 Then
                            .Sort = 2
                            g_sens_tri_bas = 2
                        Else
                            .Sort = 1
                            g_sens_tri_bas = 1
                        End If
                    Else
                        .Sort = 1
                        g_sens_tri_bas = 1
                    End If
                Else
                    .col = mouse_col
                    If g_col_tri_bas = .col Then
                        If g_sens_tri_bas = 1 Then
                            .Sort = 2
                            g_sens_tri_bas = 2
                        Else
                            .Sort = 1
                            g_sens_tri_bas = 1
                        End If
                    Else
                        .Sort = 1
                        g_sens_tri_bas = 1
                    End If
                    If .Rows > 1 Then
                        .TopRow = 1
                        .col = 0
                        .Row = 1
                    End If
                End If
                ' ***********************************************************
                Call colorer_colonne_triee(GRD_BAS, g_col_tri_bas, mouse_col)
                ' ***********************************************************
                g_col_tri_bas = mouse_col
            ElseIf mouse_row <> 0 Then ' n'importe où sur le GRD_BAS sauf la ligne fixe
                Call chercher(Index, mouse_row, vbKeySpace)
            End If
        End If
    End With

End Sub

Private Sub grd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    With grd(Index)
        If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or (KeyCode >= vbKeyA And KeyCode <= vbKeyZ) And .Row > 0 Then
            Call chercher(Index, .Row, KeyCode)
        'ElseIf KeyCode = vbKeyReturn Then
            'Call chercher(Index, .Row, KeyCode)
        ElseIf KeyCode = vbKeyEscape Then
            'Call quitter(False)
        End If
    End With

End Sub

Private Sub grd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim mouse_row As Integer, mouse_col As Integer

    Select Case Index
        Case GRD_HAUT ' GRID DU HAUT ******************************************************************************
            With grd(GRD_HAUT)
                If .MouseRow = 0 Or .Rows = 1 Then
                    .ToolTipText = ""
                    Exit Sub
                End If
                mouse_row = .MouseRow
                mouse_col = .MouseCol
                If mouse_col = GRDC_HAUT_MODIFIER Then
                    .ToolTipText = ""
                ElseIf mouse_col = GRDC_HAUT_IMPORTE Then
                    If .TextMatrix(mouse_row, GRDC_HAUT_VERT_ROUGE) = IMPORTE Then
                        .ToolTipText = .TextMatrix(mouse_row, GRDC_HAUT_NOM) _
                            & " " & .TextMatrix(mouse_row, GRDC_HAUT_PRENOM) & " est déjà importé(e)."
                    ElseIf .TextMatrix(mouse_row, GRDC_HAUT_VERT_ROUGE) = MATRICULE_REDONDANT Then
                        .ToolTipText = .TextMatrix(mouse_row, GRDC_HAUT_NOM) _
                            & " " & .TextMatrix(mouse_row, GRDC_HAUT_PRENOM) & " est importé(e) mais" _
                            & " son matricule redondant: " & .TextMatrix(mouse_row, GRDC_HAUT_MATRICULE) & "."
                    ElseIf .TextMatrix(mouse_row, GRDC_HAUT_VERT_ROUGE) = NON_IMPORTE Then
                        .ToolTipText = .TextMatrix(mouse_row, GRDC_HAUT_NOM) _
                            & " " & .TextMatrix(mouse_row, GRDC_HAUT_PRENOM) & " n'est pas encore importé(e)."
                    End If
                ElseIf mouse_col = GRDC_HAUT_ENCORE_EMPLOI Then
                    If .TextMatrix(mouse_row, GRDC_HAUT_PLUSIEURS_POSTES) Then
                        .ToolTipText = .TextMatrix(mouse_row, GRDC_HAUT_NOM) & " " _
                            & .TextMatrix(mouse_row, GRDC_HAUT_PRENOM) & " est affecté(e) à plus d'un poste."
                    Else
                        .ToolTipText = ""
                    End If
                ElseIf Len(.TextMatrix(.MouseRow, .MouseCol)) * 79 > .ColWidth(.MouseCol) Then
                        .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
                Else
                    .ToolTipText = ""
                End If
            End With
        Case GRD_BAS ' GRID DU BAS ********************************************************************************
            With grd(GRD_BAS)
                If .MouseRow = 0 Or .Rows = 1 Then
                    .ToolTipText = ""
                    Exit Sub
                End If
                mouse_row = .MouseRow
                mouse_col = .MouseCol
                ' le colonne des pastilles verts/rouges
                If mouse_col = GRDC_BAS_IMPORTE Then
                    If .TextMatrix(mouse_row, GRDC_BAS_VERT_ROUGE) = IMPORTE Then
                        .ToolTipText = .TextMatrix(mouse_row, GRDC_BAS_NOM) _
                            & " " & .TextMatrix(mouse_row, GRDC_BAS_PRENOM) & " est déjà importé(e)."
                    ElseIf .TextMatrix(mouse_row, GRDC_BAS_VERT_ROUGE) = NON_IMPORTE Then
                        .ToolTipText = .TextMatrix(mouse_row, GRDC_BAS_NOM) _
                            & " " & .TextMatrix(mouse_row, GRDC_BAS_PRENOM) & " n'est pas encore importé(e)."
                    End If
                ' le reste de colonnes si le texte dépasse la cellule
                ElseIf Len(.TextMatrix(.MouseRow, .MouseCol)) * 100 > .ColWidth(.MouseCol) Then
                    .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
                Else
                    .ToolTipText = ""
                End If
            End With
    End Select

End Sub

Private Sub grd_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim mouse_row As Integer
    Dim mon_y As Single

    mon_y = Y
    If Button = vbRightButton Then
        With grd(Index)
            mouse_row = .MouseRow
            'If Not ((mon_y > (.RowHeight(0) * (.Row - .TopRow + 1))) _
                And (mon_y < (.RowHeight(0) * (.Row - .TopRow + 2)))) Then
            If mouse_row <> .Row Then
                'mnuAcceder.Enabled = False
                mnuAssocier.Enabled = False
            Else ' clique-droit sur la bonne ligne
                'mnuAcceder.Enabled = True
                mnuAssocier.Enabled = cmd(CMD_ASSOCIER).Enabled
            End If
            mnuAcceder.Enabled = True
            mnuAcceder.Visible = (Index = GRD_HAUT)
            Call PopupMenu(mnuContextuel)
        End With
    End If

End Sub

Private Sub grd_RowColChange(Index As Integer)

    cmd(CMD_ASSOCIER).Enabled = False

End Sub

Private Sub mnuAcceder_Click()

    Dim frm As Form

    With grd(GRD_HAUT)
        Set frm = PrmPersonne
        If PrmPersonne.AppelFrm(.TextMatrix(.Row, GRDC_HAUT_U_NUM), "") Then
            ' Mettre à jour la ligne du grid(GRD_HAUT) après des modifications
            Call maj_ligne(.Row, .TextMatrix(.Row, GRDC_HAUT_U_NUM))
        End If
        Set frm = Nothing
    End With

End Sub

Private Sub mnuAssocier_Click()

    Call associer(grd(GRD_HAUT).Row, grd(GRD_BAS).Row)

End Sub


Private Sub txt_Change(Index As Integer)

    Dim autre_txt As Integer

    ' Déterminer l'autre_txt
    If Index = TXT_NOM Then
        autre_txt = TXT_MATRICULE
    Else
        autre_txt = TXT_NOM
    End If

    If Len(txt(Index)) = 0 Then
        If Len(txt(autre_txt)) = 0 Then
            cmd(CMD_RECHERCHER).Enabled = False
        Else
            cmd(CMD_RECHERCHER).Enabled = True
        End If
    Else ' peu importe txt(autre_txt), txt(Index) n'est pas vide
        cmd(CMD_RECHERCHER).Enabled = True
    End If

End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Call rechercher_personne
        txt(TXT_NOM).Text = ""
        txt(TXT_MATRICULE).Text = ""
    End If

End Sub


