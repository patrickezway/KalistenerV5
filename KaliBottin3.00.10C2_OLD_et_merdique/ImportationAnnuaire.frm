VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form ImportationAnnuaire 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11895
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Importation à partir du fichier : "
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
      TabIndex        =   7
      Top             =   0
      Width           =   11895
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
         Height          =   500
         Index           =   6
         Left            =   1200
         Picture         =   "ImportationAnnuaire.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Envoyer une question"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtPopUp 
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Text            =   "txt caché"
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
      End
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
         Height          =   1815
         Left            =   720
         TabIndex        =   12
         Top             =   6000
         Visible         =   0   'False
         Width           =   11535
         Begin ComctlLib.ProgressBar pgb 
            Height          =   495
            Left            =   120
            TabIndex        =   13
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
            Left            =   4920
            TabIndex        =   15
            Top             =   1080
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
            TabIndex        =   16
            Top             =   1080
            Width           =   4815
         End
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
         Height          =   495
         Index           =   2
         Left            =   240
         Picture         =   "ImportationAnnuaire.frx":0496
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Actualiser les données du tableau"
         Top             =   840
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   735
      End
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
         Left            =   2040
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   9375
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
            Height          =   500
            Index           =   1
            Left            =   8520
            Picture         =   "ImportationAnnuaire.frx":0979
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Rechercher"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   1
            Left            =   5280
            TabIndex        =   1
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   0
            Top             =   360
            Width           =   3015
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
            Index           =   2
            Left            =   4320
            TabIndex        =   10
            Top             =   360
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
            Index           =   1
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   495
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   6090
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   10742
         _Version        =   393216
         Rows            =   1
         Cols            =   26
         ForeColor       =   0
         BackColorFixed  =   12648447
         ForeColorFixed  =   0
         AllowUserResizing=   1
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "ici le compteur des lignes"
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
         Left            =   8760
         TabIndex        =   11
         Top             =   350
         Visible         =   0   'False
         Width           =   2655
      End
      Begin ComctlLib.ImageList imglst 
         Left            =   1440
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   17
         ImageHeight     =   17
         MaskColor       =   16777215
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   7
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ImportationAnnuaire.frx":0EE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ImportationAnnuaire.frx":12AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ImportationAnnuaire.frx":1672
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ImportationAnnuaire.frx":1A38
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ImportationAnnuaire.frx":1DFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ImportationAnnuaire.frx":21C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ImportationAnnuaire.frx":26E6
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
      TabIndex        =   6
      Top             =   7580
      Width           =   11895
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Associations ?"
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
         Index           =   5
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Liste des Associations déjà existantes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1785
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Désactiver toutes les personnes à désactiver"
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
         Index           =   4
         Left            =   4890
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   4065
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Corriger tous les mouvements identiques"
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
         Index           =   3
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   3705
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
         Left            =   11160
         Picture         =   "ImportationAnnuaire.frx":2CA0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Quitter l'application en cours"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
   Begin VB.Menu mnuMenuContextuel 
      Caption         =   "menu contextuel"
      Visible         =   0   'False
      Begin VB.Menu mnuActualiser 
         Caption         =   "Actualiser le tableau"
      End
      Begin VB.Menu mnuAction 
         Caption         =   "une action à faire"
      End
   End
End
Attribute VB_Name = "ImportationAnnuaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ***************************************************************************
' CE QUI A ÉTÉ DÉSACTIVÉ/MODIFIÉ:
' ===============================
' - le bouton Actualiser                 dans remplir_grid()
' - envoyer_code_mpasse(v_row, new_code) dans creer_cette_personne()
' - selectionner_ligne(v_row)            dans associer_cette_personne()
' - desactiver_ligne(v_row, ...)         dans creation()
' - desactiver_ligne(v_row, ...)         dans desactiver_personne()
' - desactiver_ligne(v_row, ...)         dans modification()
' - maj_grid(u_num)                      dans rechercher_personne()
' - maj_ligne(v_row, True)               dans afficher_personne_inactive()
' - .ColWidth(GRDC_PASTILLE) = 0         dans initialiser()
' - .ColWidth(GRDC_LIB_SRV_FICH) = 2920  dans initialiser()
' - .ColWidth(GRDC_LIB_POSTE_FICH)= 2720 dans initialiser()
' ***************************************************************************

' Index des colonnes du grid
Private Const GRDC_U_NUM = 0
Private Const GRDC_ACTION = 1
Private Const GRDC_PASTILLE = 2
Private Const GRDC_MATRICULE = 3
Private Const GRDC_CIVILITE = 4
Private Const GRDC_NOM = 5
Private Const GRDC_NJF = 6
Private Const GRDC_PRENOM = 7
Private Const GRDC_CODE_SRV_FICH = 8
Private Const GRDC_LIB_SRV_FICH = 9
Private Const GRDC_CODE_POSTE_FICH = 10
Private Const GRDC_LIB_POSTE_FICH = 11
Private Const GRDC_NUM_SRV_KB = 12
Private Const GRDC_NUM_POSTE = 13
Private Const GRDC_PERSONNE_INACTIVE = 14
Private Const GRDC_INFO_PERSO = 15
Private Const GRDC_METTRE_A_JOUR_NOM = 16
Private Const GRDC_POSTE_EN_GRAS = 17
Private Const GRDC_LISTE_PERSONNE_INACTIVE = 18
Private Const GRDC_SPM_KB_A_SYNCHRONISER = 19
Private Const GRDC_LIGNE_LUE = 20
Private Const GRDC_ETAT_AVANT = 21
' Nom et Prénom Junon (cachés)
Private Const GRDC_NOM_JUNON = 22
Private Const GRDC_PRENOM_JUNON = 23
' Ancien et Nouveaux postes secondaires (cachés)
Private Const GRDC_ANC_PSTSECOND = 24
Private Const GRDC_NEW_PSTSECOND = 25

' Index des TextBox
Private Const TXT_NOM = 0
Private Const TXT_MATRICULE = 1

' Index des FRAMES
Private Const FRM_PRINCIPALE = 0
Private Const FRM_CMD_QUITTER = 1
Private Const FRM_RECHERCHER = 2

Private tbcaractere_nontraite()

' Index des couleurs
Private Const COLOR_DESACTIVE = &HE0E0E0         ' GRIS
Private Const COLOR_A_DETRUIRE = &H8080FF        ' ROSE
Private Const COLOR_A_CREER = &H80FF80           ' BLEUE
Private Const COLOR_A_MODIFIER = &HFFFFFF        ' BLANC
Private Const COLOR_COLONNE_TRIEE = &HFFFF&      ' JAUNE FONCE &H0000C0C0&
Private Const COLOR_COLONNE_NON_TRIEE = &HC0FFFF ' couleur de la ligne fixe par défaut

' Index des IMAGES
Private Const IMG_PASTILLE_VERTE = 4
Private Const IMG_PASTILLE_ROUGE = 5
Private Const IMG_PERSONNE_INACTIVE = 3
Private Const IMG_VOIR = 6
Private Const IMG_MODIF_POSTESECOND = 7

' Index des positions des images
Private Const POS_GAUCHE = flexAlignLeftCenter
Private Const POS_CENTRE = flexAlignCenterCenter
Private Const POS_DROITE = flexAlignRightCenter

' Index des actions:
Private Const DETRUIRE = "DETRUIRE"
Private Const CREER = "CREER"
Private Const MODIFIER = "MODIFIER"

' Constatantes de séparation
Private Const SEPARATEUR_SERVICE_POSTE = " <=> "

' Index des CMD
Private Const CMD_QUITTER = 0
Private Const CMD_RECHERCHER = 1
Private Const CMD_ACTUALISER = 2
Private Const CMD_CORRIGER_TOUS = 3
Private Const CMD_DESACTIVER_TOUS = 4
Private Const CMD_LISTE_ASSOC = 5
Private Const CMD_QUESTION = 6

' Index des LIBELLES
Private Const LBL_COMPTEUR = 3

Private Const PRM_OUI = 1
Private Const PRM_NON = 2
Private Const PRM_JAMAIS = 3
Private Const PRM_TOUJOURS = 4
Private Const PRM_QUEST = 5

Private g_aller_dans_prm As Integer

Private g_form_active As Boolean
' Sens des tris
Private g_col_tri As Integer
Private g_sens_tri As Integer
' Nombre des premiers caractères du PRENOM à comparer
Private Const g_nbr_car_prenom = 1
Private g_actualiser As Boolean
Private g_row_context_menu As Integer
Private g_importation_ok As Integer
Private g_nbr_lignes As Long

' Les informations supplémentaires à valider
Private g_nbr_infoSuppl As Integer

Private g_poste As String
'Private g_code_emploi_auto  As String
'Private g_code_service_auto  As String
Private g_GRDC_CODE_SRV_FICH As String
Private g_GRDC_CODE_POSTE_FICH As String
Private g_GRDC_LIB_SRV_FICH As String
Private g_GRDC_LIB_POSTE_FICH As String

Private Type T_U_INFO_SUPPL
    unum As Long
    unom As String
    umatricule As String
    infosuppl As String
    tis_num As Long
    tis_alimente As Long
    tis_value As String
    tis_pour_creer As Boolean
End Type

Private g_LISTE_U_INFO_SUPPL() As T_U_INFO_SUPPL

Private Sub a_corriger(ByVal v_ligne_en_cours As String, ByVal v_rs As rdoResultset, _
                                  ByVal v_str As String, ByVal v_spm As String)
' *****************************************************************
' Remplir le GRID avec la ligne à METTRE A JOUR
' v_str = l'une des chaines suivantes {"NOM", "NOM_POSTE", "POSTE"}
' *****************************************************************
    Dim sql As String, mon_nom As String
    Dim rs As rdoResultset
    Dim I As Integer
    Dim NewPoste As Long, NewSrv As Long

    With grd
        mon_nom = v_rs("U_Nom").Value
        .AddItem ""
        .TextMatrix(.Rows - 1, GRDC_U_NUM) = v_rs("U_Num").Value
        .col = GRDC_ACTION
        .Row = .Rows - 1
        .TextMatrix(.Rows - 1, GRDC_ACTION) = "Corriger"
        .CellFontBold = True
        .CellBackColor = vbWhite
        .col = GRDC_PASTILLE
        .CellBackColor = vbWhite
        '''.TextMatrix(.Rows - 1, GRDC_MATRICULE) = Trim$(STR_GetChamp(v_ligne_en_cours, p_separateur, p_pos_matricule))
        .TextMatrix(.Rows - 1, GRDC_MATRICULE) = Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_matricule, p_long_matricule, "matricule"))
        If p_pos_civilite <> -1 Then
            .TextMatrix(.Rows - 1, GRDC_CIVILITE) = Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_civilite, p_long_civilite, "civilité"))
        Else
            .TextMatrix(.Rows - 1, GRDC_CIVILITE) = ""
        End If
        .TextMatrix(.Rows - 1, GRDC_NOM) = UCase$(P_ChangerCar(Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_nom, p_long_nom, "nom")), tbcaractere_nontraite))
        If left$(v_str, 3) = "NOM" Then ' mettre la cellule NOM en GRAS
            .col = GRDC_NOM
            .CellFontBold = True
            ' stocker le NOM trouvé dans KaliBottin
            .TextMatrix(.Rows - 1, GRDC_METTRE_A_JOUR_NOM) = UCase$(mon_nom)
        End If
        If p_pos_njf <> -1 Then
            .TextMatrix(.Rows - 1, GRDC_NJF) = Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_njf, p_long_njf, "nom de jeune fille"))
        Else
            .TextMatrix(.Rows - 1, GRDC_NJF) = ""
        End If
        .TextMatrix(.Rows - 1, GRDC_PRENOM) = Trim$(P_ChangerCar(formater_prenom(Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_prenom, p_long_prenom, "prénom"))), tbcaractere_nontraite))
        .TextMatrix(.Rows - 1, GRDC_NOM_JUNON) = Trim$(v_rs("U_NomJunon").Value & "")
        .TextMatrix(.Rows - 1, GRDC_PRENOM_JUNON) = Trim$(v_rs("U_PrenomJunon").Value & "")
        .TextMatrix(.Rows - 1, GRDC_CODE_SRV_FICH) = Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_code_section, p_long_code_section, "code section"))
        .TextMatrix(.Rows - 1, GRDC_LIB_SRV_FICH) = Trim$(P_ChangerCar(Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_lib_section, p_long_lib_section, "libellé section")), tbcaractere_nontraite))
        .TextMatrix(.Rows - 1, GRDC_CODE_POSTE_FICH) = Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_code_emploi, p_long_code_emploi, "code emploi"))
        .TextMatrix(.Rows - 1, GRDC_LIB_POSTE_FICH) = Trim$(P_ChangerCar(Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_lib_emploi, p_long_lib_emploi, "libellé emploi")), tbcaractere_nontraite))
        .col = GRDC_INFO_PERSO
            Set .CellPicture = imglst.ListImages(IMG_VOIR).Picture
                .CellPictureAlignment = flexAlignCenterCenter
        .CellFontBold = True
        If v_str = "POSTE_SECONDAIRE" Then
            .TextMatrix(.Rows - 1, GRDC_NEW_PSTSECOND) = STR_GetChamp(v_spm, "!", 1)
            .TextMatrix(.Rows - 1, GRDC_ANC_PSTSECOND) = STR_GetChamp(v_spm, "!", 2)
            v_spm = STR_GetChamp(v_spm, "!", 0)
        End If
        .TextMatrix(.Rows - 1, GRDC_NUM_SRV_KB) = P_get_num_srv_poste(v_spm, P_SERVICE)
        .TextMatrix(.Rows - 1, GRDC_NUM_POSTE) = P_get_num_srv_poste(v_spm, P_POSTE)
        ' Service de remplacement
        NewSrv = .TextMatrix(.Rows - 1, GRDC_NUM_SRV_KB)
        NewPoste = .TextMatrix(.Rows - 1, GRDC_NUM_POSTE)
        If FctTransposePoste(NewSrv, NewPoste, False) > 0 Then
            .TextMatrix(.Rows - 1, GRDC_NUM_SRV_KB) = NewSrv
            .TextMatrix(.Rows - 1, GRDC_NUM_POSTE) = NewPoste
        End If
        ' modification des postes secondaires
        If v_str = "POSTE_SECONDAIRE" Then
            .ToolTipText = "Les postes secondaires de cette personne ont été modifiés depuis la dernière synchronisation"
            .col = GRDC_INFO_PERSO
                Set .CellPicture = imglst.ListImages(IMG_MODIF_POSTESECOND).Picture
                    .CellPictureAlignment = flexAlignCenterCenter
            .TextMatrix(.Rows - 1, GRDC_ACTION) = "Postes Sec."
            .TextMatrix(.Rows - 1, GRDC_ETAT_AVANT) = "Postes Sec."
        Else
            .TextMatrix(.Rows - 1, GRDC_ETAT_AVANT) = "Corriger"
        End If
        
        If Right$(v_str, 5) = "POSTE" Then ' mettre les cellules CODE/LIBELLE SERVICE/POSTE en GRAS
            .col = GRDC_CODE_SRV_FICH
            .CellFontBold = True
            .col = GRDC_LIB_SRV_FICH
            .CellFontBold = True
            .col = GRDC_CODE_POSTE_FICH
            .CellFontBold = True
            .col = GRDC_LIB_POSTE_FICH
            .CellFontBold = True
            .TextMatrix(.Rows - 1, GRDC_POSTE_EN_GRAS) = P_OUI
        Else
            .TextMatrix(.Rows - 1, GRDC_POSTE_EN_GRAS) = P_NON
        End If
        ' stocker le POSTE trouvé dans KaliBottin
        .TextMatrix(.Rows - 1, GRDC_SPM_KB_A_SYNCHRONISER) = v_spm
    End With

End Sub


Private Function FctTransposePoste(ByRef r_NewSrv As Long, ByRef r_NewPoste As Long, ByVal v_bMessage As Boolean) As Integer
    Dim sql As String, rs As rdoResultset, rsTmp As rdoResultset
    Dim FT_NivRemplace As String, srv_num As Long, ft_num As Long, srv_num_Rempl As Long
    Dim strRemplace As String
        
    sql = "select FT_NivRemplace, PO_SrvNum, FT_Num from Poste, FctTrav" _
        & " where PO_Num=" & r_NewPoste _
        & " and FT_Num=PO_FTNum"
    If Odbc_RecupVal(sql, FT_NivRemplace, srv_num, ft_num) <> P_ERREUR Then
        If FT_NivRemplace = "" Then FT_NivRemplace = 0
        If FT_NivRemplace > 0 Then
            If FctNivRemplace(FT_NivRemplace, ft_num, srv_num, r_NewPoste, strRemplace) < 0 Then
                If Not p_traitement_background Then
                    MsgBox "!!! Attention : " & strRemplace
                Else
                    p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> " & strRemplace
                End If
                r_NewSrv = 0
                r_NewPoste = 0
                FctTransposePoste = P_ERREUR
            Else
                If Not p_traitement_background Then
                    If v_bMessage Then
                        MsgBox "Une Transposition est définie pour ce poste : " & strRemplace
                    End If
                Else
                    p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> " & "Une Transposition est définie pour le poste " & P_get_lib_srv_poste(r_NewPoste, P_POSTE) & " : " & strRemplace
                End If
                r_NewSrv = srv_num
                FctTransposePoste = 1
            End If
        End If
    End If

End Function

Private Sub a_creer(ByVal v_ligne_en_cours As String, ByVal v_row As Integer)
' ***************************************
' Remplir le grid avec les lignes à créer
' ***************************************
    Dim matricule_en_cours As String, nomimg As String
    Dim I As Integer
    

    With grd
        matricule_en_cours = Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_matricule, p_long_matricule, "matricule"))
        '.AddItem ""
        .Row = v_row
        .col = GRDC_ACTION
        .TextMatrix(v_row, GRDC_ACTION) = "Créer"
        .CellFontBold = True
        .TextMatrix(v_row, GRDC_MATRICULE) = matricule_en_cours
        If p_pos_civilite <> -1 Then
            .TextMatrix(v_row, GRDC_CIVILITE) = Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_civilite, p_long_civilite, "civilité"))
        Else
            .TextMatrix(v_row, GRDC_CIVILITE) = ""
        End If
        .TextMatrix(v_row, GRDC_NOM) = Trim$(UCase$(P_ChangerCar(Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_nom, p_long_nom, "nom")), tbcaractere_nontraite)))
        If p_pos_njf <> -1 Then
            .TextMatrix(v_row, GRDC_NJF) = Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_njf, p_long_njf, "nom de jeune fille"))
        Else
            .TextMatrix(v_row, GRDC_NJF) = ""
        End If
        .TextMatrix(v_row, GRDC_PRENOM) = Trim$(P_ChangerCar(formater_prenom(Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_prenom, p_long_prenom, "prénom"))), tbcaractere_nontraite))
        .TextMatrix(v_row, GRDC_NOM_JUNON) = ""
        .TextMatrix(v_row, GRDC_PRENOM_JUNON) = ""
        .TextMatrix(v_row, GRDC_CODE_SRV_FICH) = Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_code_section, p_long_code_section, "code section"))
        .TextMatrix(v_row, GRDC_LIB_SRV_FICH) = Trim$(P_ChangerCar(Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_lib_section, p_long_lib_section, "libellé section")), tbcaractere_nontraite))
        .TextMatrix(v_row, GRDC_CODE_POSTE_FICH) = Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_code_emploi, p_long_code_emploi, "code emploi"))
        .TextMatrix(v_row, GRDC_LIB_POSTE_FICH) = Trim$(P_ChangerCar(Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_lib_emploi, p_long_lib_emploi, "libellé emploi")), tbcaractere_nontraite))
        '.TextMatrix(.Rows - 1, GRDC_INFO_PERSO) = "!"
        .col = GRDC_INFO_PERSO
          Set .CellPicture = imglst.ListImages(IMG_VOIR).Picture
              .CellPictureAlignment = flexAlignCenterCenter
            
        .CellFontBold = True
        .TextMatrix(v_row, GRDC_CODE_SRV_FICH) = Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_code_section, p_long_code_section, "code section"))
        .TextMatrix(v_row, GRDC_CODE_POSTE_FICH) = Trim$(P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_code_emploi, p_long_code_emploi, "code emploi"))
        .TextMatrix(v_row, GRDC_NUM_SRV_KB) = ""
        .TextMatrix(v_row, GRDC_NUM_POSTE) = ""
        .TextMatrix(v_row, GRDC_POSTE_EN_GRAS) = P_NON
        .TextMatrix(v_row, GRDC_LIGNE_LUE) = v_ligne_en_cours
        .TextMatrix(v_row, GRDC_ETAT_AVANT) = "Créer"
        ' vérifier les personnes inactives ou externe au fichier d'importation
        If personne_inactive_existe(v_row) Then
            .col = GRDC_PERSONNE_INACTIVE
            Set .CellPicture = imglst.ListImages(IMG_PERSONNE_INACTIVE).Picture
        Else
            .TextMatrix(v_row, GRDC_LISTE_PERSONNE_INACTIVE) = ""
        End If
        For I = 0 To .Cols - 1
            .col = I
            .CellBackColor = COLOR_A_CREER
        Next I
        .col = GRDC_PASTILLE
        Set .CellPicture = LoadPicture("")

        ' info suppl ARRET
        For I = 0 To p_nbr_lstInfoSuppl - 1
            .TextMatrix(v_row, .Cols - p_nbr_lstInfoSuppl + I) = LISTE_TIS_POS(I).prmgenb_tis_num & ";" & P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, LISTE_TIS_POS(I).prmgenb_tis_pos, LISTE_TIS_POS(I).prmgenb_tis_long, LISTE_TIS_POS(I).prmgenb_tis_lib)
        Next I
    End With

End Sub
Private Sub a_detruire(ByVal v_rs As rdoResultset)
' *****************************************
' Se charge de remplir les ligne à DETRUIRE
' *****************************************
    Dim I As Integer

    With grd
        .AddItem ""
        .col = GRDC_ACTION
        .Row = .Rows - 1
        .TextMatrix(.Rows - 1, GRDC_ACTION) = "Désactiver"
        .CellFontBold = True
        .TextMatrix(.Rows - 1, GRDC_MATRICULE) = v_rs("U_Matricule").Value
        If p_pos_civilite <> -1 Then
            .TextMatrix(.Rows - 1, GRDC_CIVILITE) = v_rs("U_Prefixe").Value
        Else
            .TextMatrix(.Rows - 1, GRDC_CIVILITE) = ""
        End If
        .TextMatrix(.Rows - 1, GRDC_U_NUM) = v_rs("U_Num").Value
        .TextMatrix(.Rows - 1, GRDC_NOM) = v_rs("U_Nom").Value
        .TextMatrix(.Rows - 1, GRDC_NJF) = ""
        .TextMatrix(.Rows - 1, GRDC_PRENOM) = formater_prenom(v_rs("U_Prenom").Value)
        Call remplir_srv_poste(v_rs("U_SPM").Value, "DETRUIRE")
        .TextMatrix(.Rows - 1, GRDC_POSTE_EN_GRAS) = P_NON
        .TextMatrix(.Rows - 1, GRDC_ETAT_AVANT) = "Désactiver"
        .col = GRDC_INFO_PERSO
        Set .CellPicture = imglst.ListImages(IMG_VOIR).Picture
            .CellPictureAlignment = flexAlignCenterCenter
            .CellFontBold = True
        For I = 0 To .Cols - 1
            .col = I
            .CellBackColor = COLOR_A_DETRUIRE
        Next I
    End With

End Sub

Private Sub action(ByVal v_row As Integer)

    Dim frm As Form
    Dim I As Integer, nb As Integer
    Dim ret As Integer
    Dim s As String
    Dim sListe As String, sPoste As String, sService As String
    
    Me.cmd(CMD_CORRIGER_TOUS).Visible = False
    If v_row = 0 Then Exit Sub
    With grd
        If .TextMatrix(v_row, GRDC_ACTION) = "Créer" Then
            Call selectionner_ligne(v_row)
            If .TextMatrix(v_row, GRDC_LISTE_PERSONNE_INACTIVE) <> "" Then
                s = "Il existe au moins une personne inactive ou externe au fichier d'importation" & vbCrLf _
                        & " qui porte le nom: " & vbTab & .TextMatrix(v_row, GRDC_NOM)
                If p_traitement_background Then
                    p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> " & Replace(s, vbCrLf, " ")
                Else
                    p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> " & s
                    Call MsgBox(s & vbCrLf & vbCrLf & vbTab & "Voici la liste de cette/ces personne(s).", _
                        vbInformation + vbOKOnly, "Attention")
                    Call afficher_personne_inactive(v_row)
                End If
            Else
                Call creation(v_row)
            End If
        ElseIf .TextMatrix(v_row, GRDC_ACTION) = "Désactiver" Then
            Call selectionner_ligne(v_row)
            Call desactiver_personne(v_row, False)
        ElseIf .TextMatrix(v_row, GRDC_ACTION) = "Corriger" Then
            Call selectionner_ligne(v_row)
            sPoste = grd.TextMatrix(v_row, GRDC_CODE_POSTE_FICH)
            sService = grd.TextMatrix(v_row, GRDC_CODE_SRV_FICH)
            Call modification(v_row)
            nb = 0
            If Not p_traitement_background Then
                If cmd(CMD_CORRIGER_TOUS).Visible Then
                    For I = 0 To grd.Rows - 1
                        grd.Row = I
                        If sPoste = grd.TextMatrix(I, GRDC_CODE_POSTE_FICH) Then
                            If sService = grd.TextMatrix(I, GRDC_CODE_SRV_FICH) Then
                                If grd.TextMatrix(I, GRDC_ACTION) = "Corriger" Then
                                    nb = nb + 1
                                    sListe = sListe & Chr(13) & Chr(10) & grd.TextMatrix(I, GRDC_MATRICULE) & " " & grd.TextMatrix(I, GRDC_NOM) & " " & grd.TextMatrix(I, GRDC_PRENOM)
                                End If
                            End If
                        End If
                    Next I
                    If nb > 0 Then
                        ret = MsgBox("Voulez vous corriger aussi les " & nb & " autre(s) ligne(s) concernée(s) pour " & Chr(13) & Chr(10) & _
                                sListe, vbYesNo + vbQuestion)
                        grd.Row = v_row
                        If ret = vbYes Then
                            Call cmd_Click(CMD_CORRIGER_TOUS)
                        End If
                    Else
                        cmd(CMD_CORRIGER_TOUS).Visible = False
                    End If
                End If
            End If
        ElseIf .TextMatrix(v_row, GRDC_ACTION) = "Postes Sec." Then
            Call selectionner_ligne(v_row)
            Call modification(v_row)
        ElseIf .TextMatrix(v_row, GRDC_ACTION) = "Accéder" Then
            Call selectionner_ligne(v_row)
            ' Modification d'une personne
            Set frm = PrmPersonne
            If PrmPersonne.AppelFrm(.TextMatrix(v_row, GRDC_U_NUM), "") Then
            ' il y a eu des changements importatnts
                cmd(CMD_ACTUALISER).Enabled = True
                ' Mettre à jour la ligne du grid(GRD_HAUT) après cet modification
                Call maj_ligne(v_row, True)
            Else ' il n'y a pas eu des changements importants
            End If
            Set frm = Nothing
        End If
    End With
    If p_mess_fait_background <> "" Then Print #g_fd1, p_mess_fait_background
    If p_mess_pasfait_background <> "" Then Print #g_fd2, p_mess_pasfait_background
    p_mess_fait_background = ""
    p_mess_pasfait_background = ""

    g_row_context_menu = 0
    grd.SetFocus

End Sub

Private Sub actualiser()
'*****************************************************************************
' Actualiser le grid après les mouvements: créaion/modifications/désactivation
'*****************************************************************************

    grd.Visible = False
    cmd(CMD_RECHERCHER).Enabled = False
    g_actualiser = True
    Sleep (2000)
    Call initialiser
    cmd(CMD_ACTUALISER).Enabled = False

End Sub

Private Sub actualiser_compteur()
    g_nbr_lignes = g_nbr_lignes - 1
    lbl(LBL_COMPTEUR).Caption = g_nbr_lignes & " mouvement(s) à traiter"
End Sub

Private Function afficher_infoSuppl() As Integer
' Afficher les infosuppl des personnes non affichées dans le grid
    Dim tis_value As String, sql As String
    Dim I As Integer, j As Integer, nbr As Integer, n As Integer
    Dim lng As Long, tis_num As Long, lnb As Long
    Dim i_INFO_SUPPL As Long

    Call CL_Init
    Call CL_InitMultiSelect(True, False)
    Call CL_InitGererTousRien(True)
    Call CL_InitTaille(0, -15)
    Call CL_InitTitreHelp("Informations supplémentaires modifiées", "")
    Call CL_AddBouton("Mettre à jour", "", vbKeyM, 0, 2500)
    Call CL_AddBouton("Ne pas mettre à jour et Continuer", "", vbKeyC, 0, 2500)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 550)
    n = STR_GetNbchamp(g_LISTE_U_INFO_SUPPL(0).unom, vbTab)
    For I = 0 To n - 1
        Call CL_AddLabel(STR_GetChamp(g_LISTE_U_INFO_SUPPL(0).unom, vbTab, I))
    Next I
    For I = 1 To g_nbr_infoSuppl
        If g_LISTE_U_INFO_SUPPL(I).tis_pour_creer = False Then
            Call CL_AddLigne(g_LISTE_U_INFO_SUPPL(I).unom & vbTab & g_LISTE_U_INFO_SUPPL(I).infosuppl, I, "", False)
        End If
    Next I

lab_reafficher:
    ChoixListe.Show 1

    If CL_liste.retour = 2 Then ' quitter
        Call quitter("")
        GoTo lab_erreur
    ElseIf CL_liste.retour = 0 Then ' Enregistrer / MAJ
        If MsgBox("Confirmez-vous la mise à jour de ces informations supplémentaires ?", vbYesNo + vbQuestion, _
                  "Confirmation") = vbYes Then
            ' renseigner la table des Info Suppl
            For I = 0 To UBound(CL_liste.lignes)
                If Not CL_liste.lignes(I).selected Then GoTo LabNextI
                i_INFO_SUPPL = CL_liste.lignes(I).num
                nbr = STR_GetNbchamp(g_LISTE_U_INFO_SUPPL(i_INFO_SUPPL).tis_value, "|")
                For j = 0 To nbr - 1
                    tis_num = STR_GetChamp(STR_GetChamp(g_LISTE_U_INFO_SUPPL(i_INFO_SUPPL).tis_value, "|", j), ";", 0)
                    tis_value = STR_GetChamp(STR_GetChamp(g_LISTE_U_INFO_SUPPL(i_INFO_SUPPL).tis_value, "|", j), ";", 1)
                    If g_LISTE_U_INFO_SUPPL(i_INFO_SUPPL).tis_alimente > 0 Then
                        ' on crée dans zoneutil
                        If g_LISTE_U_INFO_SUPPL(i_INFO_SUPPL).tis_alimente > 0 Then
                            Call AlimenteZoneUtil("M", g_LISTE_U_INFO_SUPPL(i_INFO_SUPPL).tis_alimente, g_LISTE_U_INFO_SUPPL(i_INFO_SUPPL).unum, tis_value)
                        End If
                    Else
                        sql = " select count(*) from InfoSupplEntite WHERE ISE_Type='U' " _
                            & " AND ISE_TypeNum=" & g_LISTE_U_INFO_SUPPL(i_INFO_SUPPL).unum _
                            & " AND ISE_TisNum=" & tis_num
                        If Odbc_Count(sql, lnb) = P_ERREUR Then
                            GoTo lab_erreur
                        End If
                        If lnb > 0 Then
                            ' MAJ si info existe
                            If Odbc_Update("InfoSupplEntite", "ISE_Num", _
                                           "WHERE ISE_Type='U' AND ISE_TypeNum=" & g_LISTE_U_INFO_SUPPL(i_INFO_SUPPL).unum _
                                                & " AND ISE_TisNum=" & tis_num, _
                                            "ISE_Valeur", tis_value) = P_ERREUR Then
                                GoTo lab_erreur
                            End If
                        Else
                            ' l'info n'existe pas, on l'ajoute
                            If Odbc_AddNew("InfoSupplEntite", "ISE_Num", "ISE_Seq", False, lng, _
                                            "ISE_TisNum", tis_num, _
                                            "ISE_TypeNum", g_LISTE_U_INFO_SUPPL(i_INFO_SUPPL).unum, _
                                            "ISE_Type", "U", _
                                            "ISE_Valeur", tis_value) = P_ERREUR Then
                                GoTo lab_erreur
                            End If
                        End If
                    End If
                Next j
LabNextI:
            Next I
        Else
            GoTo lab_reafficher
        End If
    End If

    afficher_infoSuppl = P_OK
    Exit Function

    ' quitter
lab_erreur:
    afficher_infoSuppl = P_ERREUR

End Function

Private Function AlimenteZoneUtil(Trait, v_alimente, v_unum, v_value)
    Dim sql As String
    Dim rs As rdoResultset
    Dim lng As Long
    Dim mess As String
            
    sql = "SELECT * FROM UtilCoordonnee WHERE UC_Type='U' and UC_TypeNum=" & v_unum & " and UC_ZuNum=" & v_alimente & " and UC_Par_Synchro"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    If rs.EOF Then
        ' on la créee
        If Odbc_AddNew("UtilCoordonnee", "UC_Num", "UC_Seq", False, lng, _
                       "UC_Type", "U", _
                       "UC_TypeNum", v_unum, _
                       "UC_ZUNum", v_alimente, _
                       "UC_Valeur", v_value, _
                       "UC_Comm", "", _
                       "UC_Niveau", "0", _
                       "UC_Par_Synchro", True, _
                       "UC_Principal", True, _
                       "UC_LstPoste", "") = P_ERREUR Then
                GoTo lab_erreur
        End If
        mess = "Création de coordonnée : " & v_value
        Print #g_fd1, mess
    Else
        If Odbc_Update("UtilCoordonnee", "UC_Num", _
                       "WHERE UC_Num=" & rs("UC_Num"), _
                       "UC_Valeur", v_value) = P_ERREUR Then
                GoTo lab_erreur
        End If
        mess = "Modification de coordonnée : " & v_value
        Print #g_fd1, mess
    End If
    Exit Function
lab_erreur:
    MsgBox "Erreur SQL " & Error$
End Function

Private Function afficher_personne_inactive(ByVal v_row As Integer) As Integer
'*******************************************************************************
' Afficher la liste des personnes inactives et/ou externes au fichiers d'import.
' La colonne contient une information comme: "IE:213;E:20;I:881;E:625;"
' où "I" pour inactive et "E" pour externe au fichier ("I" tjs avant le "E")
'*******************************************************************************

    Dim sql As String, liste As String, mon_srv As String, mon_poste As String, etat As String, str As String
    Dim spm_choisi As String, s As String
    Dim I As Integer, nbr As Integer
    Dim numposte As Long, numutil As Long
    Dim rs As rdoResultset
    Dim srv_num As Long
    Dim srv_num_Rempl As Long
    Dim ft_num As Long
    Dim rsTmp As rdoResultset
    Dim strRemplace As String, FT_NivRemplace As Integer

    Call CL_Init
    Call CL_InitTaille(0, -15)
    Call CL_InitTitreHelp("Liste des personnes se nommant: " & grd.TextMatrix(v_row, GRDC_NOM), "")
    Call CL_AddBouton("Associer à cette personne", "", vbKeyN, 0, 2500)
    Call CL_AddBouton("Créer nouvelle personne", "", vbKeyN, 0, 2300)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 550)
    With grd
        liste = .TextMatrix(v_row, GRDC_LISTE_PERSONNE_INACTIVE)
        nbr = STR_GetNbchamp(liste, ";")
        For I = 0 To nbr - 1
            etat = STR_GetChamp(STR_GetChamp(liste, ";", I), ":", 0)
            sql = "SELECT * FROM Utilisateur WHERE U_Num=" & STR_GetChamp(STR_GetChamp(liste, ";", I), etat & ":", 1)
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                afficher_personne_inactive = P_NON
                Exit Function
            End If
            If Not rs.EOF Then
                mon_srv = P_get_lib_srv_poste(P_get_num_srv_poste(rs("U_SPM").Value, P_SERVICE), P_SERVICE)
                mon_poste = P_get_lib_srv_poste(P_get_num_srv_poste(rs("U_SPM").Value, P_POSTE), P_POSTE)
                str = ""
                str = IIf(InStr(etat, "I") <> 1, "[]", "[Inactive]")
                str = str & " " & IIf(InStr(etat, "E") <> 1 And InStr(etat, "E") <> 2, "[]", "[Externe au fichier]")
                Call CL_AddLigne(str & vbTab & rs("U_Nom").Value & " " & formater_prenom(rs("U_Prenom").Value) _
                                 & vbTab & rs("U_Matricule").Value & vbTab & mon_poste & " - " & mon_srv, _
                                 rs("U_Num").Value, _
                                 rs("U_Matricule").Value, _
                                 False)
            End If
            rs.Close
        Next I
    End With
    ' Afficher la liste
    ChoixListe.Show 1
    ' Tester le code de retour
    If CL_liste.retour = 1 Then '     CREER
        Call creation(v_row)
    ElseIf CL_liste.retour = 0 Then ' ASSOCIER
        numutil = CL_liste.lignes(CL_liste.pointeur).num
        numposte = choisir_poste_synchro_struct(grd.TextMatrix(v_row, GRDC_CODE_SRV_FICH), "", grd.TextMatrix(v_row, GRDC_CODE_POSTE_FICH))
        If numposte = 0 Then ' on n'a rien selectionné
            afficher_personne_inactive = P_NON
            Exit Function
        End If
        
        spm_choisi = build_arbor_srv(numposte)
        If spm_choisi = "" Then
            afficher_personne_inactive = P_NON
            Exit Function
        End If
        spm_choisi = spm_choisi & "|"
        grd.TextMatrix(v_row, GRDC_U_NUM) = numutil
        Call modifier_cette_personne(v_row, spm_choisi, "", False)
        Call associer_cette_personne(v_row, numutil)
        ' Lancement KD pour maj des diffusions
        s = p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";KaliDoc;" & p_nom_fichier_ini_kalidoc & ";CONNEXION=" & p_NumUtil & ";ACTIVER_PERS=" & numutil & "[WAIT];KBAUTO"
        Call SYS_ExecShell(s, True, True)
    ElseIf CL_liste.retour = 2 Then ' QUITTER
        'afficher_personne_inactive = P_NON
' A NE PAS SUPPRIMER
        'Call maj_ligne(v_row, True)
        afficher_personne_inactive = P_OUI
    End If

End Function

Private Function FctNivRemplace(ByVal v_FT_NivRemplace As Integer, ByVal v_FTNum As Long, ByRef r_numsrv As Long, ByRef r_numpo As Long, ByRef r_strRemplace As String) As Integer
    ' Voir si pour ce service il a un père du niveau de remplacement indiqué
    Dim sql As String, rs As rdoResultset
    Dim encore As Boolean
    Dim ilya As Boolean
    Dim srvnum As Long
    Dim srvnom_prem As String
    Dim numpo As Long
    Dim strNiveau As String
    Dim anc_ftnum As Long
    Dim anc_lnum As Long
    Dim anc_libelle As String
    Dim lib_srv As String
    Dim lib_niveau As String
    
    If v_FT_NivRemplace = 0 Then
        FctNivRemplace = 0
    Else
        srvnum = r_numsrv
        encore = True
        While encore
            sql = "select SRV_Num, SRV_NumPere, SRV_Nom, SRV_NivsNum from Service where SRV_Num=" & srvnum
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                r_strRemplace = "Erreur " & sql
                FctNivRemplace = P_ERREUR
                Exit Function
            ElseIf rs.EOF Then
                Call Odbc_RecupVal("select nivs_nom from niveau_structure where nivs_num='" & v_FT_NivRemplace & "'", strNiveau)
                r_strRemplace = "Il n'y a pas de niveau de remplacement (" & strNiveau & ")" '    & " pour " & srvnom_prem
                FctNivRemplace = P_ERREUR
                rs.Close
                Exit Function
            ElseIf rs("SRV_NivsNum") = v_FT_NivRemplace Then
                Call Odbc_RecupVal("select nivs_nom from niveau_structure where nivs_num='" & v_FT_NivRemplace & "'", strNiveau)
                r_strRemplace = "Niveau de remplacement : " & strNiveau & " => " & rs("SRV_Nom")
                lib_srv = rs("SRV_Nom")
                FctNivRemplace = 1
                r_numsrv = rs("SRV_Num")
                rs.Close
                ' Voir si ce poste existe
                sql = "select * from Poste where PO_FTNum=" & v_FTNum & " and PO_SrvNum=" & r_numsrv
                If Odbc_SelectV(sql, rs) = P_ERREUR Then
                    MsgBox "Erreur SQL"
                ElseIf rs.EOF Then
                    r_strRemplace = "Le poste " & recup_PSLib(r_numpo) & " n'existe pas dans le service de remplacement" & r_strRemplace
                    FctNivRemplace = P_ERREUR
                    ' le créer automatiquement
                    Call Odbc_RecupVal("Select PO_Ftnum, PO_lNum, PO_Libelle from Poste where po_num=" & r_numpo, anc_ftnum, anc_lnum, anc_libelle)
                    If Odbc_AddNew("Poste", _
                                     "PO_Num", _
                                     "po_seq", _
                                     True, _
                                     numpo, _
                                     "PO_SRVNum", r_numsrv, _
                                     "PO_FTNum", anc_ftnum, _
                                     "PO_LNum", anc_lnum, _
                                     "PO_NumResp", -1, _
                                     "PO_Libelle", anc_libelle, _
                                     "PO_Actif", True) = P_ERREUR Then
                        r_strRemplace = r_strRemplace & Chr(13) & Chr(10) & "Le poste n'a pas pu être créé"
                    Else
                        Call Odbc_RecupVal("select nivs_nom from niveau_structure where nivs_num='" & v_FT_NivRemplace & "'", lib_niveau)
                        r_strRemplace = r_strRemplace & Chr(13) & Chr(10) & "Super KaliBottin l'a créé pour vous au niveau " & lib_niveau & " de " & lib_srv
                        ' On prend ce poste nouvellement créé
                        r_numpo = numpo
                        FctNivRemplace = 1
                    End If
                    Exit Function
                Else
                    ' On prend ce nouveau poste
                    r_numpo = rs("PO_Num")
                End If
                rs.Close
                Exit Function
            Else    ' voir son pere
                srvnum = rs("SRV_NumPere")
                If srvnom_prem = "" Then srvnom_prem = rs("SRV_Nom")
            End If
        Wend
        rs.Close
    End If
End Function

Private Function ajout_dans_synchro(ByVal v_row As Integer, Optional ByVal v_um_spnum As Long) As Boolean
'**********************************************************************
' Appelée depuis creer_cette_personne() ou associer_cette_personne()
' Déterminer si on peut ajouter un enregistrement dans la table Synchro
'**********************************************************************
    Dim rs As rdoResultset

    With grd
        If Odbc_SelectV("SELECT * FROM Synchro" _
                      & " WHERE SYNC_Section=" & Odbc_String(.TextMatrix(v_row, GRDC_CODE_SRV_FICH)) _
                      & " AND SYNC_Emploi=" & Odbc_String(.TextMatrix(v_row, GRDC_CODE_POSTE_FICH)) _
                      & " AND SYNC_SPNum=" & IIf(v_um_spnum = 0, _
                                                .TextMatrix(v_row, GRDC_NUM_POSTE), _
                                                v_um_spnum), _
                      rs) = P_ERREUR Then
            ajout_dans_synchro = False
            Exit Function
        End If
        If Not rs.EOF Then ' le triplé existe
            rs.Close
            ajout_dans_synchro = False
            Exit Function
        End If
        rs.Close
    End With

    ajout_dans_synchro = True

End Function

Public Function AppelFrm() As Integer

    Call FRM_ResizeForm(Me, Me.width, Me.Height)

    g_form_active = False

    Me.Show 1

    AppelFrm = g_importation_ok

    SendKeys "%{A}"

End Function

Private Sub associer_cette_personne(ByVal v_row As Integer, ByVal v_u_num As Long)
'**************************************************************************************
' La personne choisie dans la liste des personnes inactives ayant le même NOM et
'  g_nbr_car_prenom premiers caractères du PRENOM sera associée à la personne en cours:
'  elle aura le même MATRICULE, sera active et non externe au fichier d'importation.
'**************************************************************************************
    Dim sql As String
    Dim lng As Long, um_spnum As Long, u_numposte As Long
    Dim rs As rdoResultset

    With grd
        If Odbc_Update("Utilisateur", "U_Num", _
                       "WHERE U_Num=" & v_u_num, _
                       "U_Prenom", .TextMatrix(v_row, GRDC_PRENOM), _
                       "U_Matricule", .TextMatrix(v_row, GRDC_MATRICULE), _
                       "U_Actif", True, _
                       "U_KB_Actif", True, _
                       "U_ExterneFich", False, _
                       "U_Importe", True) = P_ERREUR Then
            Exit Sub
        End If
        If P_InsertIntoUtilmouvement(v_u_num, "A", "", 0) = P_ERREUR Then
            Exit Sub
        End If
        If Odbc_RecupVal("SELECT U_Po_Princ FROM Utilisateur where U_Num=" & v_u_num, u_numposte) = P_ERREUR Then
            Exit Sub
        End If
        .TextMatrix(v_row, GRDC_NUM_POSTE) = u_numposte
        If ajout_dans_synchro(v_row, um_spnum) Then
' Supprimé RV car possibilité de mettre une mauvaise synchro si ce n'est pas le même poste (13/03/2010)
            'If Odbc_AddNew("Synchro", "Sync_Num", "Sync_Seq", False, lng, _
            '               "Sync_Section", .TextMatrix(v_row, GRDC_CODE_SRV_FICH), _
            '               "Sync_Emploi", .TextMatrix(v_row, GRDC_CODE_POSTE_FICH), _
            '               "Sync_SPNum", .TextMatrix(v_row, GRDC_NUM_POSTE), _
            '               "Sync_auto", False) = P_ERREUR Then
            '    Exit Sub
            'End If
        End If
        .TextMatrix(v_row, GRDC_U_NUM) = v_u_num
        .TextMatrix(v_row, GRDC_LISTE_PERSONNE_INACTIVE) = ""
        .col = GRDC_PERSONNE_INACTIVE
        Set .CellPicture = LoadPicture("")
' A NE PAS SUPPRIMER
        'Call desactiver_ligne(v_row, IMG_PASTILLE_VERTE)
        'Call selectionner_ligne(v_row)
        If .Rows = 2 Then
            .Rows = 1
            Call actualiser_compteur
        Else
            If Not p_traitement_background Then
                .RemoveItem (v_row)
            Call actualiser_compteur
            End If
        End If
    End With

End Sub

Private Function build_arbor_srv(ByVal v_numposte As Long) As String

    Dim s_srv As String, sql As String
    Dim concatener As Boolean
    Dim numsrv As Long, numsrv_pere As Long
    
    s_srv = ""
    
    sql = "SELECT po_srvnum FROM poste WHERE po_Num=" & v_numposte
    If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
        build_arbor_srv = ""
        Exit Function
    End If
    
    ' Récupérer la hiérarchie de mon_poste "Sx;Sy;Pz;"
    s_srv = "S" & numsrv & ";" & "P" & v_numposte & ";|"
    concatener = False
    Do
        sql = "SELECT SRV_NumPere FROM Service WHERE SRV_Num=" & numsrv
        If Odbc_RecupVal(sql, numsrv_pere) = P_ERREUR Then
            build_arbor_srv = ""
            Exit Function
        End If
        If concatener Then
            s_srv = "S" & numsrv & ";" & s_srv
        End If
        If numsrv_pere <> 0 Then
            numsrv = numsrv_pere
            concatener = True
        End If
    Loop Until numsrv_pere = 0
    
    build_arbor_srv = s_srv
    
End Function

Private Sub chercher(ByVal v_row As Integer, ByVal v_keycode As Integer)

    Dim I As Integer

    With grd
        .SelectionMode = flexSelectionByRow
        If v_row < .Rows - 1 Then
            For I = v_row + 1 To .Rows - 1
                If Mid$(.TextMatrix(I, GRDC_NOM), 1, 1) = Chr(v_keycode) Then
                    .Row = I
                    If I > 4 Then
                        .TopRow = I - 4
                    Else
                        .TopRow = I
                    End If
                    .Row = I
                    .col = GRDC_U_NUM
                    .RowSel = I
                    .ColSel = .Cols - 1
                    .SetFocus
                    Exit Sub
                End If
            Next I
            For I = 1 To v_row ' On recommence depuis le début jusqu'à la ligne en cours
                If Mid$(.TextMatrix(I, GRDC_NOM), 1, 1) = Chr(v_keycode) Then
                    .Row = I
                    If I > 4 Then
                        .TopRow = I - 4
                    Else
                        .TopRow = I
                    End If
                    .col = GRDC_U_NUM
                    .RowSel = I
                    .ColSel = .Cols - 1
                    .SetFocus
                    Exit Sub
                End If
            Next I
            ' On a pas trouvé la lettre recherchée => on déselectionne la ligne en cours
            .RowSel = .Row
            .ColSel = GRDC_U_NUM
        Else ' v_row = .Rows - 1
            For I = 1 To v_row ' On recommence depuis le début jusqu'à la ligne en cours
                If Mid$(.TextMatrix(I, GRDC_NOM), 1, 1) = Chr(v_keycode) Then
                    .Row = I
                    If I > 4 Then
                        .TopRow = I - 4
                    Else
                        .TopRow = I
                    End If
                    .col = GRDC_U_NUM
                    .RowSel = I
                    .ColSel = .Cols - 1
                    .SetFocus
                    Exit Sub
                End If
            Next I
        End If
    End With

End Sub

Private Function choisir_poste_synchro_struct(ByVal v_code_srv As String, _
                                              ByRef r_message_background As String, _
                                              ByVal v_code_poste As String) As Long

    Dim sql As String, lib_fct As String, lib_srv As String
    Dim bAssocAuto As Boolean
    Dim lib_poste As String
    Dim nbligne As Integer
    Dim numposte As Long, numsrv As Long, lnb As Long
    Dim rs As rdoResultset
    Dim s As String, ancnumposte As Long
    
Lab_Debut:
    p_background_synchro_auto = False
    ' chercher si on a une synchro à afficher, sinon demander le(s) poste(s) KaliBottin
    sql = "SELECT * FROM Synchro WHERE SYNC_Section=" & Odbc_String(v_code_srv) _
        & " AND SYNC_Emploi=" & Odbc_String(v_code_poste)
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        choisir_poste_synchro_struct = 0
        Exit Function
    End If
    bAssocAuto = False
    If rs.EOF Then ' on n'a pas trouvé de synchro
        rs.Close
        If p_traitement_background Then
            s = "Pas de Synchronisation pour "
            Call MAJ_Corps_Background(v_code_srv, v_code_poste, s, p_corps_background)
            p_background_synchro_auto = False
            Exit Function
        Else
            s = "Pas de Synchronisation pour "
            Call MAJ_Corps_Background(v_code_srv, v_code_poste, s, p_corps_background)
        End If
        numposte = choisir_service()
        If numposte = 0 Then
            choisir_poste_synchro_struct = 0
            Exit Function
        End If
        If FctTransposePoste(0, numposte, True) <= 0 Then
            'numposte = rs("sync_spnum").Value
        End If
        choisir_poste_synchro_struct = numposte
        Exit Function
    Else
        rs.MoveLast
        rs.MoveFirst
        If rs.RowCount = 1 Then
            If rs("sync_auto").Value Then
                ' Associer en Auto si un seul
                p_background_synchro_auto = True
                sql = "SELECT po_libelle FROM poste Where po_Num=" & rs("sync_spnum")
                If Odbc_RecupVal(sql, lib_poste) = P_ERREUR Then
                End If
                If True Or p_traitement_background Then
                    s = "Synchronisation Automatique (" & lib_poste & ") pour "
                    Call MAJ_Corps_Background(v_code_srv, v_code_poste, s, p_corps_background)
                End If
                bAssocAuto = True
                Me.cmd(CMD_CORRIGER_TOUS).Visible = True
                numposte = rs("sync_spnum").Value
                If FctTransposePoste(0, numposte, True) <= 0 Then
                    numposte = rs("sync_spnum").Value
                Else
                    If True Or p_traitement_background Then
                        sql = "SELECT po_libelle FROM poste Where po_Num=" & numposte
                        If Odbc_RecupVal(sql, lib_poste) = P_ERREUR Then
                        End If
                        r_message_background = r_message_background & " : Poste transposé en " & lib_poste
                        s = "Poste transposé en " & lib_poste
                        Call MAJ_Corps_Background(v_code_srv, v_code_poste, s, p_corps_background)
                    End If
                End If
                choisir_poste_synchro_struct = numposte
                Exit Function
            Else
                p_background_synchro_auto = False
                s = "Synchronisations non automatique pour "
                Call MAJ_Corps_Background(v_code_srv, v_code_poste, s, p_corps_background)
            End If
        Else
            p_background_synchro_auto = False
            s = "Plusieurs synchronisations possibles "
            Call MAJ_Corps_Background(v_code_srv, v_code_poste, s, p_corps_background)
        End If
    End If
    If p_traitement_background Then
        If p_background_synchro_auto Then
            GoTo Lab_Synchro_Auto_Background
        Else
            Exit Function
        End If
    End If

    Call CL_Init
    Call CL_InitTaille(0, -15)
    Call CL_InitTitreHelp("Postes associés à [ " & v_code_srv & " - " _
                        & v_code_poste & " ]", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    'Call CL_AddBouton("", p_chemin_appli + "\btnAssocierAutrePoste.bmp", vbKeyA, vbKeyF2, 2500)
    Call CL_AddBouton("Associer à un autre poste", "", vbKeyN, 0, 2500)
    Call CL_AddBouton("Supprimer cette association", "", vbKeyS, 1, 2500)
    Call CL_AddBouton("à partir de maintenant" & Chr(13) & Chr(10) & "Associer automatiquement", "", vbKeyS, 1, 2500)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    nbligne = 0
    ' Proposer en premier le SYNC_Auto s'il y en a un
    rs.MoveFirst
    While Not rs.EOF
        If rs("SYNC_Auto").Value Then
            Me.cmd(CMD_CORRIGER_TOUS).Visible = True
            ' il existe dans la table Synchro la paire (SECTION, EMPLOI) => l'ajouter dans la liste
            numposte = rs("sync_spnum").Value
            If FctTransposePoste(0, numposte, True) <= 0 Then
                numposte = rs("sync_spnum").Value
            End If
            sql = "SELECT FT_Libelle, PO_SRVNum FROM FctTrav, Poste" _
                & " WHERE FT_Num=PO_FTNum AND PO_Num=" & numposte
            If Odbc_RecupVal(sql, lib_fct, numsrv) = P_ERREUR Then
                choisir_poste_synchro_struct = 0
                Exit Function
            End If
            sql = "SELECT SRV_Nom FROM Service WHERE SRV_Num=" & numsrv
            If Odbc_RecupVal(sql, lib_srv) = P_ERREUR Then
                choisir_poste_synchro_struct = 0
                Exit Function
            End If
            Call CL_AddLigne("Auto" & vbTab & lib_fct & vbTab & lib_srv, _
                         rs("SYNC_Num"), _
                         numposte, _
                         rs("SYNC_Auto").Value)
            nbligne = nbligne + 1
        End If
        rs.MoveNext
    Wend
    ' Puis mettre les autres
    rs.MoveFirst
    While Not rs.EOF
        If Not rs("SYNC_Auto").Value Then
            Me.cmd(CMD_CORRIGER_TOUS).Visible = True
            ' il existe dans la table Synchro la paire (SECTION, EMPLOI) => l'ajouter dans la liste
            numposte = rs("sync_spnum").Value
            If FctTransposePoste(0, numposte, True) <= 0 Then
                numposte = rs("sync_spnum").Value
            End If
            sql = "SELECT FT_Libelle, PO_SRVNum FROM FctTrav, Poste" _
                & " WHERE FT_Num=PO_FTNum AND PO_Num=" & numposte
            If Odbc_RecupVal(sql, lib_fct, numsrv) = P_ERREUR Then
                choisir_poste_synchro_struct = 0
                Exit Function
            End If
            sql = "SELECT SRV_Nom FROM Service WHERE SRV_Num=" & numsrv
            If Odbc_RecupVal(sql, lib_srv) = P_ERREUR Then
                choisir_poste_synchro_struct = 0
                Exit Function
            End If
            Call CL_AddLigne(" " & vbTab & lib_fct & vbTab & lib_srv, _
                             rs("SYNC_Num"), _
                             numposte, _
                             rs("SYNC_Auto").Value)
            nbligne = nbligne + 1
        End If
        rs.MoveNext
    Wend
    
afficher_liste:
    ' Afficher la liste de tous les postes
    
    ChoixListe.Show 1
    ' Tester le choix:

    If CL_liste.retour = 2 Then ' Supprimer la synchro
        ' Supprimé car supprimait TOUTES lesassoc. (pas seulement cette ligne)
        'If Odbc_Delete("Synchro", "Sync_num", _
        '            "WHERE SYNC_Section=" & Odbc_String(.TextMatrix(v_row, GRDC_CODE_SRV_FICH)) _
        '            & " AND SYNC_Emploi=" & Odbc_String(.TextMatrix(v_row, GRDC_CODE_POSTE_FICH)), _
        '            lng) = P_ERREUR Then
        '    Exit Sub
        'End If
        If Odbc_Delete("Synchro", "Sync_num", _
                    "WHERE SYNC_Num=" & CL_liste.lignes(CL_liste.pointeur).num, _
                    lnb) = P_ERREUR Then
            choisir_poste_synchro_struct = 0
            Exit Function
        End If
        GoTo Lab_Debut
    End If
    If CL_liste.retour = 3 Then ' Associer en auto
        Call Odbc_Update("synchro", "SYNC_Num", _
                         "WHERE SYNC_Num=" & CL_liste.lignes(CL_liste.pointeur).num, _
                         "SYNC_Auto", True)
        MsgBox "L'association pour cette fonction - dans ce service, est devenue automatique"
        bAssocAuto = True
        choisir_poste_synchro_struct = numposte
        Exit Function
    End If
    If CL_liste.retour = 4 Then ' QUITTER
        choisir_poste_synchro_struct = 0
        Exit Function
    End If
    If CL_liste.retour = 1 Then ' AJOUTER
        numposte = choisir_service()
        If numposte = 0 Then
            GoTo afficher_liste
        End If
        Call FctTransposePoste(0, numposte, True)
        choisir_poste_synchro_struct = numposte
    End If
    
    ' CL_liste.retour = 0  ' OK
    If CL_liste.retour = 0 Or CL_liste.retour = 3 Then
        Me.cmd(CMD_CORRIGER_TOUS).Visible = True
    End If
    
    numposte = CL_liste.lignes(CL_liste.pointeur).tag
    
Lab_Synchro_Auto_Background:
    ' Service de remplacement
    ancnumposte = numposte
    If FctTransposePoste(0, numposte, True) < 0 Then
        choisir_poste_synchro_struct = 0
    Else
        If True Or p_traitement_background Then
            If numposte <> ancnumposte Then
                sql = "SELECT po_libelle FROM poste Where po_Num=" & numposte
                If Odbc_RecupVal(sql, lib_poste) = P_ERREUR Then
                End If
                s = " - Poste transposé en Poste=" & lib_poste
                r_message_background = r_message_background & Chr(13) & Chr(10) & "==> " & s
                s = "Poste transposé en " & lib_poste
                Call MAJ_Corps_Background(v_code_srv, v_code_poste, s, p_corps_background)
            End If
        End If
        choisir_poste_synchro_struct = numposte
    End If
    
End Function

Private Function MAJ_Corps_Background(ByVal v_codeServ, ByVal v_codeEmploi, ByVal v_string, ByRef r_Corps)
    Dim s As String
    
    If v_codeServ & v_codeEmploi = "" Then
        s = "==>  " & v_string
        r_Corps = r_Corps & Chr(13) & Chr(10) & s
    ElseIf InStr(r_Corps, v_codeServ & "-" & v_codeEmploi) = 0 Then
        s = v_codeServ & "-" & v_codeEmploi & "==>  " & v_string & " service=" & v_codeServ & " emploi=" & v_codeEmploi
        r_Corps = r_Corps & Chr(13) & Chr(10) & s
    End If
    'Debug.Print s
End Function

Private Function choisir_service() As Long

    Dim poste_selectionne As String
    Dim frm As Form

    Set frm = PrmService
    poste_selectionne = PrmService.AppelFrm("Choix d'un poste", "S", False, "1;", "P", False)
    Set frm = Nothing

    If poste_selectionne = "" Then
        choisir_service = 0
        Exit Function
    End If

    choisir_service = P_get_num_srv_poste(poste_selectionne, P_POSTE)

End Function

Private Sub creation(ByVal v_row As Integer)

' ************************************************************************************
' Créer le nouveau utilisateur selectionné, avec une possibilité d'annuler l'opération
' ************************************************************************************
    Dim sql As String, ma_civilite As String, lib_srv As String, lib_fct As String
    Dim s_srvpo As String, njf As String
    Dim concatener As Boolean
    Dim cr As Integer
    Dim bAssocAuto As Boolean
    Dim srv_num As Long
    Dim nbligne As Integer
    Dim cle_sync As Long, numposte As Long
    Dim mon_num_service As Long, mon_num_poste As Long, num_srv As Long, num_srv_pere As Long, lng As Long
    Dim rs As rdoResultset
    Dim po_actif As Boolean
    Dim lib_poste As String
    Dim libSrv As String
    Dim s As String
    Dim mess_bck As String
    Dim ok As Boolean
    Dim retKB As Integer
    
    With grd
        ok = True
        If True Or p_traitement_background Then
            mess_bck = "Création de " & .TextMatrix(v_row, GRDC_NOM) & "." & .TextMatrix(v_row, GRDC_PRENOM) & " (" & .TextMatrix(v_row, GRDC_MATRICULE) & " " & .TextMatrix(v_row, GRDC_LIB_SRV_FICH) & " " & .TextMatrix(v_row, GRDC_LIB_POSTE_FICH) & ") "
        End If
        p_background_synchro_auto = False
        numposte = choisir_poste_synchro_struct(.TextMatrix(v_row, GRDC_CODE_SRV_FICH), mess_bck, .TextMatrix(v_row, GRDC_CODE_POSTE_FICH))
        If numposte = 0 Then
            mess_bck = mess_bck & " pas de synchro "
            GoTo LabFin
        End If
        Call Odbc_RecupVal("select srv_nom from poste, service  where po_num=" & numposte & " and srv_num=po_srvnum", lib_srv)
        Call P_RecupPosteNomfct(numposte, lib_fct)
            
        If p_pos_civilite <> -1 Then
            ma_civilite = " * CIVILITE:" & vbTab & .TextMatrix(v_row, GRDC_CIVILITE) & vbCrLf
        Else
            ma_civilite = ""
        End If
        If p_pos_njf <> -1 Then
            njf = " * NJF:" & vbTab & vbTab & .TextMatrix(v_row, GRDC_NJF) & vbCrLf
        Else
            njf = ""
        End If
        ' On ajoute à la liste ce qui a été choisi (après confirmation, sauf si bAssocAuto)
        Call Odbc_RecupVal("select srv_num from poste, service  where po_num=" & numposte & " and srv_num=po_srvnum", srv_num)
        libSrv = P_FctRecupNiveau(srv_num)
        libSrv = IIf(libSrv = "", "SERVICE", libSrv)
        
        s = "Créer un compte pour : " & ma_civilite & " NOM: " & .TextMatrix(v_row, GRDC_NOM) _
                & njf _
                & " PRENOM: " & .TextMatrix(v_row, GRDC_PRENOM) _
                & " MATRICULE:" & .TextMatrix(v_row, GRDC_MATRICULE) _
                & " " & libSrv & ":" & lib_srv _
                & " POSTE:" & lib_fct _
                & " (Fichier : " & .TextMatrix(v_row, GRDC_LIB_SRV_FICH) & " / " & .TextMatrix(v_row, GRDC_LIB_POSTE_FICH) & ")"
        If p_traitement_background Then
            If p_background_synchro_auto Then
                GoTo lab_ok
            Else
                p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> " & mess_bck & " pas de synchro "
                GoTo LabFin
            End If
        'Else
        '    p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> " & mess_bck
        End If
        If MsgBox("Vous allez créer un compte pour la personne suivante :" & vbCrLf & vbCrLf _
                & ma_civilite _
                & " * NOM:" & vbTab & vbTab & .TextMatrix(v_row, GRDC_NOM) & vbCrLf _
                & njf _
                & " * PRENOM:" & vbTab & .TextMatrix(v_row, GRDC_PRENOM) & vbCrLf _
                & " * MATRICULE:" & vbTab & .TextMatrix(v_row, GRDC_MATRICULE) & vbCrLf _
                & " * " & libSrv & ":" & vbTab & lib_srv & vbCrLf _
                & " * POSTE:" & vbTab & lib_fct & vbCrLf _
                & " (Fichier : " & .TextMatrix(v_row, GRDC_LIB_SRV_FICH) & " / " & .TextMatrix(v_row, GRDC_LIB_POSTE_FICH) & ")" & vbCrLf & vbCrLf _
                & "Confirmez-vous cette opération ?", _
                vbQuestion + vbYesNo, "Création d'une personne") = vbYes Then
lab_ok:
            ' Récupérer la hiérarchie de mon_poste "Sx;Sy;Pz;"
            ' Réactiver seulement kb_actif si false
            retKB = RéactiverKB(.TextMatrix(v_row, GRDC_MATRICULE))
            If retKB Then
                mess_bck = "Le MATRICULE " & .TextMatrix(v_row, GRDC_MATRICULE) & " (" & .TextMatrix(v_row, GRDC_PRENOM) & " " & .TextMatrix(v_row, GRDC_NOM) & ")" & " a été Réactivé"
                If Not p_traitement_background Then MsgBox mess_bck
                Call actualiser_compteur
                grd.RemoveItem (v_row)
                GoTo LabFin
            End If
            s_srvpo = build_arbor_srv(numposte)
            If creer_cette_personne(v_row, s_srvpo, mess_bck) <> P_OUI Then
                'mess_bck = mess_bck & " non faite"
                ok = False
                'p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> " & mess_bck & " non faite"
                GoTo LabFin
            Else
                'mess_bck = mess_bck & " faite"
                'p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> " & mess_bck
                ' si le poste n'est pas actif => le ré-activer
                sql = "SELECT po_actif, po_libelle FROM poste Where po_Num=" & numposte
                If Odbc_RecupVal(sql, po_actif, lib_poste) = P_ERREUR Then
                    MsgBox "Erreur SQL fonction creation SQL=" & sql
                Else
                    If Not po_actif Then
                        Call Odbc_Update("Poste", "PO_Num", _
                         "WHERE PO_num=" & numposte, _
                         "PO_Actif", True)
                         s = "Le poste " & lib_poste & " (inactif) a été ré-activé"
                        If p_traitement_background Then
                            mess_bck = mess_bck & Chr(13) & Chr(10) & "==> " & s
                        Else
                            MsgBox s
                            mess_bck = mess_bck & Chr(13) & Chr(10) & "==> " & s
                        End If
                    End If
                End If
            End If
        End If
        'Call desactiver_ligne(v_row, IMG_PASTILLE_VERTE)
        If grd.Rows = 2 Then
            grd.Rows = 0
        Else
            If Not p_traitement_background Then
                Call actualiser_compteur
                grd.RemoveItem (v_row)
            End If
        End If
LabFin:
        If ok Then
            p_mess_fait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> " & mess_bck & " faite"
        Else
            p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> " & mess_bck & " non faite"
        End If
    End With
    
End Sub

Private Function RéactiverKB(ByVal v_matricule)
    '
    Dim sql As String, rs As rdoResultset
    Dim u_actif As Boolean, u_kb_actif As Boolean
    
    sql = "select u_actif,u_kb_actif from utilisateur where u_Matricule='" & v_matricule & "'"
    Call Odbc_SelectV(sql, rs)
    If Not rs.EOF Then
        Call Odbc_RecupVal("select u_actif,u_kb_actif from utilisateur where u_Matricule='" & v_matricule & "'", u_actif, u_kb_actif)
        If rs("u_actif").Value Then
            If Not rs("u_kb_actif").Value Then
                Call Odbc_Update("Utilisateur", "U_Num", "WHERE U_matricule='" & v_matricule & "'", "U_KB_Actif", True)
                RéactiverKB = True
                Exit Function
            End If
        End If
    End If
    rs.Close
    
    RéactiverKB = False
End Function
Private Function creer_cette_personne(ByVal v_row As Integer, _
                                      ByVal v_spm As String, _
                                      ByRef r_mess_bck As String) As Integer
' ************************************************************************
' Appelée depuis creation() afin de céer le compte d'une nouvelle personne
' ************************************************************************
    Dim mon_nom As String, ma_civilite As String, mon_prenom As String, mon_spm As String, mon_matricule As String, _
        sql As String, new_code As String, new_mdp As String
    Dim ajouter_dans_synchro As Boolean
    Dim lib_poste As String
    Dim new_u_num As Long, lng As Long
    Dim frm As Form
    Dim s As String
    Dim NewSrv As Long, NewPoste As Long

    ajouter_dans_synchro = False
    With grd
        Call Odbc_BeginTrans
        ' Tester si on gere la civilité
        If Odbc_AddNew("Utilisateur", "U_Num", "U_Seq", True, new_u_num, _
                       "U_Nom", .TextMatrix(v_row, GRDC_NOM), _
                       "U_Prenom", .TextMatrix(v_row, GRDC_PRENOM), _
                       "U_NomJunon", .TextMatrix(v_row, GRDC_NOM_JUNON), _
                       "U_PrenomJunon", .TextMatrix(v_row, GRDC_PRENOM_JUNON), _
                       "U_Prefixe", .TextMatrix(v_row, GRDC_CIVILITE), _
                       "U_Actif", True, "U_Externe", False, "U_ExterneFich", False, _
                       "U_Importe", True, _
                       "U_AR", True, _
                       "U_SPM", v_spm, "U_FctTrav", P_get_fcttrav(v_spm), _
                       "U_Labo", "L1;", "U_LNumPrinc", 1, _
                       "U_Matricule", .TextMatrix(v_row, GRDC_MATRICULE), _
                       "U_DONumLast", 0, _
                       "U_LstDocs", "", _
                       "U_CATPNum", 0, _
                       "U_DateDebEmbauche", Null, "U_DateFinEmbauche", Null, _
                       "U_CTRAVNum", 0, "U_NbHeures", 0, _
                       "U_BaseHeures", 0, _
                       "U_POTNumNext", 0, "U_LNumNext", 0, "U_NoSemNext", 0, _
                       "U_Po_Princ", P_get_num_srv_poste(v_spm, P_POSTE), _
                       "U_kw_mailauth", True, _
                       "U_kb_actif", True, _
                       "U_Fictif", False) = P_ERREUR Then
            Call Odbc_RollbackTrans
            creer_cette_personne = P_ERREUR
            Exit Function
        End If
        cmd(CMD_ACTUALISER).Enabled = True
        new_code = get_code_mdp(v_row, "CODE")
        If p_mdp = p_format_code Then
            new_mdp = new_code
        Else
            new_mdp = get_code_mdp(v_row, "MDP")
        End If
        If new_code = "" Or new_mdp = "" Then
            Call Odbc_RollbackTrans
            creer_cette_personne = P_NON
            Exit Function
        End If
        r_mess_bck = r_mess_bck & " avec le Code : " & new_code & " mdPasse : " & new_mdp
        If Odbc_AddNew("UtilAppli", "UAPP_Num", "UAPP_Seq", False, lng, _
                       "UAPP_UNum", new_u_num, _
                       "UAPP_APPNum", p_appli_kalidoc, _
                       "UAPP_Code", new_code, _
                       "UAPP_TypeCrypt", p_Mode_Auth_UtilAppli, _
                       "UAPP_MotPasse", STR_Crypter_New(new_mdp)) = P_ERREUR Then
            Call Odbc_RollbackTrans
            creer_cette_personne = P_ERREUR
            Exit Function
        End If
        ' utilmouvement
        If P_InsertIntoUtilmouvement(new_u_num, "C", "", 0) = P_ERREUR Then
            Call Odbc_RollbackTrans
            creer_cette_personne = P_ERREUR
            Exit Function
        End If
        ' informations supplémentaires
        If insertInfoSuppl(new_u_num, v_row) = P_ERREUR Then
            Call Odbc_RollbackTrans
            creer_cette_personne = P_ERREUR
            Exit Function
        End If
        ' Renseigner les champs manquants
        .TextMatrix(v_row, GRDC_NUM_SRV_KB) = P_get_num_srv_poste(v_spm, P_SERVICE)
        .TextMatrix(v_row, GRDC_NUM_POSTE) = P_get_num_srv_poste(v_spm, P_POSTE)
        
        ' Service de remplacement
        NewSrv = .TextMatrix(v_row, GRDC_NUM_SRV_KB)
        NewPoste = .TextMatrix(v_row, GRDC_NUM_POSTE)
        If FctTransposePoste(NewSrv, NewPoste, True) > 0 Then
            .TextMatrix(v_row, GRDC_NUM_SRV_KB) = NewSrv
            .TextMatrix(v_row, GRDC_NUM_POSTE) = NewPoste
            If True Or p_traitement_background Then
                sql = "SELECT po_libelle FROM poste Where po_Num=" & NewPoste
                If Odbc_RecupVal(sql, lib_poste) = P_ERREUR Then
                End If
                s = " - Poste transposé en Srv=" & NewSrv & " Poste=" & lib_poste
                p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> " & s
                s = "Poste transposé en " & lib_poste
                Call MAJ_Corps_Background(NewSrv, NewPoste, s, p_corps_background)
            End If
        End If
        
        .TextMatrix(v_row, GRDC_U_NUM) = new_u_num
        If ajout_dans_synchro(v_row) Then
            If Odbc_AddNew("Synchro", "Sync_Num", "Sync_Seq", False, lng, _
                           "Sync_Section", .TextMatrix(v_row, GRDC_CODE_SRV_FICH), _
                           "Sync_Emploi", .TextMatrix(v_row, GRDC_CODE_POSTE_FICH), _
                           "Sync_SPNum", .TextMatrix(v_row, GRDC_NUM_POSTE), _
                           "Sync_auto", False) = P_ERREUR Then
                Call Odbc_RollbackTrans
                creer_cette_personne = P_ERREUR
                Exit Function
            End If
        End If
        Call Odbc_CommitTrans
        
        ' Gérer les postes secondaires
        Call GererPosteSecondaire("C", .TextMatrix(v_row, GRDC_PRENOM) & " " & .TextMatrix(v_row, GRDC_NOM), NewSrv, NewPoste, new_u_num, .TextMatrix(v_row, GRDC_MATRICULE), .TextMatrix(v_row, GRDC_NEW_PSTSECOND), .TextMatrix(v_row, GRDC_ANC_PSTSECOND))
        
        s = p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";KaliDoc;" & p_nom_fichier_ini_kalidoc & ";CONNEXION=" & p_NumUtil & ";CREERPERS=" & new_u_num & "[WAIT];KBAUTO"
        Call SYS_ExecShell(s, True, True)
        If p_traitement_background Then
            g_aller_dans_prm = PRM_JAMAIS
        Else
            If g_aller_dans_prm = PRM_QUEST Then
                Call demander_si_aller_dans_prm
            End If
        End If
        If g_aller_dans_prm = PRM_OUI Then
            Set frm = PrmPersonne
            Call PrmPersonne.AppelFrm(new_u_num, "")
            Set frm = Nothing
        End If
        
        ' Detrminer la couleur de la pastille
        Call maj_ligne(v_row, False)
        ' Prevenir la personne à qui envoyer par mail Code et MPasse de cette personne
' A NE PAS SUPPRIMER
'        Call envoyer_code_mpasse(v_row, new_code)
        Call selectionner_ligne(v_row)
    End With

    creer_cette_personne = P_OUI
End Function

Private Function GererPosteSecondaire(v_trait As String, v_titre As String, v_NumSrvPrinc As Long, v_NumPostePrinc As Long, v_unum As Long, v_umatricule As String, v_new_pst_second As String, v_anc_pst_second As String)
    ' Supprimer tous les postes de la personne qui sont dans la table historique (kb_poste_secondaire)
    ' reconstruire ses postes avec : poste principal + ce qui reste + ceux de la table temporaire
    Dim I As Integer, j As Integer
    Dim strManuel As String
    Dim sN As String, sA As String
    Dim sql As String
    Dim strPS As String
    Dim strSuppr As String
    Dim trouvé As Boolean
    Dim rs As rdoResultset
    Dim s As String
    Dim n As Integer
    Dim nb As Integer
    Dim rs1 As rdoResultset
    Dim rsHisto As rdoResultset
    Dim rsTmp As rdoResultset
    Dim new_spm As String
    Dim anc_spm As String, po_princ As String
    Dim str_po_princ As String
    Dim new_po_princ As Long
    Dim str_new_po_princ As String
    Dim str_new_spm_second_fichier As String
    Dim str_ponum As String, table_tempo As String
    Dim déjà As Boolean
    Dim laS As String
    
    table_tempo = "temp_utilisateur_postes_secondaires_" & p_NumUtil
    
    If v_trait = "D" Then
        ' Supprimer tous les postes de la personne qui sont dans la table historique (kb_poste_secondaire)
        sql = "delete from kb_poste_secondaire where psu_matricule='" & v_umatricule & "'"
        Call Odbc_Cnx.Execute(sql)
        ' Supprimer de la table temporaire
        sql = "delete from " & table_tempo & " where ttd_matricule='" & v_umatricule & "'"
        Call Odbc_Cnx.Execute(sql)
        Exit Function
    End If
    ' Lire l'utilisateur
    If Odbc_RecupVal("Select u_spm, u_po_princ from utilisateur where u_num = " & v_unum, anc_spm, po_princ) = P_ERREUR Then
        p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==>  " & "Erreur dans GererPosteSecondaire v_trait=" & v_trait & " v_NumSrvPrinc=" & v_NumSrvPrinc & " v_NumPostePrinc=" & v_NumPostePrinc & " v_unum=" & v_unum & " v_umatricule=" & v_umatricule
        Exit Function
    End If
    ' Reformer correctement v_anc_pst_second (enlever les doublons)
    s = ""
    For I = 0 To STR_GetNbchamp(v_anc_pst_second, "|")
        sN = STR_GetChamp(v_anc_pst_second, "|", I)
        If sN <> "" And InStr(s, sN) = 0 Then
            s = s & sN & "|"
        End If
    Next I
    v_anc_pst_second = s
    ' Supprimer dans anc_spm tous les postes de la personne qui sont dans la table historique
    ' ce sont ses anciens postes secondaires (kb_poste_secondaire)
    sql = "select * from kb_poste_secondaire where psu_unum=" & v_unum
    If Odbc_SelectV(sql, rsHisto) = P_ERREUR Then
        p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==>  " & "Erreur dans GererPosteSecondaire sql=" & sql
    End If
    While Not rsHisto.EOF
        anc_spm = Replace(anc_spm, rsHisto("psu_poste") & "|", "")
        rsHisto.MoveNext
    Wend
    str_po_princ = build_arbor_srv(v_NumPostePrinc)
    ' conserver ceux qui ont été mis manuellement
    anc_spm = Replace(anc_spm, str_po_princ, "")
    v_new_pst_second = Replace(v_new_pst_second, str_po_princ, "")
    rsHisto.Close
    
    Call CL_Init
    Call CL_InitTitreHelp(v_titre, "")
    Call CL_InitMultiSelect(True, True)
    Call CL_InitGererTousRien(True)
    nb = 0
    ' les nouveaux postes
    For I = 0 To STR_GetNbchamp(v_new_pst_second, "|")
        sN = STR_GetChamp(v_new_pst_second, "|", I)
        If sN <> "" Then
            trouvé = False
            s = STR_GetChamp(sN, ";P", 1)
            s = Replace(s, "|", "")
            s = Replace(s, ";", "")
            s = recup_PSLib(s)
            For j = 0 To STR_GetNbchamp(v_anc_pst_second, "|")
                sA = STR_GetChamp(v_anc_pst_second, "|", j)
                If sA <> "" Then
                    If sN = sA Then ' ce nouveau poste y était déja : le mettre en coché direct
                        trouvé = True
                        Call CL_AddLigne("Conserver => " & s, 0, sN, True)
                        nb = nb + 1
                    End If
                End If
            Next j
            If Not trouvé Then  ' ce nouveau poste n'y était pas : le proposer en coché
                Call CL_AddLigne("Ajouter => " & s, 0, sN, True)
                nb = nb + 1
            End If
        End If
    Next I
    ' les postes qui ont disparu
    For I = 0 To STR_GetNbchamp(v_anc_pst_second, "|")
        sA = STR_GetChamp(v_anc_pst_second, "|", I)
        If sA <> "" Then
            trouvé = False
            s = STR_GetChamp(sA, ";P", 1)
            s = Replace(s, "|", "")
            s = Replace(s, ";", "")
            s = recup_PSLib(s)
            For j = 0 To STR_GetNbchamp(v_new_pst_second, "|")
                sN = STR_GetChamp(v_new_pst_second, "|", j)
                If sN <> "" Then
                    If sN = sA Then ' cet ancien poste est toujours là : le mettre en coché
                        trouvé = True
                        déjà = False
                        For n = 0 To nb - 1
                            If CL_liste.lignes(n).tag = sA Then
                                déjà = True
                            End If
                        Next n
                        If Not déjà Then
                            Call CL_AddLigne("Conserver => " & s, 0, sA, True)
                            nb = nb + 1
                        End If
                        Exit For
                    End If
                End If
            Next j
            If Not trouvé Then  ' cet ancien poste a disparu : le proposer en non coché
                Call CL_AddLigne("Supprimer => " & s, 0, sA, True)
                nb = nb + 1
            End If
        End If
    Next I
    ' Ceux qui ont été mis manuellement
    For I = 0 To STR_GetNbchamp(anc_spm, "|")
        sA = STR_GetChamp(anc_spm, "|", I)
        If sA <> "" Then
            trouvé = False
            s = STR_GetChamp(sA, ";P", 1)
            s = Replace(s, "|", "")
            s = Replace(s, ";", "")
            s = recup_PSLib(s)
            déjà = False
            For n = 0 To nb - 1
                If CL_liste.lignes(n).tag = sA Then
                    déjà = True
                    CL_liste.lignes(n).texte = Replace(CL_liste.lignes(n).texte, "Ajouter", "Fusionner avec poste manuel")
                End If
            Next n
            If Not déjà Then
                Call CL_AddLigne("Poste manuel => " & s, 0, sA, True)
                nb = nb + 1
            End If
        End If
    Next I
    
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)

    If nb = 0 Then Exit Function
    Call CL_InitTaille(0, -15)
    
    ' Si seulement des conserver et postes manuel => rien à faire
    For n = 0 To nb - 1
        If Mid(CL_liste.lignes(n).texte, 1, 12) <> "Poste manuel" And Mid(CL_liste.lignes(n).texte, 1, 9) <> "Conserver" Then
            GoTo Suite
        End If
    Next n
    GoTo Fin
Suite:
    If Not p_traitement_background Then
        ChoixListe.Show 1
    ElseIf p_traitement_background_semiauto Then
        ChoixListe.Show 1
    Else    ' en auto
        CL_liste.retour = 0
    End If
    ' Sortie
    strPS = ""
    strSuppr = ""
    strManuel = ""
    If CL_liste.retour = 1 Then
        bFaireRemove = False
        Exit Function
    End If
    For I = 0 To nb - 1
        If CL_liste.lignes(I).selected Then
            If Mid(CL_liste.lignes(I).texte, 1, 12) = "Poste manuel" Then
                strManuel = strManuel & CL_liste.lignes(I).tag & "|"
            Else
                If CL_liste.lignes(I).selected Then
                    If Mid(CL_liste.lignes(I).texte, 1, 7) = "Ajouter" Then
                        strPS = strPS & CL_liste.lignes(I).tag & "|"
                    ElseIf Mid(CL_liste.lignes(I).texte, 1, 9) = "Fusionner" Then
                        strPS = strPS & CL_liste.lignes(I).tag & "|"
                    ElseIf Mid(CL_liste.lignes(I).texte, 1, 9) = "Supprimer" Then
                        strSuppr = strSuppr & CL_liste.lignes(I).tag & "|"
                    ElseIf Mid(CL_liste.lignes(I).texte, 1, 9) = "Conserver" Then
                        strPS = strPS & CL_liste.lignes(I).tag & "|"
                    End If
                End If
            End If
        End If
    Next I
    str_po_princ = build_arbor_srv(po_princ)
    strPS = Replace(strPS, str_po_princ, "")
    new_spm = str_po_princ & strPS & strManuel
    new_spm = Controle_doublon_dans_spm(new_spm)
    ' Supprimer tous les postes de la personne qui sont dans la table historique (kb_poste_secondaire)
    sql = "delete from kb_poste_secondaire where psu_matricule='" & v_umatricule & "'"
    Call Odbc_Cnx.Execute(sql)
    ' Reconstruire la table historique (kb_poste_secondaire)
    For I = 0 To STR_GetNbchamp(strPS, "|")
        s = STR_GetChamp(strPS, "|", I)
        If s <> "" Then
            If s & "|" <> str_po_princ Then
                sql = "INSERT INTO kb_poste_secondaire(psu_unum, psu_matricule, psu_poste, psu_faire_maj) VALUES(" & v_unum & ", '" & v_umatricule & "', '" & s & "', 't' )"
                Call Odbc_Cnx.Execute(sql)
                str_ponum = STR_GetChamp(s, ";P", 1)
                str_ponum = Replace(str_ponum, "|", "")
                str_ponum = Replace(str_ponum, ";", "")
                ' mettre le message si pas Conserver
                déjà = False
                For n = 0 To nb - 1
                    If CL_liste.lignes(n).tag = s Then
                        If Mid(CL_liste.lignes(n).texte, 1, 7) = "Ajouter" Or Mid(CL_liste.lignes(n).texte, 1, 9) = "Fusionner" Or Mid(CL_liste.lignes(n).texte, 1, 9) = "Supprimer" Then
                            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==>  " & "Matricule:" & v_umatricule & " => Poste secondaire : " & CL_liste.lignes(n).texte
                        End If
                    End If
                Next n
            End If
        End If
    Next I
    ' Modifier l'utilisateur
    sql = "update utilisateur set u_spm ='" & new_spm & "' where u_num=" & v_unum
    Call Odbc_Cnx.Execute(sql)
    ' Supprimer de la table temporaire
    sql = "delete from " & table_tempo & " where ttd_matricule='" & v_umatricule & "'"
    Call Odbc_Cnx.Execute(sql)
    ' supprimer les postes à supprimer
    For I = 0 To STR_GetNbchamp(strSuppr, "|")
        s = STR_GetChamp(strSuppr, "|", I)
        If s <> "" Then
            sql = "DELETE FROM kb_poste_secondaire where psu_unum=" & v_unum & " AND psu_poste='" & s & "'"
            Call Odbc_Cnx.Execute(sql)
        End If
    Next I
Fin:
End Function

Private Function recup_PSLib(ByVal v_numposte As Long) As String
    
    Dim s As String, lib As String, sql As String
    Dim numsrv As Long
    Dim rs As rdoResultset
    
    lib = ""

    sql = "select PO_SRVNum, FT_Libelle from Poste, FctTrav" _
        & " where PO_Num=" & v_numposte _
        & " and FT_Num=PO_FTNum"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        recup_PSLib = "???"
        Exit Function
    End If
    If rs.EOF Then
        rs.Close
        recup_PSLib = "???"
        Exit Function
    End If
    lib = rs("FT_Libelle").Value
    numsrv = rs("PO_SRVNum").Value
    rs.Close
    sql = "select SRV_Nom from Service" _
        & " where SRV_Num=" & numsrv
    If Odbc_SelectV(sql, rs) = P_OK Then
        If Not rs.EOF Then
            lib = lib & " - " & rs("SRV_Nom").Value
        End If
        rs.Close
    End If
    
    recup_PSLib = lib

End Function

Private Sub demander_si_aller_dans_prm()

    Dim tbl_libelle(3) As String, tbl_tooltip(3) As String, mess As String
    Dim cr As Integer
    Dim frm As Form
    
lab_choix:
    tbl_libelle(0) = "Oui"
    tbl_tooltip(0) = ""
    tbl_libelle(1) = "Non"
    tbl_tooltip(1) = ""
    tbl_libelle(2) = "Jamais"
    tbl_tooltip(2) = ""
    tbl_libelle(3) = "Toujours"
    tbl_tooltip(3) = ""
    Set frm = Com_Message
    mess = "Voulez-vous aller sur la fiche des personnes créées ?"
    cr = Com_Message.AppelFrm(mess, _
                              "", _
                              tbl_libelle(), _
                              tbl_tooltip())
    Set frm = Nothing
    
    If cr = 0 Then
        g_aller_dans_prm = PRM_OUI
    ElseIf cr = 1 Then
        g_aller_dans_prm = PRM_NON
    ElseIf cr = 2 Then
        g_aller_dans_prm = PRM_JAMAIS
    ElseIf cr = 3 Then
        g_aller_dans_prm = PRM_TOUJOURS
    End If
    
End Sub

Private Sub desactiver_ligne(ByVal v_row As Integer, ByVal v_etat_pastille As Integer)
' **********************************************************************************
' Mettre la couleur de la ligne en COLOR_DESACTIVE et enlèver l'image de GRDC_ACTION
' **********************************************************************************
    Dim I As Integer

    With grd
        .Row = v_row
        For I = 0 To .Cols - 1
            .col = I
            .CellBackColor = COLOR_DESACTIVE
            .CellFontBold = False
        Next I
        .TextMatrix(v_row, GRDC_ACTION) = "Accéder"
        .TextMatrix(.Rows - 1, GRDC_ETAT_AVANT) = "Accéder"
        .col = GRDC_ACTION
        .CellFontBold = True
        .col = GRDC_PASTILLE
        Set .CellPicture = imglst.ListImages(v_etat_pastille).Picture
        .col = GRDC_U_NUM
    End With

End Sub

Private Function desactiver_personne(ByVal v_row As Integer, _
                                     ByVal v_bauto As Boolean) As Boolean
    
    Dim util_est_actif As Boolean
    Dim s As String
    Dim mess_bck As String
    
    With grd
        If Not v_bauto Then
            s = " désactiver " & .TextMatrix(v_row, GRDC_NOM) _
                    & " " & .TextMatrix(v_row, GRDC_PRENOM)
            If p_traitement_background Then GoTo Lab_Auto_Background
            If MsgBox("Etes-vous sûr de vouloir " & s & " ?", _
                    vbQuestion + vbYesNo, _
                    "Demande de confirmation") = vbNo Then
                desactiver_personne = False
                Exit Function
            End If
        End If
Lab_Auto_Background:
        If P_RemplacerResponsableAppli(.TextMatrix(v_row, GRDC_U_NUM), .TextMatrix(v_row, GRDC_NOM), _
                                       .TextMatrix(v_row, GRDC_PRENOM), .TextMatrix(v_row, GRDC_MATRICULE) _
                                       ) = P_ERREUR Then
            desactiver_personne = False
            Exit Function
        End If
        cmd(CMD_ACTUALISER).Enabled = True
        Call Odbc_Update("Utilisateur", "U_Num", _
                         "WHERE U_Num=" & .TextMatrix(v_row, GRDC_U_NUM), _
                         "U_KB_Actif", False, _
                         "U_Actif", False)
        If P_InsertIntoUtilmouvement(.TextMatrix(v_row, GRDC_U_NUM), "I", "", 0) = P_ERREUR Then
            desactiver_personne = False
            Exit Function
        End If
        s = p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";KaliDoc;" & p_nom_fichier_ini_kalidoc & ";CONNEXION=" & p_NumUtil & ";DESACTIVER_PERS=" & .TextMatrix(v_row, GRDC_U_NUM) & "[WAIT];KBAUTO"
        Call SYS_ExecShell(s, True, True)
        ' Voir si l'utilisateur a été désactivé par KaliDoc (ou retour arrière)
        If Odbc_RecupVal("SELECT U_Actif FROM Utilisateur" _
                        & " Where U_Num = " & .TextMatrix(v_row, GRDC_U_NUM), _
                        util_est_actif) = P_ERREUR Then
            desactiver_personne = False
            Exit Function
        End If
        
        If .TextMatrix(v_row, GRDC_ETAT_AVANT) = "Désactiver" Then
            ' la personne ne se trouve pas dans le fichier d'importation et est devenue INACTIVE
            If .Rows = .FixedRows + 1 Then
                .Rows = .FixedRows
            Else
                If util_est_actif Then
                    If p_traitement_background Then
                        s = .TextMatrix(v_row, GRDC_NOM) & "." & .TextMatrix(v_row, GRDC_PRENOM) & " (" & .TextMatrix(v_row, GRDC_MATRICULE) & " " & .TextMatrix(v_row, GRDC_LIB_SRV_FICH) & " " & .TextMatrix(v_row, GRDC_LIB_POSTE_FICH) & ") " & " n'a pas été désactivé(e) car elle a des rôles dans KaliDoc"
                        p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> " & s
                        'Exit Function
                    ElseIf Not v_bauto Then
                        s = .TextMatrix(v_row, GRDC_NOM) & "." & .TextMatrix(v_row, GRDC_PRENOM) & " (" & .TextMatrix(v_row, GRDC_MATRICULE) & " " & .TextMatrix(v_row, GRDC_LIB_SRV_FICH) & " " & .TextMatrix(v_row, GRDC_LIB_POSTE_FICH) & ") " & " n'a pas été désactivé(e) car elle a des rôles dans KaliDoc"
                        p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> " & s
                        MsgBox "Cette personne n'a pas été désactivée"
                    End If
                    desactiver_personne = False
                    Exit Function
                Else
                    If p_traitement_background Then
                        s = .TextMatrix(v_row, GRDC_NOM) & "." & .TextMatrix(v_row, GRDC_PRENOM) & " (" & .TextMatrix(v_row, GRDC_MATRICULE) & " " & .TextMatrix(v_row, GRDC_LIB_SRV_FICH) & " " & .TextMatrix(v_row, GRDC_LIB_POSTE_FICH) & ") " & " a été désactivé(e)"
                        p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> " & s
                        'Exit Function
                    Else
                        s = .TextMatrix(v_row, GRDC_NOM) & "." & .TextMatrix(v_row, GRDC_PRENOM) & " (" & .TextMatrix(v_row, GRDC_MATRICULE) & " " & .TextMatrix(v_row, GRDC_LIB_SRV_FICH) & " " & .TextMatrix(v_row, GRDC_LIB_POSTE_FICH) & ") " & " a été désactivé(e)"
                        p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> " & s
                    End If
                    
                    ' Gérer les postes secondaires
                    Call GererPosteSecondaire("D", .TextMatrix(v_row, GRDC_NOM) & "." & .TextMatrix(v_row, GRDC_PRENOM), 0, 0, .TextMatrix(v_row, GRDC_U_NUM), .TextMatrix(v_row, GRDC_MATRICULE), "", "")
                    If Not p_traitement_background Then
                        .RemoveItem (v_row)
                    End If
                    'Call Odbc_Update("Utilisateur", "U_Num", _
                    '                 "WHERE U_Num=" & .TextMatrix(v_row, GRDC_U_NUM), _
                    '                 "U_kb_Actif", False)
                    
                End If
            End If
            If Not p_traitement_background Then
                Call actualiser_compteur
            End If
            desactiver_personne = True
            Exit Function
        End If
' A NE PAS SUPPRIMER
        'Call desactiver_ligne(v_row, IMG_PASTILLE_ROUGE)
        If .Rows = .FixedRows + 1 Then
            .Rows = .FixedRows
        Else
            .RemoveItem (v_row)
        End If
        If Not p_traitement_background Then Call actualiser_compteur
        desactiver_personne = True
    End With

End Function

Private Sub desactiver_tous()

    Dim I As Integer
    
    I = 1
    While I < grd.Rows
        grd.Row = I
        If grd.TextMatrix(I, GRDC_ACTION) = "Désactiver" Then
            If Not desactiver_personne(I, True) Then
                I = I + 1
            End If
        Else
            I = I + 1
        End If
    Wend
    cmd(CMD_DESACTIVER_TOUS).Visible = False

End Sub

Private Sub envoyer_question()

    Dim sql As String, choix As String, nom As String, prenom As String, _
        nom_pers As String, prenom_pers As String, message As String
    Dim unum As Long
    Dim frm As Form
    Dim cr As Integer
    Dim rs As rdoResultset
    Dim question As String
    Dim MATRICULE As String, fonction As String
    Dim lesRoles As String, nomdata As String
    Dim fd As Integer, I As Integer
    
    unum = val(cmd(CMD_QUESTION).tag)
    Call SAIS_Init
    Call SAIS_InitTitreHelp("Votre question", "")
    Call SAIS_AddBouton("", p_chemin_appli + "\btnEnregistrer.gif", vbKeyE, vbKeyF1, 0)
    Call SAIS_AddBouton("Annuler", "", 0, 0, 1000)
    Call SAIS_AddChamp("Question", 180, SAIS_TYP_TOUT_CAR, False, " ?")
    Saisie.Show 1
    If SAIS_Saisie.retour = 1 Then
        Exit Sub
    End If
    ' Valeur retournée
    question = SAIS_Saisie.champs(0).sval
    If unum > 0 Then
        lesRoles = liste_roles(cmd(CMD_QUESTION).tag)
    End If
    If lesRoles = "" Then
        nomdata = ""
    Else
        nomdata = p_chemin_appli & "\tmp\KbDate.txt"
        If FICH_OuvrirFichier(nomdata, FICH_ECRITURE, fd) = P_ERREUR Then
            MsgBox "Erreur pour " & nomdata
        End If
        For I = 0 To STR_GetNbchamp(lesRoles, Chr(13) & Chr(10))
            Print #fd, STR_GetChamp(lesRoles, Chr(13) & Chr(10), I)
        Next I
        Close #fd
    End If
Lab_Debut:
    ' --------------------------- CHOISIR LA PERSONNE --------------------------
    cr = P_SelectionnerPersonne("La personne à qui envoyer le Mail", unum, nom, prenom)
    If cr = P_ERREUR Then
        Exit Sub
    ElseIf cr = P_PERSONNE_SELECTIONNEE_NON Then
        GoTo Lab_Debut
    End If
    ' --------------------------- ENVOYER LE MAIL ------------------------------
    sql = "SELECT U_Matricule FROM Utilisateur WHERE U_Num=" & unum
    If Odbc_RecupVal(sql, MATRICULE) = P_ERREUR Then
        MsgBox "Erreur"
    End If
    message = "Concernant " & prenom & " " & nom & "( matricule : " & MATRICULE & ") :" _
          & vbCrLf & vbCrLf & " Liste de ses rôles"
    message = message & vbCrLf & vbCrLf & lesRoles
    Call P_EnvoyerMessage(unum, "", question, message, nomdata)

End Sub
Private Function liste_roles(v_numutil)

    Dim nom As String, prenom As String, sql As String, tblib_docs() As String
    Dim tblib_for() As String, saction As String, s As String, lib As String
    Dim sDest As String, sfct As String
    Dim a_plusieurs_postes As Boolean
    Dim lig As Integer, ndocs As Integer, idocs As Integer, icycle As Integer
    Dim nfor As Integer, ifor As Integer, n As Integer, I As Integer, ispm As Integer
    Dim nspm As Integer, ifct As Integer, nfct As Integer
    Dim lnb As Long, tbnum_docs() As Long, tbnum_for() As Long, largeur As Long
    Dim numposte As Long, numfct As Long
    Dim spm As Variant
    Dim rs As rdoResultset, rs2 As rdoResultset
    Dim strret As String
    
    Call P_ChargerCycles

    If Odbc_RecupVal("select U_Nom, U_Prenom, u_spm, u_fcttrav from Utilisateur" _
                        & " where U_Num=" & v_numutil, _
                     nom, prenom, spm, sfct) = P_ERREUR Then
        Call quitter(True)
        Exit Function
    End If
    a_plusieurs_postes = IIf(STR_GetNbchamp(spm, "|") > 1, True, False)
    
    ' Groupe
    sql = "select GU_Num, GU_Nom from GroupeUtil" _
        & " where GU_Lst like '%U" & v_numutil & "|%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Call quitter(True)
        Exit Function
    End If
    While Not rs.EOF
        Call faire_une_ligne(v_numutil, strret, "G", rs("GU_Num").Value, 0, "", "Groupe : " & rs("GU_Nom").Value, 0)
        rs.MoveNext
    Wend
    rs.Close
    
    ' Référentiels
    sql = "select REF_Num, REF_Nom from Referentiel" _
        & " where REF_LstAutor like '%U" & v_numutil & ";%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Call quitter(True)
        Exit Function
    End If
    While Not rs.EOF
        Call faire_une_ligne(v_numutil, strret, "R", rs("REF_Num").Value, 0, "", "Référentiel : " & rs("REF_Nom").Value, 0)
        rs.MoveNext
    Wend
    rs.Close
    
    ' Documents
    sql = "select DO_Num, DO_Titre from documentation" _
        & " order by do_titre"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Call quitter(True)
        Exit Function
    End If
    ndocs = -1
    While Not rs.EOF
        ndocs = ndocs + 1
        ReDim Preserve tbnum_docs(ndocs) As Long
        tbnum_docs(ndocs) = rs("do_num").Value
        ReDim Preserve tblib_docs(ndocs) As String
        tblib_docs(ndocs) = rs("do_titre").Value
        rs.MoveNext
    Wend
    rs.Close
    For idocs = 0 To ndocs
        ' Superviseur docs
        sql = "select count(*) from Documentation" _
            & " where DO_Num=" & tbnum_docs(idocs) _
            & " and DO_lstsuperv like '%U" & v_numutil & ";%'"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            Call quitter(True)
            Exit Function
        End If
        If lnb > 0 Then
            lib = "Superviseur de " & tblib_docs(idocs)
            Call faire_une_ligne(v_numutil, strret, "D", tbnum_docs(idocs), -1, "", lib, 0)
        End If
        ' Resp documents
        sql = "select count(*) from Document" _
            & " where D_DONum=" & tbnum_docs(idocs) _
            & " and (D_UNumResp=" & v_numutil _
            & " or D_LstResp like '%U" & v_numutil & ";%')"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            Call quitter(True)
            Exit Function
        End If
        If lnb > 0 Then
            lib = "Responsable de documents dans " & tblib_docs(idocs)
            Call faire_une_ligne(v_numutil, strret, "D", tbnum_docs(idocs), 0, "D", lib, 0)
            GoTo lab_doc_acteur
        End If
        ' Resp dossiers
        sql = "select count(*) from Dossier" _
            & " where DS_DONum=" & tbnum_docs(idocs) _
            & " and DS_lstresp like '%U" & v_numutil & ";%'"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            Call quitter(True)
            Exit Function
        End If
        If lnb > 0 Then
            lib = "Responsable de dossiers dans " & tblib_docs(idocs)
            Call faire_une_ligne(v_numutil, strret, "D", tbnum_docs(idocs), 0, "S", lib, 0)
            GoTo lab_doc_acteur
        End If
        ' Resp docs
        sql = "select count(*) from Documentation" _
            & " where DO_Num=" & tbnum_docs(idocs) _
            & " and DO_lstresp like '%U" & v_numutil & ";%'"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            Call quitter(True)
            Exit Function
        End If
        If lnb > 0 Then
            lib = "Responsable de " & tblib_docs(idocs)
            Call faire_une_ligne(v_numutil, strret, "D", tbnum_docs(idocs), 0, "O", lib, 0)
        End If
            
lab_doc_acteur:
        For icycle = 1 To UBound(p_scycledocs())
            If p_scycledocs(icycle).acteur = "" Then
                GoTo lab_doc_cycle_suiv
            End If
            ' Acteurs paramétrés de documents
            sql = "select du_ponum, count(*) from Document, DocUtil" _
                & " where D_DONum=" & tbnum_docs(idocs) _
                & " and DU_DNum=D_Num" _
                & " and DU_UNum=" & v_numutil _
                & " and DU_Cyordre=" & icycle _
                & " group by du_ponum"
            If Odbc_SelectV(sql, rs2) = P_ERREUR Then
                Call quitter(True)
                Exit Function
            End If
            If Not rs2.EOF Then
                lib = p_scycledocs(icycle).acteur & " de documents dans " & tblib_docs(idocs)
                While Not rs2.EOF
                    Call faire_une_ligne(v_numutil, strret, "D", tbnum_docs(idocs), icycle, "D", lib, rs2("du_ponum").Value)
                    rs2.MoveNext
                Wend
                rs2.Close
                GoTo lab_doc_cycle_suiv
            Else
                rs2.Close
            End If
            ' Acteurs temporaires de documents
            sql = "select dac_ponum, count(*) from Document, Docaction" _
                & " where D_DONum=" & tbnum_docs(idocs) _
                & " and DAC_DNum=D_Num" _
                & " and DAC_UNum=" & v_numutil _
                & " and DAC_Cyordre=" & icycle _
                & " group by dac_ponum"
            If Odbc_SelectV(sql, rs2) = P_ERREUR Then
                Call quitter(True)
                Exit Function
            End If
            If Not rs2.EOF Then
                lib = p_scycledocs(icycle).acteur & " de documents dans " & tblib_docs(idocs)
                While Not rs2.EOF
                    Call faire_une_ligne(v_numutil, strret, "D", tbnum_docs(idocs), icycle, "D", lib, rs2("dac_ponum").Value)
                    rs2.MoveNext
                Wend
                rs2.Close
                GoTo lab_doc_cycle_suiv
            Else
                rs2.Close
            End If
            ' Acteurs paramétrés de dossiers
            sql = "select dsu_ponum, count(*) from Dossier, DosUtil" _
                & " where DS_DONum=" & tbnum_docs(idocs) _
                & " and DSU_DSNum=DS_Num" _
                & " and DSU_UNum=" & v_numutil _
                & " and DSU_Cyordre=" & icycle _
                & " group by dsu_ponum"
            If Odbc_SelectV(sql, rs2) = P_ERREUR Then
                Call quitter(True)
                Exit Function
            End If
            If Not rs2.EOF Then
                lib = p_scycledocs(icycle).acteur & " de dossiers dans " & tblib_docs(idocs)
                While Not rs2.EOF
                    Call faire_une_ligne(v_numutil, strret, "D", tbnum_docs(idocs), icycle, "D", lib, rs2("dsu_ponum").Value)
                    rs2.MoveNext
                Wend
                rs2.Close
                GoTo lab_doc_cycle_suiv
            Else
                rs2.Close
            End If
            ' Acteurs paramétrés de la documentation
            sql = "select dou_ponum from DocsUtil" _
                & " where DOU_DONum=" & tbnum_docs(idocs) _
                & " and DOU_UNum=" & v_numutil _
                & " and DOU_Cyordre=" & icycle
            If Odbc_SelectV(sql, rs2) = P_ERREUR Then
                Call quitter(True)
                Exit Function
            End If
            If Not rs2.EOF Then
                lib = p_scycledocs(icycle).acteur & " de " & tblib_docs(idocs)
                Call faire_une_ligne(v_numutil, strret, "D", tbnum_docs(idocs), icycle, "O", lib, rs2("dou_ponum").Value)
            End If
            rs2.Close
            ' Acteurs pour sous traitance documents
            sql = "select count(*) from Document, Docaction" _
                & " where D_DONum=" & tbnum_docs(idocs) _
                & " and DAC_DNum=D_Num" _
                & " and DAC_UNumModif=" & v_numutil _
                & " and DAC_Cyordre=" & icycle
            If Odbc_Count(sql, lnb) = P_ERREUR Then
                Call quitter(True)
                Exit Function
            End If
            If lnb > 0 Then
                lib = p_scycledocs(icycle).acteur & " de documents dans " & tblib_docs(idocs)
                Call faire_une_ligne(v_numutil, strret, "D", tbnum_docs(idocs), icycle, "A", lib, 0)
                GoTo lab_doc_cycle_suiv
            End If
lab_doc_cycle_suiv:
        Next icycle
        
        ' Destinataire
        sql = "select count(*) from Document" _
            & " where D_DONum=" & tbnum_docs(idocs) _
            & " and D_Dest like '%U" & v_numutil & "|%'"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            Call quitter(True)
            Exit Function
        End If
        If lnb > 0 Then
            Call faire_une_ligne(v_numutil, strret, "D", tbnum_docs(idocs), P_DESTINATAIRE, "D", "Destinataire personnel de documents dans " & tblib_docs(idocs), 0)
            GoTo lab_docs_suiv
        End If
        ' L'util possède des originaux
        sql = "select count(*) from Document, DocDiffusion" _
                & " where D_DONum=" & tbnum_docs(idocs) _
                & " and DD_DNum=D_Num" _
                & " and DD_UNum=" & v_numutil _
                & " and DD_Exemplaire<>''"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            Call quitter(True)
            Exit Function
        End If
        If lnb > 0 Then
            Call faire_une_ligne(v_numutil, strret, "D", tbnum_docs(idocs), P_DESTINATAIRE, "D", "Destinataire personnel avec original de documents dans " & tblib_docs(idocs), 0)
            GoTo lab_docs_suiv
        End If
        sql = "select count(*) from Document, DocPrmDiffusion" _
                & " where D_DONum=" & tbnum_docs(idocs) _
                & " and DPD_DNum=D_Num" _
                & " and DPD_UNum=" & v_numutil _
                & " and DPD_Exemplaire<>''"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            Call quitter(True)
            Exit Function
        End If
        If lnb > 0 Then
            Call faire_une_ligne(v_numutil, strret, "D", tbnum_docs(idocs), P_DESTINATAIRE, "D", "Destinataire personnel avec original de documents dans " & tblib_docs(idocs), 0)
            GoTo lab_docs_suiv
        End If
        
        sql = "select count(*) from Dossier, DosUtil" _
            & " where DS_DONum=" & tbnum_docs(idocs) _
            & " and DSU_DSNum=DS_Num" _
            & " and DSU_UNum=" & v_numutil _
            & " and DSU_Cyordre=" & P_DESTINATAIRE
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            Call quitter(True)
            Exit Function
        End If
        If lnb > 0 Then
            Call faire_une_ligne(v_numutil, strret, "D", tbnum_docs(idocs), P_DESTINATAIRE, "S", "Destinataire personnel de dossiers dans " & tblib_docs(idocs), 0)
            GoTo lab_docs_suiv
        End If
        sql = "select count(*) from DocsUtil" _
            & " where DOU_DONum=" & tbnum_docs(idocs) _
            & " and DOU_UNum=" & v_numutil _
            & " and DOU_Cyordre=" & P_DESTINATAIRE
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            Call quitter(True)
            Exit Function
        End If
        If lnb > 0 Then
            Call faire_une_ligne(v_numutil, strret, "D", tbnum_docs(idocs), P_DESTINATAIRE, "O", "Destinataire personnel de " & tblib_docs(idocs), 0)
        End If
lab_docs_suiv:
    Next idocs
    
    ' Formulaires
    sql = "select FOR_Num, FOR_Lib from formulaire" _
        & " order by for_lib"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Call quitter(True)
        Exit Function
    End If
    nfor = -1
    While Not rs.EOF
        nfor = nfor + 1
        ReDim Preserve tbnum_for(nfor) As Long
        tbnum_for(nfor) = rs("for_num").Value
        ReDim Preserve tblib_for(nfor) As String
        tblib_for(nfor) = rs("for_lib").Value
        rs.MoveNext
    Wend
    rs.Close
    For ifor = 0 To nfor
        ' Resp form
        sql = "select count(*) from Formulaire" _
            & " where FOR_Num=" & tbnum_for(ifor) _
            & " and (FOR_Lstresp like '%U" & v_numutil & ";%'" _
            & " or FOR_Lstrespkw like '%U" & v_numutil & ";%')"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            Call quitter(True)
            Exit Function
        End If
        If lnb > 0 Then
            lib = "Responsable du formulaire " & tblib_for(ifor)
            Call faire_une_ligne(v_numutil, strret, "F", tbnum_for(ifor), 0, "", lib, 0)
        End If
        
        ' Destinataire form etape 1
        sql = "select * from Formetape" _
            & " where FORE_FORNum=" & tbnum_for(ifor) _
            & " and FORE_numetape=1" _
            & " and FORE_Dest like '%U" & v_numutil & "|%'"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Call quitter(True)
            Exit Function
        End If
        If Not rs.EOF Then
            lib = "Personne nominative pouvant saisir le formulaire " & tblib_for(ifor)
            Call faire_une_ligne(v_numutil, strret, "F", tbnum_for(ifor), 1, "S", lib, 0)
        End If
        rs.Close
        
        ' Destinataire form etape > 1
        sql = "select * from Formaction" _
            & " where FORA_FORNum=" & tbnum_for(ifor) _
            & " and FORA_Type='D'" _
            & " and FORA_Action like '%U" & v_numutil & "|%'" _
            & " order by fora_numetape"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Call quitter(True)
            Exit Function
        End If
        If Not rs.EOF Then
            While Not rs.EOF
                lib = "Destinataire nominatif de l'étape " & rs("FORA_Numetape").Value & " du formulaire " & tblib_for(ifor)
                Call faire_une_ligne(v_numutil, strret, "F", tbnum_for(ifor), rs("FORA_Numetape").Value, "D", lib, 0)
                rs.MoveNext
            Wend
            rs.Close
            GoTo lab_form_prevenir
        Else
            rs.Close
        End If
        nspm = STR_GetNbchamp(spm, "|")
        For ispm = 0 To nspm - 1
            s = STR_GetChamp(spm, "|", ispm)
            n = STR_GetNbchamp(s, ";")
            numposte = Mid$(STR_GetChamp(s, ";", n - 1), 2)
            sql = "select * from Formaction" _
                & " where FORA_FORNum=" & tbnum_for(ifor) _
                & " and FORA_Type='D'" _
                & " and FORA_Action like '%P" & numposte & ";|%'" _
                & " order by fora_numetape"
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                Call quitter(True)
                Exit Function
            End If
            If Not rs.EOF Then
                While Not rs.EOF
                    lib = "Destinataire par son poste de l'étape " & rs("FORA_Numetape").Value & " du formulaire " & tblib_for(ifor)
                    Call faire_une_ligne(v_numutil, strret, "F", tbnum_for(ifor), rs("FORA_Numetape").Value, "Dp", lib, 0)
                    rs.MoveNext
                Wend
                rs.Close
                GoTo lab_form_prevenir
            Else
                rs.Close
            End If
        Next ispm
        nfct = STR_GetNbchamp(sfct, ";")
        For ifct = 0 To nfct - 1
            numfct = Mid$(STR_GetChamp(sfct, ";", ifct), 2)
            sql = "select * from Formaction" _
                & " where FORA_FORNum=" & tbnum_for(ifor) _
                & " and FORA_Type='D'" _
                & " and FORA_Action like '%F" & numfct & "|%'" _
                & " order by fora_numetape"
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                Call quitter(True)
                Exit Function
            End If
            If Not rs.EOF Then
                While Not rs.EOF
                    lib = "Destinataire par sa fonction de l'étape " & rs("FORA_Numetape").Value & " du formulaire " & tblib_for(ifor)
                    Call faire_une_ligne(v_numutil, strret, "F", tbnum_for(ifor), rs("FORA_Numetape").Value, "Df", lib, 0)
                    rs.MoveNext
                Wend
                rs.Close
                GoTo lab_form_prevenir
            Else
                rs.Close
            End If
        Next ifct
        
        ' Dest en cours form
        sql = "select donec_numetape, count(*) from donnees_encours" _
            & " where donec_fornum=" & tbnum_for(ifor) _
            & " and donec_unum=" & v_numutil _
            & " group by donec_numetape" _
            & " order by donec_numetape"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Call quitter(True)
            Exit Function
        End If
        While Not rs.EOF
            lib = "Acteur en cours de l'étape " & rs("donec_Numetape").Value & " du formulaire " & tblib_for(ifor)
            Call faire_une_ligne(v_numutil, strret, "F", tbnum_for(ifor), rs("donec_Numetape").Value, "d", lib, 0)
            rs.MoveNext
        Wend
        rs.Close
        
lab_form_prevenir:
        ' Prévenir Form
        sql = "select * from Formaction" _
            & " where FORA_FORNum=" & tbnum_for(ifor) _
            & " and FORA_Type='A'" _
            & " and FORA_Action like '%MAIL%'" _
            & " order by fora_numetape"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Call quitter(True)
            Exit Function
        End If
        While Not rs.EOF
            saction = rs("FORA_Action").Value
            n = STR_GetNbchamp(saction, "|")
            For I = 0 To n - 1
                s = STR_GetChamp(saction, "|", I)
                If STR_GetChamp(s, "%", 0) = "MAIL" Then
                    s = STR_GetChamp(s, "%", 1)
                    sDest = STR_GetChamp(s, ";", 1)
                    If InStr(sDest, "U" & v_numutil & "~") > 0 Then
                        Call faire_une_ligne(v_numutil, strret, "F", tbnum_for(ifor), rs("FORA_Numetape").Value, "M", "Mail pour l'étape " & rs("fora_numetape").Value & " du formulaire " & tblib_for(ifor), 0)
                    End If
                End If
            Next I
            rs.MoveNext
        Wend
        rs.Close
    Next ifor
    
    'Alertes
    ' kav_compte.kav_cpt_resp -> numutil;
    ' categalrt.catal_lstdiff -> Unum;
    ' categalrt.catal_resp -> Unum;
    
    'Projets
    'paq.paq_pilote -> Unum;x:|
    'paq.paq_grptrav -> Unum;x:|
    
    'Plans d'action
    'projtachrisk.pjtar_resp -> numutil
    'projtachrisk.pjtar_lstacteur -> Unumutil;..
    
    'Réunions
    'typereunion.treu_unumorg -> numutil
    'typereunion.treu_unumresp -> numutil
    'typereunion.treu_lstacteur_oj -> Unumutil;...
    'typereunion.treu_lstacteur_cr -> Unumutil;...
    'typereunion.treu_lstparticipant -> Unumutil;x;|...
    'typereunion.treu_lstdest_cr -> à vérifier
    'reunion.reu_unumorg -> numutil
    'reunion.reu_unumresp -> numutil
    'reunion.reu_lstacteur_oj -> Unumutil;...
    'reunion.reu_lstacteur_cr -> Unumutil;...
    'reunion.reu_lstparticipant -> Unumutil;x;|...
    'reunion.reu_lstparticipant_rdv -> Unumutil;...
    'reunion.reu_lstdest_cr -> à vérifier
    
    'Indicateurs
    'ki_param.kip_resp -> Unum;....
    'ki_declinaison.kide_actpot -> Unum;...
    
    'KaliBottin
    'application.app_lstresp -> Unum;...
    
    'piece.pc_resp -> Unum;...
    
    liste_roles = strret
    
End Function

' Charge les cycles associés à la documentation en cours dans p_scycledocs
Public Function P_ChargerCycles() As Integer

    Dim sql As String
    Dim I As Integer
    Dim rs As rdoResultset
    
    Erase p_scycledocs()
    p_cycle_relecture = -1
    p_cycle_verifliens = -1
    p_cycle_diffusion = -1
    p_cycle_consultable = -1
    
    ' Récupère les paramètres CYCLE de la documentation
    sql = "select * from Cycle" _
        & " order by CY_Ordre"
    If Odbc_Select(sql, rs) = P_ERREUR Then
        P_ChargerCycles = P_ERREUR
        Exit Function
    End If
    I = 0
    While Not rs.EOF
        ReDim Preserve p_scycledocs(I) As SCYCLEDOCS
        p_scycledocs(I).etape = rs("CY_Etape").Value
        p_scycledocs(I).acteur = rs("CY_Acteur").Value
        p_scycledocs(I).action = rs("CY_Action").Value
        p_scycledocs(I).ordre_si_refus = rs("CY_OrdreSiRefus").Value
        p_scycledocs(I).informer_si_refus = rs("CY_InformerSiRefus").Value
        p_scycledocs(I).modifiable = rs("CY_Modifiable").Value
        If rs("CY_Relecture").Value Then p_cycle_relecture = I
        If rs("CY_VerifLiens").Value Then p_cycle_verifliens = I
        If rs("CY_Diffusion").Value Then p_cycle_diffusion = I
        If rs("CY_Consultable").Value Then p_cycle_consultable = I
        rs.MoveNext
        I = I + 1
    Wend
    rs.Close
    
    P_ChargerCycles = P_OK

End Function

Private Sub faire_une_ligne(ByVal v_numutil, _
                        ByRef r_str, _
                        ByVal v_stype As String, _
                        ByVal v_num As Long, _
                        ByVal v_numetape As Long, _
                        ByVal v_sdetail As String, _
                        ByVal v_lib As String, _
                        ByVal v_numposte As Long)

    Dim sql As String, nomposte As String, nomutil As String
    Dim fsuppr As Boolean
    Dim lig As Integer
    Dim lnb As Long
    
    r_str = r_str & v_stype & " " & v_num & " " & v_numetape & " " & v_sdetail & " " & v_lib & " " & v_numposte & Chr(13) & Chr(10)
    If v_numposte > 0 Then
        Call Odbc_RecupVal("select po_libelle from poste where po_num=" & v_numposte, _
                           nomposte)
        r_str = r_str & " " & nomposte
    End If
    If P_RecupUtilNomP(v_numutil, nomutil) = P_ERREUR Then
        Exit Sub
    End If
    r_str = r_str & " " & nomutil
    
End Sub

Private Sub envoyer_code_mpasse(ByVal v_row As Integer, ByVal v_code As String)

    Dim sql As String, choix As String, nom As String, prenom As String, _
        nom_pers As String, prenom_pers As String, message As String
    Dim unum As Long
    Dim cr As Integer
    Dim rs As rdoResultset

    nom_pers = grd.TextMatrix(v_row, GRDC_NOM)
    prenom_pers = grd.TextMatrix(v_row, GRDC_PRENOM)
Lab_Debut:
    ' --------------------------- CHOISIR LA PERSONNE --------------------------
    cr = P_SelectionnerPersonne("La personne à qui envoyer le Mail", unum, nom, prenom)
    If cr = P_ERREUR Then
        Exit Sub
    ElseIf cr = P_PERSONNE_SELECTIONNEE_NON Then
        GoTo Lab_Debut
    End If
    ' --------------------------- ENVOYER LE MAIL ------------------------------
    message = "Voici le Code et le Mot de Passe de " & nom_pers & " " & prenom_pers & ":" _
          & vbCrLf & vbCrLf & "CODE:" & vbTab & vbTab & v_code _
          & vbCrLf & vbCrLf & "MOT DE PASSE:" & vbTab & p_mdp
    Call P_EnvoyerMessage(unum, "", _
                        "CODE et MOT DE PASSE du nouvel utilisateur", _
                        message)

End Sub

Private Function formater_prenom(ByVal v_prenom As String) As String
' **********************************************************
' Mettre la 1° lettre en majiscule, et le reste en miniscule
' **********************************************************
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
                sous_str = UCase$(Mid$(STR_GetChamp(v_prenom, " ", I), 1, 1)) _
                         & LCase$(Mid$(STR_GetChamp(v_prenom, " ", I), 2))
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
                sous_str = UCase$(Mid$(STR_GetChamp(v_prenom, "-", I), 1, 1)) _
                         & LCase$(Mid$(STR_GetChamp(v_prenom, "-", I), 2))
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

Private Function get_code_mdp(ByVal v_row As Integer, _
                              ByVal v_code_mdp As String) As String
' *************************************************************
' Retourne un code temporaire qui n'existe pas dans la base
' Un code qui ne comporte ni accents (`éëËÂò'...) ni vides
' *************************************************************
    Dim code_tempo As String, srecup As String, nom As String, prenom As String, stest As String
    Dim str As String, str2 As String, MATRICULE As String, format As String, njf As String
    Dim I As Integer, posdeb As Integer, n As Integer, pos As Integer
    Dim s As String
' **************************************************************************************
    
    ' Mot de passe en dur
    If v_code_mdp = "MDP" Then
        If InStr(p_mdp, "<") = 0 Then
            get_code_mdp = UCase(p_mdp)
            Exit Function
        End If
    End If
    
    nom = UCase$(grd.TextMatrix(v_row, GRDC_NOM))
    prenom = UCase$(grd.TextMatrix(v_row, GRDC_PRENOM))
    njf = UCase$(grd.TextMatrix(v_row, GRDC_NJF))
    MATRICULE = UCase$(grd.TextMatrix(v_row, GRDC_MATRICULE))
    
    ' LE NOM
    str = ""
    For I = 1 To Len(nom)
        ' éliminer les blancs et les apostrophes
        If Mid$(nom, I, 1) <> " " And Mid$(nom, I, 1) <> "'" And Mid$(nom, I, 1) <> "`" Then
            str = str & P_ChangerCar(Mid$(nom, I, 1), tbcaractere_nontraite)
        End If
    Next I
    nom = str
    
    ' LE NOM JF
    str = ""
    For I = 1 To Len(njf)
        ' éliminer les blancs et les apostrophes
        If Mid$(njf, I, 1) <> " " And Mid$(njf, I, 1) <> "'" And Mid$(njf, I, 1) <> "`" Then
            str = str & P_ChangerCar(Mid$(njf, I, 1), tbcaractere_nontraite)
        End If
    Next I
    njf = str
    
    ' LE PRENOM
    str = ""
    prenom = Replace(prenom, " ", "-")
    For I = 1 To Len(prenom)
        ' éliminer les blancs et les apostrophes
        If Mid$(prenom, I, 1) <> " " And Mid$(prenom, I, 1) <> "'" And Mid$(prenom, I, 1) <> "`" Then
            str = str & P_ChangerCar(Mid$(prenom, I, 1), tbcaractere_nontraite)
        End If
    Next I
    prenom = str
    
    ' LE Matricule
    str = ""
    For I = 1 To Len(MATRICULE)
        ' éliminer les blancs et les apostrophes
        If Mid$(MATRICULE, I, 1) <> " " And Mid$(MATRICULE, I, 1) <> "'" And Mid$(MATRICULE, I, 1) <> "`" Then
            str = str & P_ChangerCar(Mid$(MATRICULE, I, 1), tbcaractere_nontraite)
        End If
    Next I
    MATRICULE = str
    
    I = 1
    code_tempo = ""
    posdeb = 0
    If v_code_mdp = "CODE" Then
        format = p_format_code
    Else
        format = p_mdp
    End If
    While I <= Len(format)
        If Mid$(format, I, 1) = "<" Then
            If posdeb > 0 Then
                If p_traitement_background Then
                    s = "Le format du " & v_code_mdp & "'" & format & "' n'a pû être interprété."
                    GoTo lab_fin
                Else
                    Call MsgBox("Le format du " & v_code_mdp & "'" & format & "' n'a pû être interprété.", vbExclamation + vbOKOnly, "")
                    s = "Le format du " & v_code_mdp & "'" & format & "' n'a pû être interprété."
                    GoTo lab_saisie
                End If
            End If
            posdeb = I
        ElseIf Mid$(format, I, 1) = ">" Then
            If posdeb = 0 Then
                If p_traitement_background Then
                    s = "Le format du " & v_code_mdp & "'" & format & "' n'a pû être interprété."
                    GoTo lab_fin
                Else
                    Call MsgBox("Le format du " & v_code_mdp & "'" & format & "' n'a pû être interprété.", vbExclamation + vbOKOnly, "")
                    s = "Le format du " & v_code_mdp & "'" & format & "' n'a pû être interprété."
                    GoTo lab_saisie
                End If
            End If
            stest = Mid$(format, posdeb + 1, I - posdeb - 1)
            If left$(stest, 2) = "P=" Then
                srecup = RecupPrenom(prenom, stest)
            ElseIf left$(stest, 1) = "P" Then
                srecup = Replace(prenom, "-", "")
            ElseIf left$(stest, 1) = "p" Then
                srecup = prenom
            ElseIf left$(stest, 1) = "N" Then
                srecup = nom
            ElseIf left$(stest, 1) = "J" Then
                srecup = njf
            ElseIf left$(stest, 1) = "M" Then
                srecup = MATRICULE
            Else
                If p_traitement_background Then
                    s = "Le format du " & v_code_mdp & "'" & format & "' n'a pû être interprété."
                    GoTo lab_fin
                Else
                    s = "Le format du " & v_code_mdp & "'" & format & "' n'a pû être interprété."
                    Call MsgBox("Le format du " & v_code_mdp & "'" & format & "' n'a pû être interprété.", vbExclamation + vbOKOnly, "")
                    GoTo lab_saisie
                End If
            End If
            If InStr(stest, "=") > 0 Then
                code_tempo = code_tempo + srecup
            ElseIf Len(stest) > 1 Then
                If Not IsNumeric(Mid$(stest, 2)) Then
                    If p_traitement_background Then
                        s = "Le format du " & v_code_mdp & "'" & format & "' n'a pû être interprété."
                        GoTo lab_fin
                    Else
                        Call MsgBox("Le format du " & v_code_mdp & "'" & format & "' n'a pû être interprété.", vbExclamation + vbOKOnly, "")
                        s = "Le format du " & v_code_mdp & "'" & format & "' n'a pû être interprété."
                        GoTo lab_saisie
                    End If
                End If
                n = Mid$(stest, 2)
                If Len(srecup) >= n Then
                    If left$(stest, 1) = "p" And n = 1 Then
                        code_tempo = code_tempo + left$(srecup, 1)
                        pos = InStr(srecup, "-")
                        If pos > 0 Then
                            code_tempo = code_tempo + Mid$(srecup, pos + 1, 1)
                        End If
                    Else
                        code_tempo = code_tempo + left$(srecup, n)
                    End If
                Else
                    code_tempo = code_tempo + srecup
                End If
            Else
                code_tempo = code_tempo + srecup
            End If
            posdeb = 0
        Else
            If posdeb = 0 Then
                code_tempo = code_tempo + Mid$(p_format_code, I, 1)
            End If
        End If
        I = I + 1
    Wend
    
    If v_code_mdp = "CODE" Then
        ' Verification dans la base
        If P_ValiderCode(0, code_tempo, s) Then
            get_code_mdp = UCase$(code_tempo)
            Exit Function
        Else
            If p_traitement_background Then
                If p_traitement_background_semiauto Then
                    MsgBox s
                End If
                get_code_mdp = ""
                Exit Function
            End If
        End If
    Else
        If Len(code_tempo) > 15 Then
            code_tempo = left$(code_tempo, 15)
        End If
        get_code_mdp = UCase$(code_tempo)
        Exit Function
    End If

' **************************************************************************************
lab_saisie:
    str = ""
    str2 = code_tempo
lab_nouvel_essai:
    ' Le champ de saisie
    Call SAIS_Init
    Call SAIS_InitTitreHelp("Saisir le code" & str, "")
    Call SAIS_AddBouton("", p_chemin_appli + "\btnEnregistrer.gif", vbKeyE, vbKeyF1, 0)
    Call SAIS_AddBouton("Annuler", "", 0, 0, 1000)
    Call SAIS_AddChamp("Personne", -50, 0, True, nom & " " & grd.TextMatrix(v_row, GRDC_PRENOM))
    Call SAIS_AddChamp("Code", 15, SAIS_TYP_CODE, False, "" & str2)
    str2 = ""
    Saisie.Show 1
    If SAIS_Saisie.retour = 1 Then
        get_code_mdp = ""
        Exit Function
    End If
    ' Valeur retournée
    code_tempo = SAIS_Saisie.champs(1).sval
    If Len(code_tempo) = 0 Or Len(code_tempo) > 15 Then
        If Len(code_tempo) = 0 Then
            str = " (NON VIDE)"
        Else '  Len(code_tempo) > 15
            str = " (contenant au maximum 15 caractères)"
            str2 = left$(code_tempo, 15)
        End If
        GoTo lab_nouvel_essai
    End If
    ' Verification dans le dictionnaire
    If P_ValiderCode(0, code_tempo, s) Then
        get_code_mdp = UCase$(code_tempo)
    Else
        str = " ( " & code_tempo & " déjà existant)"
        GoTo lab_nouvel_essai
    End If
    p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==>  " & s
    Exit Function
' **************************************************************************************
lab_fin:
    p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==>  " & s
End Function

Private Function RecupPrenom(v_prenom, v_masque)
    Dim s1 As String, s2 As String
    Dim sOut As String
    Dim s As String
    Dim n1 As String, n2 As String
    Dim boolSimple As Boolean
    Dim n As Integer
    
    If Mid(v_masque, 2, 1) = "=" Then
        s = STR_GetChamp(v_masque, "=", 1)
        s1 = STR_GetChamp(s, ";", 0)
        s2 = STR_GetChamp(s, ";", 1)
        boolSimple = InStr(v_prenom, "-") = 0
        If left(s1, 2) = "PS" Then
            n1 = Replace(s1, "PS", "")
            n = n1
            If boolSimple Then
                sOut = Mid(v_prenom, 1, n)
            End If
        End If
        If left(s2, 2) = "PC" Then
            n2 = Replace(s2, "PC", "")
            If Not boolSimple Then
                sOut = Mid(STR_GetChamp(v_prenom, "-", 0), 1, 1) & Mid(STR_GetChamp(v_prenom, "-", 1), 1, 1)
            End If
        End If
    Else
        sOut = Replace(v_prenom, "-", "")
    End If
    RecupPrenom = sOut
End Function

Private Function RecupNom(v_nom, v_masque)

End Function

Private Function get_nom_infoSuppl(ByVal v_tis_num) As String

    Dim tis_libelle As String

    If Odbc_RecupVal("SELECT KB_TisLibelle FROM KB_TypeInfoSuppl" _
                    & " Where KB_TisNum = " & v_tis_num, _
                    tis_libelle) = P_ERREUR Then
        GoTo lab_erreur
    End If

    get_nom_infoSuppl = tis_libelle
    Exit Function

lab_erreur:

End Function

Private Function infosuppl_existant(ByVal v_tis_num As Long, _
                                    ByVal v_unum As Long, _
                                    ByVal v_infoSuppl As String) As Boolean
' chercher si cette info suppl existe déjà pour cette personne
    Dim sql As String
    Dim lng As Long

    sql = "SELECT COUNT(*) FROM InfoSupplEntite" _
      & " WHERE ISE_TisNum=" & v_tis_num & " AND ISE_Type='U'" _
      & " AND ISE_TypeNum=" & v_unum & " AND UPPER(ISE_Valeur)='" & UCase(v_infoSuppl) & "'"

    If Odbc_Count(sql, lng) = P_ERREUR Then
        infosuppl_existant = False
        Exit Function
    End If

    infosuppl_existant = (lng > 0)

End Function

Private Sub initialiser()

    Dim sql As String
    Dim I As Integer
    Dim lnb As Long
    
    g_aller_dans_prm = PRM_QUEST
    
    ' Suppression ds synchro des associations qui n'existent plus
    sql = "delete from synchro where sync_spnum not in (select po_num from poste)"
    Call Odbc_Cnx.Execute(sql)
    sql = "delete from synchro where sync_spnum in (select po_num from poste where po_actif='f')"
    Call Odbc_Cnx.Execute(sql)
    
    g_importation_ok = P_OK
    frm(FRM_PRINCIPALE).Caption = "Importation des personnes depuis le fichier: " _
                            & p_nom_fichier_importation & " " & IIf(p_est_sur_serveur, "(serveur)", "(local)")
    If Not g_actualiser Then
        frmPatience.left = (Me.width / 2) - (frmPatience.width / 2)
        frmPatience.Top = (Me.Height / 2) - (frmPatience.Height / 2)
        frmPatience.Visible = True
    End If
    g_row_context_menu = 0

    cmd(CMD_QUESTION).Visible = False
    With grd
        .Rows = 1
        .FormatString = "u_num|ACTION||MATRIC.|CIVILITE|NOM|NJF|PRENOM|CODE|SERVICE|CODE|POSTE|numSRV|numPOST" _
                      & "|||majNOM|postGRAS|PersonInnactiv|SPMaSYNCHRO|ligneLuePourA_Créer"
        .ScrollTrack = True
        .SelectionMode = flexSelectionByRow
        .ColWidth(GRDC_U_NUM) = 0
        .ColWidth(GRDC_ACTION) = 1050
        .ColWidth(GRDC_PASTILLE) = 0 ' A NE PAS SUPPRIMER 245
        .ColWidth(GRDC_MATRICULE) = 900
        .ColWidth(GRDC_CIVILITE) = 0
        .ColWidth(GRDC_NOM) = 1600
        .ColWidth(GRDC_NJF) = 0
        .ColWidth(GRDC_PRENOM) = 1200
        .ColWidth(GRDC_CODE_SRV_FICH) = 800
        .ColWidth(GRDC_LIB_SRV_FICH) = 2180 ' A NE PAS SUPPRIMER 2800
        .ColWidth(GRDC_CODE_POSTE_FICH) = 800
        .ColWidth(GRDC_LIB_POSTE_FICH) = 2180 ' A NE PAS SUPPRIMER 2600
        .ColWidth(GRDC_NUM_SRV_KB) = 0
        .ColWidth(GRDC_NUM_POSTE) = 0
        .ColWidth(GRDC_PERSONNE_INACTIVE) = 200
        .ColWidth(GRDC_INFO_PERSO) = 285
        .ColWidth(GRDC_METTRE_A_JOUR_NOM) = 0
        .ColWidth(GRDC_POSTE_EN_GRAS) = 0
        .ColWidth(GRDC_LISTE_PERSONNE_INACTIVE) = 0
        .ColWidth(GRDC_SPM_KB_A_SYNCHRONISER) = 0
        .ColWidth(GRDC_LIGNE_LUE) = 0
        .ColWidth(GRDC_ETAT_AVANT) = 0
        .ColWidth(GRDC_NOM_JUNON) = 0
        .ColWidth(GRDC_PRENOM_JUNON) = 0
        .ColWidth(GRDC_ANC_PSTSECOND) = 0
        .ColWidth(GRDC_NEW_PSTSECOND) = 0
        .ColAlignment(GRDC_ACTION) = POS_CENTRE
        .ColAlignment(GRDC_MATRICULE) = POS_GAUCHE
        .ColAlignment(GRDC_NOM) = POS_GAUCHE
        .ColAlignment(GRDC_PRENOM) = POS_GAUCHE
        .ColAlignment(GRDC_CODE_SRV_FICH) = POS_GAUCHE
        .ColAlignment(GRDC_LIB_SRV_FICH) = POS_GAUCHE
        .ColAlignment(GRDC_CODE_POSTE_FICH) = POS_GAUCHE
        .ColAlignment(GRDC_LIB_POSTE_FICH) = POS_GAUCHE
        .ColAlignment(GRDC_INFO_PERSO) = POS_CENTRE
        .Row = 0
        For I = 0 To .Cols - 1
            .col = I
            .CellFontBold = True
        Next I
        ' Les informations supplémentaires ARRET
        .Cols = .Cols + p_nbr_lstInfoSuppl
        For I = .Cols - p_nbr_lstInfoSuppl - 1 To .Cols - 1
            .ColWidth(I) = 0
        Next I
    End With

    Me.Refresh

    ' info suppl
    g_nbr_infoSuppl = 0
    ReDim g_LISTE_U_INFO_SUPPL(0)
    g_LISTE_U_INFO_SUPPL(0).unom = "Nom" & vbTab & "Prénom"
    For I = 0 To p_nbr_lstInfoSuppl - 1
        g_LISTE_U_INFO_SUPPL(0).unom = g_LISTE_U_INFO_SUPPL(0).unom & vbTab & get_nom_infoSuppl(LISTE_TIS_POS(I).prmgenb_tis_num)
    Next I

    Call remplir_grid

    frmPatience.Visible = False
    
    ' tester si on a des infos suppl à afficher
    If g_nbr_infoSuppl > 0 Then
        If afficher_infoSuppl <> P_OK Then
            Call quitter("Pb sur ImportationAnnuaire.afficher_infoSuppl")
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "*************************************************************"
            Print #g_fd1, p_mess_fait_background
            Exit Sub
        End If
    End If

    g_nbr_lignes = grd.Rows - 1
    Call actualiser_compteur
    lbl(LBL_COMPTEUR).Visible = True

    On Error Resume Next
    grd.SetFocus
    On Error GoTo 0
    
    If p_traitement_background Then
        Call Traitement_Background("")
    End If
End Sub

Public Function Traitement_Background(ByVal v_s As String)
    Dim I As Integer
    Dim FichLog As String
    Dim lstresp As String, sDest As String
    Dim fd As Integer
    Dim fd1 As Integer, fd2 As Integer
    Dim fpLog As Integer
    Dim nomfich As String
    Dim mouse_row As Integer, mouse_col As Integer
    Dim sAnc As String
    Dim sql As String, rs As rdoResultset
    Dim sMessage As String
    Dim nbfait As Integer, nbpasfait As Integer
    Dim cote As String, dblcote As String
    Dim s As String
    Dim ListeDest As String
    Dim iDest As Integer
    Dim titre As String
    Dim snom As String
    Dim sPrenom As String
    
    If v_s <> "" Then GoTo Lab_Fichier
    If p_traitement_background Then
        ' Faire toutes les actions
        I = 1
        While I < grd.Rows
            grd.Row = I
            mouse_row = grd.Row
            mouse_col = GRDC_ACTION
            grd.col = mouse_col
            g_poste = ""
            If mouse_row <> 0 Then
                grd.RowSel = mouse_row
                p_mess_fait_background = ""
                p_mess_pasfait_background = ""
                'p_corps_background = ""
                Call action(mouse_row)
                If p_mess_fait_background <> "" Then
                    p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "*************************************************************"
                    Print #g_fd1, p_mess_fait_background
                End If
                If p_mess_pasfait_background <> "" Then
                    p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "*************************************************************"
                    Print #g_fd2, p_mess_pasfait_background
                End If
            End If
            I = I + 1
        Wend
    End If
Lab_Fichier:
    If p_traitement_background Then
        titre = "Traitement automatique"
    Else
        If p_NumUtil = 1 Then
            snom = "ROOT"
        Else
            sql = "select u_nom, u_prenom from utilisateur where u_num=" & p_NumUtil
            If Odbc_RecupVal(sql, snom, sPrenom) <> P_ERREUR Then
                titre = "Traitement par " & sPrenom & " " & snom
            End If
        End If
    End If
    Call EnvoiLog(v_s, titre)
    If p_traitement_background And v_s <> "" Then
        End
    Else
        nomfich = p_chemin_appli & "\rapport\rapport_kalibottin_fait_" & p_NumUtil & ".txt"
        Call FICH_EffacerFichier(nomfich, False)
        Open nomfich For Append As g_fd1
        
        nomfich = p_chemin_appli & "\rapport\rapport_kalibottin_pasfait_" & p_NumUtil & ".txt"
        Call FICH_EffacerFichier(nomfich, False)
        Open nomfich For Append As g_fd2
        
        nomfich = p_chemin_appli & "\rapport\rapport_kalibottin_assoc_" & p_NumUtil & ".txt"
        Call FICH_EffacerFichier(nomfich, False)
        Open nomfich For Append As g_fd3
    End If
End Function

Private Function EnvoiLog(v_s As String, v_titre As String)
    Dim fpLog As Integer
    Dim fdh As Integer
    Dim nomfichHTML As String
    Dim nomfich As String, s As String
    Dim nbfait As Integer, nbpasfait As Integer
    Dim sql As String, rs As rdoResultset
    Dim dblcote As String, cote As String
    Dim I As Integer
    Dim FichLog As String, ListeDest As String, titre As String, sMessage As String, lstresp As String, sDest As String
    Dim iDest As Integer
    
    Close #g_fd1
    Close #g_fd2
    Print #g_fd3, p_corps_background
    Close #g_fd3
    
    nomfich = p_chemin_appli & "\rapport\rapport_kalibottin_fait_" & p_NumUtil & ".txt"
    Open nomfich For Input As g_fd1
    nbfait = 0
    While Not EOF(g_fd1)
        Line Input #g_fd1, s
        If Replace(s, "*", "") <> "" Then
            nbfait = nbfait + 1
        End If
    Wend
    Close #g_fd1
    
    nomfich = p_chemin_appli & "\rapport\rapport_kalibottin_pasfait_" & p_NumUtil & ".txt"
    Open nomfich For Input As g_fd2
    nbpasfait = 0
    While Not EOF(g_fd2)
        Line Input #g_fd2, s
        If Replace(s, "*", "") <> "" Then
            nbpasfait = nbpasfait + 1
        End If
    Wend
    Close #g_fd2
    
    nomfichHTML = p_chemin_appli & "\rapport"
    If FICH_EstFichierOuRep(nomfichHTML) = FICH_REP Then
    Else
        MkDir nomfichHTML
    End If
    g_nomfichHTML = p_chemin_appli & "\rapport\rapport_kalibottin_" & p_NumUtil & "_" & Replace(Date, "/", "-") & "_" & Replace(Time, ":", "_") & ".html"
    If FICH_OuvrirFichier(g_nomfichHTML, FICH_ECRITURE, fdh) = P_ERREUR Then
        Call MsgBox("impossible d'ouvrir " & nomfich)
        Exit Function
    End If
    Print #fdh, "<Head>"
    Print #fdh, " <Body>"
    
    If v_s <> "" Then
        Print #fdh, "  <B>" & v_s & "</B><BR>"
        GoTo Lab_Suite1
    End If
    
    dblcote = """"
    cote = "'"
    ' entete
    Print #fdh, "  <B><Span OnClick=" & dblcote & "window.document.getElementById(" & cote & "IdCorps" & cote & ").style.display = " & cote & cote & "; " & dblcote & "><font color=green>" & " Rapport des synchronisations</font></Span></B><BR>"
    If nbpasfait > 0 Then
        Print #fdh, "  <B><Span OnClick=" & dblcote & "window.document.getElementById(" & cote & "IdPasFaits" & cote & ").style.display = " & cote & cote & "; " & dblcote & "><font color=red>" & nbpasfait & " Mouvements Non effectués</font></Span></B><BR>"
    End If
    If nbfait > 0 Then
        Print #fdh, "  <B><Span OnClick=" & dblcote & "window.document.getElementById(" & cote & "IdFaits" & cote & ").style.display = " & cote & cote & "; " & dblcote & "><font color=blue>" & nbfait & " Mouvements effectués</font></Span></B><BR>"
    End If
    
    ' tableaux
    If nbfait > 0 Then
        nomfich = p_chemin_appli & "\rapport\rapport_kalibottin_fait_" & p_NumUtil & ".txt"
        If FICH_FichierExiste(nomfich) Then
            Open nomfich For Input As g_fd1
            Print #fdh, "  <Span id='IdFaits' style='display:none'>"
            Print #fdh, "   <Table border='1'>"
            Print #fdh, "    <Tr style:background:#3366FF><Td align=center><font color=blue><B>Modifications faites</B></font></Td></Tr>"
            While Not EOF(g_fd1)
                Line Input #g_fd1, s
                If Replace(s, "*", "") <> "" Then
                    s = Replace(s, vbCrLf, "")
                    s = Replace(s, "*", "")
                    Print #fdh, "    <Tr><Td>" & s & "</Td></Tr>"
                End If
            Wend
            Close #g_fd1
            Print #fdh, "   </Table>"
            Print #fdh, "  </Span>"
        End If
    End If
    
    If nbpasfait > 0 Then
        nomfich = p_chemin_appli & "\rapport\rapport_kalibottin_pasfait_" & p_NumUtil & ".txt"
        If FICH_FichierExiste(nomfich) Then
            Open nomfich For Input As g_fd2
            Print #fdh, "  <Span id='IdPasFaits' style='display:none'>"
            Print #fdh, "   <Table border='1'>"
            Print #fdh, "    <Tr style:background:#3366FF><Td align=center><font color=red><B>Modifications NON faites</B></font></Td></Tr>"
            While Not EOF(g_fd2)
                Line Input #g_fd2, s
                If Replace(s, "*", "") <> "" Then
                    s = Replace(s, vbCrLf, "")
                    s = Replace(s, "*", "")
                    Print #fdh, "    <Tr><Td>" & s & "</Td></Tr>"
                End If
            Wend
            Close #g_fd2
            Print #fdh, "   </Table>"
            Print #fdh, "  </Span>"
        End If
    End If
    
    nomfich = p_chemin_appli & "\rapport\rapport_kalibottin_assoc_" & p_NumUtil & ".txt"
    If FICH_FichierExiste(nomfich) Then
        Open nomfich For Input As g_fd3
        Print #fdh, "  <Span id='IdCorps' style='display:none'>"
        Print #fdh, "   <Table border='1'>"
        Print #fdh, "    <Tr style:background:#3366FF><Td align=center><font color=green><B>Rapport des associations</B></font></Td></Tr>"
        While Not EOF(g_fd3)
            Line Input #g_fd3, s
            If s <> "" Then
                's = STR_GetChamp(p_corps_background, vbCrLf, i)
                If s <> vbCrLf And s <> "" Then
                    s = Replace(s, vbCrLf, "")
                    s = Replace(s, "*", "")
                    Print #fdh, "   <Tr><Td>" & s & "</Td></Tr>"
                End If
            End If
        Wend
        Close #g_fd3
        Print #fdh, "   </Table>"
        Print #fdh, "  </Span>"
    End If
Lab_Suite1:
    
    FichLog = SYS_GetIni("PREIMPORT", "FICHIER_LOG", p_nomini)
    If FichLog <> "" Then
        'FichLog = "\\192.168.101.20\kalidoc\Sources_VB\Outils\Import_Kalidoc\CH_Mulhouse\RECUP_PERSONNEL\Structure_KaliBottin.err"
        If FICH_FichierExiste(FichLog) Then
            fpLog = FreeFile
            FICH_OuvrirFichier FichLog, FICH_LECTURE, fpLog
            ' pré_import
            Print #fdh, "  <B><Span OnClick=" & dblcote & "window.document.getElementById(" & cote & "IdPreImport" & cote & ").style.display = " & cote & cote & "; " & dblcote & "><font color=red>" & " Rapport des erreurs du Pré_import</font></Span></B><BR>"
            ' le fichier
            Print #fdh, "  <Span id='IdPreImport' style='display:none'>"
            Print #fdh, "   <Table border='1'>"
            Print #fdh, "    <Tr style:background:#3366FF><Td align=center><font color=blue><B>Rapport des erreurs du Pré_import</B></font></Td></Tr>"
                
            While Not EOF(fpLog)
                Line Input #fpLog, s
                Print #fdh, "   <Tr><Td>" & s & "</Td></Tr>"
            Wend
            Close #fpLog
            ' Call FICH_EffacerFichier(FichLog, False)
            Print #fdh, "   </Table>"
            Print #fdh, "  </Span>"
        End If
    End If
Lab_Finir:
    Print #fdh, " </Body>"
    Print #fdh, "</Head>"
    Close #fdh
    
    'HTTP_Appel_PutFile
    ' à qui envoyer ?
    p_ouvrir_log = False
    'If Not p_traitement_background Then
        ListeDest = SYS_GetIni("MAILDEST", "MAILDEST", p_nomini)
        sql = "select APP_lstresp from Application where APP_Code='KALIBOTTIN'"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Exit Function
        End If
        If v_s <> "" Then
            sMessage = "Impossible d'effectuer l'intégration automatique par KaliBottin le " & Date & " à " & Time & " (" & titre & ")" & vbCrLf & vbCrLf
        ElseIf nbpasfait > 0 Then
            sMessage = "Rapport des intégrations effectuées par KaliBottin le " & Date & " à " & Time & " (" & titre & ")" & vbCrLf & vbCrLf
            sMessage = sMessage & nbfait & " mouvements ont été traités"
            sMessage = sMessage & " - " & nbpasfait & " mouvements n'ont pas été traités"
        Else
            sMessage = "Rapport des intégrations effectuées par KaliBottin le " & Date & " à " & Time & " (" & titre & ")" & vbCrLf & vbCrLf
            sMessage = sMessage & nbfait & " mouvements ont été traités"
        End If
        If Not rs.EOF Then
            lstresp = rs("APP_lstresp")
            For I = 0 To STR_GetNbchamp(lstresp, ";")
                sDest = STR_GetChamp(lstresp, ";", I)
                sDest = Replace(sDest, "U", "")
                If sDest <> "" Then
                    Call P_EnvoyerMessage(val(sDest), "", "Rapport d'intégration de KaliBottin", sMessage, g_nomfichHTML)
                End If
            Next I
        End If
        For iDest = 0 To STR_GetNbchamp(ListeDest, ";")
            s = STR_GetChamp(ListeDest, ";", iDest)
            If s <> "" Then
                Call P_EnvoyerMessage(0, s, "Rapport d'intégration de KaliBottin", sMessage, g_nomfichHTML)
            End If
        Next iDest
    'Else
    If Not p_traitement_background Then
        If nbfait + nbpasfait > 0 Then
            p_ouvrir_log = True
        End If
    End If
    ' Vider
    p_mess_fait_background = ""
    p_mess_pasfait_background = ""
    p_corps_background = ""
    nbfait = 0
    nbpasfait = 0

End Function

Private Function insertInfoSuppl(ByVal v_unum As Long, ByVal v_row As Integer) As Integer
' ajouter l'information supplémentaire pour cette personne
    Dim tis_value As String
    Dim lng As Long, tis_num As Long
    Dim I As Integer
    Dim MATRICULE As String
    Dim sql As String
    Dim lnb As Long
    
    With grd
        MATRICULE = .TextMatrix(v_row, GRDC_MATRICULE)
        For I = 0 To UBound(g_LISTE_U_INFO_SUPPL)
            If g_LISTE_U_INFO_SUPPL(I).umatricule = MATRICULE Then
                If g_LISTE_U_INFO_SUPPL(I).tis_pour_creer Then
                    If g_LISTE_U_INFO_SUPPL(I).tis_alimente > 0 Then
                        ' on crée dans zoneutil
                        Call AlimenteZoneUtil("M", g_LISTE_U_INFO_SUPPL(I).tis_alimente, v_unum, g_LISTE_U_INFO_SUPPL(I).infosuppl)
                    Else
                        tis_num = g_LISTE_U_INFO_SUPPL(I).tis_num
                        sql = " select count(*) from InfoSupplEntite WHERE ISE_Type='U' " _
                            & " AND ISE_TypeNum=" & v_unum _
                            & " AND ISE_TisNum=" & tis_num
                        If Odbc_Count(sql, lnb) = P_ERREUR Then
                            GoTo lab_erreur
                        End If
                        tis_value = g_LISTE_U_INFO_SUPPL(I).infosuppl
                        If lnb > 0 Then
                            ' MAJ si info existe
                            If Odbc_Update("InfoSupplEntite", "ISE_Num", _
                                           "WHERE ISE_Type='U' AND ISE_TypeNum=" & v_unum _
                                                & " AND ISE_TisNum=" & tis_num, _
                                            "ISE_Valeur", tis_value) = P_ERREUR Then
                                GoTo lab_erreur
                            End If
                        Else
                            ' l'info n'existe pas, on l'ajoute
                            If Odbc_AddNew("InfoSupplEntite", "ISE_Num", "ISE_Seq", False, lng, _
                                            "ISE_TisNum", tis_num, _
                                            "ISE_TypeNum", v_unum, _
                                            "ISE_Type", "U", _
                                            "ISE_Valeur", tis_value) = P_ERREUR Then
                                GoTo lab_erreur
                            End If
                        End If
                    End If
                End If
            End If
        Next I
        
        'For I = .Cols - p_nbr_lstInfoSuppl To .Cols - 1
        '    tis_num = STR_GetChamp(.TextMatrix(v_row, I), ";", 0)
        '    tis_value = STR_GetChamp(.TextMatrix(v_row, I), ";", 1)
        '    If Len(tis_value) > 0 Then
        '        If Odbc_AddNew("InfoSupplEntite", "ISE_Num", "ISE_Seq", False, lng, _
        '                        "ISE_TisNum", tis_num, _
        '                        "ISE_TypeNum", v_unum, _
        '                        "ISE_Type", "U", _
        '                        "ISe_Valeur", tis_value) = P_ERREUR Then
        '            GoTo lab_erreur
        '        End If
        '    End If
        'Next I
    End With

    insertInfoSuppl = P_OK
    Exit Function

lab_erreur:
    insertInfoSuppl = P_ERREUR

End Function

Private Sub maj_grid(ByVal v_u_num As Long)
' *********************************************************
' Metter à jour les lignes du grid suite à une modification
' dans PrmPersonne apres la recherche dans le dictionnaire.
' *********************************************************
End Sub

Private Sub maj_ligne(ByVal v_row As Integer, ByVal v_toucher_pastille As Boolean)
'***************************************************************
' Appelée après l'appel de PrmPersonne (cellule GRDC_RE_MODIFER)
'   ou après la création d'un nouvel utilisateur
' Mettre-à-jour la ligne depuis la table Utilisateur
'***************************************************************
    Dim sql As String, ancien_matricule As String, str As String
    Dim I As Integer, nbr_spm As Integer, nbr As Integer, j As Integer, p As Integer
    Dim rs As rdoResultset

    With grd
        ancien_matricule = .TextMatrix(v_row, GRDC_MATRICULE)
        sql = "SELECT * FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & .TextMatrix(v_row, GRDC_U_NUM)
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Exit Sub
        End If
        If Not rs.EOF Then
            ' doit-on modifier la pastille rouge/verte ?
            If v_toucher_pastille Then
                .col = GRDC_PASTILLE
                If rs("U_Actif").Value Then
                    If Not rs("U_ExterneFich").Value Then
                        Set .CellPicture = imglst.ListImages(IMG_PASTILLE_VERTE).Picture
                    End If
                Else ' rs("U_Actif").Value = False
                    If .TextMatrix(v_row, GRDC_ACTION) = "Désactiver" Then
                        GoTo lab_supprimer_ligne
                    ElseIf .TextMatrix(v_row, GRDC_ACTION) = "Accéder" Then
                        Call a_creer(.TextMatrix(v_row, GRDC_LIGNE_LUE), v_row)
                        Exit Sub
                    End If
                    'Set .CellPicture = imglst.ListImages(IMG_PASTILLE_ROUGE).Picture
                End If
            End If
            .TextMatrix(v_row, GRDC_MATRICULE) = rs("U_Matricule").Value
            .TextMatrix(v_row, GRDC_NOM) = rs("U_Nom").Value
            .TextMatrix(v_row, GRDC_PRENOM) = formater_prenom(rs("U_Prenom").Value)
            ' ====================
            ' str sous forme "num_srv:lib_srv <=> num_poste:lib_poste"
            str = remplir_srv_poste(rs("U_SPM").Value, "AUCUN")
            .TextMatrix(v_row, GRDC_CODE_SRV_FICH) = STR_GetChamp(STR_GetChamp(str, SEPARATEUR_SERVICE_POSTE, 0), ":", 0)
            .TextMatrix(v_row, GRDC_LIB_SRV_FICH) = STR_GetChamp(STR_GetChamp(str, SEPARATEUR_SERVICE_POSTE, 0), ":", 1)
            .TextMatrix(v_row, GRDC_CODE_POSTE_FICH) = STR_GetChamp(STR_GetChamp(str, SEPARATEUR_SERVICE_POSTE, 1), ":", 0)
            .TextMatrix(v_row, GRDC_LIB_POSTE_FICH) = STR_GetChamp(STR_GetChamp(str, SEPARATEUR_SERVICE_POSTE, 1), ":", 1)
            ' ====================
            .col = GRDC_ACTION
        Else
            GoTo lab_supprimer_ligne
        End If
        rs.Close
    End With

    Exit Sub

lab_supprimer_ligne:
    
    If Not p_traitement_background Then
        grd.RemoveItem (v_row)
        Call actualiser_compteur
    End If
End Sub

Private Sub modification(ByVal v_row As Integer)

' *****************************************************************************
' dans le grid on trouve le verbe "Corriger"
' Mettre-à-jour le ligne qui comporte des "anomalies" sur le NOM et/ou le POSTE
' *****************************************************************************
    Dim new_nom As String, old_nom As String, ma_civilite As String, new_prenom As String, old_prenom As String, _
        new_poste As String, old_poste As String, new_service As String, old_service As String, _
        sql As String, spm_choisi As String, old_spm As String, _
        mess As String, mess_avt As String, mess_apr As String, s As String, sret As String
    Dim code_poste As String, code_srv As String
    Dim afficher_srv_poste As Boolean
    Dim reponse As Integer, nbr_old_spm As Integer, n As Integer, I As Integer, lig As Integer
    Dim lng As Long, numposte As Long
    Dim frm As Form
    Dim mess_bck As String
    Dim libSrv As String
    Dim libNivSrv As String
    Dim rs As rdoResultset
    Dim ss As String
    
    With grd
        sql = "SELECT U_Prenom, U_SPM FROM Utilisateur" _
            & " WHERE U_kb_actif=True AND U_Matricule='" & .TextMatrix(v_row, GRDC_MATRICULE) & "'"
        If Odbc_RecupVal(sql, old_prenom, old_spm) = P_ERREUR Then
            Exit Sub
        End If
        If (p_pos_civilite <> -1) Then
            ma_civilite = "CIVILITE:" & vbTab & vbTab & .TextMatrix(v_row, GRDC_CIVILITE)
        Else
            ma_civilite = ""
        End If
        new_prenom = .TextMatrix(v_row, GRDC_PRENOM)
        old_prenom = formater_prenom(old_prenom)
        new_nom = .TextMatrix(v_row, GRDC_NOM)
        old_nom = .TextMatrix(v_row, GRDC_METTRE_A_JOUR_NOM)
        old_service = .TextMatrix(v_row, GRDC_LIB_SRV_FICH)
        old_poste = .TextMatrix(v_row, GRDC_LIB_POSTE_FICH)
        .Row = v_row
        .col = GRDC_CODE_SRV_FICH ' GRDC_LIB_SRV_FICH, GRDC_CODE_POSTE_FICH ou GRDC_LIB_POSTE_FICH
        If .CellFontBold Then
            new_service = P_get_lib_srv_poste(.TextMatrix(v_row, GRDC_NUM_SRV_KB), P_SERVICE)
            new_poste = P_get_lib_srv_poste(.TextMatrix(v_row, GRDC_NUM_POSTE), P_POSTE)
            ' 1° Demander si on associe un autre poste à cette personne:
            selectionner_ligne (v_row)
lab_remessage:
            libSrv = P_FctRecupNiveau(.TextMatrix(v_row, GRDC_NUM_SRV_KB))
            libSrv = IIf(libSrv = "", "SERVICE", libSrv)
            
            s = " Modifier le poste actuel de: " & vbCrLf _
                            & ma_civilite _
                            & " NOM:" & new_nom _
                            & " PRENOM:" & new_prenom & vbCrLf _
                            & " POSTE:" & P_get_lib_srv_poste(.TextMatrix(v_row, GRDC_NUM_POSTE), P_POSTE) & vbCrLf _
                            & " " & libSrv & ":" & P_get_lib_srv_poste(.TextMatrix(v_row, GRDC_NUM_SRV_KB), P_SERVICE) & vbCrLf _
                            & " (Fichier : " & .TextMatrix(v_row, GRDC_LIB_SRV_FICH) & " / " & .TextMatrix(v_row, GRDC_LIB_POSTE_FICH) & ")"
            If p_traitement_background Then
                mess_bck = Replace(s, vbCrLf, "")
                GoTo Lab_Modifier
            Else
                mess_bck = Replace(s, vbCrLf, "")
            End If
            reponse = MsgBox("Voulez-vous " & s, vbQuestion + vbYesNoCancel, "Attention")
            If reponse = vbCancel Then ' on ne fait rien
                cmd(CMD_CORRIGER_TOUS).Visible = False
                Exit Sub
            End If
            ' on propose une liste de postes à choisir
            If reponse = vbYes Then
Lab_Modifier:
                numposte = choisir_poste_synchro_struct(.TextMatrix(v_row, GRDC_CODE_SRV_FICH), mess_bck, .TextMatrix(v_row, GRDC_CODE_POSTE_FICH))
                If p_traitement_background Then
                    If Not p_background_synchro_auto Then
                        p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==>  " & mess_bck
                        Exit Sub
                    Else
                        'p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==>  " & mess_bck
                        afficher_srv_poste = True
                        GoTo Lab_Continuer
                    End If
                End If
                afficher_srv_poste = True
                If numposte = 0 Then ' on n'a rien selectionné
                    GoTo lab_remessage
                Else
Lab_Continuer:
                    spm_choisi = build_arbor_srv(numposte)
                    If spm_choisi = "" Then
                        GoTo lab_remessage
                    Else
                        spm_choisi = ControleFormat(spm_choisi)
                        g_poste = numposte
                        g_GRDC_CODE_SRV_FICH = .TextMatrix(v_row, GRDC_CODE_SRV_FICH)
                        g_GRDC_CODE_POSTE_FICH = .TextMatrix(v_row, GRDC_CODE_POSTE_FICH)
                        g_GRDC_LIB_SRV_FICH = .TextMatrix(v_row, GRDC_LIB_SRV_FICH)
                        g_GRDC_LIB_POSTE_FICH = .TextMatrix(v_row, GRDC_LIB_POSTE_FICH)
                        Me.cmd(CMD_CORRIGER_TOUS).Visible = True
                        'g_code_emploi_auto = new_poste
                        'g_code_service_auto = new_service
                        If cmd(CMD_CORRIGER_TOUS).Visible Then
                            cmd(CMD_CORRIGER_TOUS).Caption = "Corriger pour tous les" & Chr(13) & Chr(10)
                            cmd(CMD_CORRIGER_TOUS).Caption = cmd(CMD_CORRIGER_TOUS).Caption & g_GRDC_LIB_POSTE_FICH & " -> " & g_GRDC_LIB_SRV_FICH
                        End If
                    End If
                End If
            Else ' on a cliqué sur NON
                spm_choisi = .TextMatrix(v_row, GRDC_SPM_KB_A_SYNCHRONISER) & "|"
                ' Voir si existe pas déjà
                sql = "SELECT * FROM Synchro WHERE Sync_Section=" & Odbc_String(.TextMatrix(v_row, GRDC_CODE_SRV_FICH))
                sql = sql & " And Sync_Emploi = " & Odbc_String(.TextMatrix(v_row, GRDC_CODE_POSTE_FICH))
                sql = sql & " And Sync_SPNum = " & Odbc_String(.TextMatrix(v_row, GRDC_NUM_POSTE))
                If Odbc_SelectV(sql, rs) = P_ERREUR Then
                    Exit Sub
                ElseIf rs.RowCount > 1 Then
                    MsgBox "Il y a des doublons  (" & rs.RowCount & ") pour l'association Poste : " & .TextMatrix(v_row, GRDC_CODE_SRV_FICH) & " - UF : " & .TextMatrix(v_row, GRDC_CODE_POSTE_FICH)
                ElseIf rs.EOF Then
                    If Odbc_AddNew("Synchro", "Sync_Num", "Sync_Seq", False, lng, _
                                   "Sync_Section", .TextMatrix(v_row, GRDC_CODE_SRV_FICH), _
                                   "Sync_Emploi", .TextMatrix(v_row, GRDC_CODE_POSTE_FICH), _
                                   "Sync_SPNum", .TextMatrix(v_row, GRDC_NUM_POSTE), _
                                   "Sync_auto", False) = P_ERREUR Then
                        MsgBox "Erreur AddNew"
                        Exit Sub
                    End If
                End If
                lig = 1
                code_srv = .TextMatrix(v_row, GRDC_CODE_SRV_FICH)
                code_poste = .TextMatrix(v_row, GRDC_CODE_POSTE_FICH)
                numposte = .TextMatrix(v_row, GRDC_NUM_POSTE)
                While lig < .Rows - 1
                    If .TextMatrix(lig, GRDC_ACTION) = "Corriger" And _
                       .TextMatrix(lig, GRDC_CODE_SRV_FICH) = code_srv And _
                       .TextMatrix(lig, GRDC_CODE_POSTE_FICH) = code_poste And _
                       .TextMatrix(lig, GRDC_NUM_POSTE) = numposte & "" Then
                        If .Rows = 2 Then
                            .Rows = 1
                        Else
                            .RemoveItem (lig)
                        End If
                    Else
                        lig = lig + 1
                    End If
                Wend
                Exit Sub
            End If
            .col = GRDC_NOM
            If .CellFontBold Then ' NOM=> gras, SRV_POSTE_FICH=> gras ------------------------------
                ' on a déjà new_nom et new_prenom (cf. plus haut)
            Else ' NOM=> normal, SRV_POSTE_FICH=> gras ---------------------------------------------
                old_nom = .TextMatrix(v_row, GRDC_NOM)
                old_prenom = .TextMatrix(v_row, GRDC_PRENOM)
            End If
        Else ' NOM=> gras (puisqu'on est là !), SRV_POSTE_FICH=> normal ----------------------------
            old_service = new_service
            old_poste = new_poste
            spm_choisi = .TextMatrix(v_row, GRDC_SPM_KB_A_SYNCHRONISER) & "|"
            s = STR_GetChamp(.TextMatrix(v_row, GRDC_SPM_KB_A_SYNCHRONISER), ";P", 1)
            g_poste = Replace(s, ";", "")
        End If
        If .TextMatrix(v_row, GRDC_ACTION) = "Postes Sec." Then ' on ne traite que les postes secondaires
            GoTo LabModifierPostesSecondaire
        End If
        ' message de confirmation
        mess = "Modification de " & .TextMatrix(v_row, GRDC_NOM) _
                & " " & .TextMatrix(v_row, GRDC_PRENOM) _
                & " (MATRICULE " & .TextMatrix(v_row, GRDC_MATRICULE) & ")" & vbCrLf & vbCrLf
        mess_avt = ""
        mess_apr = ""
        If old_nom <> new_nom Then
            mess_avt = " * NOM:" & vbTab & vbTab & new_nom & vbCrLf
            mess_apr = " * NOM:" & vbTab & vbTab & old_nom & vbCrLf
        End If
        If old_prenom <> new_prenom Then
            mess_avt = mess_avt & " * PRENOM:" & vbTab & new_prenom & vbCrLf
            mess_apr = mess_apr & " * PRENOM:" & vbTab & old_prenom & vbCrLf
        End If
        If afficher_srv_poste Then ' on a cliqué sur le bouton OUI
            libNivSrv = P_FctRecupNiveau(P_get_num_srv_poste(spm_choisi, P_SERVICE))
            mess_avt = mess_avt & " * " & IIf(libNivSrv = "", "SERVICE", libNivSrv) & ":" & vbTab & P_get_lib_srv_poste(P_get_num_srv_poste(spm_choisi, P_SERVICE), P_SERVICE) & vbCrLf
            mess_avt = mess_avt & " * POSTE:" & vbTab & P_get_lib_srv_poste(P_get_num_srv_poste(spm_choisi, P_POSTE), P_POSTE) & vbCrLf
            libNivSrv = P_FctRecupNiveau(.TextMatrix(v_row, GRDC_NUM_SRV_KB))
            mess_apr = mess_apr & " * " & IIf(libNivSrv = "", "SERVICE", libNivSrv) & ":" & vbTab & new_service & vbCrLf
            mess_apr = mess_apr & " * POSTE:" & vbTab & new_poste & vbCrLf
        End If
        mess = mess & "- Informations se trouvant actuellement dans le dictionnaire -" & vbCrLf _
                    & mess_apr & vbCrLf & vbCrLf _
                    & "- Nouvelles informations -" & vbCrLf & mess_avt & vbCrLf & vbCrLf _
                    & "Confirmez-vous cette opération ?" & vbCrLf & vbCrLf
        selectionner_ligne (v_row)
        If mess_avt <> "" And mess_apr <> "" Then
            If p_traitement_background Then
                mess_bck = Replace(mess_avt & " <font color=red><B> Vers </B></font> " & mess_apr, vbCrLf, "")
            Else
                mess_bck = Replace(mess_avt & " <font color=red><B> Vers </B></font> " & mess_apr, vbCrLf, "")
                If MsgBox(mess, vbQuestion + vbYesNo, "") <> vbYes Then
                    Exit Sub
                End If
            End If
        End If
        
        If Not p_traitement_background Then
            ' Mettre à jour la table de synchro
            ' Voir si existe pas déjà
            sql = "SELECT * FROM Synchro WHERE Sync_Section=" & Odbc_String(.TextMatrix(v_row, GRDC_CODE_SRV_FICH))
            sql = sql & " And Sync_Emploi = " & Odbc_String(.TextMatrix(v_row, GRDC_CODE_POSTE_FICH))
            sql = sql & " And Sync_SPNum = " & Odbc_String(numposte)
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==>  " & mess_bck
                Exit Sub
            ElseIf rs.RowCount > 1 Then
                MsgBox "Il y a des doublons  (" & rs.RowCount & ") pour l'association Poste : " & .TextMatrix(v_row, GRDC_CODE_SRV_FICH) & " - UF : " & .TextMatrix(v_row, GRDC_CODE_POSTE_FICH)
                p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==>  " & mess_bck
                p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==>  " & "Il y a des doublons  (" & rs.RowCount & ") pour l'association Poste : " & .TextMatrix(v_row, GRDC_CODE_SRV_FICH) & " - UF : " & .TextMatrix(v_row, GRDC_CODE_POSTE_FICH)
            ElseIf rs.EOF Then
                If Odbc_AddNew("Synchro", "Sync_Num", "Sync_Seq", False, lng, _
                               "Sync_Section", .TextMatrix(v_row, GRDC_CODE_SRV_FICH), _
                               "Sync_Emploi", .TextMatrix(v_row, GRDC_CODE_POSTE_FICH), _
                               "Sync_SPNum", numposte, _
                               "Sync_auto", False) = P_ERREUR Then
                    MsgBox "Erreur AddNew"
                    p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==>  " & mess_bck
                    Exit Sub
                End If
            End If
        End If
LabModifierPostesSecondaire:
        bFaireRemove = True
        Call modifier_cette_personne(v_row, spm_choisi, mess_bck, True)
' A NE PAS SUPPRIMER
        ' Call desactiver_ligne(v_row, IMG_PASTILLE_VERTE)
        On Error Resume Next
        If Not bFaireRemove Then
            Exit Sub
        End If
        If Not p_traitement_background Then
            .RemoveItem (v_row)
            Call actualiser_compteur
        End If
    End With

End Sub

Private Sub modifier_cette_personne(ByVal v_row As Integer, _
                                    ByVal v_spm As String, _
                                    ByRef v_mess_bck As String, _
                                    ByVal v_lancer_kd As Boolean)
'v_lancer_kd = False
' **********************************************************************************
' Appelée uniquement depuis modification() afin de corriger les infos de la personne
' **********************************************************************************
    Dim fcttrav As String, col_nom As String, col_prenom As String
    Dim old_fct As String, old_spm As String, scmd As String, sfct As String
    Dim new_spm As String, new_fct As String, s As String
    Dim n As Integer, I As Integer
    Dim lng As Long, old_poste As Long
    Dim numposte As Long
    Dim sql As String, rs As rdoResultset
    Dim po_actif As Boolean, lib_poste As String
    Dim ok As Boolean
    
    With grd
        ' enregistrer les anciennes infos avant MAJ Utilisateur pour renseigner UtilMouvement
        If remplir_utilmouvement(v_row, v_spm) = P_ERREUR Then
            GoTo LabFin
        End If
        If .TextMatrix(v_row, GRDC_NOM_JUNON) <> "" Then
            col_nom = "U_NomJunon"
        Else
            col_nom = "U_Nom"
        End If
        If .TextMatrix(v_row, GRDC_PRENOM_JUNON) <> "" Then
            col_prenom = "U_PrenomJunon"
        Else
            col_prenom = "U_Prenom"
        End If
        If .TextMatrix(v_row, GRDC_ACTION) = "Postes Sec." Then
            GoTo LabGérerPostesSecond
        End If
        
        If v_spm = .TextMatrix(v_row, GRDC_SPM_KB_A_SYNCHRONISER) & "|" Then
            If Odbc_Update("Utilisateur", "U_Num", _
                               "WHERE U_Num=" & .TextMatrix(v_row, GRDC_U_NUM), _
                               col_nom, .TextMatrix(v_row, GRDC_NOM), _
                               col_prenom, .TextMatrix(v_row, GRDC_PRENOM), _
                               "U_Matricule", .TextMatrix(v_row, GRDC_MATRICULE), _
                               "U_Importe", True, _
                               "U_Actif", True, _
                               "U_kb_Actif", True) = P_ERREUR Then
                If True Or p_traitement_background Then
                    ok = False
                    v_mess_bck = v_mess_bck & Chr(13) & Chr(10) & "==> Erreur SQL " & sql
                    GoTo LabFin
                End If
            Else
                ' si le poste n'est pas actif => le ré-activer
                numposte = g_poste  ' P_get_num_srv_poste(new_spm, P_POSTE)
                sql = "SELECT po_actif, po_libelle, po_srvnum FROM poste Where po_Num=" & numposte
                If Odbc_RecupVal(sql, po_actif, lib_poste) = P_ERREUR Then
                    MsgBox "Erreur SQL fonction creation SQL=" & sql
                    ok = False
                    v_mess_bck = v_mess_bck & Chr(13) & Chr(10) & "==> Erreur SQL " & sql
                    GoTo LabFin
                Else
                    If Not po_actif Then
                        Call Odbc_Update("Poste", "PO_Num", _
                         "WHERE PO_num=" & numposte, _
                         "PO_Actif", True)
                         s = "Le poste (" & numposte & ") " & lib_poste & " (inactif) a été ré-activé"
                         v_mess_bck = v_mess_bck & Chr(13) & Chr(10) & "==> " & s
                         If p_traitement_background Then
                            Call MAJ_Corps_Background("", "", s, p_corps_background)
                         Else
                            Call MAJ_Corps_Background("", "", s, p_corps_background)
                            MsgBox s
                         End If
                    End If
                End If
                ok = True
                GoTo LabFin
            End If
        Else
            sql = "select u_po_princ, u_fcttrav, u_spm from utilisateur where u_num=" & .TextMatrix(v_row, GRDC_U_NUM)
            If Odbc_RecupVal(sql, _
                                old_poste, old_fct, old_spm) = P_ERREUR Then
                v_mess_bck = v_mess_bck & Chr(13) & Chr(10) & "==> Erreur SQL " & sql
                ok = False
                GoTo LabFin
            End If
            v_spm = ControleFormat(v_spm)
            fcttrav = P_get_fcttrav(v_spm)
            old_spm = ControleFormat(old_spm)
            If STR_GetNbchamp(old_spm, "|") = 1 Then
                new_spm = v_spm
                new_fct = fcttrav
            Else
                new_spm = ""
                new_fct = ""
                n = STR_GetNbchamp(old_spm, "|")
                For I = 0 To n - 1
                    s = STR_GetChamp(old_spm, "|", I)
                    If STR_GetChamp(s, ";", STR_GetNbchamp(s, ";") - 1) <> "P" & old_poste Then
                        new_spm = new_spm & s & "|"
                        sfct = P_get_fcttrav(s)
                        If InStr(new_fct, sfct) = 0 Then
                            new_fct = new_fct + sfct
                        End If
                    End If
                Next I
                new_spm = v_spm & new_spm
                new_spm = Controle_doublon_dans_spm(new_spm)
                If InStr(new_fct, fcttrav) = 0 Then
                    new_fct = new_fct + fcttrav
                End If
            End If
            If Odbc_Update("Utilisateur", "U_Num", _
                               "WHERE U_Num=" & .TextMatrix(v_row, GRDC_U_NUM), _
                               col_nom, .TextMatrix(v_row, GRDC_NOM), _
                               col_prenom, .TextMatrix(v_row, GRDC_PRENOM), _
                               "U_Matricule", .TextMatrix(v_row, GRDC_MATRICULE), _
                               "U_Importe", True, _
                               "U_Actif", True, _
                               "U_KB_Actif", True, _
                               "U_SPM", new_spm, _
                               "U_FctTrav", new_fct, _
                               "U_Po_Princ", P_get_num_srv_poste(new_spm, P_POSTE)) = P_ERREUR Then
                If True Or p_traitement_background Then
                    v_mess_bck = v_mess_bck & " pour " & .TextMatrix(v_row, GRDC_MATRICULE) & " Nom : " & .TextMatrix(v_row, GRDC_NOM) & " " & .TextMatrix(v_row, GRDC_PRENOM)
                    ok = False
                    GoTo LabFin
                    'p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> " & v_mess_bck
                End If
                'Exit Sub
            Else
                If True Or p_traitement_background Then
                    v_mess_bck = v_mess_bck & " pour " & .TextMatrix(v_row, GRDC_MATRICULE) & " Nom : " & .TextMatrix(v_row, GRDC_NOM) & " " & .TextMatrix(v_row, GRDC_PRENOM)
                    'p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> " & v_mess_bck
                End If
            End If
            ' si le poste n'est pas actif => le ré-activer
            numposte = P_get_num_srv_poste(new_spm, P_POSTE)
            sql = "SELECT po_actif, po_libelle, po_srvnum FROM poste Where po_Num=" & numposte
            If Odbc_RecupVal(sql, po_actif, lib_poste) = P_ERREUR Then
                MsgBox "Erreur SQL fonction creation SQL=" & sql
                v_mess_bck = v_mess_bck & Chr(13) & Chr(10) & "==> Erreur SQL fonction creation SQL=" & sql
                ok = False
                GoTo LabFin
            Else
                If Not po_actif Then
                    Call Odbc_Update("Poste", "PO_Num", _
                     "WHERE PO_num=" & numposte, _
                     "PO_Actif", True)
                     s = "Le poste " & lib_poste & " (inactif) a été ré-activé"
                     v_mess_bck = v_mess_bck & Chr(13) & Chr(10) & "==> " & s
                     If Not p_traitement_background Then
                     '   Call MAJ_Corps_Background("", "", s, p_corps_background)
                     'Else
                     '   Call MAJ_Corps_Background("", "", s, p_corps_background)
                        MsgBox s
                     End If
                End If
            End If
            If v_lancer_kd Then
                ' Lance KD pour gérer le changement de poste
                scmd = p_chemin_appli & "\Lance.exe " _
                     & p_chemin_appli & ";KaliDoc;" & p_nom_fichier_ini_kalidoc & ";CONNEXION=" & p_NumUtil & ";CHGTPOSTE_PERS=" _
                     & .TextMatrix(v_row, GRDC_U_NUM) & "|" _
                     & Replace(old_fct, ";", "-") & "|" _
                     & Replace(old_spm, ";", "-") & "[WAIT];[KBAUTO]"
                Call SYS_ExecShell(scmd, True, True)
            End If
        End If
LabGérerPostesSecond:
        ' Gérer les postes secondaires
        s = STR_GetChamp(v_spm, ";P", 1)
        s = Replace(s, ";", "")
        s = Replace(s, "|", "")
        Call GererPosteSecondaire("M", .TextMatrix(v_row, GRDC_PRENOM) & " " & .TextMatrix(v_row, GRDC_NOM), 0, CLng(s), .TextMatrix(v_row, GRDC_U_NUM), .TextMatrix(v_row, GRDC_MATRICULE), .TextMatrix(v_row, GRDC_NEW_PSTSECOND), .TextMatrix(v_row, GRDC_ANC_PSTSECOND))
        ' ***********************************
        If Not bFaireRemove Then
            Exit Sub
        End If
        cmd(CMD_ACTUALISER).Enabled = True
        Call remplir_srv_poste(v_spm, "AUCUN")
        Call maj_ligne(v_row, True)
        Call selectionner_ligne(v_row)
        .TextMatrix(v_row, GRDC_METTRE_A_JOUR_NOM) = ""
        .TextMatrix(v_row, GRDC_SPM_KB_A_SYNCHRONISER) = v_spm
        ok = True
    End With
LabFin:
    If ok Then
        p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> " & v_mess_bck
    Else
        p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> " & v_mess_bck
    End If
End Sub

Private Function Controle_doublon_dans_spm(ByVal v_spm)
    Dim n As Integer, I As Integer, j As Integer
    Dim s1 As String, s2 As String
    Dim sOut As String
    Dim bDéjà As Boolean
    
    n = STR_GetNbchamp(v_spm, "|")
    For I = 0 To n - 1
        bDéjà = False
        s1 = STR_GetChamp(v_spm, "|", I)
        For j = 0 To STR_GetNbchamp(sOut, "|") - 1
            s2 = STR_GetChamp(sOut, "|", j)
            If s1 = s2 Then
                bDéjà = True
            End If
        Next j
        If Not bDéjà Then
            sOut = sOut & s1 & "|"
        End If
    Next I
    Controle_doublon_dans_spm = sOut
End Function

Private Function personne_inactive_existe(ByVal v_row As Integer) As Boolean
'*************************************************************************
' Determine si la personne en cous d'ajout dans le grid ET à créer n'a pas
' le même NOM et g_nbr_car_prenom d'une personne dans le dictionnaire
' Remplir la colonne GRDC_LISTE_PERSONNE_INACTIVE = "I:789;IE:610;E:55;"
'       où "I" pour U_Inactif=False et "E" pour U_ExterneFich=True
'*************************************************************************
    Dim sql As String
    Dim rs As rdoResultset

    personne_inactive_existe = False
    With grd
        ' Recherche des personnes inactives et/ou externes au fichiers ayant le même nom que celui du fichier
        sql = "SELECT U_Num, U_Prenom, U_Actif, U_ExterneFich FROM Utilisateur" _
            & " where (U_Nom=" & Odbc_String(.TextMatrix(v_row, GRDC_NOM)) & " OR u_matricule = " & Odbc_String(.TextMatrix(v_row, GRDC_MATRICULE)) & ")" _
            & " AND (U_Actif=False OR U_ExterneFich=True)"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Exit Function
        End If
        ' Parcourir la liste de ces personnes
        While Not rs.EOF
            ' Recherche si les g_nbr_car_prenom premiers caractères du prenom se resemblent
            If UCase$(Mid$(rs("U_Prenom").Value, 1, g_nbr_car_prenom)) = UCase$(Mid$(.TextMatrix(v_row, GRDC_PRENOM), 1, g_nbr_car_prenom)) Then
                ' Construction de la liste avec les U_Num séparés par des ";"
                .TextMatrix(v_row, GRDC_LISTE_PERSONNE_INACTIVE) = .TextMatrix(v_row, GRDC_PERSONNE_INACTIVE) _
                                                                 & IIf(Not rs("U_Actif").Value, "I", "") _
                                                                 & IIf(rs("U_ExterneFich").Value, "E", "") & ":" _
                                                                 & rs("U_Num").Value & ";"
                personne_inactive_existe = True
            End If
            rs.MoveNext
        Wend
        rs.Close
    End With

End Function

Private Sub quitter(ByVal v_s As String)

    If p_traitement_background Then
        If v_s <> "" Then
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "==> " & v_s
        End If
    Else
        ' enregistrer le fichier de log
        Call Traitement_Background("")
        Unload Me
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
    sql = "SELECT * FROM Utilisateur WHERE U_kb_actif=True"
    If text_nom <> "" Then
        If text_matricule <> "" Then ' on a un NOM et un MATRICULE
            sql = sql & " AND U_Nom LIKE " & Odbc_upper() & "(" & Odbc_String(text_nom) & ")" _
                      & " AND U_Matricule LIKE '%" & text_matricule & "%'ORDER BY U_Nom"
        Else ' que le NOM
            sql = sql & " AND U_Nom LIKE " & Odbc_upper() & "(" & Odbc_String(text_nom & "%") & ") ORDER BY U_Nom"
        End If
    Else ' que le MATRICULE
        sql = sql & " AND U_Matricule LIKE " & Odbc_String(text_matricule & "%") & " ORDER BY U_Nom"
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
        txt(TXT_NOM).Text = ""
        txt(TXT_MATRICULE).Text = ""
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
        txt(TXT_NOM).Text = ""
        txt(TXT_MATRICULE).Text = ""
        Exit Sub
    Else                        ' OK
        Set frm = PrmPersonne
        If Not PrmPersonne.AppelFrm(CL_liste.lignes(CL_liste.pointeur).num, "") Then
        ' il n'y a pas eu des changements importants
            Set frm = Nothing
            GoTo lab_afficher
        Else ' il y a eu des changements importants
            ' forcer le passage par la verification
            ' ****************************************************************************************************
            Set frm = Nothing
            g_importation_ok = P_ERREUR
            Call quitter("")
            ' ****************************************************************************************************
' A NE PAS SUPPRIMER
            'Call maj_grid(CL_liste.lignes(CL_liste.pointeur).num)
            'reafficher = True
            'If nbr_affiche > 1 Then GoTo lab_afficher
        End If
    End If

End Sub

Private Function remplir_srv_poste(ByVal v_spm As String, ByVal v_mode As String) As String
' **************************************************************************************************
' 1° La chaine v_mode determine les opération attendues par cette fonction:
'  a_detruire => "DETRUIRE" ou a_corriger => "MAJ" , afficher tooltip => "TOOLTIP", sinon => "AUCUN"
' 2° Formater le SPM en: "NUM_SRV" : "LIB_SRV" SEPARATEUR "NUM_POSTE" : "LIB_POSTE"
' **************************************************************************************************
    Dim mon_service As String, mon_poste As String, spm_a_afficher As String
    Dim nbr_spm As Integer, I As Integer
    Dim mon_num_service As Long, mon_num_poste As Long
    Dim rs As rdoResultset

    With grd
        ' Le nombre de postes
        nbr_spm = STR_GetNbchamp(v_spm, "|")
        If nbr_spm = 1 Then ' ***************** UN SEUL POSTE ******************
            spm_a_afficher = STR_GetChamp(v_spm, "|", 0)
            mon_num_service = P_get_num_srv_poste(spm_a_afficher, P_SERVICE)
            If v_mode = "DETRUIRE" Or v_mode = "MAJ" Then
                .TextMatrix(.Rows - 1, GRDC_NUM_SRV_KB) = mon_num_service
            End If
            mon_num_poste = P_get_num_srv_poste(spm_a_afficher, P_POSTE)
            If v_mode = "DETRUIRE" Or v_mode = "MAJ" Then
                .TextMatrix(.Rows - 1, GRDC_NUM_POSTE) = mon_num_poste
            End If
            mon_service = P_get_lib_srv_poste(mon_num_service, P_SERVICE)
            mon_poste = P_get_lib_srv_poste(mon_num_poste, P_POSTE)
            If v_mode = "TOOLTIP" Then
                remplir_srv_poste = mon_poste & " - " & mon_service
            Else
                remplir_srv_poste = mon_num_service & ":" & mon_service & SEPARATEUR_SERVICE_POSTE _
                                  & mon_num_poste & ":" & mon_poste
            End If
            If v_mode = "DETRUIRE" Then
                .TextMatrix(.Rows - 1, GRDC_CODE_SRV_FICH) = mon_num_service
                .TextMatrix(.Rows - 1, GRDC_LIB_SRV_FICH) = mon_service
                .TextMatrix(.Rows - 1, GRDC_CODE_POSTE_FICH) = mon_num_poste
                .TextMatrix(.Rows - 1, GRDC_LIB_POSTE_FICH) = mon_poste
            End If
        Else ' ******************************** PLUSIEURS POSTES ******************
            For I = 0 To nbr_spm - 1
                spm_a_afficher = STR_GetChamp(v_spm, "|", I)
                mon_num_poste = P_get_num_srv_poste(spm_a_afficher, P_POSTE)
                If v_mode = "DETRUIRE" Or v_mode = "MAJ" Then
                    .TextMatrix(.Rows - 1, GRDC_NUM_POSTE) = mon_num_poste
                End If
                If Odbc_SelectV("SELECT SYNC_SPNum FROM Synchro", rs) = P_ERREUR Then
                    remplir_srv_poste = ""
                    Exit Function
                End If
                While Not rs.EOF
                    If mon_num_poste = rs("SYNC_SPNum").Value Then
                        rs.Close
                        mon_num_service = P_get_num_srv_poste(spm_a_afficher, P_SERVICE)
                        If v_mode = "DETRUIRE" Or v_mode = "MAJ" Then
                            .TextMatrix(.Rows - 1, GRDC_NUM_SRV_KB) = mon_num_service
                        End If
                        mon_service = P_get_lib_srv_poste(mon_num_service, P_SERVICE)
                        mon_poste = P_get_lib_srv_poste(mon_num_poste, P_POSTE)
                        If v_mode = "TOOLTIP" Then
                            remplir_srv_poste = mon_poste & " - " & mon_service
                        Else
                            remplir_srv_poste = mon_num_service & ":" & mon_service & SEPARATEUR_SERVICE_POSTE _
                                              & mon_num_poste & ":" & mon_poste
                        End If
                        If v_mode = "DETRUIRE" Then
                            .TextMatrix(.Rows - 1, GRDC_CODE_SRV_FICH) = mon_num_service
                            .TextMatrix(.Rows - 1, GRDC_LIB_SRV_FICH) = mon_service
                            .TextMatrix(.Rows - 1, GRDC_CODE_POSTE_FICH) = mon_num_poste
                            .TextMatrix(.Rows - 1, GRDC_LIB_POSTE_FICH) = mon_poste
                        End If
                        Exit Function
                    End If
                    rs.MoveNext
                Wend
                rs.Close
            Next I

            ' On n'a pas trouvé dans la table Synchro, on affiche le dernier poste trouvé
            mon_num_service = P_get_num_srv_poste(spm_a_afficher, P_SERVICE)
            If v_mode = "DETRUIRE" Or v_mode = "MAJ" Then
                .TextMatrix(.Rows - 1, GRDC_NUM_SRV_KB) = mon_num_service
            End If
            mon_service = P_get_lib_srv_poste(mon_num_service, P_SERVICE)
            mon_poste = P_get_lib_srv_poste(mon_num_poste, P_POSTE)
            If v_mode = "TOOLTIP" Then
                remplir_srv_poste = mon_poste & " - " & mon_service
            Else
                remplir_srv_poste = mon_num_service & ":" & mon_service & SEPARATEUR_SERVICE_POSTE _
                                  & mon_num_poste & ":" & mon_poste
            End If
            If v_mode = "DETRUIRE" Then
                .TextMatrix(.Rows - 1, GRDC_CODE_SRV_FICH) = mon_num_service
                .TextMatrix(.Rows - 1, GRDC_LIB_SRV_FICH) = mon_service
                .TextMatrix(.Rows - 1, GRDC_CODE_POSTE_FICH) = mon_num_poste
                .TextMatrix(.Rows - 1, GRDC_LIB_POSTE_FICH) = mon_poste
            End If
        End If
    End With

End Function

Private Function VoirSiDansSynchroDoublon(v_ligne_en_cours, ByRef r_spm_doublon)
    Dim section As String
    Dim emploi As String
    Dim sql As String, rs As rdoResultset
    Dim MATRICULE As String
    Dim s As String
    Dim s_poprinc As String
    
    section = P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_code_section, p_long_code_section, "code section")
    section = Trim$(section)
    emploi = P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_code_emploi, p_long_code_emploi, "code emploi")
    emploi = Trim$(emploi)
    MATRICULE = P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_matricule, p_long_matricule, "matricule")
    s = P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_nom, p_long_nom, "nom") & " " & P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_prenom, p_long_prenom, "prénom")
    s = s & " (matricule=" & MATRICULE & ")"
    ' Parcourir la table Synchro pour la paire (SECTION, EMPLOI)
    sql = "SELECT Sync_Num, Sync_Auto, Sync_emploi, Sync_Section, Sync_Spnum, Po_Num, Srv_Num, Srv_Nom, Ft_Num, Ft_Code, Ft_Libelle"
    sql = sql & " FROM Synchro, Poste, FctTrav, Service "
    sql = sql & " Where Poste.po_num = Synchro.sync_spnum"
    sql = sql & " And Poste.po_srvnum = Service.srv_num"
    sql = sql & " And Poste.po_ftnum = Fcttrav.ft_num"
    sql = sql & " And Synchro.Sync_emploi='" & emploi & "'"
    sql = sql & " And Synchro.Sync_Section='" & section & "'"
    'MsgBox sql
    r_spm_doublon = ""
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        VoirSiDansSynchroDoublon = False
        p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> Pas de synchro pour code section=" & section & " code emploi=" & emploi & " " & s
        Exit Function
    End If
    If rs.EOF Then
        VoirSiDansSynchroDoublon = False
        p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "==> Pas de synchro pour code section=" & section & " code emploi=" & emploi & " " & s
        Exit Function
    Else
        rs.MoveLast
        If rs.RowCount > 1 Then
            Call Odbc_RecupVal("select u_po_princ from utilisateur where u_matricule='" & MATRICULE & "'", s_poprinc)
            rs.MoveFirst
            While Not rs.EOF
                If rs("Sync_Spnum") = s_poprinc Then
                    ' construire_spm
                    r_spm_doublon = build_arbor_srv(rs("Sync_Spnum"))
                End If
                rs.MoveNext
            Wend
        Else
            ' construire_spm
            r_spm_doublon = build_arbor_srv(rs("Sync_Spnum"))
        End If
    End If
    VoirSiDansSynchroDoublon = True
End Function

Private Sub remplir_grid()

' ********************************************************************************************************
' A CREER:    les personnes de la base qu'on ne trouve pas dans le fichier (en se basant sur le MATRICULE)
' A DETRUIRE: les personnes du fichier qui ne se trouvent pas dans la base
' A MODIFIER: ...
' ********************************************************************************************************
    Dim sql As String, table_temporaire As String, matricule_en_cours As String, _
        spm_trouve As String, nom_a_comparer As String, pgb_lstposinfoautre As String
    Dim sinfo As String, mesPasTraité As String, s As String, sext As String
    Dim nomfich As String, sutil As String, nomutil As String
    Dim I As Integer, fd As Integer, pos As Integer, iBoucle As Integer
    Dim Anc_Po_Princ As Long
    Dim lib1 As String, lib2 As String
    Dim response As Integer
    Dim num_unique As Long, numutil As Long
    Dim rs As rdoResultset, rs2 As rdoResultset
    Dim ligne_lu As Variant
    Dim matricule_ancien As String
    Dim Anc_ligne_lu As String
    Dim prem As Boolean
    Dim table_temporaire_doublons As String
    Dim spm_doublon As String
    Dim prem_ligne_matricule As Integer
    Dim str_new_poste_secondaire As String
    Dim str_anc_poste_secondaire As String
    Dim laS As String
    Dim déjàTraité As Boolean
    Dim str_new_assoc As String
    Dim sN_PS As String, sA_PS As String
    Dim new_spm_princ As String
    Dim sret As Boolean
    Dim new_nom As String
    
    table_temporaire = "utilisateur_dans_fichier_" & p_NumUtil
    
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
    pgb2.Max = 1
    pgb2.Value = 0
    While Not EOF(fd)
        Line Input #fd, ligne_lu
        pgb2.Max = pgb2.Max + 1
        Me.LbGauge2.Caption = "Remplir le tableau des personnes"
        Me.Refresh
    Wend
    Close #fd
    
    ' Ouverture du fichier en lecture seule
    If FICH_OuvrirFichier(nomfich, FICH_LECTURE, fd) = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If

    ' Remplir le GRID
    ' Créer la table pour historique des postes secondaires
    If Not Odbc_TableExiste("kb_poste_secondaire") Then
        'Call Odbc_Cnx.Execute("drop TABLE kb_poste_secondaire")
        Call Odbc_Cnx.Execute("CREATE TABLE kb_poste_secondaire (psu_unum INTEGER, psu_matricule VARCHAR(20), psu_poste VARCHAR(40), psu_faire_maj BOOLEAN)")
    End If
    
    ' Y a t il d'autres utilisateurs en cours d'import
    If numutil = 0 Then
        numutil = P_SUPER_UTIL
    End If
    If Not p_traitement_background Or p_traitement_background_semiauto Then
        sql = "SELECT tablename FROM PG_Tables WHERE TableName like 'utilisateur_dans_fichier_%'"
        If Odbc_SelectV(sql, rs) = P_OK Then
            If Not rs.EOF Then
                sutil = ""
                While Not rs.EOF
                    pos = InStrRev(rs("tablename").Value, "_")
                    numutil = Mid$(rs("tablename").Value, pos + 1)
                    Call P_RecupUtilPpointNom(numutil, nomutil)
                    If sutil <> "" Then
                        sutil = sutil + ", "
                    End If
                    sutil = sutil + nomutil
                    rs.MoveNext
                Wend
                response = MsgBox("ATTENTION : Quelqu'un est déjà en cours d'importation ce qui peut générer des conflits." & vbCrLf & "   " & sutil & vbCrLf & "Souhaitez-vous quand même lancer l'importation ?", vbQuestion + vbYesNo, "")
                If response = vbNo Then
                    Call quitter("")
                    Exit Sub
                End If
            End If
            rs.Close
        End If
    End If
    ' créer une table temporaire
    If Odbc_TableExiste(table_temporaire) Then
        Call Odbc_Cnx.Execute("DELETE FROM " & table_temporaire)
    Else
        Call Odbc_Cnx.Execute("CREATE TABLE " & table_temporaire _
                            & " (tt_num INTEGER PRIMARY KEY, tt_matricule VARCHAR(25))")
    End If
    
    ' créer une table temporaire pour les doublons (postes secondaires)
    table_temporaire_doublons = "temp_utilisateur_postes_secondaires_" & p_NumUtil
    If Odbc_TableExiste(table_temporaire_doublons) Then
        Call Odbc_Cnx.Execute("DELETE FROM " & table_temporaire_doublons)
    Else
        Call Odbc_Cnx.Execute("CREATE TABLE " & table_temporaire_doublons _
                            & " (ttd_matricule VARCHAR(25), ttd_spm VARCHAR(150))")
    End If
    
    num_unique = 1
    With grd
        ' *************************** Remplir les lignes à CREER ***************************
        matricule_ancien = ""
        prem_ligne_matricule = 0
        While Not EOF(fd)
            If pgb.Value = pgb.Max Then
                pgb.Value = 0
            End If
            pgb.Value = pgb.Value + 1
            pgb2.Value = pgb2.Value + 1
            Anc_ligne_lu = ligne_lu
            Line Input #fd, ligne_lu
            If ligne_lu <> "" Then
                p_mess_fait_background = ""
                p_mess_pasfait_background = ""
                matricule_en_cours = P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_matricule, p_long_matricule, "matricule")
                'Debug.Print matricule_en_cours
                'If matricule_en_cours = "M373067" Or matricule_en_cours = "M353650" Then
                '    I = I
                'End If
                If matricule_en_cours <> matricule_ancien And matricule_ancien <> "" Then
                    ' Voir si modification poste secondaire
                    s = verifier_coherence_assoc(str_new_poste_secondaire, str_new_assoc, matricule_ancien)
                    If Not déjàTraité And RecupAncPstSecond(Anc_ligne_lu, str_anc_poste_secondaire, matricule_ancien, rs) Then
                        If verifier_spm(Anc_ligne_lu, rs("U_SPM").Value, rs("U_Po_Princ").Value, spm_trouve) = False Then
                        End If
                        str_new_poste_secondaire = controlePS(str_new_poste_secondaire)
                        'If str_new_poste_secondaire <> str_anc_poste_secondaire Then
                        '    Debug.Print str_new_poste_secondaire
                        '    Debug.Print str_anc_poste_secondaire
                        '    i = i
                        'End If
                        ' spm_trouve est l'ancien poste principal
                        sN_PS = ""
                        For I = 1 To STR_GetNbchamp(s, "|") - 1
                            sN_PS = sN_PS & STR_GetChamp(s, "|", I) & "|"
                        Next I
                        sA_PS = ""
                        For I = 1 To STR_GetNbchamp(str_anc_poste_secondaire, "|") - 1
                            sA_PS = sA_PS & STR_GetChamp(str_anc_poste_secondaire, "|", I) & "|"
                        Next I
                        If Not Egaux_postes_secondaires(sN_PS, sA_PS) Then
                            Call a_corriger(Anc_ligne_lu, rs, "POSTE_SECONDAIRE", spm_trouve & "!" & sN_PS & "!" & sA_PS)
                        End If
                    End If
                    ' continuer pour le prochain
                    déjàTraité = False
                    prem_ligne_matricule = 0
                    str_new_poste_secondaire = ""
                    str_new_assoc = ""
                End If
                ' Voir si dans synchro, transformer en spm
                If VoirSiDansSynchroDoublon(ligne_lu, spm_doublon) Then
                    ' on ne met pas le poste principal (la première ligne)
                    If prem_ligne_matricule > 0 Then
                        sql = "INSERT INTO " & table_temporaire_doublons & " VALUES(" _
                                                & "'" & matricule_en_cours & "', '" & spm_doublon & "' )"
                        Call Odbc_Cnx.Execute(sql)
                    Else
                        new_spm_princ = spm_doublon
                    End If
                    str_new_poste_secondaire = str_new_poste_secondaire & spm_doublon
                    str_new_assoc = str_new_assoc & P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_code_section, p_long_code_section, "section")
                    str_new_assoc = str_new_assoc & ";" & P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_code_emploi, p_long_code_emploi, "emploi") & "|"
                End If
                ' on ne traite vraiment que le premier
                prem_ligne_matricule = prem_ligne_matricule + 1
                If matricule_en_cours = matricule_ancien Then
                    GoTo lab_suivant
                End If
                '
                sql = "SELECT U_prefixe, U_Prenom, U_Num, U_Nom, U_Matricule, U_Prenom, U_Po_Princ, U_SPM, U_NomJunon, U_PrenomJunon" _
                    & " FROM Utilisateur" _
                    & " WHERE U_kb_actif=True AND U_Actif=TRUE AND U_ExterneFich=FALSE" _
                    & " AND U_Matricule=" & Odbc_String(matricule_en_cours)
                If Odbc_SelectV(sql, rs) = P_ERREUR Then
                    Call quitter("Pb sur ImportationAnnuaire.remplir_grid sql=" & sql)
                    Exit Sub
                End If
                If rs.EOF Then
                    ' ligne (MATRICULE) du fichier ne se trouve pas dans la base => à CREER
                    ' Voir si a des infos. supplémentaires
                    ' info suppl ARRET
                    For I = 0 To p_nbr_lstInfoSuppl - 1
                        sinfo = P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, LISTE_TIS_POS(I).prmgenb_tis_pos, LISTE_TIS_POS(I).prmgenb_tis_long, LISTE_TIS_POS(I).prmgenb_tis_lib)
                        If sinfo <> "" Then
                                'If g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).umatricule <> matricule_en_cours Then
                                    g_nbr_infoSuppl = g_nbr_infoSuppl + 1
                                    ReDim Preserve g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl)
                                    g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).unum = 0
                                    g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).unom = P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_nom, p_long_nom, "nom") _
                                                                                              & vbTab & P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_prenom, p_long_prenom, "prénom")
                                    g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).umatricule = matricule_en_cours
                                    g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).tis_pour_creer = True
                                'End If
                                g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).tis_alimente = LISTE_TIS_POS(I).prmgenb_tis_lien
                                g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).tis_num = LISTE_TIS_POS(I).prmgenb_tis_num
                                If g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).infosuppl <> "" Then
                                    g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).infosuppl = g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).infosuppl & vbTab
                                End If
                                g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).infosuppl = g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).infosuppl _
                                                                                  & P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, LISTE_TIS_POS(I).prmgenb_tis_pos, LISTE_TIS_POS(I).prmgenb_tis_long, LISTE_TIS_POS(I).prmgenb_tis_lib)
                                g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).tis_value = g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).tis_value _
                                                                                    & LISTE_TIS_POS(I).prmgenb_tis_num & ";" & P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, LISTE_TIS_POS(I).prmgenb_tis_pos, LISTE_TIS_POS(I).prmgenb_tis_long, LISTE_TIS_POS(I).prmgenb_tis_lib) & ";|"
                        End If
                    Next I
                    .AddItem ""
                    Call a_creer(ligne_lu, .Rows - 1)
                    déjàTraité = True
                Else
                    ' ligne (MATRICULE) se trouve dans la base et dans le fichier
                    If rs.RowCount > 1 Then
                        ' Des doublons !!
                        lib1 = ""
                        While Not rs.EOF
                            lib1 = lib1 & rs("U_Matricule").Value & " : " & rs("U_Nom").Value & " " & rs("U_PreNom").Value & Chr(13) & Chr(10)
                            rs.MoveNext
                        Wend
                        rs.Close
                        s = "ERREUR ! => La vérification n'est pas encore terminée." _
                                  & " Il reste des matricules redondants." & Chr(13) & Chr(10) & lib1
                        Call MsgBox(s, vbCritical + vbOKOnly, "Attention")
                        p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "*************************************************************" & s
                        Print #g_fd1, p_mess_fait_background
                        Close #fd
                        Call quitter(s)
                        Exit Sub
                    End If
                    ' *************************** Remplir les lignes à METTRE_A_JOUR ***************************
                    spm_trouve = ""
                    nom_a_comparer = IIf(rs("U_NomJunon").Value <> "", rs("U_NomJunon").Value, rs("U_Nom").Value)
                    new_nom = P_ChangerCar(Trim$(UCase$(P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_nom, p_long_nom, "nom"))), tbcaractere_nontraite)
                    If new_nom <> UCase$(nom_a_comparer) Then
                        ' Le NOM n'est pas le même => METTRE_A_JOUR
                        ' tester si (SECTION, EMPLOI et NUM_POSTE) se trouve dans Synchro
                        If Not verifier_spm(ligne_lu, rs("U_SPM").Value, rs("U_Po_Princ").Value, spm_trouve) Then
                            Call a_corriger(ligne_lu, rs, "NOM_POSTE", spm_trouve)
                            déjàTraité = True
                        Else
                            Call a_corriger(ligne_lu, rs, "NOM", spm_trouve)
                            déjàTraité = True
                        End If
                    Else ' Le NOM est le même
                        ' tester si (SECTION, EMPLOI et NUM_POSTE) se trouve dans Synchro
                        sret = verifier_spm(ligne_lu, rs("U_SPM").Value, rs("U_Po_Princ").Value, spm_trouve)
                        If sret Then
                            If new_spm_princ <> spm_trouve & "|" Then
                                Call a_corriger(ligne_lu, rs, "POSTE", spm_trouve)
                                déjàTraité = True
                            End If
                        End If
                    End If
                    ' info suppl ARRET
                    For I = 0 To p_nbr_lstInfoSuppl - 1
                        sinfo = P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, LISTE_TIS_POS(I).prmgenb_tis_pos, LISTE_TIS_POS(I).prmgenb_tis_long, LISTE_TIS_POS(I).prmgenb_tis_lib)
                        If sinfo <> "" Then
                            ' a-t-on déjà cette information dans la table ?
                            If Not infosuppl_existant(LISTE_TIS_POS(I).prmgenb_tis_num, _
                                                      rs("U_Num").Value, _
                                                      sinfo) Then
                                If g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).unum <> rs("U_Num").Value Then
                                    g_nbr_infoSuppl = g_nbr_infoSuppl + 1
                                    ReDim Preserve g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl)
                                    g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).unum = rs("U_Num").Value
                                    g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).unom = P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_nom, p_long_nom, "nom") _
                                                                                              & vbTab & P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_prenom, p_long_prenom, "prénom")
                                    g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).umatricule = matricule_en_cours
                                    g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).tis_pour_creer = False
                                End If
                                g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).tis_alimente = LISTE_TIS_POS(I).prmgenb_tis_lien
                                g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).tis_num = LISTE_TIS_POS(I).prmgenb_tis_num
                                If g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).infosuppl <> "" Then
                                    g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).infosuppl = g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).infosuppl & vbTab
                                End If
                                g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).infosuppl = g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).infosuppl _
                                                                                  & P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, LISTE_TIS_POS(I).prmgenb_tis_pos, LISTE_TIS_POS(I).prmgenb_tis_long, LISTE_TIS_POS(I).prmgenb_tis_lib)
                                g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).tis_value = g_LISTE_U_INFO_SUPPL(g_nbr_infoSuppl).tis_value _
                                                                                    & LISTE_TIS_POS(I).prmgenb_tis_num & ";" & P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, LISTE_TIS_POS(I).prmgenb_tis_pos, LISTE_TIS_POS(I).prmgenb_tis_long, LISTE_TIS_POS(I).prmgenb_tis_lib) & ";|"
                            End If
                        End If
                    Next I
                    num_unique = num_unique + 1
                    sql = "INSERT INTO " & table_temporaire & " VALUES(" & num_unique _
                                        & ", '" & matricule_en_cours & "')"
                    Call Odbc_Cnx.Execute(sql)
                End If
                rs.Close
            End If
lab_suivant:
            matricule_ancien = matricule_en_cours
            
            If p_mess_fait_background <> "" Then
                p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "*************************************************************"
                Print #g_fd1, p_mess_fait_background
            End If
            If p_mess_pasfait_background <> "" Then
                p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "*************************************************************"
                Print #g_fd2, p_mess_pasfait_background
            End If
        Wend ' fin de lecture du fichier
        ' Fermer le fichier
        Close #fd
        
        ' message pour les caractères non traités
        On Error Resume Next
        iBoucle = UBound(tbcaractere_nontraite)
        If iBoucle > 0 Then
            mesPasTraité = "Certains caratères ne sont pas pris en compte" & Chr(13) & Chr(10)
            For I = 0 To iBoucle
                If tbcaractere_nontraite(I) <> "" Then
                    'mesPasTraité = mesPasTraité & Chr(13) & Chr(10) & Chr(STR_GetChamp(tbcaractere_nontraite(I), " ", 0)) & " (Code ASCII " & tbcaractere_nontraite(I) & ")"
                    mesPasTraité = mesPasTraité & Chr(13) & Chr(10) & tbcaractere_nontraite(I)
                    'MsgBox mesPasTraité
                End If
            Next I
            mesPasTraité = mesPasTraité & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Veuillez Prévenir KaliTech ou ajouter ce caractère dans le fichier de paramétrage"
            If p_traitement_background Then
                p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "*************************************************************"
                p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & mesPasTraité
                Print #g_fd2, p_mess_pasfait_background
            Else
                MsgBox mesPasTraité
            End If
        End If
        ' *************************** Remplir les lignes à DETRUIRE ***************************
        sql = "SELECT * FROM Utilisateur WHERE U_kb_actif=True AND U_Actif=True AND U_ExterneFich=False" _
            & " AND U_Matricule NOT IN (SELECT tt_matricule FROM " & table_temporaire & ")"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Call quitter("Pb sur ImportationAnnuaire.remplir_grid sql=" & sql)
            Exit Sub
        End If
                
        Me.LbGauge2.Caption = "Supprimer"
        Me.Refresh
        If Not rs.EOF Then
            cmd(CMD_DESACTIVER_TOUS).Visible = True
            rs.MoveLast
            rs.MoveFirst
            pgb2.Max = pgb2.Max + rs.RowCount
        End If

        While Not rs.EOF
            pgb2.Value = pgb2.Value + 1
            p_mess_fait_background = ""
            p_mess_pasfait_background = ""
            p_corps_background = ""
            Call a_detruire(rs)
            If p_mess_fait_background <> "" Then
                p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "*************************************************************"
                Print #g_fd1, p_mess_fait_background
            End If
            If p_mess_pasfait_background <> "" Then
                p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & "*************************************************************"
                Print #g_fd2, p_mess_pasfait_background
            End If
            'If p_corps_background <> "" Then
            '    p_corps_background = p_corps_background & Chr(13) & Chr(10) & "*************************************************************"
            '    Print #g_fd3, p_corps_background
            'End If
            rs.MoveNext
        Wend
        rs.Close
        
        If .Rows - 1 = 0 Then
            Call MsgBox("Il n'y a aucune modification à traiter.", vbInformation + vbOKOnly, "")
            Call quitter("Il n'y a aucune modification à traiter.")
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "Il n 'y a aucune modification à traiter."
            p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "*************************************************************"
            Print #g_fd1, p_mess_fait_background
            Exit Sub
        End If
        
        ' Trier le grid
        .Row = 0
        .col = GRDC_NOM
        .Sort = 1
        g_sens_tri = 1
        g_col_tri = GRDC_NOM
        .CellBackColor = COLOR_COLONNE_TRIEE
        ' Selection de la première ligne (si elle existe)
        .col = 0
        .ColSel = .Cols - 1
        .Row = 1
        .RowSel = 1

        ' Fermer le fichier
        Close #fd
        If p_est_sur_serveur Then
            Call FICH_EffacerFichier(nomfich, False)
        End If
        
        ' détruire la table temporaire
        Call Odbc_Cnx.Execute("DROP TABLE " & table_temporaire)

        If .Rows - 1 = 0 Then
            .Enabled = False
        ElseIf grd.Rows <= 22 Then
            .ColWidth(GRDC_LIB_POSTE_FICH) = .ColWidth(GRDC_LIB_POSTE_FICH) + 255
        End If
        txt(TXT_MATRICULE).Text = ""
        txt(TXT_NOM).Text = ""
        frm(FRM_RECHERCHER).Visible = True
        .Visible = True
' A NE PAS SUPPRIMER
        'cmd(CMD_ACTUALISER).Visible = True
    End With
End Sub

Private Function controlePS(v_PS As String)
    Dim I As Integer, s As String
    Dim j As Integer
    Dim s1 As String, deja As Boolean
    Dim sret As String
    
    For I = 0 To STR_GetNbchamp(v_PS, "|") - 1
        deja = False
        s = STR_GetChamp(v_PS, "|", I)
        If s <> "" Then
            For j = 0 To STR_GetNbchamp(sret, "|") - 1
                s1 = STR_GetChamp(v_PS, "|", j)
                If s = s1 Then
                    deja = True
                    Exit For
                End If
            Next j
        End If
        If s <> "" And Not deja Then
            sret = sret & s & "|"
        End If
    Next I
    If sret <> v_PS Then
        I = I
    End If
    controlePS = sret
End Function

Private Function RecupAncPstSecond(v_ligne_lu As String, ByRef r_str_anc_poste_secondaire As String, v_matricule_ancien As String, ByRef r_rs As rdoResultset)
    Dim sql As String
    Dim Anc_Po_Princ As Long, laS As String
    Dim str_anc_poste_secondaire As String
    '
    r_str_anc_poste_secondaire = construit_anc_poste_secondaire(v_matricule_ancien)
    sql = "SELECT * FROM Utilisateur WHERE U_Matricule='" & v_matricule_ancien & "'"
    Call Odbc_SelectV(sql, r_rs)
    If Not r_rs.EOF Then
        Anc_Po_Princ = r_rs("U_Po_Princ")
        laS = build_arbor_srv(Anc_Po_Princ)
        r_str_anc_poste_secondaire = laS & r_str_anc_poste_secondaire
        RecupAncPstSecond = True
    Else
        RecupAncPstSecond = False
    End If
End Function

Private Function construit_anc_poste_secondaire(v_matricule)
    Dim sql As String, rs As rdoResultset
    Dim sOut As String
    
    sql = "select * from kb_poste_secondaire where psu_matricule='" & v_matricule & "'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Function
    End If
    While Not rs.EOF
        sOut = sOut & rs("psu_poste") & "|"
        rs.MoveNext
    Wend
    construit_anc_poste_secondaire = sOut
End Function

Private Function Egaux_postes_secondaires(v_new_PS, v_anc_PS)
    Dim n As Integer, j As Integer
    Dim I As Integer, m As Integer
    Dim bTrouvé As Boolean
    Dim sAnc As String, sNew As String
    Dim yEstPas As Boolean
    Dim bool1 As Boolean, bool2 As Boolean
    
    If STR_GetNbchamp(v_anc_PS, "|") <> STR_GetNbchamp(v_new_PS, "|") Then
        Egaux_postes_secondaires = False
        Exit Function
    ElseIf v_new_PS = v_anc_PS Then
        Egaux_postes_secondaires = True
        Exit Function
    End If
    bool1 = True
    bool2 = True
    ' voir si différence de postes (pas le même nombre ou diff.)
    Egaux_postes_secondaires = True
    n = STR_GetNbchamp(v_anc_PS, "|")
    For I = 0 To n
        sAnc = STR_GetChamp(v_anc_PS, "|", I)
        If sAnc <> "" Then
            m = STR_GetNbchamp(v_new_PS, "|")
            bTrouvé = False
            For j = 0 To m
                sNew = STR_GetChamp(v_new_PS, "|", j)
                If sNew <> "" Then
                    If sAnc = sNew Then
                        bTrouvé = True
                        Exit For
                    End If
                End If
            Next j
            If Not bTrouvé Then
                bool1 = False
                Egaux_postes_secondaires = False
                Exit Function
            End If
        End If
    Next I
    '
    bTrouvé = False
    n = STR_GetNbchamp(v_new_PS, "|")
    For I = 0 To n
        sNew = STR_GetChamp(v_new_PS, "|", I)
        If sNew <> "" Then
            m = STR_GetNbchamp(v_anc_PS, "|")
            bTrouvé = False
            For j = 0 To m
                sAnc = STR_GetChamp(v_anc_PS, "|", j)
                If sAnc <> "" Then
                    If sNew = sAnc Then
                        bTrouvé = True
                        Exit For
                    End If
                End If
            Next j
            If Not bTrouvé Then
                bool2 = False
            End If
        End If
    Next I
    If Not bool1 And Not bool2 Then
        Egaux_postes_secondaires = False
    End If
    
End Function
Private Function remplir_utilmouvement(ByVal v_row As Integer, ByVal v_spm As String) As Integer
' ********************************************************
' Appelée depuis modifier_cette_personne()
' Eenseigner la table UtilMouvement avec les modifications
' ********************************************************
    Dim lng As Long
    Dim sql As String, old_nom As String, old_prenom As String, old_matricule As String, old_spm As String

    With grd
        sql = "SELECT U_Nom, U_Prenom, U_Matricule,U_SPM FROM Utilisateur WHERE U_Num=" & .TextMatrix(v_row, GRDC_U_NUM)
        If Odbc_RecupVal(sql, old_nom, old_prenom, old_matricule, old_spm) = P_ERREUR Then
            GoTo lab_erreur
        End If
        ' modifications
        If old_nom <> .TextMatrix(v_row, GRDC_NOM) Then ' -------------------------  NOM
            If P_InsertIntoUtilmouvement(.TextMatrix(v_row, GRDC_U_NUM), "M", "NOM=" & old_nom & ";" & .TextMatrix(v_row, GRDC_NOM) & ";", 0) = P_ERREUR Then
                GoTo lab_erreur
            End If
        End If
        If UCase$(old_prenom) <> UCase$(.TextMatrix(v_row, GRDC_PRENOM)) Then ' ---  PRENOM
            If P_InsertIntoUtilmouvement(.TextMatrix(v_row, GRDC_U_NUM), "M", "PRENOM=" & old_prenom & ";" & .TextMatrix(v_row, GRDC_PRENOM) & ";", 0) = P_ERREUR Then
                GoTo lab_erreur
            End If
        End If
        If old_matricule <> .TextMatrix(v_row, GRDC_MATRICULE) Then ' -------------  MATRICULE
            If P_InsertIntoUtilmouvement(.TextMatrix(v_row, GRDC_U_NUM), "M", "MATRICULE=" & old_matricule & ";" & .TextMatrix(v_row, GRDC_MATRICULE) & ";", 0) = P_ERREUR Then
                GoTo lab_erreur
            End If
        End If
        If old_spm <> v_spm Then ' ------------------------------------------------  POSTE
            If P_InsertIntoUtilmouvement(.TextMatrix(v_row, GRDC_U_NUM), "M", P_get_poste_modif(old_spm, v_spm), 0) = P_ERREUR Then
                GoTo lab_erreur
            End If
        End If
    End With

    remplir_utilmouvement = P_OK
    Exit Function

lab_erreur:
    remplir_utilmouvement = P_ERREUR

End Function

Private Sub selectionner_ligne(ByVal v_row As Integer)
' *******************************
' Mettre la ligne en surbrillance
' *******************************
    With grd
        .col = GRDC_U_NUM
        .Row = v_row
        .ColSel = .Cols - 1
        .RowSel = v_row
        If p_traitement_background_semiauto Then    ' auto demandé par le menu
            .TopRow = .Row
        End If
    End With

End Sub

Private Function verifier_spm(ByVal v_ligne_en_cours As String, _
                              ByVal v_spm As String, _
                              ByVal v_numpo As Long, _
                              ByRef r_spm As String) As Boolean
' *************************************************************************
' Verifier si la (SECTION, EMPLOI et POSTE) se trouve dans la table Synchro
' *************************************************************************
    Dim sql As String, spm_en_cours As String, section As String, emploi As String
    Dim nbr_spm As Integer, poste_en_cours As Long, I As Integer, nbr As Integer
    Dim lnb As Long
    
    v_spm = ControleFormat(v_spm)

    nbr_spm = STR_GetNbchamp(v_spm, "|")
    r_spm = ""
    For I = 0 To nbr_spm - 1
        spm_en_cours = STR_GetChamp(v_spm, "|", I)
        nbr = STR_GetNbchamp(spm_en_cours, ";")
        poste_en_cours = Mid$(STR_GetChamp(spm_en_cours, ";", nbr - 1), 2)
        If poste_en_cours = v_numpo Then
            r_spm = spm_en_cours
            Exit For
        End If
    Next I
    
    section = P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_code_section, p_long_code_section, "code section")
    section = Trim$(section)
    emploi = P_lire_valeur(p_type_fichier, v_ligne_en_cours, p_separateur, p_pos_code_emploi, p_long_code_emploi, "code emploi")
    emploi = Trim$(emploi)
    ' Parcourir la table Synchro pour la paire (SECTION, EMPLOI)
    sql = "SELECT count(*) FROM Synchro" _
        & " WHERE Sync_Section=" & Odbc_String(section) _
        & " AND Sync_Emploi=" & Odbc_String(emploi) _
        & " and sync_spnum=" & v_numpo
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        verifier_spm = False
        p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "Erreur SQL pour " & sql
        Exit Function
    End If
    If lnb = 0 Then
        verifier_spm = True
        p_mess_fait_background = p_mess_fait_background & Chr(13) & Chr(10) & "Pas d'association pour section=" & section & " et emploi=" & emploi
        Exit Function
    End If

    verifier_spm = True

End Function

Private Function verifier_coherence_assoc(ByVal v_new_spm As String, ByVal v_new_assoc As String, ByVal v_matricule As String)
' ***********************************************************************
' Verifier si pas d'associations redondantes et retourner la bonne chaine
' ***********************************************************************
    Dim sql As String, spm As String, spm_en_cours As String, section As String, emploi As String
    Dim nbr_spm As Integer, poste_en_cours As Long, j As Integer, I As Integer, nbr As Integer
    Dim lnb As Long, mess As String, s1 As String, s2 As String
    Dim s_ret As String, Anc_Po_Princ As Integer, laS As String
    Dim v_Nspm As String, v_Aspm As String, sret As String
    Dim str_anc_poste_secondaire As String, messa As String
    Dim rs As rdoResultset
    Dim str_spmPrinc As String
    Dim deja As Boolean
    
    str_anc_poste_secondaire = construit_anc_poste_secondaire(v_matricule)
    sql = "SELECT * FROM Utilisateur WHERE U_Matricule='" & v_matricule & "'"
    Call Odbc_SelectV(sql, rs)
    If Not rs.EOF Then
        Anc_Po_Princ = rs("U_Po_Princ")
        laS = build_arbor_srv(Anc_Po_Princ)
        str_anc_poste_secondaire = laS & str_anc_poste_secondaire
    End If
    v_Nspm = ControleFormat(v_new_spm)
    v_Aspm = ControleFormat(str_anc_poste_secondaire)

    str_spmPrinc = STR_GetChamp(v_Nspm, "|", 0)
    nbr_spm = STR_GetNbchamp(v_Nspm, "|")
    For I = 1 To nbr_spm - 1
        spm_en_cours = STR_GetChamp(v_Nspm, "|", I)
        If spm_en_cours = str_spmPrinc Then
            s1 = STR_GetChamp(v_new_assoc, "|", 0)
            s2 = STR_GetChamp(v_new_assoc, "|", I)
            mess = mess & " => Redondance d'associations pour " & s1 & " et " & s2 & Chr(13) & Chr(10)
        End If
    Next I
    ' Voir les redondances et les enlever
    sret = str_spmPrinc & "|"
    nbr_spm = STR_GetNbchamp(v_Nspm, "|")
    For I = 1 To nbr_spm - 1
        spm_en_cours = STR_GetChamp(v_Nspm, "|", I)
        deja = False
        For j = I + 1 To nbr_spm - 1
            spm = STR_GetChamp(v_Nspm, "|", j)
            If spm_en_cours = spm Then
                s1 = STR_GetChamp(v_new_assoc, "|", I)
                s2 = STR_GetChamp(v_new_assoc, "|", j)
                deja = True
                mess = mess & " => Redondance d'associations pour " & s1 & " et " & s2
                Exit For
            End If
        Next j
        If Not deja And spm_en_cours <> str_spmPrinc Then
            sret = sret & spm_en_cours & "|"
        End If
    Next I
    verifier_coherence_assoc = sret
    If mess <> "" Then
        p_mess_pasfait_background = p_mess_pasfait_background & Chr(13) & Chr(10) & mess
    End If

End Function

Private Function ControleFormat(v_spm As String)
    Dim I As Integer, j As Integer, n As Integer, m As Integer
    Dim s As String, s1 As String
    Dim s2 As String
    Dim sOut As String
    ' S209;S279;S280;S566;P12204;S891;S896;S912;P12504;|
    
    s = v_spm
    'MsgBox "s=" & s
    n = STR_GetNbchamp(s, "|")
    For I = 0 To n
        s1 = STR_GetChamp(s, "|", I)
        If s1 <> "" Then
            m = STR_GetNbchamp(s1, ";")
            For j = 0 To m
                s2 = STR_GetChamp(s1, ";", j)
                If s2 <> "" Then
                    If Mid(s2, 1, 1) = "P" Then
                        sOut = sOut & s2 & ";|"
                    Else
                        sOut = sOut & s2 & ";"
                    End If
                End If
            Next j
        End If
    Next I
    sOut = sOut & "|"
    sOut = Replace(sOut, "||", "|")
    'MsgBox sOut
    ControleFormat = sOut
End Function

Private Sub cmd_Click(Index As Integer)
    
    Dim sPoste As String, sService As String, spm_choisi As String
    Dim I As Integer
    
    Select Case Index
        Case CMD_QUESTION
            Call envoyer_question
        Case CMD_LISTE_ASSOC
            Call Liste_Assoc("", "", "", "")
        Case CMD_QUITTER
            Call quitter("")
        Case CMD_RECHERCHER
            Call rechercher_personne
        Case CMD_ACTUALISER
            Call actualiser
        Case CMD_CORRIGER_TOUS
            'sPoste = grd.TextMatrix(grd.Row, GRDC_CODE_POSTE_FICH)
            'sService = grd.TextMatrix(grd.Row, GRDC_CODE_SRV_FICH)
            
            sPoste = g_GRDC_CODE_POSTE_FICH
            sService = g_GRDC_CODE_SRV_FICH
            
            For I = 0 To grd.Rows - 1
                If I >= grd.Rows Then Exit For
                grd.Row = I
                If sPoste = grd.TextMatrix(I, GRDC_CODE_POSTE_FICH) Then
                    If sService = grd.TextMatrix(I, GRDC_CODE_SRV_FICH) Then
                        If grd.TextMatrix(I, GRDC_ACTION) = "Corriger" Then
                            ' MsgBox "oui " & i & " " & grd.TextMatrix(i, GRDC_NOM)
                            spm_choisi = build_arbor_srv(g_poste)
                            Call modifier_cette_personne(I, spm_choisi, "", True)
                            grd.RemoveItem (I)
                            I = I - 1
                            Call actualiser_compteur
                        End If
                    End If
                End If
            Next I
            Me.cmd(CMD_CORRIGER_TOUS).Visible = False
            g_poste = ""
        Case CMD_DESACTIVER_TOUS
            Call desactiver_tous
    End Select

End Sub

Public Function Liste_Assoc(v_emploi As String, v_section As String, v_titre As String, v_trait As String) As String
    Dim sql As String
    Dim srvnum As String
    Dim I As Long
    Dim s As String
    Dim rs As rdoResultset
    Dim lnb As Long
    Dim sret As String
    Dim frm As Form
    Dim cret As String
    Dim StrFct As String
    Dim emploi As String, service As String
    Dim str_filtre As String
    Dim s1 As String, s2 As String
    Dim s3 As String
    Dim s4 As String
    Dim bDirect As Boolean
    Dim lst_ponum As String
    Dim nb As Long
    Dim selected As Boolean
Lab_Début:

    If v_emploi <> "" And v_section <> "" Then
        bDirect = True
        If v_trait = "New" Then
            GoTo Lab_Nouvelle
        End If
    End If
    
    Call CL_Init
    Call CL_InitMultiSelect(True, True) ' (selection multiple=True, retourner la ligne courante=False)
    Call CL_InitGererTousRien(True)
    If bDirect Then
        Call CL_InitTitreHelp("Associations pour : " & v_titre, "")
    Else
        Call CL_InitTitreHelp("Liste des associations déjà réalisées", "")
    End If
    Call CL_InitTaille(0, -15)
    
    sql = "SELECT Sync_Num, Sync_Auto, Sync_emploi, Sync_Section, Sync_Spnum, Po_Num, Srv_Num, Srv_Nom, Ft_Num, Ft_Code, Ft_Libelle"
    sql = sql & " FROM Synchro, Poste, FctTrav, Service "
    sql = sql & " Where Poste.po_num = Synchro.sync_spnum"
    sql = sql & " And Poste.po_srvnum = Service.srv_num"
    sql = sql & " And Poste.po_ftnum = Fcttrav.ft_num"
    If bDirect Then
        sql = sql & " And Sync_emploi = '" & v_emploi & "' And Sync_section = '" & v_section & "'"
    End If
    sql = sql & " Order by Sync_Section, Sync_emploi"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Function
    End If
    If bDirect And rs.EOF Then
        Liste_Assoc = ""
        Exit Function
    End If
    s = UCase(str_filtre)
    lst_ponum = ""
    rs.MoveLast
    rs.MoveFirst
    nb = rs.RowCount
    selected = False
    While Not rs.EOF
        s1 = rs("Sync_emploi") & " : " & rs("Sync_Section")
        s2 = rs("Srv_Nom")
        s3 = rs("Ft_Libelle") & " (" & rs("Ft_Code") & " poste n° " & rs("PO_Num") & ")"
        'If s = "" Or (InStr(UCase(s1), s) > 0 Or InStr(UCase(s2), s) > 0 Or InStr(UCase(s3), s) > 0) Then
            s4 = Mid(s2 & " - - - - - - - - - - - - - - - - - ", 1, 50) & " : " & s3
            If nb = 1 Then selected = True
            Call CL_AddLigne(IIf(rs("Sync_auto"), "Auto.", "") & vbTab & s1 & vbTab & s4, rs("PO_Num"), rs("Sync_Num"), selected)
            lst_ponum = lst_ponum & IIf(lst_ponum = "", "", "|") & rs("PO_Num") & "!" & rs("SYNC_Auto")
        'End If
        rs.MoveNext
    Wend
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("Voir les Personnes", "", vbKeyO, vbKeyF1, 2000)
    Call CL_AddBouton("Supprimer", "", vbKeyO, vbKeyF1, 1400)
    Call CL_AddBouton("Rendre Auto", "", 0, 0, 1400)
    Call CL_AddBouton("Supprimer Auto", "", 0, 0, 1400)
    Call CL_AddBouton("Nouvelle Assoc.", "", 0, 0, 1400)
    'Call CL_AddBouton("Filtrer", "", 0, 0, 1400)
    'Call CL_AddBouton("initialiser à partir du fichier", "", 0, 0, 2000)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
Lab_Retour:
    ChoixListe.Show 1
    
    If CL_liste.retour = 6 Or CL_liste.retour = 0 Then ' --------------- QUITTER
        If bDirect Then
            Liste_Assoc = lst_ponum
        End If
        Exit Function
    End If
    
    ' Filtre
    'If CL_liste.retour = 5 Then ' --------------- Filtre
    '    str_filtre = InputBox("Rechercher sur", "Rechercher", str_filtre)
    '    If str_filtre = "" Then
    '        GoTo Lab_Retour
    '    Else
    '        GoTo Lab_Début
    '    End If
    'End If
    
    ' à partir du fichier
    'If CL_liste.retour = 5 Then ' --------------- à partir du fichier
    '    Set frm = AnalyseFichier
    '    cret = AnalyseFichier.AppelFrm
    'End If
    
    ' Nouvelle
    If CL_liste.retour = 5 Then ' --------------- Nouvelle
        If bDirect Then
Lab_Nouvelle:
            v_trait = ""
            Set frm = PrmService
            sret = PrmService.AppelFrm("Choix d'un poste", "S", False, "", "P", False)
            Set frm = Nothing
            If sret <> "" Then
                sret = STR_GetChamp(sret, ";P", 1)
                sret = Replace(sret, "|", "")
                sret = Replace(sret, ";", "")
                sql = "Insert into Synchro "
                sql = sql & "(SYNC_Section, SYNC_Emploi, SYNC_SPNum, SYNC_Auto)"
                sql = sql & "Values ('" & v_section & "', '" & v_emploi & "', " & sret & ", 'f')"
                Call Odbc_Cnx.Execute(sql)
            End If
        Else
            emploi = InputBox("Code emploi dans le fichier GRH", "Nouvelle association")
            If emploi <> "" Then
                service = InputBox("Code service dans le fichier GRH", "Nouvelle association")
                If service <> "" Then
                    Set frm = PrmService
                    sret = PrmService.AppelFrm("Choix d'un poste", "S", False, "", "P", False)
                    Set frm = Nothing
                    If sret <> "" Then
                        sret = STR_GetChamp(sret, ";P", 1)
                        sret = Replace(sret, "|", "")
                        sret = Replace(sret, ";", "")
                        sql = "Insert into Synchro "
                        sql = sql & "(SYNC_Section, SYNC_Emploi, SYNC_SPNum, SYNC_Auto)"
                        sql = sql & "Values ('" & service & "', '" & emploi & "', " & sret & ", 'f')"
                        Call Odbc_Cnx.Execute(sql)
                    End If
                End If
            End If
        End If
        GoTo Lab_Début
    End If
    
    ' Auto
    If CL_liste.retour = 3 Then ' --------------- Mettre en auto
        For I = 0 To UBound(CL_liste.lignes)
            If (CL_liste.lignes(I).selected) Then
                ' Supprimer CL_liste.lignes(CL_liste.pointeur).tag
                If Odbc_Update("Synchro", _
                                "Sync_Num", _
                               "WHERE Sync_Num=" & CL_liste.lignes(I).tag, _
                                "Sync_Auto", True) = P_ERREUR Then
                    MsgBox "Impossible de Supprimer"
                End If
            End If
        Next I
    End If
    
    ' Non Auto
    If CL_liste.retour = 4 Then ' --------------- Mettre en Non auto
        For I = 0 To UBound(CL_liste.lignes)
            If (CL_liste.lignes(I).selected) Then
                ' Supprimer CL_liste.lignes(CL_liste.pointeur).tag
                If Odbc_Update("Synchro", _
                                "Sync_Num", _
                               "WHERE Sync_Num=" & CL_liste.lignes(I).tag, _
                                "Sync_Auto", False) = P_ERREUR Then
                    MsgBox "Impossible de Supprimer"
                End If
            End If
        Next I
    End If
    
    ' Voir les personnes
    If CL_liste.retour = 1 Then ' --------------- Voir les Personnes
        lnb = 0
        StrFct = ""
        For I = 0 To UBound(CL_liste.lignes)
            If (CL_liste.lignes(I).selected) Then
                StrFct = StrFct & ";" & CL_liste.lignes(I).num
            End If
        Next I
        Call CL_Init
        Call CL_InitMultiSelect(False, False) ' (selection multiple=True, retourner la ligne courante=False)
        Call CL_InitTitreHelp("Liste des Personnes", "")
        Call CL_InitTaille(0, -15)
        Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
        For I = 0 To STR_GetNbchamp(StrFct, ";")
            If STR_GetChamp(StrFct, ";", I) <> "" Then
                sql = "SELECT PO_Num, SRV_Num, SRV_Nom, FT_Num, Ft_Code, Ft_Libelle FROM FctTrav, Poste, Service" _
                    & " WHERE PO_Ftnum = FT_Num And PO_Srvnum = SRV_Num And PO_Num = " & STR_GetChamp(StrFct, ";", I)
                'MsgBox sql
                If Odbc_SelectV(sql, rs) = P_ERREUR Then
                    Exit Function
                End If
                If rs.EOF Then
                    Call CL_AddLigne("Poste " & STR_GetChamp(StrFct, ";", I) & " non trouvée", 0, 0, False)
                Else
                    srvnum = rs("Srv_Num")
                    Call CL_AddLigne(rs("SRV_Nom") & vbTab & rs("Ft_Libelle") & " (" & rs("Ft_Code") & ")", 0, 0, False)
                    sql = "SELECT U_Nom, U_Prenom, U_Matricule, U_kb_actif FROM Utilisateur " _
                        & " WHERE U_Spm LIKE '%S" & srvnum & ";%P" & rs("PO_Num") & ";%' ORDER BY U_Nom"
                    'MsgBox sql
                    If Odbc_SelectV(sql, rs) = P_ERREUR Then
                        Exit Function
                    End If
                    While Not rs.EOF
                        lnb = lnb + 1
                        Call CL_AddLigne("   " & IIf(rs("U_kb_actif"), "Actif", "") & vbTab & rs("U_Nom") & " " & rs("U_Prenom") & " (" & rs("U_Matricule") & ")", 0, 0, False)
                        rs.MoveNext
                    Wend
                End If
            End If
        Next I
        
        If lnb > 0 Then
            ChoixListe.Show 1
        Else
            MsgBox "Aucune personne trouvée"
        End If
        
        GoTo Lab_Début

    End If

    ' supprimer
    If CL_liste.retour = 2 Then ' --------------- SUPPRIMER
        For I = 0 To UBound(CL_liste.lignes)
            If (CL_liste.lignes(I).selected) Then
                ' Supprimer CL_liste.lignes(CL_liste.pointeur).tag
                If Odbc_Delete("Synchro", _
                                "Sync_Num", _
                               "WHERE Sync_Num=" & CL_liste.lignes(I).tag, _
                                0) = P_ERREUR Then
                    MsgBox "Impossible de Supprimer"
                End If
            End If
        Next I
        str_filtre = ""
    End If
    GoTo Lab_Début

End Function

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Call quitter("")
    End If
    
End Sub

Private Sub Form_Load()

    g_form_active = False

End Sub

Private Sub Form_Activate()

    If g_form_active Then Exit Sub

    g_form_active = True
    g_actualiser = False
    Call initialiser
    
    If p_traitement_background Then
        If p_traitement_background_semiauto Then    ' auto demandé par le menu
            p_traitement_background_semiauto = False
            p_traitement_background = False
            ' renommer le fichier
            Call P_RenommerFichierImportation
        Else
            ' renommer le fichier
            Call P_RenommerFichierImportation
            End
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Unload Me
    End If

End Sub

Private Sub grd_Click()

    Dim mouse_row As Integer, mouse_col As Integer
    Dim ma_civilite As String, njf As String, nom As String, s_srv As String
    Dim titre_compl As String
    Dim s As String
    Dim s_numposte As String
    Dim I As Integer

    With grd
        mouse_col = .MouseCol
        mouse_row = .MouseRow
        .col = GRDC_U_NUM
        If mouse_row <> 0 Then
            .Row = mouse_row
            .RowSel = mouse_row
            .ColSel = .Cols - 1
            cmd(CMD_QUESTION).Visible = True
            cmd(CMD_QUESTION).Enabled = True
            cmd(CMD_QUESTION).tag = grd.TextMatrix(mouse_row, GRDC_U_NUM)
            If mouse_col = GRDC_ACTION Then ' *** C'est ici que sont regroupées toutes les actions ***
                Call action(mouse_row)
            ElseIf mouse_col = GRDC_INFO_PERSO Then
                If grd.TextMatrix(mouse_row, GRDC_ACTION) = "Postes Sec." Then
                    s = grd.TextMatrix(mouse_row, GRDC_ANC_PSTSECOND)
                    If s <> "" Then
                        titre_compl = "Anciens postes secondaires" & Chr(13) & Chr(10)
                        For I = 0 To STR_GetNbchamp(s, "|")
                            s_numposte = STR_GetChamp(s, "|", I)
                            If s_numposte <> "" Then
                                s_numposte = STR_GetChamp(s_numposte, ";P", 1)
                                s_numposte = Replace(s_numposte, ";", "")
                                If s_numposte <> "" Then
                                    titre_compl = titre_compl & IIf(s_numposte = "", "", " -> " & recup_PSLib(s_numposte) & Chr(13) & Chr(10))
                                End If
                            End If
                        Next I
                    Else
                        titre_compl = "Pas d'anciens postes secondaires" & Chr(13) & Chr(10)
                    End If
                    s = grd.TextMatrix(mouse_row, GRDC_NEW_PSTSECOND)
                    If s <> "" Then
                        titre_compl = titre_compl & Chr(13) & Chr(10) & "Nouveaux postes secondaires" & Chr(13) & Chr(10)
                        For I = 0 To STR_GetNbchamp(s, "|")
                            s_numposte = STR_GetChamp(s, "|", I)
                            If s_numposte <> "" Then
                                s_numposte = STR_GetChamp(s_numposte, ";P", 1)
                                s_numposte = Replace(s_numposte, ";", "")
                                If s_numposte <> "" Then
                                    titre_compl = titre_compl & IIf(s_numposte = "", "", " -> " & recup_PSLib(s_numposte) & Chr(13) & Chr(10))
                                End If
                            End If
                        Next I
                    Else
                        titre_compl = titre_compl & "Pas de nouveau poste secondaire" & Chr(13) & Chr(10)
                    End If
                    MsgBox titre_compl
                    Exit Sub
                End If
                If p_pos_civilite <> -1 Then
                    ma_civilite = " * CIVILITE:" & vbTab & .TextMatrix(mouse_row, GRDC_CIVILITE) & vbCrLf
                Else
                    ma_civilite = ""
                End If
                If p_pos_njf <> -1 Then
                    njf = " * NJF:" & vbTab & vbTab & .TextMatrix(mouse_row, GRDC_NJF) & vbCrLf
                Else
                    njf = ""
                End If
                ' Nom
                nom = " * NOM:" & vbTab & vbTab
                If .TextMatrix(mouse_row, GRDC_METTRE_A_JOUR_NOM) <> "" Then
                    nom = nom & .TextMatrix(mouse_row, GRDC_METTRE_A_JOUR_NOM) _
                        & " - (Fichier : " & .TextMatrix(mouse_row, GRDC_NOM) & ")"
                Else
                    nom = nom & .TextMatrix(mouse_row, GRDC_NOM)
                End If
                ' Service et Poste
                If .TextMatrix(mouse_row, GRDC_NUM_SRV_KB) <> "" Then
                    s_srv = " * SERVICE:" & vbTab & P_get_lib_srv_poste(.TextMatrix(mouse_row, GRDC_NUM_SRV_KB), P_SERVICE) & vbCrLf _
                            & " * POSTE:" & vbTab & P_get_lib_srv_poste(.TextMatrix(mouse_row, GRDC_NUM_POSTE), P_POSTE)
                    If .TextMatrix(mouse_row, GRDC_POSTE_EN_GRAS) = P_OUI Then
                        s_srv = s_srv & vbCrLf & " Fichier : " & .TextMatrix(mouse_row, GRDC_LIB_SRV_FICH) & " / " _
                                & .TextMatrix(mouse_row, GRDC_LIB_POSTE_FICH)
                    End If
                    titre_compl = ""
                Else
                    s_srv = " * SERVICE:" & vbTab & .TextMatrix(mouse_row, GRDC_LIB_SRV_FICH) & vbCrLf _
                            & " * POSTE:" & vbTab & .TextMatrix(mouse_row, GRDC_LIB_POSTE_FICH)
                    titre_compl = " à créer"
                End If

                MsgBox " Informations sur la personne" & titre_compl & " :" & vbCrLf & vbCrLf _
                        & ma_civilite _
                        & nom & vbCrLf _
                        & njf _
                        & " * PRENOM:" & vbTab & .TextMatrix(mouse_row, GRDC_PRENOM) & vbCrLf _
                        & " * MATRICULE:" & vbTab & .TextMatrix(mouse_row, GRDC_MATRICULE) & vbCrLf & vbCrLf _
                        & s_srv & vbCrLf & vbCrLf
            ElseIf mouse_col = GRDC_PERSONNE_INACTIVE Then
                Call selectionner_ligne(mouse_row)
                If .TextMatrix(mouse_row, GRDC_LISTE_PERSONNE_INACTIVE) <> "" Then
                    Call afficher_personne_inactive(mouse_row)
                End If
            Else ' N'importe où sur le grid sauf la colonne GRDC_ACTION ni GRDC_PERSONNE_INACTIVE
            End If
        Else ' un clic sur la ligne fixe
            .Row = 1
            .ColSel = .Cols - 1
            .RowSel = 1
        End If
    End With

End Sub

Private Sub grd_DblClick()

    Dim mouse_row As Integer, mouse_col As Integer, I As Integer

    With grd
        If .Rows - 1 = 0 Then Exit Sub ' pour eviter les surprises
        mouse_row = .MouseRow
        mouse_col = .MouseCol
        ' Trier selon la colonne choisie dans la ligne fixe uniquement
        If mouse_row = 0 Then
            ' pas de tri possible avec les colonnes ACTION ni PASTILLE
            If mouse_col = GRDC_PASTILLE Or mouse_col = GRDC_PERSONNE_INACTIVE Then Exit Sub
            .Row = 0
            .col = mouse_col
            If g_col_tri = .col Then
                If g_sens_tri = 1 Then
                    .Sort = 2
                    g_sens_tri = 2
                Else
                    .Sort = 1
                    g_sens_tri = 1
                End If
            Else
                .Sort = 1
                g_sens_tri = 1
            End If
            ' colorer la colonne triée
            .Row = 0
            .col = g_col_tri
            .CellBackColor = COLOR_COLONNE_NON_TRIEE
            .col = mouse_col
            .CellBackColor = COLOR_COLONNE_TRIEE
            .TopRow = 1
            .col = 0
            .Row = 1
            .RowSel = 1
            .ColSel = .Cols - 1
            ' renseigner la colonne nouvellement triée
            g_col_tri = mouse_col
        ElseIf mouse_row <> 0 Then ' n'importe où sur le GRID sauf la 1° ligne
            'txt(TXT_NOM) = .TextMatrix(mouse_row, GRDC_NOM)
            'txt(TXT_MATRICULE) = .TextMatrix(mouse_row, GRDC_MATRICULE)
        End If
    End With

End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)

    With grd
        If (KeyCode >= vbKeyA And KeyCode <= vbKeyZ) And .Row > 0 Then
            Call chercher(.Row, KeyCode)
        ElseIf KeyCode = vbKeyEscape Then
            Call quitter("")
        End If
    End With

End Sub

Private Sub grd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim message As String, mon_service As String, mon_poste As String
    Dim mouse_row As Integer, mouse_col As Integer

    With grd
        If grd.Rows = 1 Then
            Exit Sub
        End If
        mouse_row = .MouseRow
        mouse_col = .MouseCol
        If mouse_row > 0 Then
            If mouse_col = GRDC_NOM Then
                If .TextMatrix(mouse_row, GRDC_METTRE_A_JOUR_NOM) <> "" Then
                    .ToolTipText = "Le nom trouvé dans le dictionnaire est: " _
                                & .TextMatrix(mouse_row, GRDC_METTRE_A_JOUR_NOM) & "."
                Else
                    .ToolTipText = ""
                End If
            ElseIf mouse_col = GRDC_CODE_SRV_FICH Or mouse_col = GRDC_LIB_SRV_FICH _
                  Or mouse_col = GRDC_CODE_POSTE_FICH Or mouse_col = GRDC_LIB_POSTE_FICH Then
                If .TextMatrix(mouse_row, GRDC_POSTE_EN_GRAS) = P_OUI Then
                    .ToolTipText = "Le poste trouvé dans le dictionnaire est: " _
                             & remplir_srv_poste(.TextMatrix(mouse_row, GRDC_SPM_KB_A_SYNCHRONISER), "TOOLTIP") & "."
                Else
                    .ToolTipText = ""
                End If
            ElseIf mouse_col = GRDC_PERSONNE_INACTIVE Then
                If .TextMatrix(mouse_row, GRDC_LISTE_PERSONNE_INACTIVE) = "" Then
                    .ToolTipText = ""
                Else
                    .ToolTipText = "Il y a au moins une personne inactive ou externe au fichier d'importation" _
                                  & " qui porte le nom " & .TextMatrix(mouse_row, GRDC_NOM) & "..."
                End If
            Else ' pour le reste du grid
                .ToolTipText = ""
            End If
        End If
    End With

End Sub

Private Sub grd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim str As String
    Dim mouse_row As Integer
    Dim mon_y As Single

    With grd
        If .Rows - 1 = 0 Then Exit Sub ' par précaution
        mon_y = Y
        ' Récupérer la ligne sur laquelle on a fait un click-droit
        str = (mon_y / .RowHeight(0))
        mouse_row = (mon_y / .RowHeight(0))
        On Error Resume Next
        mouse_row = IIf(left$(STR_GetChamp(str, ",", 1), 1) > 4, (mon_y / .RowHeight(0)) - 1, (mon_y / .RowHeight(0)))
        If Button = vbRightButton Then
            If .MouseCol = GRDC_ACTION Then
                Call grd_Click
                Exit Sub
            End If
            Call selectionner_ligne(.MouseRow)
            mnuActualiser.Visible = cmd(CMD_ACTUALISER).Enabled
            Select Case .TextMatrix(.Row, GRDC_ACTION)
                Case "Créer"
                    mnuAction.Caption = "Créer un compte pour " & .TextMatrix(.Row, GRDC_NOM) _
                                            & " " & .TextMatrix(.Row, GRDC_PRENOM) & "."
                Case "Désactiver"
                    mnuAction.Caption = "Rendre le compte de " & .TextMatrix(.Row, GRDC_NOM) _
                                        & " " & .TextMatrix(.Row, GRDC_PRENOM) & " inactif."
                Case "Accéder"
                    mnuAction.Caption = "Modifier les coordonnées de " & .TextMatrix(.Row, GRDC_NOM) _
                                        & " " & .TextMatrix(.Row, GRDC_PRENOM) & "."
                Case "Corriger"
                    mnuAction.Caption = "Mettre à jour les coordonnées de " & .TextMatrix(.Row, GRDC_NOM) _
                                        & " " & .TextMatrix(.Row, GRDC_PRENOM) & "."
            End Select
            g_row_context_menu = .Row
            Call PopupMenu(mnuMenuContextuel)
'            g_row_context_menu = 0
        End If
    End With

End Sub

Private Sub mnuAction_Click()

    txtPopUp.Visible = True
    txtPopUp.SetFocus
'    Call action(g_row_context_menu)

End Sub

Private Sub mnuActualiser_Click()

    Me.frmPatience.Visible = True
    Call actualiser
    Me.frmPatience.Visible = False

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
    Else ' txt(Index) n'est pas vide
        cmd(CMD_RECHERCHER).Enabled = True
    End If

End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If cmd(CMD_RECHERCHER).Enabled Then
            Call rechercher_personne
        End If
    End If

End Sub

Private Sub txtPopUp_GotFocus()
'MsgBox g_row_context_menu
    Dim j As Integer
    
    txtPopUp.Visible = False
    Call action(g_row_context_menu)

End Sub
