VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form AnalyseFichier 
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
      Caption         =   "Analyse des associations à partir du fichier d'import "
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
      TabIndex        =   4
      Top             =   0
      Width           =   11895
      Begin VB.TextBox txtPopUp 
         Height          =   375
         Left            =   5400
         TabIndex        =   8
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
         Left            =   240
         TabIndex        =   6
         Top             =   5880
         Visible         =   0   'False
         Width           =   11535
         Begin ComctlLib.ProgressBar pgb 
            Height          =   495
            Left            =   120
            TabIndex        =   7
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
            TabIndex        =   9
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
            TabIndex        =   10
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
         Picture         =   "AnalyseFichier.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Actualiser les données du tableau"
         Top             =   840
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   6090
         Left            =   120
         TabIndex        =   0
         Top             =   1440
         Visible         =   0   'False
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   10742
         _Version        =   393216
         Rows            =   1
         Cols            =   24
         ForeColor       =   0
         BackColorFixed  =   12648447
         ForeColorFixed  =   0
         AllowUserResizing=   1
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "ici le compteur des boules"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "ici le compteur des lignes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   7320
         TabIndex        =   5
         Top             =   345
         Visible         =   0   'False
         Width           =   4095
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
            NumListImages   =   3
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "AnalyseFichier.frx":04E3
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "AnalyseFichier.frx":08A9
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "AnalyseFichier.frx":0C6F
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
      TabIndex        =   3
      Top             =   7580
      Width           =   11895
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Associer dans Structure "
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
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Associer en parcourant la structure"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   3555
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rechercher"
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
         Index           =   3
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Rechercher pour Associer"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   3555
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
         Picture         =   "AnalyseFichier.frx":1035
         Style           =   1  'Graphical
         TabIndex        =   1
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
      Begin VB.Menu mnuActionP 
         Caption         =   "Associer"
         Begin VB.Menu mnuAction 
            Caption         =   "une action à faire"
         End
         Begin VB.Menu mnurechercher 
            Caption         =   "Rechercher dans KaliWeb"
         End
      End
   End
End
Attribute VB_Name = "AnalyseFichier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Index des colonnes du grid
Private Const GRDC_SYNC_NUM = 0
Private Const GRDC_SYNC_PASTILLE = 1
Private Const GRDC_SYNC_EMPLOI_SECTION = 2
Private Const GRDC_SYNC_LIB_EMPLOI = 3
Private Const GRDC_SYNC_LIB_SECTION = 4
Private Const GRDC_SYNC_SERVICE = 5
Private Const GRDC_SYNC_POSTE = 6
Private Const GRDC_SYNC_PONUM = 7
Private Const GRDC_SYNC_LST_SRV_PONUM = 8
Private Const GRDC_SYNC_AUTO = 9

' Index des FRAMES
Private Const FRM_PRINCIPALE = 0
Private Const FRM_CMD_QUITTER = 1

' Index des couleurs
Private Const COLOR_DESACTIVE = &HE0E0E0         ' GRIS
Private Const COLOR_COLONNE_TRIEE = &HFFFF&      ' JAUNE FONCE &H0000C0C0&
Private Const COLOR_COLONNE_NON_TRIEE = &HC0FFFF ' couleur de la ligne fixe par défaut

' Index des IMAGES
Private Const IMG_PASTILLE_PB = 1
Private Const IMG_PASTILLE_VERTE = 2
Private Const IMG_PASTILLE_ROUGE = 3

' Index des positions des images
Private Const POS_GAUCHE = flexAlignLeftCenter
Private Const POS_CENTRE = flexAlignCenterCenter
Private Const POS_DROITE = flexAlignRightCenter

' Constatantes de séparation
Private Const SEPARATEUR_SERVICE_POSTE = " <=> "

' Index des CMD
Private Const CMD_QUITTER = 0
Private Const CMD_ASSOCIER = 1
Private Const CMD_ACTUALISER = 2
Private Const CMD_RECHERCHER = 3

' Index des LIBELLES
Private Const LBL_COMPTEUR = 3
Private Const LBL_COMPTEUR_BOULES = 0

Private Const PRM_OUI = 1
Private Const PRM_NON = 2

Private g_form_active As Boolean
' Sens des tris
Private g_col_tri As Integer
Private g_sens_tri As Integer

Private g_actualiser As Boolean
Private g_row_context_menu As Integer
Private g_nbr_lignes As Long
Private g_nbRouge As Integer, g_nbVerte As Integer, g_nbPb As Integer

Private g_poste As String
Private g_GRDC_CODE_SRV_FICH As String
Private g_GRDC_CODE_POSTE_FICH As String
Private g_GRDC_LIB_SRV_FICH As String
Private g_GRDC_LIB_POSTE_FICH As String

Private Sub action(ByVal v_row As Integer)

    Dim frm As Form
    Dim spo_num As String
    Dim I As Integer, nb As Integer
    Dim ret As Integer, iImg As Integer
    Dim nbSynchro As Integer
    Dim section As String
    Dim emploi As String
    Dim titre As String
    Dim sret As String
    Dim s As String, lib_service As String, lib_poste As String
    Dim Cas As Integer, po_num As Integer
    
    If v_row = 0 Then Exit Sub
    With grd
        emploi = STR_GetChamp(.TextMatrix(v_row, GRDC_SYNC_EMPLOI_SECTION), "|", 0)
        section = STR_GetChamp(.TextMatrix(v_row, GRDC_SYNC_EMPLOI_SECTION), "|", 1)
        titre = .TextMatrix(v_row, GRDC_SYNC_LIB_EMPLOI) & " - " & .TextMatrix(v_row, GRDC_SYNC_LIB_SECTION)
        If Me.txtPopUp.tag = 0 Then
            ' créer une association
            Cas = 1
            sret = ImportationAnnuaire.Liste_Assoc(emploi, section, titre, "New")
            GoTo Lab_Fait
        ElseIf Me.txtPopUp.tag > 0 Then
            ' afficher les association
            Cas = 2
            sret = ImportationAnnuaire.Liste_Assoc(emploi, section, titre, "")
Lab_Fait:
            If sret = "" Then
                ' pastille rouge
                nbSynchro = 0
                iImg = IMG_PASTILLE_ROUGE
            ElseIf InStr(sret, "|") = 0 Then
                ' pastille verte
                nbSynchro = 1
                iImg = IMG_PASTILLE_VERTE
            Else
                ' pastille Pb
                nbSynchro = 2
                iImg = IMG_PASTILLE_PB
            End If
            '
            If Cas = 1 Then
                ' Création
                If nbSynchro > 0 Then
                    Call actualiser_compteur("MOINS")
                    cmd(CMD_ACTUALISER).Enabled = True
                    cmd(CMD_ACTUALISER).Visible = True
                End If
            ElseIf Cas = 2 Then
                ' Modification
                If nbSynchro = 0 Then
                    Call actualiser_compteur("PLUS")
                    cmd(CMD_ACTUALISER).Enabled = True
                    cmd(CMD_ACTUALISER).Visible = True
                End If
            End If
            '
            .col = GRDC_SYNC_PASTILLE
            .Row = v_row
            Set .CellPicture = imglst.ListImages(iImg).Picture
            ' le reste
            If sret = "" Then
                .TextMatrix(.Row, GRDC_SYNC_PONUM) = ""
                grd.TextMatrix(.Row, GRDC_SYNC_SERVICE) = ""
                grd.TextMatrix(.Row, GRDC_SYNC_POSTE) = ""
                grd.TextMatrix(.Row, GRDC_SYNC_LST_SRV_PONUM) = ""
            Else
                .TextMatrix(.Row, GRDC_SYNC_PONUM) = sret
                spo_num = STR_GetChamp(sret, "|", 0)
                po_num = STR_GetChamp(spo_num, "!", 0)
                s = recup_PSLib(val(po_num), lib_service, lib_poste)
                grd.TextMatrix(.Row, GRDC_SYNC_SERVICE) = lib_service & IIf(nbSynchro > 1, " => ...... (" & nbSynchro & ")", "")
                grd.TextMatrix(.Row, GRDC_SYNC_POSTE) = lib_poste
                grd.TextMatrix(.Row, GRDC_SYNC_LST_SRV_PONUM) = sret
            End If
        End If
    End With

    g_row_context_menu = 0
    grd.SetFocus

End Sub

Private Sub actualiser()
'*****************************************************************************
' Actualiser le grid après les mouvements: créaion/modifications/désactivation
'*****************************************************************************

    grd.Visible = False
    g_actualiser = True
    Sleep (100)
    Call initialiser
    cmd(CMD_ACTUALISER).Enabled = False
    cmd(CMD_ACTUALISER).Visible = False
End Sub

Private Sub actualiser_compteur(v_trait)
    If v_trait = "PLUS" Then
        g_nbr_lignes = g_nbr_lignes + 1
    ElseIf v_trait = "MOINS" Then
        g_nbr_lignes = g_nbr_lignes - 1
    Else
        'MsgBox "actualiser_compteur v_trait=" & v_trait & " ?"
    End If
    lbl(LBL_COMPTEUR).Caption = g_nbr_lignes & " associations à traiter"
End Sub

Public Function AppelFrm() As Integer

    Call FRM_ResizeForm(Me, Me.width, Me.Height)

    g_form_active = False

    Me.Show 1

    AppelFrm = 0

    SendKeys "%{A}"

End Function

Private Function recup_PSLib(ByVal v_numposte As Long, ByRef r_lib_service, ByRef r_lib_poste) As String
    
    Dim s As String, lib As String, sql As String
    Dim numsrv As Long
    Dim rs As rdoResultset
    
    r_lib_service = ""
    r_lib_poste = ""
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
    r_lib_service = rs("FT_Libelle").Value
    numsrv = rs("PO_SRVNum").Value
    rs.Close
    sql = "select SRV_Nom from Service" _
        & " where SRV_Num=" & numsrv
    If Odbc_SelectV(sql, rs) = P_OK Then
        If Not rs.EOF Then
            r_lib_poste = rs("SRV_Nom").Value
        End If
        rs.Close
    End If

End Function

Private Sub initialiser()

    Dim sql As String
    Dim I As Integer
    Dim lnb As Long
    
    ' Suppression ds synchro des associations qui n'existent plus
    sql = "delete from synchro where sync_spnum not in (select po_num from poste)"
    Call Odbc_Cnx.Execute(sql)
    sql = "delete from synchro where sync_spnum in (select po_num from poste where po_actif='f')"
    Call Odbc_Cnx.Execute(sql)
    
    Call P_InitFichierImportation
           
    frm(FRM_PRINCIPALE).Caption = "Analyse des associations à partir du fichier : " _
                            & p_nom_fichier_importation & " " & IIf(p_est_sur_serveur, "(serveur)", "(local)")
    If Not g_actualiser Then
        frmPatience.left = (Me.width / 2) - (frmPatience.width / 2)
        frmPatience.Top = (Me.Height / 2) - (frmPatience.Height / 2)
        frmPatience.Visible = True
    End If
    g_row_context_menu = 0

    cmd(CMD_ACTUALISER).Visible = False
    cmd(CMD_ACTUALISER).Enabled = False
    With grd
        .Rows = 1
        .FormatString = "sync_num|||EMPLOI|SECTION|SERVICE|POSTE|"
        .ScrollTrack = True
        .SelectionMode = flexSelectionByRow
        .ColWidth(GRDC_SYNC_NUM) = 0
        .ColWidth(GRDC_SYNC_PASTILLE) = 245
        .ColWidth(GRDC_SYNC_EMPLOI_SECTION) = 0
        .ColWidth(GRDC_SYNC_LIB_EMPLOI) = 2700
        .ColWidth(GRDC_SYNC_LIB_SECTION) = 2600
        .ColWidth(GRDC_SYNC_SERVICE) = 2600
        .ColWidth(GRDC_SYNC_POSTE) = 2600
        .ColWidth(GRDC_SYNC_PONUM) = 0
        .ColWidth(GRDC_SYNC_LST_SRV_PONUM) = 0
        .ColWidth(GRDC_SYNC_AUTO) = 500
        .ColAlignment(GRDC_SYNC_PASTILLE) = POS_CENTRE
        .ColAlignment(GRDC_SYNC_LIB_SECTION) = POS_GAUCHE
        .ColAlignment(GRDC_SYNC_LIB_EMPLOI) = POS_GAUCHE
        .ColAlignment(GRDC_SYNC_SERVICE) = POS_GAUCHE
        .ColAlignment(GRDC_SYNC_POSTE) = POS_GAUCHE
        .Row = 0
        For I = 0 To .Cols - 1
            .col = I
            .CellFontBold = True
        Next I
    End With

    Me.Refresh

    Call remplir_grid

    frmPatience.Visible = False
    
    'g_nbr_lignes = grd.Rows - 1
    Call actualiser_compteur("")
    lbl(LBL_COMPTEUR).Visible = True


    On Error Resume Next
    grd.SetFocus
    On Error GoTo 0
    
End Sub

Private Sub quitter(ByVal v_s As String)
    Unload Me
End Sub


Private Sub remplir_grid()

    Dim sinfo As String, s As String, sext As String
    Dim bDéjà As Boolean
    Dim lst_ponum As String
    Dim nomfich As String, ligne_lu As String
    Dim I As Integer, fd As Integer, pos As Integer, iBoucle As Integer
    Dim rs As rdoResultset
    Dim sql As String
    Dim num_unique As Long
    Dim emploi As String, section As String
    Dim emploi_section As String, lib_emploi As String, lib_section As String
    Dim lib_service As String, lib_poste As String, po_num As String
    Dim nbSynchro As Integer
    Dim iImg As Integer

    cmd(CMD_RECHERCHER).Visible = False
    cmd(CMD_ASSOCIER).Visible = False
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
        Me.LbGauge2.Caption = "Remplir le tableau des emplois et services GRH"
        Me.Refresh
    Wend
    Close #fd
    
    ' Ouverture du fichier en lecture seule
    If FICH_OuvrirFichier(nomfich, FICH_LECTURE, fd) = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    
    g_nbr_lignes = 0
    num_unique = 1
    While Not EOF(fd)
        If pgb.Value = pgb.Max Then
            pgb.Value = 0
        End If
        pgb.Value = pgb.Value + 1
        pgb2.Value = pgb2.Value + 1
        Line Input #fd, ligne_lu
        If ligne_lu <> "" Then
            emploi = P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_code_emploi, p_long_code_emploi, "code emploi")
            section = P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_code_section, p_long_code_section, "code section")
            ' voir si pas déjà dans grid
            bDéjà = False
            For I = 1 To grd.Rows - 1
                If grd.TextMatrix(I, GRDC_SYNC_EMPLOI_SECTION) = emploi & "|" & section Then
                    bDéjà = True
                    Exit For
                End If
            Next I
            If Not bDéjà Then
                ' le mettre
                grd.AddItem ""
                nbSynchro = VoirSiDansSynchro(emploi, section, lst_ponum)
                lib_emploi = P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_lib_emploi, p_long_lib_emploi, "libellé emploi")
                lib_section = P_lire_valeur(p_type_fichier, ligne_lu, p_separateur, p_pos_lib_section, p_long_lib_section, "libellé section")
                If InStr(lib_emploi, "ERREUR :") > 0 Then
                    ' pastille Pb
                    iImg = IMG_PASTILLE_PB
                    g_nbPb = g_nbPb + 1
                ElseIf nbSynchro = 0 Then
                    ' pastille rouge
                    iImg = IMG_PASTILLE_ROUGE
                    g_nbr_lignes = g_nbr_lignes + 1
                    g_nbRouge = g_nbRouge + 1
                ElseIf nbSynchro = 1 Then
                    ' pastille verte
                    iImg = IMG_PASTILLE_VERTE
                    g_nbVerte = g_nbVerte + 1
                Else
                    ' pastille Pb
                    iImg = IMG_PASTILLE_PB
                    g_nbPb = g_nbPb + 1
                End If
                grd.col = GRDC_SYNC_PASTILLE
                grd.Row = grd.Rows - 1
                I = grd.Row
                Set grd.CellPicture = imglst.ListImages(iImg).Picture
                grd.TextMatrix(I, GRDC_SYNC_PONUM) = lst_ponum
                grd.TextMatrix(I, GRDC_SYNC_EMPLOI_SECTION) = emploi & "|" & section
                grd.TextMatrix(I, GRDC_SYNC_LIB_EMPLOI) = emploi & " (" & lib_emploi & ")"
                grd.TextMatrix(I, GRDC_SYNC_LIB_SECTION) = section & " (" & lib_section & ")"
                If nbSynchro = 1 Then
                    If STR_GetChamp(UCase(lst_ponum), "!", 1) = "TRUE" Then
                        grd.TextMatrix(I, GRDC_SYNC_AUTO) = "Auto."
                    End If
                End If
                If nbSynchro > 0 Then
                    po_num = STR_GetChamp(lst_ponum, "|", 0)
                    s = recup_PSLib(val(po_num), lib_service, lib_poste)
                    grd.TextMatrix(I, GRDC_SYNC_SERVICE) = lib_service & IIf(nbSynchro > 1, " => ...... (" & nbSynchro & ")", "")
                    grd.TextMatrix(I, GRDC_SYNC_POSTE) = lib_poste
                    grd.TextMatrix(I, GRDC_SYNC_LST_SRV_PONUM) = lst_ponum
                End If
            End If
        End If
    Wend ' fin de lecture du fichier
        
    With grd
        ' Trier le grid
        .Row = 0
        .col = GRDC_SYNC_POSTE
        .Sort = 1
        g_sens_tri = 1
        g_col_tri = GRDC_SYNC_POSTE
        .CellBackColor = COLOR_COLONNE_TRIEE
        ' Selection de la première ligne (si elle existe)
        .col = 0
        .ColSel = .Cols - 1
        .Row = 1
        .RowSel = 1
    End With
    
    ' Fermer le fichier
    Close #fd
    If p_est_sur_serveur Then
        Call FICH_EffacerFichier(nomfich, False)
    End If
    
    If grd.Rows - 1 = 0 Then
        grd.Enabled = False
    End If
    grd.Visible = True
    lbl(LBL_COMPTEUR_BOULES).Visible = True
    lbl(LBL_COMPTEUR_BOULES).Caption = g_nbRouge & " rouges / " & g_nbVerte & " vertes / " & g_nbPb & " problèmes"
    cmd(CMD_RECHERCHER).Visible = True
    cmd(CMD_ASSOCIER).Visible = True


End Sub

Private Function VoirSiDansSynchro(emploi, section, ByRef r_lst_ponum)
    Dim sql As String, rs As rdoResultset
    
    r_lst_ponum = ""
    sql = "SELECT * FROM Synchro WHERE SYNC_Section='" & section & "' and SYNC_Emploi='" & emploi & "'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        VoirSiDansSynchro = 0
    End If
    If rs.EOF Then
        VoirSiDansSynchro = 0
    Else
        VoirSiDansSynchro = 0
        While Not rs.EOF
            VoirSiDansSynchro = VoirSiDansSynchro + 1
            r_lst_ponum = r_lst_ponum & IIf(r_lst_ponum = "", "", "|") & rs("SYNC_SPNum") & "!" & rs("SYNC_Auto")
            rs.MoveNext
        Wend
    End If
End Function

Private Sub selectionner_ligne(ByVal v_row As Integer)
' *******************************
' Mettre la ligne en surbrillance
' *******************************
    With grd
        .col = GRDC_SYNC_NUM
        .Row = v_row
        .ColSel = .Cols - 1
        .RowSel = v_row
    End With

End Sub

Private Sub cmd_Click(Index As Integer)
    
    Dim sPoste As String, sService As String, spm_choisi As String
    Dim I As Integer
    
    Select Case Index
        Case CMD_QUITTER
            Call quitter("")
        Case CMD_ACTUALISER
            Call actualiser
        Case CMD_RECHERCHER
            Call mnurechercher_Click
        Case CMD_ASSOCIER
            Me.txtPopUp.tag = 0
            Call mnuAction_Click
    End Select
End Sub


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
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Unload Me
    End If
End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)

    With grd
        If KeyCode = vbKeyEscape Then
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
        If True Or Button = vbRightButton Then
            Call selectionner_ligne(.MouseRow)
            mnuActualiser.Visible = cmd(CMD_ACTUALISER).Enabled
            ' si pastille rouge
            If .TextMatrix(.Row, GRDC_SYNC_LST_SRV_PONUM) = "" Then
                mnuAction.Caption = "Définir l'association (choix dans les services-postes)"
                Me.txtPopUp.tag = 0
            ElseIf InStr(.TextMatrix(.Row, GRDC_SYNC_LST_SRV_PONUM), "|") > 0 Then
                mnuAction.Caption = "Modifier les associations"
                Me.txtPopUp.tag = 2
            Else
                mnuAction.Caption = "Modifier l'association"
                Me.txtPopUp.tag = 1
            End If
            g_row_context_menu = .Row
            Call PopupMenu(mnuMenuContextuel)
        End If
    End With

End Sub

Private Sub mnuAction_Click()

    txtPopUp.Visible = True
    txtPopUp.SetFocus

End Sub

Private Sub mnuActualiser_Click()

    Me.frmPatience.Visible = True
    Call actualiser
    Me.frmPatience.Visible = False

End Sub

Private Sub mnurechercher_Click()
    Dim sret As String
    Dim I As Integer
    Dim mot1 As String, mot2 As String
    Dim sql As String, rs As rdoResultset
    Dim mot As String
    Dim section As String, emploi As String
    Dim po_num As Long
    
Lab_Debut:
    sret = InputBox("vous cherchez ?", "Recherche", sret)
    If sret = "" Then
        Exit Sub
    Else
        mot1 = STR_GetChamp(sret, " ", 0)
        mot2 = STR_GetChamp(sret, " ", 1)
        
        sql = "select * from service,poste,fcttrav where po_ftnum=ft_num and po_srvnum=srv_num"
        For I = 1 To 2
            If I = 1 Then mot = mot1
            If I = 2 Then mot = mot2
            If mot <> "" Then
                sql = sql & " and ("
                sql = sql & "upper(ft_libelle) like '%" & UCase(mot) & "%'"
                sql = sql & " or upper(srv_nom) like '%" & UCase(mot) & "%'"
                sql = sql & " or upper(srv_code) like '%" & UCase(mot) & "%'"
                sql = sql & " or upper(srv_libcourt) like '%" & UCase(mot) & "%'"
                sql = sql & " or upper(po_libelle) like '%" & UCase(mot) & "%'"
                sql = sql & " or upper(ft_libelle) like '%" & UCase(mot) & "%'"
                sql = sql & ")"
            End If
        Next I
        ' MsgBox sql
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            MsgBox "SQL : " & sql
            Exit Sub
        End If
        If rs.EOF Then
            MsgBox "Aucun trouvé"
            GoTo Lab_Debut
        Else
            Call CL_Init
            Call CL_InitTitreHelp("Liste des postes trouvés", "")
            Call CL_InitTaille(0, -15)

            While Not rs.EOF
                Call CL_AddLigne(rs("po_libelle").Value & " - " & rs("srv_nom").Value, rs("PO_Num").Value, rs("PO_Num").Value, False)
                rs.MoveNext
            Wend

            Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
            Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
        
            ChoixListe.Show 1
            ' ******** QUITTER ********
            If CL_liste.retour = 0 Then
                section = STR_GetChamp(grd.TextMatrix(g_row_context_menu, GRDC_SYNC_EMPLOI_SECTION), "|", 0)
                emploi = STR_GetChamp(grd.TextMatrix(g_row_context_menu, GRDC_SYNC_EMPLOI_SECTION), "|", 1)
                po_num = CL_liste.lignes(CL_liste.pointeur).tag
                If section = "" Or emploi = "" Then
                    Exit Sub
                End If
                sql = "Insert into Synchro "
                sql = sql & "(SYNC_Emploi, SYNC_Section, SYNC_SPNum, SYNC_Auto)"
                sql = sql & "Values ('" & section & "', '" & emploi & "', " & po_num & ", 't')"
                Call Odbc_Cnx.Execute(sql)
                txtPopUp.Visible = False
                Me.txtPopUp.tag = po_num
                Call action(g_row_context_menu)
                Exit Sub
            End If
        End If
    End If
    GoTo Lab_Debut
End Sub

Private Sub txtPopUp_GotFocus()
    txtPopUp.Visible = False
    Call action(g_row_context_menu)
End Sub

