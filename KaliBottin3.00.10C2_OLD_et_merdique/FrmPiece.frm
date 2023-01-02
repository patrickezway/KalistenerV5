VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form PrmPiece 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
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
      Height          =   5415
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10935
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Index           =   2
         Left            =   4680
         TabIndex        =   13
         Text            =   "caché"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   10290
         Picture         =   "FrmPiece.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Supprimer la coordonnée selectionnée pour cette personne."
         Top             =   4935
         Width           =   310
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   10290
         Picture         =   "FrmPiece.frx":0447
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Ajouter une coordonnée."
         Top             =   2040
         Width           =   310
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   10
         Top             =   480
         Width           =   5450
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   0
         Top             =   1050
         Width           =   5450
      End
      Begin MSFlexGridLib.MSFlexGrid grdCoord 
         Height          =   3195
         Left            =   240
         TabIndex        =   1
         Top             =   2040
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   5636
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   8388608
         BackColorFixed  =   8454143
         ForeColorFixed  =   0
         BackColorSel    =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   4194304
         GridColorFixed  =   4194304
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   2
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ComctlLib.ImageList ImageListe 
         Left            =   9120
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   13
         ImageHeight     =   13
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   1
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FrmPiece.frx":089E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Coordonnées de la pièce"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   2505
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rattachée au service"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1545
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Intitulé"
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
         Top             =   1080
         Width           =   705
      End
   End
   Begin VB.Frame frm 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   855
      Index           =   1
      Left            =   -120
      TabIndex        =   7
      Top             =   5280
      Width           =   11055
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "FrmPiece.frx":0AF8
         Height          =   510
         Index           =   2
         Left            =   5040
         Picture         =   "FrmPiece.frx":1087
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   550
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
         Index           =   1
         Left            =   9960
         Picture         =   "FrmPiece.frx":161C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "FrmPiece.frx":1BD5
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
         Left            =   600
         Picture         =   "FrmPiece.frx":2131
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "PrmPiece"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Index des TXT
Private Const TXT_SRV = 0
Private Const TXT_NOM = 1
Private Const TXT_CACHE = 2

Private Const LBL_NOM = 1

Private Const IMG_COCHE = 1

' Constantes pour le GRID des coordonnées
Private Const NBRMAX_ROWS = 11
Private Const LARGEUR_GRID_PAR_DEFAUT = 10035
Private Const LEFT_CMD_PAR_DEFAUT = 10350

' Index des CMD
Private Const CMD_OK = 0
Private Const CMD_QUITTER = 1
Private Const CMD_SUPP = 2
Private Const CMD_PLUS = 3
Private Const CMD_MOINS = 4

'Index des colonnes du GridCoord
Private Const GRDC_ZUNUM = 0
Private Const GRDC_TYPE = 1
Private Const GRDC_VALEUR = 2
Private Const GRDC_NIVEAU = 3
Private Const GRDC_PRINCIPAL = 4
Private Const GRDC_COMMENTAIRE = 5

' Variables globales pour stocker des données utiles
Private g_position_txt_cache As String
Private g_sret As String
Private g_ancien_nom As String ' util lors d'une modification
Private g_num_srv As Long
Private g_num_piece As Long
Private g_form_active As Integer
Private g_mode_creation As Boolean ' True=Création et False=modification

Private Function afficher_coordonnee() As Integer

    Dim sql As String
    Dim I As Integer
    Dim rs As rdoResultset

    sql = "SELECT * FROM UtilCoordonnee, ZoneUtil " _
        & " WHERE UC_TypeNum=" & g_num_piece _
        & " AND UC_Type='C' AND ZU_Type='C'" _
        & " AND UC_ZUNum=ZU_Num" _
        & " ORDER BY ZU_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If

    I = 1

    With grdCoord
        ' Affichage des lignes du gridCoord
        While Not rs.EOF
            .AddItem rs("ZU_Num").Value & vbTab & rs("ZU_Libelle") & vbTab & rs("UC_Valeur") _
                    & vbTab & rs("UC_Niveau") & vbTab & vbTab & rs("UC_Comm")
            '.Row = i
            .Row = .Rows - 1
            .col = GRDC_PRINCIPAL
            If rs("UC_Principal").Value Then
                Set .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture
            Else
                Set .CellPicture = LoadPicture("")
            End If
            rs.MoveNext
            I = I + 1
        Wend

        ' Redimensionner le grid est les boutons +/- si besoin est
        If .Rows > NBRMAX_ROWS And .width = LARGEUR_GRID_PAR_DEFAUT Then
            .width = .width + 255
            cmd(CMD_PLUS).left = LEFT_CMD_PAR_DEFAUT + 255
            cmd(CMD_MOINS).left = LEFT_CMD_PAR_DEFAUT + 255
        End If
        .Enabled = IIf(.Rows - 1 = 0, False, True)
    End With
    rs.Close

    afficher_coordonnee = P_OK
    Exit Function

lab_erreur:
    afficher_coordonnee = P_ERREUR

End Function

Private Sub ajouter_coord()

    Dim sql As String, sret As String
    Dim insere_milieu As Boolean
    Dim I As Integer, col_num As Integer, nbrmax As Integer, j As Integer
    Dim num As Long
    Dim rs As rdoResultset

    I = 0
    j = 0
    num = 0
    nbrmax = 0
    insere_milieu = False

lab_debut: ' réinitialiser l'affichage
    Call CL_Init
    Call CL_InitMultiSelect(False, False) ' (selection multiple=True, retourner la ligne courante=False)
    Call CL_InitTitreHelp("Liste des types de coordonnée", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_d_fonction.htm")
    Call CL_InitTaille(0, -15)

    ' boucle SQL d'ajout dans la liste des choix
    sql = "SELECT * FROM ZoneUtil WHERE ZU_Type='C'"
    If Odbc_Select(sql, rs) = P_ERREUR Then
    End If
    While Not rs.EOF
        ' Compter si NOMBRE MAX est atteint
        nbrmax = 0
        With grdCoord
            For j = 1 To .Rows - 1
                If .TextMatrix(j, GRDC_ZUNUM) = rs("ZU_Num") Then
                    nbrmax = nbrmax + 1
                End If
            Next j
        End With
        ' AJOUT DANS LA LISTE DES CHOIX
        If nbrmax < rs("ZU_NBREMax") Then
            Call CL_AddLigne(rs("ZU_Code").Value, rs("ZU_Num").Value, rs("ZU_Libelle").Value, False)
        End If
        rs.MoveNext
    Wend
    num = 0
    rs.Close

    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("&Ajouter", "", 0, 0, 1000)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)

    ChoixListe.Show 1

    If CL_liste.retour = 2 Then ' --------------- QUITTER
        Exit Sub
    End If

    If CL_liste.retour = 1 Then ' --------------- AJOUTER
        sret = PrmTypeCoordonnees.AppelFrm(0)
        If sret <> "" Then
            num = STR_GetChamp(sret, "|", 0)
        End If
        GoTo lab_debut
    End If

    ' Enregistrer
    If CL_liste.retour = 0 Then ' --------------- ENREGISTRER
        With grdCoord
            For j = 1 To .Rows - 1
                If UCase$(CL_liste.lignes(CL_liste.pointeur).tag) < UCase$(.TextMatrix(j, GRDC_TYPE)) Then
                    .Row = j
                    .AddItem (CL_liste.lignes(I).num & vbTab), j
                    insere_milieu = True
                    Exit For
                End If
            Next j
            If Not insere_milieu Then
                .AddItem (CL_liste.lignes(I).num & vbTab)
            End If
            For I = 1 To .Rows - 2
                If .TextMatrix(I, GRDC_ZUNUM) = CL_liste.lignes(CL_liste.pointeur).num Then
                    .Row = I
                    .col = GRDC_PRINCIPAL
                    GoTo coord_existant:
                End If
            Next I
            .Row = j
            .col = GRDC_PRINCIPAL
            Set .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture
            If Not insere_milieu Then ' on insert à la fin du grid
                .Row = .Rows - 1
            Else ' on insert dans la ligne J
                .Row = j
            End If
coord_existant:     ' et est déjà coché
            .TextMatrix(j, GRDC_ZUNUM) = CL_liste.lignes(CL_liste.pointeur).num
            .TextMatrix(j, GRDC_TYPE) = CL_liste.lignes(CL_liste.pointeur).tag
            cmd(CMD_OK).Enabled = True
            If .Rows > NBRMAX_ROWS Then ' gérer la bare de défilement vertical
                If .width = LARGEUR_GRID_PAR_DEFAUT Then
                    .width = .width + 255
                    cmd(CMD_PLUS).left = LEFT_CMD_PAR_DEFAUT + 255
                    cmd(CMD_MOINS).left = LEFT_CMD_PAR_DEFAUT + 255
                Else
                    .ScrollBars = flexScrollBarVertical
                End If
            End If

            .Row = j
            .col = GRDC_VALEUR
            txt(TXT_CACHE).Text = ""
            g_position_txt_cache = ""
            .Enabled = True
        End With
    End If

End Sub

Public Function AppelFrm(ByVal v_num_piece As Long, ByVal v_num_srv As Long, _
                         ByVal v_mode_creation As Boolean) As String
' ********************************************************************************
' Retourner une chaine composée du numéro de la pièce et son nom
' g_sret = "num_piece|nom_piece"
' si création: valider=> "num_piece|nom_piece", annuler=> ""
' si modification: valider=> "num_piece|nom_piece", supprimer=> "0|", annuler=> ""
' ********************************************************************************
    g_num_piece = v_num_piece
    g_num_srv = v_num_srv
    g_mode_creation = v_mode_creation

    Me.Show 1

    AppelFrm = g_sret

End Function

Private Sub basculer_colonne_principal(ByVal v_row As Integer, ByVal v_col As Integer)

    Dim I As Integer, type_en_cours As Integer, nbr As Integer, ma_row As Integer

    With grdCoord
        ' ne pas tenter de basculer si on est sur la ligne fixe
        If v_row = 0 Then Exit Sub

        type_en_cours = .TextMatrix(v_row, GRDC_ZUNUM)
        If .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture Then
        ' le cellule elle est cochée
            .Row = v_row
            .col = v_col
            Set .CellPicture = LoadPicture("") ' on décoche cette ligne
            nbr = 0
            For I = 1 To .Rows - 1 ' on chrche le nombre de lignes du même type
                If .TextMatrix(I, GRDC_ZUNUM) = type_en_cours And I <> v_row Then
                    ma_row = I
                    nbr = nbr + 1
                End If
            Next I
            If nbr = 1 Then ' s'il y en a deux, on coche la deuxième
                .Row = ma_row
                .col = GRDC_PRINCIPAL
                Set .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture
            End If
            cmd(CMD_OK).Enabled = True
        Else ' elle n'est pas cochée
            ' vérifier s'il n'y pas d'autre principale pour le même type de coordonnée
            For I = 1 To .Rows - 1
                If .TextMatrix(I, GRDC_ZUNUM) = type_en_cours And I <> v_row Then
                    .col = GRDC_PRINCIPAL
                    .Row = I
                    If .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture Then
                        Set .CellPicture = LoadPicture("")
                    End If
                End If
            Next I
            .Row = v_row
            .col = v_col
            Set .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture
            cmd(CMD_OK).Enabled = True
        End If
    End With

End Sub


Private Sub deplacer_txt(ByVal v_direction As String, ByVal v_row As Integer, ByVal v_col As Integer)
' ********************************************************************
' Sert à déplacer le TextBox(TXT_CACHE) selon les flèches de direction
' ********************************************************************
    If v_col = GRDC_NIVEAU Then ' COLCOLCOL
        txt(TXT_CACHE).MaxLength = 1
    Else
        txt(TXT_CACHE).MaxLength = 0
    End If
    With grdCoord
        Select Case v_direction
        Case "HAUT" '   ########################### VERS LE HAUT ##########################################"
            If v_row > 1 Then ' on n'est pas au TOP du grid
                .Row = .Row - 1
                Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
            End If
        Case "BAS" '    ########################### VERS LE BAS ##########################################"
            If v_row < .Rows - 1 Then ' on n'est pas au BOTTOM du grid
                .Row = .Row + 1
                Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
            End If
        Case "DROITE" ' ########################## VERS LA DROITE ########################################"
            'If v_col <> .Cols - 1 Then ' on n'est pas sur la derniere colonne
            If v_col <> GRDC_COMMENTAIRE Then ' on n'est pas sur la derniere colonne
                .col = .col + 1
                If .col <> GRDC_PRINCIPAL Then
                    Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
                Else
                    .SetFocus
                End If
            'ElseIf v_col = .Cols - 1 Then  ' on est sur la derniere colonne
            ElseIf v_col = GRDC_COMMENTAIRE Then  ' on est sur la derniere colonne
                If v_row < .Rows - 1 Then ' on n'est pas sur la derniere ligne
                    .Row = .Row + 1
                    .col = GRDC_VALEUR
                    Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
                ElseIf v_row = .Rows - 1 Then ' on est sur la derniere ligne
                    .Row = 1
                    .col = GRDC_VALEUR
                    Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
                End If
            End If
        Case "GAUCHE" ' ########################### VERS LA GAUCHE ########################################"
            If v_col > GRDC_VALEUR Then
                .col = .col - 1
                If .col <> GRDC_PRINCIPAL Then
                    Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
                Else
                    .SetFocus
                End If
            Else
                If v_row > 1 Then
                    .Row = .Row - 1
                    .col = .Cols - 1
                    .SetFocus
                End If
            End If
        End Select
    End With

End Sub

Private Function enregistrer_coordonnees(ByVal v_num_piece As Long) As Integer
' *****************************************
' Enregistrer chaque coordonnée de la pièce
' *****************************************
    Dim uc_principal_val As Boolean
    Dim I As Integer
    Dim lng As Long

    ' Supprimer les anciennes coordonnées si mode = modification
    If Not g_mode_creation Then
        If Odbc_Delete("UtilCoordonnee", _
                       "UC_Num", _
                       "WHERE UC_TypeNum=" & v_num_piece & " AND UC_Type='C'", _
                       lng) = P_ERREUR Then
            GoTo lab_erreur
        End If
    End If
    ' L'enregistrement des coordonnées
    With grdCoord
        For I = 1 To .Rows - 1
            .Row = I
            .col = GRDC_PRINCIPAL
            ' Si la coordonnée est principale
            If .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture Then uc_principal_val = True
            If Odbc_AddNew("UtilCoordonnee", "UC_Num", "UC_Seq", False, lng, _
                           "UC_Type", "C", _
                           "UC_TypeNum", v_num_piece, _
                           "UC_ZUNum", .TextMatrix(I, GRDC_ZUNUM), _
                           "UC_Valeur", .TextMatrix(I, GRDC_VALEUR), _
                           "UC_Comm", .TextMatrix(I, GRDC_COMMENTAIRE), _
                           "UC_Niveau", .TextMatrix(I, GRDC_NIVEAU), _
                           "UC_Principal", uc_principal_val, _
                           "UC_LstPoste", "") = P_ERREUR Then
                GoTo lab_erreur
            End If
            uc_principal_val = False
        Next I
    End With

    enregistrer_coordonnees = P_OK
    Exit Function

lab_erreur:
    enregistrer_coordonnees = P_ERREUR

End Function

Private Sub initialiser()

    Dim srv_nom As String
    Dim I As Integer

    Call FRM_ResizeForm(Me, Me.width, Me.Height)
    g_sret = ""
    g_position_txt_cache = ""
    txt(TXT_CACHE).Text = ""
    If g_num_srv = 1 Then ' nom du laboratoire
        Call Odbc_RecupVal("SELECT L_Code FROM Laboratoire WHERE L_Num = 1", srv_nom)
    Else ' nom du service
        srv_nom = P_get_lib_srv_poste(g_num_srv, P_SERVICE)
    End If
    txt(TXT_SRV).Text = srv_nom
    If Not g_mode_creation Then ' modification
        Me.frm.Item(0).Caption = "Modification d'une pièce"
        g_ancien_nom = P_get_nom_piece(g_num_piece)
        txt(TXT_NOM).Text = g_ancien_nom
        cmd(CMD_SUPP).Visible = True
    Else ' création
        Me.frm.Item(0).Caption = "Création d'une nouvelle pièce"
    End If

    With grdCoord
        txt(TXT_CACHE).left = .CellLeft
        txt(TXT_CACHE).Top = .CellTop

        Me.txt(TXT_NOM).SetFocus

        grdCoord.Rows = 1
        .FormatString = "|Type de coordonnée|Valeur|Niveau de confidentialité|Principal dans son type|Commentaire"
        .RowHeight(0) = 750
        .ColWidth(GRDC_ZUNUM) = 0
        .ColWidth(GRDC_TYPE) = 2250
        .ColWidth(GRDC_VALEUR) = 2300
        .ColWidth(GRDC_NIVEAU) = 1500
        .ColWidth(GRDC_PRINCIPAL) = 1650
        .ColWidth(GRDC_COMMENTAIRE) = 2330
        .Row = 0
        For I = 0 To .Cols - 1
            .col = I
            .CellFontBold = True
            .ColAlignment(I) = 4
        Next I
        .col = GRDC_ZUNUM
    End With

    If Not g_mode_creation Then
        If afficher_coordonnee() = P_ERREUR Then Exit Sub
    End If

    cmd(CMD_OK).Enabled = False

End Sub

Private Sub positionner_txt(ByVal v_txt_left As Long, ByVal v_txt_top As Long, _
                            ByVal v_txt_width As Long, ByVal v_txt_height As Long, _
                            ByVal v_txt_text As String)

    ' Les colonnes PRINCIPAL et COMMENTAIRE ne sont pas éditables directement
    If grdCoord.col = GRDC_PRINCIPAL Or grdCoord.col = GRDC_COMMENTAIRE Then
        txt(TXT_CACHE).Visible = False
        grdCoord.SetFocus
        Exit Sub
    End If

    With txt(TXT_CACHE)
        .left = v_txt_left
        .Top = v_txt_top
        .width = v_txt_width
        .Height = v_txt_height
        .Text = v_txt_text
        .SelLength = Len(.Text)
        .ZOrder 0
        .Visible = True
        .SetFocus
    End With

End Sub

Private Sub quitter(ByVal v_bforce As Boolean)

    Dim reponse As Integer

    If v_bforce Then
        g_sret = ""
        Unload Me
        Exit Sub
    End If

    If cmd(CMD_OK).Visible And cmd(CMD_OK).Enabled Then
        reponse = MsgBox("Des modifications ont été effectuées !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", _
                          vbYesNo + vbDefaultButton2 + vbQuestion)
        If reponse = vbNo Then
            g_sret = ""
            Exit Sub
        End If
    End If

    Unload Me
    Exit Sub

End Sub

Private Sub supprimer()

    Dim reponse As String
    Dim lng As Long

    reponse = MsgBox("Etes-vous sûr de vouloir supprimer cette pièce ?", _
                      vbYesNo + vbDefaultButton2 + vbQuestion, "Attention !")
    If reponse = vbYes Then
        If Odbc_Delete("Piece", _
                       "PC_Num", _
                       "WHERE PC_Num = " & g_num_piece, _
                       lng) = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
        If Odbc_Delete("UtilCoordonnee", _
                       "UC_Num", _
                       "WHERE UC_TypeNum=" & g_num_piece & " AND UC_Type='C'", _
                       lng) = P_ERREUR Then
            Call quitter(True)
            Exit Sub
        End If
        g_sret = "0|"
        Unload Me
    End If

End Sub

Private Sub supprimer_coord()

    Dim I As Integer, row_en_cours As Integer
    Dim num_en_cours As Long

    With grdCoord
        row_en_cours = .Row
        If .Rows = 2 Then ' On a une seule ligne + ligne fixe
            .Rows = 1     ' On ne laisse que la ligne fixe
        ElseIf .Row > 0 Then ' On a plusieurs lignes
            num_en_cours = .TextMatrix(.Row, GRDC_ZUNUM)
            .col = GRDC_PRINCIPAL
            If .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture Then
                For I = 1 To .Rows - 1
                    .Row = I
                    .col = GRDC_PRINCIPAL
                    If .TextMatrix(I, GRDC_ZUNUM) = num_en_cours And I <> row_en_cours Then
                        Set .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture
                        Exit For
                    End If
                Next I
            End If
            Call .RemoveItem(row_en_cours)
            ' Remettre la taille du grid par défaut, et la disposition des boutons
            If .Rows <= NBRMAX_ROWS And .width = LARGEUR_GRID_PAR_DEFAUT + 255 Then
                .width = .width - 255
                cmd(CMD_PLUS).left = LEFT_CMD_PAR_DEFAUT
                cmd(CMD_MOINS).left = LEFT_CMD_PAR_DEFAUT
            End If
        End If
        ' il ne reste plus de coordonnées dans le tableau
        If .Rows - 1 = 0 Then .Enabled = False
    End With

    g_position_txt_cache = ""
    cmd(CMD_OK).Enabled = True

End Sub
Private Sub valider()
' **********************************************************
' Enregistrer toutes les modifications si elles sont valides
' **********************************************************
    Dim nom_piece As String
    Dim num_piece As Long

    ' ************************ VÉRIFICATIONS ************************
    If verifier_tous_champs() = P_NON Then
        Exit Sub
    End If
    ' *********************** ENREGISTREMENTS ***********************
    If Odbc_BeginTrans() = P_ERREUR Then
        Exit Sub
    End If
    
    nom_piece = txt(TXT_NOM).Text
    If Not g_mode_creation Then ' mode modification
        If g_ancien_nom <> nom_piece Then
            Call Odbc_Update("Piece", _
                             "PC_Num", _
                             " WHERE PC_Num = " & g_num_piece, _
                             "PC_Nom", nom_piece)
            g_sret = g_num_piece & "|" & nom_piece
        Else
            g_sret = "" ' on n'a pas changé de nom
        End If
        If enregistrer_coordonnees(g_num_piece) = P_ERREUR Then
            GoTo lab_erreur
        End If
    Else ' mode création
        If g_num_srv = 1 Then g_num_srv = 0 ' si c'est un labo
        Call Odbc_AddNew("Piece", _
                         "PC_Num", _
                         "PC_Seq", _
                         True, _
                         num_piece, _
                         "PC_SRVNum", g_num_srv, _
                         "PC_Nom", nom_piece)
        g_sret = num_piece & "|" & nom_piece
        If enregistrer_coordonnees(num_piece) = P_ERREUR Then
            GoTo lab_erreur
        End If
    End If

    If Odbc_CommitTrans() = P_ERREUR Then
        ' Exit et Unload Me !
    End If
    Unload Me
    Exit Sub

lab_erreur:
    Call Odbc_RollbackTrans

    Unload Me

End Sub

Private Sub valider_txt(ByVal v_row As Integer, ByVal v_col As Integer)
' *******************************************************
' Mettre le texte du TextBox dans la cellule du gridCoord
' Donner les coordonnées de la cellule suivante
' *******************************************************

    With grdCoord
        ' activer le bouton CMD_VALIDER
        If .TextMatrix(v_row, v_col) <> txt(TXT_CACHE).Text Then
            cmd(CMD_OK).Enabled = True
        End If
        .TextMatrix(.Row, .col) = txt(TXT_CACHE).Text
        ' On positionne le TextBox(TXT_CACHE) dans la bonne cellule (la suivante)
        Call deplacer_txt("DROITE", .Row, .col)

    End With

End Sub

Private Function verifier_tous_champs() As Integer

    Dim I As Integer

    ' Vérifier la validité du NOM de la pièce
    If Len(txt(TXT_NOM).Text) = 0 Then
        Call MsgBox("Le nom de la salle est une rubrique obligatoire.", vbOKOnly, "Attention !")
        lbl(LBL_NOM).ForeColor = vbRed
        txt(TXT_NOM).SetFocus
        GoTo lab_erreur
    Else
        lbl(LBL_NOM).ForeColor = vbBlack
    End If
    ' Vérifier la validité du NIVEAU de confidentialité et la VALEUR du coordonnée
    With grdCoord
        For I = 1 To .Rows - 1
            ' la VALEUR du coordonnée est-elle vide ?
            If .TextMatrix(I, GRDC_VALEUR) = "" Then
                Call MsgBox("La VALEUR du coordonnée doit être renseignée.", vbCritical + vbOKOnly, _
                            "Erreur de validation")
                .col = GRDC_VALEUR
                .Row = I
                Call positionner_txt(.left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight, _
                                     .TextMatrix(I, GRDC_VALEUR))
                GoTo lab_erreur
            End If
            ' NIVEAU de confidentialité est-il valide ?
            If .TextMatrix(I, GRDC_NIVEAU) = "" Then
                .TextMatrix(I, GRDC_NIVEAU) = "0"
            End If
            If .TextMatrix(I, GRDC_NIVEAU) > 9 Then ' Le NIVEAU doit € [0; 9]
                Call MsgBox("Le NIVEAU de confidentialité incorrect." & "Il doit être compris entre 0 (niveau bas) et 9 (niveau haut).", _
                            vbCritical + vbOKOnly, "Erreur de validation")
                .TextMatrix(I, GRDC_NIVEAU) = ""
                .col = GRDC_NIVEAU
                GoTo lab_erreur
            ElseIf Not STR_EstEntierPos(.TextMatrix(I, GRDC_NIVEAU)) Then
                Call MsgBox("Le NIVEAU de confidentialité doit être un entier positif.", vbCritical + vbOKOnly, _
                            "Erreur de validation")
                .TextMatrix(I, GRDC_NIVEAU) = ""
                .col = GRDC_NIVEAU
                .Row = I
                Call positionner_txt(.left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight, _
                                     .TextMatrix(I, GRDC_NIVEAU))
                GoTo lab_erreur
            End If
        Next I
    End With

    verifier_tous_champs = P_OUI
    Exit Function

lab_erreur:
    verifier_tous_champs = P_NON

End Function

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case CMD_OK
            Call valider
        Case CMD_SUPP
            Call supprimer
        Case CMD_QUITTER
            Call quitter(False)
        Case CMD_PLUS
            Call ajouter_coord
        Case CMD_MOINS
            Call supprimer_coord
    End Select

End Sub

Private Sub cmd_GotFocus(Index As Integer)

    txt(TXT_CACHE).Visible = False

End Sub


Private Sub cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyE And Shift = vbAltMask) Or KeyCode = vbKeyF1 Then
        KeyCode = 0
        If cmd(CMD_OK).Enabled Then Call valider
    ElseIf (KeyCode = vbKeyS And Shift = vbAltMask) Or KeyCode = vbKeyF2 Then
        KeyCode = 0
        If Not g_mode_creation Then ' modification
            Call supprimer
        End If
    ElseIf (KeyCode = vbKeyH And Shift = vbAltMask) Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_d_fonction.htm")
    ElseIf (KeyCode = vbKeyEscape) Then
        If txt(TXT_CACHE).Visible Or grdCoord.tag = "focus_oui" Then
            With grdCoord
                ' le grdCoord a le focus, alors on ne quitte pas
                txt(TXT_CACHE).Visible = False
                grdCoord.tag = "focus_non"
            End With
            KeyCode = 0
        Else ' le grdCoord n'a pas le focus, on peut quitter
            KeyCode = 0
            Call quitter(False)
        End If
    End If

End Sub

Private Sub Form_Activate()

    If g_form_active Then Exit Sub

    g_form_active = True
    Call initialiser

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyE And Shift = vbAltMask) Or KeyCode = vbKeyF1 Then
        KeyCode = 0
        If cmd(CMD_OK).Enabled Then Call valider
    ElseIf (KeyCode = vbKeyS And Shift = vbAltMask) Or KeyCode = vbKeyF2 Then
        KeyCode = 0
        If Not g_mode_creation Then ' modification
            Call supprimer
        End If
    ElseIf (KeyCode = vbKeyH And Shift = vbAltMask) Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_d_fonction.htm")
    ElseIf (KeyCode = vbKeyEscape) Then
        If txt(TXT_CACHE).Visible Or grdCoord.tag = "focus_oui" Then
            With grdCoord
                ' le grdCoord a le focus, alors on ne quitte pas
                txt(TXT_CACHE).Visible = False
                grdCoord.tag = "focus_non"
            End With
            KeyCode = 0
        Else ' le grdCoord n'a pas le focus, on peut quitter
            KeyCode = 0
            Call quitter(False)
        End If
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Call quitter(False)
    End If

End Sub

Private Sub Form_Load()

    g_form_active = False

End Sub

Private Sub grdcoord_Click()

    Dim mouse_row As Integer

    With grdCoord
        If .Rows = 1 Then Exit Sub
        mouse_row = .MouseRow
        If mouse_row = 0 Then
            .Row = 0
            .col = 0
        End If
    End With

End Sub

Private Sub grdCoord_GotFocus()

    With grdCoord
        If .Rows - 1 = 0 Then Exit Sub
        .tag = "focus_oui"
        cmd(CMD_MOINS).Enabled = True
        If .Row = 0 Then Exit Sub
        If .col = GRDC_PRINCIPAL Then
            txt(TXT_CACHE).Visible = False
        End If
    End With

End Sub


Private Sub grdCoord_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyE And Shift = vbAltMask) Or KeyCode = vbKeyF1 Then
        KeyCode = 0
        If cmd(CMD_OK).Enabled Then Call valider
    ElseIf (KeyCode = vbKeyS And Shift = vbAltMask) Or KeyCode = vbKeyF2 Then
        KeyCode = 0
        If Not g_mode_creation Then ' modification
            Call supprimer
        End If
    ElseIf (KeyCode = vbKeyH And Shift = vbAltMask) Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_d_fonction.htm")
    ElseIf (KeyCode = vbKeyEscape) Then
        If txt(TXT_CACHE).Visible Or grdCoord.tag = "focus_oui" Then
            With grdCoord
                ' le grdCoord a le focus, alors on ne quitte pas
                txt(TXT_CACHE).Visible = False
                grdCoord.tag = "focus_non"
            End With
            KeyCode = 0
        Else ' le grdCoord n'a pas le focus, on peut quitter
            KeyCode = 0
            Call quitter(False)
        End If
    End If

End Sub

Private Sub grdCoord_KeyPress(KeyAscii As Integer)

    Dim frm As Form

    With grdCoord
        Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            Call deplacer_txt("DROITE", .Row, .col)
        Case vbKeySpace
            KeyAscii = 0
            If .col = GRDC_PRINCIPAL Then
                Call basculer_colonne_principal(.Row, .col)
            ElseIf .col = GRDC_COMMENTAIRE Then
                Set frm = SaisieCommentaire
                If SaisieCommentaire.AppelFrm(grdCoord) Then
                    cmd(CMD_OK).Enabled = True
                End If
                Set frm = Nothing
            Else
                g_position_txt_cache = .Row & ";" & .col
                If .col = GRDC_NIVEAU Then ' COLCOLCOL
                    txt(TXT_CACHE).MaxLength = 1
                Else
                    txt(TXT_CACHE).MaxLength = 0
                End If
                Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
            End If
        Case vbKeyEscape
            KeyAscii = 0
            If .col = GRDC_COMMENTAIRE Then
            End If
        End Select
    End With

End Sub


Private Sub grdCoord_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim old_row As Integer, old_col As Integer, mouse_row As Integer, mouse_col As Integer
    Dim frm As Form

    If Button = MouseButtonConstants.vbRightButton Then Exit Sub

    With grdCoord
        mouse_row = .MouseRow
        mouse_col = .MouseCol
        .tag = "focus_oui"
        Select Case mouse_col
            ' ************** PRINCIPAL OU NON PRINCIPAL ***************
            Case GRDC_PRINCIPAL
                Call basculer_colonne_principal(mouse_row, mouse_col)
                txt(TXT_CACHE).Visible = False
                g_position_txt_cache = mouse_row & ";" & mouse_col
            ' ************** PRINCIPAL OU NON PRINCIPAL ***************
            Case Else
                If mouse_row > 0 And mouse_row < .Rows Then
                    If mouse_col > 1 Then
                        If mouse_col = GRDC_COMMENTAIRE Then
                            Set frm = SaisieCommentaire
                            If SaisieCommentaire.AppelFrm(grdCoord) Then
                                cmd(CMD_OK).Enabled = True
                            End If
                            Set frm = Nothing
                            Exit Sub
                        Else ' ARRET
                            If mouse_col = GRDC_NIVEAU Then ' COLCOLCOL
                                txt(TXT_CACHE).MaxLength = 1
                            Else
                                txt(TXT_CACHE).MaxLength = 0
                            End If
                            Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, _
                                                 .CellHeight, .TextMatrix(mouse_row, mouse_col))
                        End If
                    End If
                Else
                    txt(TXT_CACHE).Visible = False
                End If
        End Select
    End With

End Sub


Private Sub grdCoord_RowColChange()

    With grdCoord
        If .Row = 0 Then Exit Sub
        If grdCoord.col < GRDC_VALEUR Then
            grdCoord.col = GRDC_VALEUR
        End If
        If .col = GRDC_VALEUR Or .col = GRDC_NIVEAU Then
            If .col = GRDC_NIVEAU Then ' COLCOLCOL
                txt(TXT_CACHE).MaxLength = 1
            Else
                txt(TXT_CACHE).MaxLength = 0
            End If
            Call positionner_txt(.CellLeft + .left, .CellTop + .Top, .CellWidth, .CellHeight, .TextMatrix(.Row, .col))
        End If
    End With

End Sub


Private Sub grdCoord_Scroll()

    txt(TXT_CACHE).Visible = False

End Sub


Private Sub txt_Change(Index As Integer)

    If Index <> TXT_CACHE Then
        If Len(txt(TXT_NOM).Text) > 0 Then
            lbl(LBL_NOM).ForeColor = vbBlack
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_OK).Enabled = False
        End If
    End If

End Sub

Private Sub txt_GotFocus(Index As Integer)

    If Index <> TXT_CACHE Then
        txt(TXT_CACHE).Visible = False
    End If

End Sub


Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyE And Shift = vbAltMask) Or KeyCode = vbKeyF1 Then
        KeyCode = 0
        If cmd(CMD_OK).Enabled Then Call valider
    ElseIf (KeyCode = vbKeyS And Shift = vbAltMask) Or KeyCode = vbKeyF2 Then
        KeyCode = 0
        If Not g_mode_creation Then ' modification
            Call supprimer
        End If
    ElseIf (KeyCode = vbKeyH And Shift = vbAltMask) Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_d_fonction.htm")
    ElseIf (KeyCode = vbKeyEscape) Then
        If txt(TXT_CACHE).Visible Or grdCoord.tag = "focus_oui" Then
            With grdCoord
                ' le grdCoord a le focus, alors on ne quitte pas
                txt(TXT_CACHE).Visible = False
                grdCoord.tag = "focus_non"
            End With
            KeyCode = 0
        Else ' le grdCoord n'a pas le focus, on peut quitter
            KeyCode = 0
            Call quitter(False)
        End If
    End If

End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)

    If Index = TXT_NOM Then
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            SendKeys "{TAB}"
        ElseIf KeyAscii = vbKeyEscape Then
            KeyAscii = 0
            Call quitter(False)
        End If
    ElseIf Index = TXT_CACHE Then
        With grdCoord
            Select Case KeyAscii
            Case vbKeyReturn
                KeyAscii = 0
                Call valider_txt(.Row, .col)
                g_position_txt_cache = .Row & ";" & .col
            Case vbKeyEscape
                KeyAscii = 0
                If txt(TXT_CACHE).Visible Then
                    txt(TXT_CACHE).Text = ""
                    txt(TXT_CACHE).Visible = False
                    .tag = "focus_non"
                End If
            End Select
        End With
    End If

End Sub


Private Sub txt_LostFocus(Index As Integer)

    ' Cacher le TXT_CACHE
    If Index = TXT_CACHE Then
        txt(TXT_CACHE).Visible = False
    End If

End Sub

