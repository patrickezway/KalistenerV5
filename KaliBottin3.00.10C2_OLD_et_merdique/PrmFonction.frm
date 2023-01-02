VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form PrmFonction 
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   5010
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11295
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   1
         ItemData        =   "PrmFonction.frx":0000
         Left            =   5520
         List            =   "PrmFonction.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2400
         Width           =   2625
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1320
         Width           =   3135
      End
      Begin VB.CommandButton cmd 
         Height          =   300
         Index           =   5
         Left            =   6870
         Picture         =   "PrmFonction.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1320
         Width           =   300
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         ItemData        =   "PrmFonction.frx":045B
         Left            =   2565
         List            =   "PrmFonction.frx":045D
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1950
         Width           =   2265
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
         Height          =   330
         Index           =   3
         Left            =   10425
         Picture         =   "PrmFonction.frx":045F
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Ajouter une coordonnée."
         Top             =   3075
         Width           =   360
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
         Left            =   10425
         Picture         =   "PrmFonction.frx":08B6
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Supprimer la coordonnée selectionnée pour cette personne."
         Top             =   4515
         Width           =   360
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Index           =   1
         Left            =   8520
         TabIndex        =   9
         Text            =   "caché"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   0
         Top             =   810
         Width           =   5685
      End
      Begin MSFlexGridLib.MSFlexGrid grdCoord 
         Height          =   1755
         Left            =   375
         TabIndex        =   1
         Top             =   3075
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   3096
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
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lors de l'import, cette fonction sera positionnée au niveau"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   18
         Top             =   2400
         Width           =   4980
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Visible dans l'annuaire"
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
         Left            =   375
         TabIndex        =   14
         Top             =   2010
         Width           =   2100
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Coordonnées"
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
         Left            =   360
         TabIndex        =   12
         Top             =   2715
         Width           =   2505
      End
      Begin ComctlLib.ImageList ImageListe 
         Left            =   8880
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   13
         ImageHeight     =   13
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   2
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmFonction.frx":0CFD
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmFonction.frx":104F
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Niveau de coordonnées accessibles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   1410
         Width           =   3225
      End
      Begin VB.Label lbl 
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
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   705
      End
   End
   Begin VB.Frame frm 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   4890
      Width           =   11295
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmFonction.frx":13A1
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
         Left            =   510
         Picture         =   "PrmFonction.frx":18FD
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   230
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
         Left            =   9840
         Picture         =   "PrmFonction.frx":1E66
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmFonction.frx":241F
         Height          =   510
         Index           =   2
         Left            =   5160
         Picture         =   "PrmFonction.frx":29AE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "PrmFonction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Entree : Rien en param direct / 0 pour création d'ailleurs
' Sortie : ZNum|Nom si création d'ailleurs

' Index des objets cmd
Private Const CMD_OK = 0
Private Const CMD_DETRUIRE = 2
Private Const CMD_QUITTER = 1
Private Const CMD_PLUS = 3
Private Const CMD_MOINS = 4
Private Const CMD_NIVEAU = 5

' Index des objets txt
Private Const TXT_NOM = 0
Private Const TXT_CACHE = 1
Private Const TXT_NIVEAU = 2
' libellés
Private Const LBL_COORDONNEES = 2
Private Const LBL_VISIBILITE = 3

' Les combobox
Private Const CBO_VISIBILITE = 0
' Les indices du Combo/ListBox la visibilité de l'affichage des services
Private Const CBO_VISIBILITE_JAMAIS = "Jamais"
Private Const CBO_VISIBILITE_VUE_DETAILLEE = "Vue détaillée seulement"
Private Const CBO_VISIBILITE_TOUJOURS = "Toujours"

' Pour modifier la position d'une fonction lors de l'import
Private Const CBO_MODIFIER_POSITION = 1
Private Const LBL_MODIFIER_POSITION = 7

'Index des colonnes du GridCoord
Private Const GRDC_ZUNUM = 0
Private Const GRDC_TYPE = 1
Private Const GRDC_VALEUR = 2
Private Const GRDC_NIVEAU = 3
Private Const GRDC_PRINCIPAL = 4
Private Const GRDC_COMMENTAIRE = 5

' Index des FRAMES
Private Const FRM_HAUT = 0
Private Const FRM_BAS = 1

Private Const IMG_PASCOCHE = 1
Private Const IMG_COCHE = 2

' Constantes pour le GRID des coordonnées
Private Const NBRMAX_ROWS = 5

' No fction en saisie (0 si nouveau)
Private g_numfct As Long
Private g_numposte As Long

Private g_stype As String
Private g_mode_direct As Boolean
Private g_sret As String
Private g_crfct_autor As Boolean

Private g_largeur_grid_init As Long
Private g_left_cmd_init As Long

' Indique si la forme a déjà été activée
Private g_form_active As Boolean

' Indique si la saisie est en-cours
Private g_mode_saisie As Boolean

' Stocke le texte avant modif pour gérer le changement
Private g_txt_avant As String

' Stocke le texte avant modif pour gérer le changement
Private g_txt_avant_nom As String
Private g_txt_avant_niveau As String

' Sotcke la position du TXT_CACHE
Private g_position_txt_cache As String
' la visibilité et le niveau de la fonction
Private g_old_visibilite As String

Public Function AppelFrm(ByVal v_stype As String, _
                         ByVal v_num As Long) As String
    
    g_stype = v_stype
    If g_stype = "P" Then
        g_mode_direct = True
        g_numposte = v_num
        g_numfct = P_get_num_fct(v_num)
    Else
        If v_num = -1 Then
            g_mode_direct = False
        Else
            g_mode_direct = True
        End If
        g_numfct = v_num
        g_numposte = 0
    End If
    
    Me.Show 1

    AppelFrm = g_sret

End Function

Private Sub ajouter_coord()

    Dim sql As String, sret As String
    Dim insere_milieu As Boolean
    Dim i As Integer, col_num As Integer, nbrmax As Integer, j As Integer
    Dim num As Long
    Dim rs As rdoResultset

    i = 0
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
    Call CL_AddBouton("&Ajouter un type de coordonnée", "", 0, 0, 2000)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)

    ChoixListe.Show 1
    ' ******** QUITTER ********
    If CL_liste.retour = 2 Then
        Exit Sub
    End If
    ' ******** AJOUTER ********
    If CL_liste.retour = 1 Then
        sret = PrmTypeCoordonnees.AppelFrm(0)
        If sret <> "" Then
            num = STR_GetChamp(sret, "|", 0)
        End If
        GoTo lab_debut
    End If
    ' ******** ENREGISTRER ********
    If CL_liste.retour = 0 Then
        With grdCoord
            For j = 1 To .Rows - 1
                If UCase$(CL_liste.lignes(CL_liste.pointeur).tag) < UCase$(.TextMatrix(j, GRDC_TYPE)) Then
                    .Row = j
                    .AddItem (CL_liste.lignes(i).num & vbTab), j
                    .col = GRDC_PRINCIPAL
                    .CellPictureAlignment = 4
                    Set .CellPicture = ImageListe.ListImages(IMG_PASCOCHE).Picture
                    insere_milieu = True
                    Exit For
                End If
            Next j
            If Not insere_milieu Then
                .AddItem (CL_liste.lignes(i).num & vbTab)
                .Row = .Rows - 1
                .col = GRDC_PRINCIPAL
                .CellPictureAlignment = 4
                Set .CellPicture = ImageListe.ListImages(IMG_PASCOCHE).Picture
            End If
            For i = 1 To .Rows - 2
                If .TextMatrix(i, GRDC_ZUNUM) = CL_liste.lignes(CL_liste.pointeur).num Then
                    .Row = i
                    .col = GRDC_PRINCIPAL
                    GoTo coord_existant:
                End If
            Next i
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
                If .width = g_largeur_grid_init Then
                    .width = .width + 255
                    cmd(CMD_PLUS).left = g_left_cmd_init + 255
                    cmd(CMD_MOINS).left = g_left_cmd_init + 255
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

Private Function afficher_coordonnee() As Integer

    Dim sql As String
    Dim i As Integer
    Dim rs As rdoResultset

    sql = "SELECT * FROM UtilCoordonnee, ZoneUtil " _
        & " WHERE UC_TypeNum=" & g_numposte _
        & " AND UC_Type='P' AND ZU_Type='C'" _
        & " AND UC_ZUNum=ZU_Num" _
        & " ORDER BY ZU_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If

    i = 1

    With grdCoord
        ' Affichage des lignes du gridCoord
        While Not rs.EOF
            .AddItem rs("ZU_Num").Value & vbTab & rs("ZU_Libelle") & vbTab & rs("UC_Valeur") _
                    & vbTab & rs("UC_Niveau") & vbTab & vbTab & rs("UC_Comm")
            '.Row = i
            .Row = .Rows - 1
            .col = GRDC_PRINCIPAL
            .CellPictureAlignment = 4
            If rs("UC_Principal").Value Then
                Set .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture
            Else
                Set .CellPicture = ImageListe.ListImages(IMG_PASCOCHE).Picture
            End If
            rs.MoveNext
            i = i + 1
        Wend

        ' Redimensionner le grid est les boutons +/- si besoin est
        If .Rows > NBRMAX_ROWS And .width = g_largeur_grid_init Then
            .width = .width + 255
            cmd(CMD_PLUS).left = g_left_cmd_init + 255
            cmd(CMD_MOINS).left = g_left_cmd_init + 255
        End If
        .Enabled = IIf(.Rows - 1 = 0, False, True)
    End With
    rs.Close

    afficher_coordonnee = P_OK
    Exit Function

lab_erreur:
    afficher_coordonnee = P_ERREUR

End Function


Private Function afficher_fct(ByVal v_numfct As Long) As Integer

    Dim sql As String, libserv As String, ft_visibilite As String, fniv_libelle As String
    Dim rs As rdoResultset
    Dim i As Integer
    Dim FT_NivRemplace As Long

    Call FRM_ResizeForm(Me, Me.width, Me.Height)

    If p_appli_kalibottin > 0 Then
        ' Position de la fonction
        If charge_niveau() > 0 Then
            cbo(CBO_MODIFIER_POSITION).Visible = True
            lbl(LBL_MODIFIER_POSITION).Visible = True
            cbo(CBO_MODIFIER_POSITION).ListIndex = 0
        Else
            cbo(CBO_MODIFIER_POSITION).Visible = False
            lbl(LBL_MODIFIER_POSITION).Visible = False
        End If
    End If
    
    g_numfct = v_numfct
    If v_numfct > 0 Then ' Fonction existante à afficher ou à modifier
        sql = "SELECT * FROM FctTrav WHERE FT_Num=" & v_numfct
        If Odbc_Select(sql, rs) = P_ERREUR Then
            GoTo lab_erreur
        End If
        If g_stype = "P" Then
            sql = "SELECT SRV_Nom FROM Poste, Service" _
                & " WHERE PO_Num=" & g_numposte _
                & " and SRV_Num=PO_SRVNum"
            If Odbc_RecupVal(sql, libserv) = P_ERREUR Then
                GoTo lab_erreur
            End If
            frm(FRM_HAUT).Caption = "Poste: " & rs("FT_Libelle").Value & " - Service: " & libserv
            ft_visibilite = rs("FT_Visible").Value
            cbo(CBO_VISIBILITE).ListIndex = CInt(ft_visibilite)
        Else
            frm(FRM_HAUT).Caption = "Fonction: " & rs("FT_Libelle").Value
            FT_NivRemplace = rs("FT_NivRemplace")
            ft_visibilite = rs("FT_Visible").Value
            g_old_visibilite = ft_visibilite
            cbo(CBO_VISIBILITE).ListIndex = CInt(ft_visibilite)
        End If
        txt(TXT_NOM).Text = rs("FT_Libelle").Value
        cmd(CMD_DETRUIRE).Visible = True
        cmd(CMD_DETRUIRE).Enabled = True
        cmd(CMD_DETRUIRE).ToolTipText = "Supprimer la fonction: " & UCase(rs("FT_Libelle").Value)
        txt(TXT_NIVEAU).tag = rs("FT_Niveau").Value
        rs.Close
        sql = "SELECT FNIV_Libelle FROM FCT_Niveau WHERE FNIV_Num=" & txt(TXT_NIVEAU).tag
        If Odbc_RecupVal(sql, fniv_libelle) = P_ERREUR Then
            GoTo lab_erreur
        End If
        txt(TXT_NIVEAU).Text = txt(TXT_NIVEAU).tag & " - " & fniv_libelle
    
        ' Niveau de remplacement
        If FT_NivRemplace > 0 Then
            For i = 0 To cbo(CBO_MODIFIER_POSITION).ListCount - 1
                If cbo(CBO_MODIFIER_POSITION).ItemData(i) = FT_NivRemplace Then
                    cbo(CBO_MODIFIER_POSITION).ListIndex = i
                End If
            Next i
        End If
    
    Else ' Nouvelle fonction à créer
        frm(FRM_HAUT).Caption = "Nouvelle fonction"
        txt(TXT_NOM).Text = ""
        cmd(CMD_DETRUIRE).Visible = False
        ft_visibilite = 2
        g_old_visibilite = 2
        cbo(CBO_VISIBILITE).ListIndex = 2
        txt(TXT_NIVEAU).tag = 0
        sql = "SELECT FNIV_Libelle FROM FCT_Niveau WHERE FNIV_Num=0"
        If Odbc_RecupVal(sql, fniv_libelle) = P_ERREUR Then
            GoTo lab_erreur
        End If
        txt(TXT_NIVEAU).Text = "0 - " & fniv_libelle
    End If

    cmd(CMD_OK).Enabled = False
    If g_stype = "F" Then
        If g_numfct = 0 Then
            txt(TXT_NOM).SetFocus
        Else
            txt(TXT_NIVEAU).SetFocus
        End If
    Else
        If afficher_coordonnee() = P_ERREUR Then
            GoTo lab_erreur
        End If
'        grdCoord.SetFocus
    End If
    
    Me.MousePointer = 0
    g_mode_saisie = True
    
    afficher_fct = P_OK
    Exit Function

lab_erreur:
    Me.MousePointer = 0
    g_mode_saisie = True
    afficher_fct = P_ERREUR

End Function

Private Sub basculer_colonne_principal(ByVal v_row As Integer, ByVal v_col As Integer)

    Dim i As Integer, type_en_cours As Integer, nbr As Integer, ma_row As Integer

    With grdCoord
        ' ne pas tenter de basculer si on est sur la ligne fixe
        If v_row = 0 Then Exit Sub

        type_en_cours = .TextMatrix(v_row, GRDC_ZUNUM)
        If .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture Then
        ' le cellule elle est cochée
            .Row = v_row
            .col = v_col
            Set .CellPicture = ImageListe.ListImages(IMG_PASCOCHE).Picture ' on décoche cette ligne
            nbr = 0
            For i = 1 To .Rows - 1 ' on chrche le nombre de lignes du même type
                If .TextMatrix(i, GRDC_ZUNUM) = type_en_cours And i <> v_row Then
                    ma_row = i
                    nbr = nbr + 1
                End If
            Next i
            If nbr = 1 Then ' s'il y en a deux, on coche la deuxième
                .Row = ma_row
                .col = GRDC_PRINCIPAL
                Set .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture
            End If
            cmd(CMD_OK).Enabled = True
        Else ' elle n'est pas cochée
            ' vérifier s'il n'y pas d'autre principale pour le même type de coordonnée
            For i = 1 To .Rows - 1
                If .TextMatrix(i, GRDC_ZUNUM) = type_en_cours And i <> v_row Then
                    .col = GRDC_PRINCIPAL
                    .Row = i
                    If .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture Then
                        Set .CellPicture = ImageListe.ListImages(IMG_PASCOCHE).Picture
                    End If
                End If
            Next i
            .Row = v_row
            .col = v_col
            Set .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture
            cmd(CMD_OK).Enabled = True
        End If
    End With

End Sub

Private Function charge_niveau() As Integer
    
    Dim sql As String, nNiv As Long, rs As rdoResultset
    Dim libNiveau As String
    Dim NivPere As Integer
    Dim encore As Boolean
    
    sql = "select count(*) from niveau_structure"
    If Odbc_Count(sql, nNiv) = P_ERREUR Then
        charge_niveau = 0
        Exit Function
    Else
        charge_niveau = nNiv
        NivPere = 0
        encore = True
        cbo(CBO_MODIFIER_POSITION).Clear
        cbo(CBO_MODIFIER_POSITION).AddItem "Inchangé"
        cbo(CBO_MODIFIER_POSITION).ItemData(cbo(CBO_MODIFIER_POSITION).ListCount - 1) = -1
        While encore
            sql = "select * from niveau_structure where Nivs_NivPere=" & NivPere
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                charge_niveau = 0
            ElseIf Not rs.EOF Then
                cbo(CBO_MODIFIER_POSITION).AddItem rs("Nivs_Nom")
                cbo(CBO_MODIFIER_POSITION).ItemData(cbo(CBO_MODIFIER_POSITION).ListCount - 1) = rs("Nivs_Num")
                NivPere = rs("Nivs_Num")
                rs.MoveNext
            Else
                encore = False
            End If
        Wend
    End If
    
End Function

Private Function choisir_fct() As Integer

    Dim sret As String, sql As String
    Dim n As Integer
    Dim nofct As Long
    Dim rs As rdoResultset
    Dim s As String
    Dim lib As String
    Dim strNiveau As String
    
    Call FRM_ResizeForm(Me, 0, 0)

lab_affiche:
    Call CL_Init
    'Choix de la fonction
    sql = "SELECT * FROM FctTrav" _
        & " ORDER BY FT_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        choisir_fct = P_ERREUR
        Exit Function
    End If
    n = 0
    If g_crfct_autor Then
        Call CL_AddLigne("<Nouvelle>", 0, "", False)
        n = 1
    End If
    While Not rs.EOF
        s = rs("FT_NivRemplace")
        lib = ""
        strNiveau = ""
        If s > "0" Then
            Call Odbc_RecupVal("select nivs_nom from niveau_structure where nivs_num='" & s & "'", strNiveau)
            lib = " => (transposée au : " & strNiveau & " )"
        End If
        Call CL_AddLigne(rs("FT_Libelle").Value & lib, rs("FT_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close

    If n = 0 Then
        Call MsgBox("Aucune fonction n'a été trouvée.", vbOKOnly + vbInformation, "")
        choisir_fct = P_NON
        Exit Function
    End If

    Call CL_InitTitreHelp("Liste des fonctions", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_d_fonction.htm")
    Call CL_InitTaille(0, -15)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnimprimer.gif", vbKeyI, vbKeyF3, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    ChoixListe.Show 1
    ' Quitter
    If CL_liste.retour = 2 Then
        choisir_fct = P_NON
        Exit Function
    End If

    ' Imprimer
    If CL_liste.retour = 1 Then
        Call imprimer
        GoTo lab_affiche
    End If

    If afficher_fct(CL_liste.lignes(CL_liste.pointeur).num) = P_ERREUR Then
        choisir_fct = P_ERREUR
        Exit Function
    End If

    choisir_fct = P_OUI

End Function

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

' ************************************
' Enregistrer les coordonnées du poste
' ************************************
Private Function enregistrer_coordonnees() As Integer
    
    Dim uc_principal_val As Boolean
    Dim i As Integer
    Dim lng As Long

    ' Supprssion des coordonnées du poste
    If Odbc_Delete("UtilCoordonnee", _
                   "UC_Num", _
                   "WHERE UC_TypeNum=" & g_numposte & " AND UC_Type='P'", _
                   lng) = P_ERREUR Then
        GoTo lab_erreur
    End If
    
    ' Ajout des coordonnées
    With grdCoord
        For i = 1 To .Rows - 1
            .Row = i
            .col = GRDC_PRINCIPAL
            ' Si la coordonnée est principale
            If .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture Then
                uc_principal_val = True
            Else
                uc_principal_val = False
            End If
            If Odbc_AddNew("UtilCoordonnee", "UC_Num", "UC_Seq", False, lng, _
                           "UC_Type", "P", _
                           "UC_TypeNum", g_numposte, _
                           "UC_ZUNum", .TextMatrix(i, GRDC_ZUNUM), _
                           "UC_Valeur", .TextMatrix(i, GRDC_VALEUR), _
                           "UC_Comm", .TextMatrix(i, GRDC_COMMENTAIRE), _
                           "UC_Niveau", .TextMatrix(i, GRDC_NIVEAU), _
                           "UC_Principal", uc_principal_val, _
                           "UC_LstPoste", "") = P_ERREUR Then
                GoTo lab_erreur
            End If
        Next i
    End With

    enregistrer_coordonnees = P_OK
    Exit Function

lab_erreur:
    enregistrer_coordonnees = P_ERREUR

End Function


Private Function enregistrer_fct() As Integer

    If g_stype = "F" Then
        If g_numfct = 0 Then
            ' Insertion dans FctTrav du nouveau enregistrement
            If Odbc_AddNew("FctTrav", _
                           "FT_Num", _
                           "FT_Seq", _
                           True, _
                           g_numfct, _
                           "FT_Libelle", txt(TXT_NOM).Text, _
                           "FT_Niveau", txt(TXT_NIVEAU).tag, _
                           "FT_EstGroupe", False, _
                           "FT_NivRemplace", cbo(CBO_MODIFIER_POSITION).ItemData(cbo(CBO_MODIFIER_POSITION).ListIndex), _
                           "FT_Visible", cbo(CBO_VISIBILITE).ListIndex) = P_ERREUR Then
                GoTo lab_erreur
            End If
        Else
            ' MAJ de la table FctTrav
            If Odbc_Update("FctTrav", _
                           "FT_Num", _
                           "WHERE FT_Num=" & g_numfct, _
                           "FT_Libelle", txt(TXT_NOM).Text, _
                           "FT_Niveau", txt(TXT_NIVEAU).tag, _
                           "FT_NivRemplace", cbo(CBO_MODIFIER_POSITION).ItemData(cbo(CBO_MODIFIER_POSITION).ListIndex), _
                           "FT_Visible", cbo(CBO_VISIBILITE).ListIndex) = P_ERREUR Then
                GoTo lab_erreur
            End If
        End If
    Else
        ' Enregistrement des coordonnées
        If enregistrer_coordonnees() = P_ERREUR Then
            GoTo lab_erreur
        End If
    End If
    
    enregistrer_fct = P_OK
    Exit Function

lab_erreur:
    enregistrer_fct = P_ERREUR

End Function

Private Function fct_dans_util() As Integer

    Dim sql As String
    Dim lnb As Long

    sql = "SELECT COUNT(*) FROM Utilisateur" _
        & " WHERE U_kb_actif=True AND U_FctTrav LIKE '%F" & g_numfct & ";%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        fct_dans_util = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        fct_dans_util = P_OUI
        Exit Function
    End If

    fct_dans_util = P_NON

End Function

Private Sub imprimer()

    Dim stexte(0) As String, sql As String
    Dim fl_fax As Boolean, fl_bid As Boolean
    Dim rs As rdoResultset

    ' Choix de l'imprimante
    If PR_ChoixImp(False, False, fl_fax, fl_bid) = False Then Exit Sub

    On Error Resume Next
    Printer.ScaleMode = vbTwips
    Printer.PaperSize = vbPRPSA4
    On Error GoTo err_printer

    stexte(0) = ""
    Call PR_InitFormat(True, _
                       "Liste des fonctions", _
                       False, _
                       "g", _
                       stexte())

    sql = "SELECT * FROM FctTrav" _
        & " ORDER BY FT_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    If rs.EOF Then
        rs.Close
        GoTo lab_erreur
    End If
    While Not rs.EOF
        stexte(0) = rs("FT_Libelle").Value
        Call PR_ImpLigne(stexte())
        rs.MoveNext
    Wend
    rs.Close

    Printer.EndDoc
    Call PR_RestoreImp
    Exit Sub

lab_erreur:
    On Error GoTo lab_fin
    Printer.KillDoc
    Call PR_RestoreImp
    On Error GoTo 0
    Exit Sub

err_printer:
    MsgBox "Erreur d'impression" & vbCr & vbLf & "Impression annulée", vbInformation + vbOKOnly, ""
    On Error GoTo lab_fin
    Printer.KillDoc
    Call PR_RestoreImp
    On Error GoTo 0
lab_fin:
    Exit Sub

End Sub

Private Sub initialiser()
' ***************************************************************************************************************
    Dim i As Integer

    g_position_txt_cache = ""
    txt(TXT_CACHE).Text = ""
    
    g_largeur_grid_init = grdCoord.width
    g_left_cmd_init = cmd(CMD_PLUS).left
    
    g_crfct_autor = True

    Call FRM_ResizeForm(Me, 0, 0)

    g_mode_saisie = False

    If g_stype = "P" Then
        lbl(LBL_VISIBILITE).Visible = True
        cbo(CBO_VISIBILITE).Enabled = False
    End If
    ' initialiser le combobox
    With cbo(CBO_VISIBILITE)
        .Clear
        .AddItem CBO_VISIBILITE_JAMAIS, 0
        .AddItem CBO_VISIBILITE_VUE_DETAILLEE, 1
        .AddItem CBO_VISIBILITE_TOUJOURS, 2
        .ListIndex = 0
    End With

    ' Modifier un poste
    If g_stype = "P" Then
        txt(TXT_NOM).Enabled = False
        cmd(CMD_NIVEAU).Visible = False
        txt(TXT_NIVEAU).Enabled = False
        With grdCoord
            txt(TXT_CACHE).left = .CellLeft
            txt(TXT_CACHE).Top = .CellTop
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
            For i = 0 To .Cols - 1
                .col = i
                .CellFontBold = True
                .ColAlignment(i) = 4
            Next i
            .col = GRDC_PRINCIPAL
            .col = GRDC_ZUNUM
        End With
        If afficher_fct(g_numfct) = P_ERREUR Then
            GoTo lab_erreur
        End If
        cmd(CMD_DETRUIRE).Visible = False
    Else
        'lbl(LBL_VISIBILITE).Visible = False
        lbl(LBL_COORDONNEES).Visible = False
        grdCoord.Visible = False
        cmd(CMD_PLUS).Visible = False
        cmd(CMD_MOINS).Visible = False
        ' Créer une fonction
        If g_numfct = 0 Then
            txt(TXT_NOM).Enabled = True
            If afficher_fct(0) = P_ERREUR Then
                GoTo lab_erreur
            End If
        ' Modifier une fonction
        Else
            If choisir_fct() <> P_OUI Then
                GoTo lab_erreur
            End If
        End If
        If g_numfct = 0 Then
            cmd(CMD_DETRUIRE).Visible = False
        Else
            cmd(CMD_DETRUIRE).Visible = True
        End If
    End If

    cmd(CMD_OK).Visible = True
    cmd(CMD_OK).Enabled = False

    Exit Sub

lab_erreur:
    Call quitter(True)

End Sub

Private Sub modifier_niveau()
' modifier le niveau d'une application
    Dim sql As String, lib As String
    Dim is_selected As Boolean
    Dim lnum As Long
    Dim i As Integer
    Dim rs As rdoResultset

    is_selected = False

lab_afficher: ' ajouter le niveau nouvellement créée
    Call CL_Init
    Call CL_InitMultiSelect(False, True) 'selection multiple=true, retourner la ligne courante=true
    Call CL_InitTitreHelp("Choix d'un niveau de coordonnées", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_d_fonction.htm")
    Call CL_InitTaille(0, -10)

    ' Boucle SQL d'ajout dans la liste des choix
    sql = "SELECT * FROM FCT_Niveau ORDER BY FNIV_Num"
    If Odbc_Select(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    While Not rs.EOF
        ' Recherche si le niveau ait déjà été selectionné
        If rs("FNIV_Num").Value = CInt(txt(TXT_NIVEAU).tag) Then
            is_selected = True
        Else ' ne pas cocher
            is_selected = False
        End If
        Call CL_AddLigne(rs("FNIV_Num").Value & vbTab & rs("FNIV_Libelle").Value, rs("FNIV_Num").Value, "", is_selected)
        rs.MoveNext
    Wend
    rs.Close

    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("&Modifier l'intitulé", "", 0, 0, 2000)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)

    ChoixListe.Show 1

    ' ******************************** QUITTER ********************************
    If CL_liste.retour = 2 Then
        Exit Sub
    End If
     ' ***************** MODIFIER NIVEAU ******************
    If CL_liste.retour = 1 Then
        lnum = CL_liste.lignes(CL_liste.pointeur).num
        lib = STR_GetChamp(CL_liste.lignes(CL_liste.pointeur).texte, vbTab, 1)
        Call SAIS_Init
        Call SAIS_InitTitreHelp("Intitulé du niveau", "")
        Call SAIS_AddBouton("", p_chemin_appli + "\btnEnregistrer.gif", vbKeyE, vbKeyF1, 0)
        Call SAIS_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
        Call SAIS_AddChamp("Intitulé du niveau " & lnum, 50, 0, False, lib)
        Saisie.Show 1
        If SAIS_Saisie.retour = 0 Then ' enregistrer
            If Odbc_Update("FCT_Niveau", "FNIV_Num", "where FNIV_Num=" & lnum, _
                            "FNIV_Libelle", SAIS_Saisie.champs(0).sval) = P_ERREUR Then
                Call quitter(True)
            End If
        End If
        GoTo lab_afficher
    End If
   ' ****************************** ENREGISTRER ******************************
    If CL_liste.retour = 0 Then
        txt(TXT_NIVEAU).tag = CL_liste.lignes(CL_liste.pointeur).num
        txt(TXT_NIVEAU).Text = STR_GetChamp(CL_liste.lignes(CL_liste.pointeur).texte, vbTab, 0) & " - " & STR_GetChamp(CL_liste.lignes(CL_liste.pointeur).texte, vbTab, 1)
        cmd(CMD_OK).Enabled = True
    End If
    ' *************************************************************************

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
        If reponse = vbNo Then Exit Sub
    End If

    If g_mode_direct Then
        g_sret = ""
        Unload Me
        Exit Sub
    End If

    If choisir_fct() <> P_OUI Then Unload Me

End Sub

Private Sub saisir_commentaire()

    Dim frm As Form

    Set frm = SaisieCommentaire
    If SaisieCommentaire.AppelFrm(grdCoord) Then
        cmd(CMD_OK).Enabled = True
    End If
    Set frm = Nothing

End Sub

Private Sub supprimer()

    Dim reponse As Integer, cr As Integer
    Dim lng As Long

    ' Utilisateur associé à cette fonction ?
    cr = fct_dans_util()
    If cr = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    If cr = P_OUI Then
        Call MsgBox("Des personnes sont associées à cette fonction." & vbLf & vbCr _
                  & "Cette fonction ne peut donc pas être supprimée.", vbExclamation + vbOKOnly, "")
        Exit Sub
    End If

    reponse = MsgBox("Confirmez-vous la suppression de cette fonction ?", vbQuestion + vbYesNo + vbDefaultButton2, "")
    If reponse = vbNo Then
        Exit Sub
    End If

    ' Supprimer les enregistrements des tables correspondantes
    If Odbc_Delete("FctTrav", _
                   "FT_Num", _
                   "WHERE FT_Num=" & g_numfct, _
                   lng) = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If

    If choisir_fct() <> P_OUI Then Call quitter(True)

End Sub

Private Sub supprimer_coord()

    Dim i As Integer, row_en_cours As Integer
    Dim num_en_cours As Long

    With grdCoord
        row_en_cours = .Row
        If .Rows = 2 Then ' On a une seule ligne + ligne fixe
            .Rows = 1     ' On ne laisse que la ligne fixe
        ElseIf .Row > 0 Then ' On a plusieurs lignes
            num_en_cours = .TextMatrix(.Row, GRDC_ZUNUM)
            .col = GRDC_PRINCIPAL
            If .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture Then
                For i = 1 To .Rows - 1
                    .Row = i
                    .col = GRDC_PRINCIPAL
                    If .TextMatrix(i, GRDC_ZUNUM) = num_en_cours And i <> row_en_cours Then
                        Set .CellPicture = ImageListe.ListImages(IMG_COCHE).Picture
                        Exit For
                    End If
                Next i
            End If
            Call .RemoveItem(row_en_cours)
            ' Remettre la taille du grid par défaut, et la disposition des boutons
            If .Rows <= NBRMAX_ROWS And .width = g_largeur_grid_init + 255 Then
                .width = .width - 255
                cmd(CMD_PLUS).left = g_left_cmd_init
                cmd(CMD_MOINS).left = g_left_cmd_init
            End If
        End If
        ' il ne reste plus de coordonnées dans le tableau
        If .Rows - 1 = 0 Then .Enabled = False
    End With

    g_position_txt_cache = ""
    cmd(CMD_OK).Enabled = True

End Sub

Private Sub valider()

    Dim cr As Integer

    cr = verif_tous_chp()
    If cr = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    If cr = P_NON Then
        Exit Sub
    End If

    Me.MousePointer = 11

    cr = enregistrer_fct()
    Me.MousePointer = 0
    If cr = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    If g_mode_direct Then
        g_sret = g_numfct & "|" & txt(TXT_NOM).Text
        Unload Me
        Exit Sub
    End If

    If choisir_fct() <> P_OUI Then Call quitter(True)

End Sub

Private Function verif_code() As Integer

    Dim lib As String, sql As String
    Dim rs As rdoResultset

    lib = txt(TXT_NOM).Text
    If lib <> "" Then
        sql = "SELECT FT_Num FROM FctTrav" _
            & " WHERE " & Odbc_upper() & "(FT_Libelle)=" & Odbc_String(UCase(lib))
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            verif_code = P_ERREUR
            Exit Function
        End If
        If Not rs.EOF Then
            If rs("FT_Num").Value <> g_numfct Then
                rs.Close
                Call MsgBox("Fonction déjà existante.", vbOKOnly + vbExclamation, "")
                verif_code = P_NON
                Exit Function
            End If
        End If
        rs.Close
    End If

    verif_code = P_OUI

End Function

Private Function verif_tous_chp() As Integer

    Dim i As Integer

    ' Verification de l'intitulé
    If txt(TXT_NOM).Text = "" Then
        Call MsgBox("L' INTITULE de la fonction est une rubrique obligatoire.", vbOKOnly + vbExclamation, "")
        txt(TXT_NOM).SetFocus
        GoTo lab_erreur
    End If
    ' aret ****************************************************************
    ' Verification du niveau d'accessibilité
    'If txt(TXT_NIVEAU).Text <> "" Then
    '    If Not STR_EstEntierPos(txt(TXT_NIVEAU).Text) Then
    '        Call MsgBox("Le NIVEAU D'ACCESSIBILITE doit être un entier positif.", vbExclamation + vbOKOnly, "")
    '        txt(TXT_NIVEAU).SetFocus
    '        GoTo lab_erreur
    '    End If
    'Else ' txt(TXT_NIVEAU).Text = ""
    '    txt(TXT_NIVEAU).Text = "0"
    'End If
    ' ************************************************************************
    ' Vérifier la validité du NIVEAU de confidentialité et la VALEUR du coordonnée
    With grdCoord
        For i = 1 To .Rows - 1
            ' la VALEUR du coordonnée est-elle vide ?
            If .TextMatrix(i, GRDC_VALEUR) = "" Then
                Call MsgBox("La VALEUR du coordonnée doit être renseignée.", vbCritical + vbOKOnly, _
                            "Erreur de validation")
                .col = GRDC_VALEUR
                .Row = i
                Call positionner_txt(.left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight, _
                                     .TextMatrix(i, GRDC_VALEUR))
                GoTo lab_erreur
            End If
            ' NIVEAU de confidentialité est-il valide ?
            If .TextMatrix(i, GRDC_NIVEAU) = "" Then
                .TextMatrix(i, GRDC_NIVEAU) = "0"
            End If
            If .TextMatrix(i, GRDC_NIVEAU) > 9 Then ' Le NIVEAU doit € [0; 9]
                Call MsgBox("Le NIVEAU de confidentialité incorrect." & "Il doit être compris entre 0 (niveau bas) et 9 (niveau haut).", _
                            vbCritical + vbOKOnly, "Erreur de validation")
                .TextMatrix(i, GRDC_NIVEAU) = ""
                .col = GRDC_NIVEAU
                GoTo lab_erreur
            ElseIf Not STR_EstEntierPos(.TextMatrix(i, GRDC_NIVEAU)) Then
                Call MsgBox("Le NIVEAU de confidentialité doit être un entier positif.", vbCritical + vbOKOnly, _
                            "Erreur de validation")
                .TextMatrix(i, GRDC_NIVEAU) = ""
                .col = GRDC_NIVEAU
                .Row = i
                Call positionner_txt(.left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight, _
                                     .TextMatrix(i, GRDC_NIVEAU))
                GoTo lab_erreur
            End If
        Next i
    End With

    verif_tous_chp = P_OUI
    Exit Function

lab_erreur:
    verif_tous_chp = P_NON

End Function

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

Private Sub cbo_Change(Index As Integer)
    cmd(CMD_OK).Enabled = True
End Sub

Private Sub cbo_DropDown(Index As Integer)
    With cbo(Index)
 '       If Index = CBO_VISIBILITE Then
            If g_old_visibilite <> .List(.ListIndex) Then
                cmd(CMD_OK).Enabled = True
            End If
'        End If
    End With
End Sub

Private Sub cbo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txt(TXT_NOM).SetFocus
    End If
End Sub

Private Sub cbo_LostFocus(Index As Integer)
    With cbo(Index)
        If Index = CBO_VISIBILITE Then
            If g_old_visibilite <> .List(.ListIndex) Then
                cmd(CMD_OK).Enabled = True
            End If
    End If
    End With
End Sub

Private Sub cbo_Validate(Index As Integer, Cancel As Boolean)
    With cbo(Index)
        If Index = CBO_VISIBILITE Then
            If g_old_visibilite <> .List(.ListIndex) Then
                cmd(CMD_OK).Enabled = True
            End If
        End If
    End With
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
    Case CMD_OK
        Call valider
    Case CMD_DETRUIRE
        Call supprimer
    Case CMD_QUITTER
        Call quitter(False)
    Case CMD_PLUS
        Call ajouter_coord
    Case CMD_MOINS
        Call supprimer_coord
    Case CMD_NIVEAU
        Call modifier_niveau
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
        If cmd(CMD_DETRUIRE).Enabled Then ' modification
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

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Index = CMD_QUITTER Then g_mode_saisie = False

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
        If cmd(CMD_DETRUIRE).Enabled Then ' modification
            Call supprimer
        End If
    ElseIf (KeyCode = vbKeyH And Shift = vbAltMask) Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_d_fonction.htm")
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Me.ActiveControl <> txt(TXT_CACHE) Then
            KeyAscii = 0
            SendKeys "{TAB}"
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        If txt(TXT_CACHE).Visible Or grdCoord.tag = "focus_oui" Then
            With grdCoord
                ' le grdCoord a le focus, alors on ne quitte pas
                txt(TXT_CACHE).Visible = False
                grdCoord.tag = "focus_non"
            End With
        Else ' le grdCoord n'a pas le focus, on peut quitter
            Call quitter(False)
        End If
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
        If .Rows = 1 Then Exit Sub
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
        If cmd(CMD_DETRUIRE).Enabled Then ' modification
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
                Call saisir_commentaire
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
        If .Rows = 1 Then Exit Sub
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

    cmd(CMD_OK).Enabled = True

End Sub

Private Sub txt_GotFocus(Index As Integer)

    If Index <> TXT_CACHE Then
        txt(TXT_CACHE).Visible = False
    End If
    g_txt_avant = txt(Index).Text
    
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyE And Shift = vbAltMask) Or KeyCode = vbKeyF1 Then
        KeyCode = 0
        If cmd(CMD_OK).Enabled Then Call valider
    ElseIf (KeyCode = vbKeyS And Shift = vbAltMask) Or KeyCode = vbKeyF2 Then
        KeyCode = 0
        If cmd(CMD_DETRUIRE).Enabled Then ' modification
            Call supprimer
        End If
    ElseIf (KeyCode = vbKeyH And Shift = vbAltMask) Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_d_fonction.htm")
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

    Dim cr As Integer

    ' Cacher le TXT_CACHE
    If Index = TXT_CACHE Then
        txt(TXT_CACHE).Visible = False
    End If

    If g_mode_saisie Then
        If txt(Index).Text <> g_txt_avant Then
            If Index = TXT_NOM Then
                cr = verif_code()
                If cr = P_ERREUR Then
                    Call quitter(True)
                    Exit Sub
                End If
                If cr = P_NON Then
                    txt(Index).Text = g_txt_avant
                    txt(Index).SetFocus
                    Exit Sub
                End If
            End If
            cmd(CMD_OK).Enabled = True
        End If
    End If

End Sub
