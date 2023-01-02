VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PrmServiceModif 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8955
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Height          =   315
      Index           =   4
      Left            =   10320
      Picture         =   "PrmServiceModif.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2520
      Width           =   375
   End
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Service"
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
      Height          =   8295
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10935
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   8400
         TabIndex        =   18
         Text            =   "caché"
         Top             =   4800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmd 
         Height          =   330
         Index           =   5
         Left            =   10320
         Picture         =   "PrmServiceModif.frx":0457
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4440
         Width           =   375
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "PrmServiceModif.frx":089E
         Left            =   2490
         List            =   "PrmServiceModif.frx":08A0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1590
         Width           =   2775
      End
      Begin MSComctlLib.ImageList ImageListGenerale 
         Left            =   9960
         Top             =   720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   13
         ImageHeight     =   13
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PrmServiceModif.frx":08A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PrmServiceModif.frx":0BF4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   1620
         TabIndex        =   10
         Top             =   1080
         Width           =   7815
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1620
         TabIndex        =   9
         Top             =   600
         Width           =   7815
      End
      Begin VB.CommandButton cmd 
         Height          =   330
         Index           =   3
         Left            =   10280
         Picture         =   "PrmServiceModif.frx":0F46
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   7680
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         Height          =   315
         Index           =   2
         Left            =   10280
         Picture         =   "PrmServiceModif.frx":138D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5280
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid grdCoordLiees 
         Height          =   2715
         Left            =   240
         TabIndex        =   6
         Top             =   5280
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   4789
         _Version        =   393216
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
      End
      Begin MSFlexGridLib.MSFlexGrid grdCoord 
         Height          =   2235
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   3942
         _Version        =   393216
         Cols            =   7
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
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   2535
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
         Left            =   360
         TabIndex        =   12
         Top             =   1635
         Width           =   2115
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Coordonnées liées"
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
         TabIndex        =   11
         Top             =   4920
         Width           =   2535
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
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rattaché à"
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
         TabIndex        =   4
         Top             =   690
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   8160
      Width           =   10940
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmServiceModif.frx":17E4
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
         Picture         =   "PrmServiceModif.frx":1D40
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         UseMaskColor    =   -1  'True
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
         Left            =   9750
         Picture         =   "PrmServiceModif.frx":22A9
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "PrmServiceModif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Indes des boutons
Private Const CMD_OK = 0
Private Const CMD_QUITTER = 1
Private Const CMD_PLUS_COORDLIEES = 2
Private Const CMD_MOINS_COORDLIEES = 3
Private Const CMD_PLUS_TYPE = 4
Private Const CMD_MOINS_TYPE = 5

' Le Combo/ListBox pour l'affichage des services
Private Const CBO_JAMAIS = "Jamais"
Private Const CBO_VUE_DETAILLEE = "Vue détaillée seulement"
Private Const CBO_TOUJOURS = "Toujours"

' Index des libéllés
Private Const LBL_RATTACHE = 0
Private Const LBL_NOM = 0
Private Const LBL_VISIBILITE = 3
' Index des zones texte
Private Const TXT_RATTACHE = 0
Private Const TXT_NOM = 1
' TextBox caché
Private Const TXT_CACHE = 4
' Images MSFlexGrid des coordonnées
Private Const IMG_PASCOCHE = 1
Private Const IMG_COCHE = 2


'Index des colonnes du GrdCoord
Private Const GRDC_ZUNUM = 0
Private Const GRDC_TYPE = 1
Private Const GRDC_VALEUR = 2
Private Const GRDC_NIVEAU = 3
Private Const GRDC_PRINCIPAL = 4
Private Const GRDC_COMMENTAIRE = 5
Private Const GRDC_UCNUM = 6
'Index des colonnes du GrdCoordLiees
Private Const GRDCL_CODE = 0
Private Const GRDCL_TYPE = 1
Private Const GRDCL_VALEUR = 2
Private Const GRDCL_NIVEAU = 3
Private Const GRDCL_PRINCIPAL = 4
Private Const GRDCL_IDENTITE = 5

' Nombre de ligne visibles dans le grdCoord (éditables + fixe)
Private Const NBRMAX_ROWS = 9
' Largeur du grdCoord par defaut
Private Const LARGEUR_GRID_PAR_DEFAUT = 10035

Private g_txt_avant As String, g_sprm As String, g_spm As String, _
        g_coordonnees_avant As String, g_coordonnees_apres As String
' La position de la forme: "ligne-colonne"
Private g_position_txt_cache As String
Private g_mode_saisie As Boolean
Private g_old_nom As String, g_srvnom As String
Private g_srvnum As Long, g_srvnumpere As Long, g_srvnum_creation As Long
Private g_form_width As Integer, g_form_height As Integer
Private g_form_active As Boolean, g_changements As Boolean
Private g_old_visibilite As String 'l'indice de visibilté du service
' Le tableau des types de coordonnées supprimées
Private g_coord_supprimees() As Long
Private g_nbr_coord_supp As Integer

Private Function afficher_coordonnee()

    Dim sql As String, poste As String
    Dim I As Integer
    Dim rs As rdoResultset

    sql = "SELECT * FROM UtilCoordonnee, ZoneUtil" _
        & " WHERE UC_TypeNum=" & g_srvnum _
        & " AND UC_Type ='S' AND ZU_Type='C'" _
        & " AND UC_ZUNum=ZU_Num" _
        & " ORDER BY ZU_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        afficher_coordonnee = P_ERREUR
        Exit Function
    End If

    I = 1

    With grdCoord
        ' Affichage des lignes du gridCoord
        While Not rs.EOF
            'Indicateur d'attribution de la coordonnee
            If rs("UC_Lstposte").Value <> "" Then
                poste = " *"
            Else
                poste = "  "
            End If
            .AddItem rs("ZU_Num").Value & vbTab & rs("ZU_Libelle") & poste & vbTab & rs("UC_Valeur") _
                    & vbTab & rs("UC_Niveau") & vbTab & vbTab & rs("UC_Comm") & vbTab & rs("UC_Num").Value
            '.Row = i
            .Row = .Rows - 1
            .col = GRDC_PRINCIPAL
            .CellPictureAlignment = 4
            If rs("UC_Principal").Value Then
                Set .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture
            Else
                Set .CellPicture = ImageListGenerale.ListImages(IMG_PASCOCHE).Picture
            End If
            rs.MoveNext
            I = I + 1
        Wend

        ' Redimensionner le grid est déplacer les boutons +/- si besoin
        If .Rows > NBRMAX_ROWS And .width = LARGEUR_GRID_PAR_DEFAUT Then
            .ColWidth(GRDC_COMMENTAIRE) = 2280
        End If
        .Enabled = IIf(.Rows - 1 = 0, False, True)
    End With

    afficher_coordonnee = P_OK
    rs.Close

End Function

Private Function afficher_coordonneeLiees()

    Dim sql As String
    Dim rs As rdoResultset

    sql = "SELECT * FROM Coordonnee_Associee, UtilCoordonnee, ZoneUtil" _
        & " WHERE UC_ZUNum=ZU_Num" _
        & " AND UC_Num=CA_UCNum" _
        & " AND CA_UCTypeNum=" & g_srvnum _
        & " AND CA_UCType ='S' AND ZU_Type='C'" _
        & " ORDER BY ZU_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then GoTo lab_erreur

    ' Affichage des lignes du gridCoord
    With grdCoordLiees
        .Rows = 1
        While Not rs.EOF
            If set_coordLiees(rs("UC_Num").Value, rs("CA_Principal").Value) = P_ERREUR Then GoTo lab_erreur
            rs.MoveNext
        Wend

        ' Redimensionner le grid est les boutons +/- si besoin est
        If .Rows > NBRMAX_ROWS And .width = LARGEUR_GRID_PAR_DEFAUT Then
            .ColWidth(GRDCL_IDENTITE) = 2310
        End If
        .Enabled = IIf(.Rows = 0, False, True)
    End With
    rs.Close

    afficher_coordonneeLiees = P_OK
    Exit Function

lab_erreur:
    afficher_coordonneeLiees = P_ERREUR

End Function

Private Function afficher_service() As Integer

    Dim sql As String, s As String, srv_nom As String
    Dim rattache As String, srv_visible As String
    Dim srv_numpere As Long
    Dim pos As Integer, nbr_mouvements As Integer
    Dim rs As rdoResultset
    
    g_mode_saisie = False
    
    Call FRM_ResizeForm(Me, 0, 0)

    If (g_srvnum > 0) Then ' Modifier
        sql = "SELECT * FROM Service WHERE SRV_Num=" & g_srvnum
        If Odbc_Select(sql, rs) = P_ERREUR Then
            GoTo lab_err
        End If
        srv_nom = rs("SRV_Nom").Value
        srv_numpere = rs("SRV_NumPere").Value
        srv_visible = rs("SRV_Visible").Value
        rs.Close
        If (srv_numpere > 0) Then ' chercher le service père
            sql = "SELECT SRV_Nom FROM Service WHERE SRV_Num=" & srv_numpere
        Else
            sql = "SELECT L_Nom FROM Laboratoire"
        End If
        If Odbc_RecupVal(sql, rattache) = P_ERREUR Then
            GoTo lab_err
        End If
        txt(TXT_RATTACHE).Text = rattache
        txt(TXT_NOM).Text = srv_nom
        cbo.ListIndex = CInt(srv_visible)
        frm.Caption = "Service " & srv_nom
        g_old_nom = srv_nom
        g_old_visibilite = srv_visible
        frm.Caption = "Service: " & srv_nom
        If afficher_coordonnee() = P_ERREUR Then
            GoTo lab_err
        End If
        If afficher_coordonneeLiees() = P_ERREUR Then
            GoTo lab_err
        End If
    Else ' Création
        If Odbc_RecupVal("SELECT L_Nom FROM Laboratoire", rattache) = P_ERREUR Then
            GoTo lab_err
        End If
        txt(TXT_RATTACHE).Text = rattache
        frm.Caption = "Création d'un nouveau service"
        g_old_visibilite = 2
        cbo.ListIndex = 2
    End If
    
    g_mode_saisie = True
    
    cmd(CMD_OK).Enabled = False

    Call FRM_ResizeForm(Me, Me.width, Me.Height)

    txt(TXT_NOM).SetFocus
    
    afficher_service = P_OK
    Exit Function

lab_err:
    afficher_service = P_ERREUR

End Function

Private Sub ajouter_coordLiee()
    Dim selection As String, sret As String, entite As String, sql As String, _
        titre_entite As String, nom As String, prenom As String, choix As String
    Dim I As Long, lng As Long, nbr_ligne As Long
    Dim uc_selected As Boolean
    Dim rs As rdoResultset
    Dim frm As Form

    With grdCoordLiees
        Call CL_Init
        Call CL_InitMultiSelect(False, False) ' (selection multiple=True, retourner la ligne courante=False)
        Call CL_InitTitreHelp("Liste des entités possibles", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_d_fonction.htm")
        Call CL_InitTaille(0, -15)
    
        Call CL_AddLigne("Personne", 0, "U", False)
        Call CL_AddLigne("Poste ou Pièce", 1, "P", False)
    
        Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
        Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    
        ChoixListe.Show 1
    
        If CL_liste.retour = 1 Then ' --------------- QUITTER
            Exit Sub
        End If
    
        ' selection
        If CL_liste.retour = 0 Then ' --------------- SELECTIONNER
            If CL_liste.pointeur = 0 Then ' personne
                ' selectionner depuis la fonction recherche
lab_reselection_personne:
                Set frm = ChoixUtilisateur
                choix = ChoixUtilisateur.AppelFrm("Choisir une personne", _
                                                    "", _
                                                    False, _
                                                    False, _
                                                    "")
                Set frm = Nothing
                If choix = "" Then ' pas de personne selectionnée
                    Exit Sub
                End If
                lng = p_tblu_sel(0)
                entite = "U"
                ' pour les utilisateur, on passe par
            Else ' CL_liste.pointeur = 1 => poste ou pièce
lab_reselection_poste_piece:
                Set frm = ChoixService
                sret = ChoixService.AppelFrm("Choix du poste/pièce", "S", False, "1;", "PC", False)
                Set frm = Nothing
                If sret = "" Then Exit Sub
                entite = STR_GetChamp(sret, ";", STR_GetNbchamp(sret, ";") - 1)
                lng = Mid$(entite, 2)
                entite = Mid$(entite, 1, 1)
            End If
            sql = "SELECT * FROM UtilCoordonnee WHERE UC_Type='" & entite & "' AND UC_TypeNum=" & lng
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                Exit Sub
            End If
            If rs.EOF Then ' pas de coordonnées pour cette entité
                If MsgBox("Cette entitié ne dispose pas de coordonnée." & vbCrLf & _
                           "Voulez-vous choisir une autre entité ?", vbQuestion + vbYesNo, "Pas de coordonnée") = vbYes Then
                    If entite = "U" Then
                        GoTo lab_reselection_personne
                    Else
                        GoTo lab_reselection_poste_piece
                    End If
                Else ' on sort
                    Exit Sub
                End If
            End If
            rs.Close
            ' choix multiple : CL_InitMultiSelect
            Select Case entite
                Case "U":
                    If Odbc_RecupVal("SELECT U_Nom, U_Prenom FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & lng, _
                                    nom, prenom) = P_ERREUR Then
                        Exit Sub
                    End If
                    titre_entite = nom & " " & prenom
                Case "P":
                    titre_entite = "le poste: " & P_get_lib_srv_poste(lng, P_POSTE)
                Case "C":
                    titre_entite = "la pièce: " & P_get_nom_piece(lng)
            End Select
            ' Afficher la liste multiselect des coordonnées de l'entité
            Call CL_Init
            Call CL_InitMultiSelect(True, False) 'selection multiple=True, retourner la ligne courante=False
            Call CL_InitTitreHelp("Liste des coordonnées pour " & titre_entite, p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_d_fonction.htm")
            Call CL_InitTaille(0, -15)
            ' Boucle d'ajout dans la liste des coordonnées de l'entité selectionnée
            sql = "SELECT * FROM ZoneUtil, UtilCoordonnee" & _
                  " WHERE UC_ZUNum=ZU_Num AND UC_Type='" & entite & "' AND UC_TypeNum=" & lng
            If Odbc_Select(sql, rs) = P_ERREUR Then
                Exit Sub
            End If
            nbr_ligne = rs.RowCount
            While Not rs.EOF ' on est sur d'avoir au moin une coordonnée !
                uc_selected = coordLiee_existe(rs("UC_Num").Value)
                Call CL_AddLigne(rs("ZU_Libelle").Value & " :" & vbTab _
                                & rs("UC_Valeur").Value, rs("UC_Num").Value, _
                                rs("ZU_Code").Value, uc_selected)
                rs.MoveNext
            Wend
            rs.Close
            Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
            Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
            ChoixListe.Show 1
            ' La réponse:
            If CL_liste.retour = 1 Then ' QUITTER
                Exit Sub
            End If
            If CL_liste.retour = 0 Then ' ENREGISTRER
                For I = 0 To nbr_ligne - 1 ' on commence par ajouter
                    cmd(CMD_OK).Enabled = True
                    If CL_liste.lignes(I).selected Then
                        If Not coordLiee_existe(CL_liste.lignes(I).num) Then
                            Call set_coordLiees(CL_liste.lignes(I).num, False)
                        End If
                    End If
                Next I
                For I = 0 To nbr_ligne - 1 ' ..puis par supprimer
                    If Not CL_liste.lignes(I).selected Then ' essayer de la supprimer du tableau
                        Call enlever_coordLiees(CL_liste.lignes(I).num)
                    End If
                Next I
                cmd(CMD_OK).Enabled = True
            End If
        End If
        .Enabled = True
    End With

End Sub

Private Sub ajouter_type_coord()

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
    Call CL_AddBouton("&Créer un type de coordonnées", "", vbKeyU, vbKeyF3, 1500)
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
                    .AddItem (CL_liste.lignes(I).num & vbTab), j
                    .Row = j
                    .col = GRDC_PRINCIPAL
                    .CellPictureAlignment = 4
                    Set .CellPicture = ImageListGenerale.ListImages(IMG_PASCOCHE).Picture
                    insere_milieu = True
                    Exit For
                End If
            Next j
            If Not insere_milieu Then
                .AddItem (CL_liste.lignes(I).num & vbTab)
                .Row = .Rows - 1
                .col = GRDC_PRINCIPAL
                .CellPictureAlignment = 4
                Set .CellPicture = ImageListGenerale.ListImages(IMG_PASCOCHE).Picture
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
            Set .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture
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
                    .ColWidth(GRDC_COMMENTAIRE) = 2310
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

Public Function AppelFrm(ByVal v_srvnum As Long, ByVal v_srvnumpere As Long, _
                        ByRef v_srvnom As String, ByRef v_srvnum_creation As Long) As Boolean

    g_srvnum = v_srvnum
    g_srvnumpere = v_srvnumpere

    Me.Show 1

    v_srvnom = g_srvnom
    v_srvnum_creation = g_srvnum_creation
    AppelFrm = g_changements

End Function

Private Function coordLiee_existe(ByVal v_ucnum As Long) As Boolean
' Vérifier si la coordonnée liée v_ucnum est déjà dans le grid
    Dim I As Integer

    With grdCoordLiees
        For I = 1 To .Rows - 1
            If .TextMatrix(I, GRDCL_CODE) = v_ucnum Then
                coordLiee_existe = True
                Exit Function
            End If
        Next I
    End With
    ' on n'a rien trouvé
    coordLiee_existe = False

End Function

Private Function enlever_coordLiees(ByVal v_ucnum As Long) As Boolean
    Dim I As Integer

    With grdCoordLiees
        For I = 1 To .Rows - 1
            .Row = I
            If .TextMatrix(I, GRDCL_CODE) = v_ucnum Then
                Call supprimer_coordLiee
                enlever_coordLiees = True
                Exit Function
            End If
        Next I
    End With

    enlever_coordLiees = True

End Function

Private Function enregistrer_coordonnees() As Integer
' *******************************************************************
' Enregistrer les coordonnées de la personne et ses coordonnées liées
' *******************************************************************
    Dim uc_principal_val As Boolean
    Dim I As Integer
    Dim lng As Long

    If (g_srvnum > 0) Then ' service existe
        For I = 0 To g_nbr_coord_supp - 1
            If Odbc_Delete("UtilCoordonnee", _
                           "UC_Num", _
                           "WHERE UC_Num=" & g_coord_supprimees(I), _
                           lng) = P_ERREUR Then
                GoTo lab_erreur
            End If
        Next I
        ' Supprimer les anciennes coordonnées associées
        If Odbc_Delete("Coordonnee_Associee", _
                        "CA_Num", _
                        "WHERE CA_UCTypeNum=" & g_srvnum & " AND CA_UCType='S'", _
                        lng) = P_ERREUR Then
                GoTo lab_erreur
        End If
    End If
    
    With grdCoord
        For I = 1 To .Rows - 1
            .Row = I
            .col = GRDC_PRINCIPAL
            ' La coordonnée est-elle principale ?
            If .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture Then uc_principal_val = True
            If .TextMatrix(I, GRDC_UCNUM) <> "" Then ' une mise à jour
                If Odbc_Update("UtilCoordonnee", "UC_Num", "WHERE UC_Num=" & .TextMatrix(I, GRDC_UCNUM), _
                                "UC_Valeur", .TextMatrix(I, GRDC_VALEUR), _
                                "UC_Comm", .TextMatrix(I, GRDC_COMMENTAIRE), _
                                "UC_Niveau", .TextMatrix(I, GRDC_NIVEAU), _
                                "UC_Principal", uc_principal_val) = P_ERREUR Then
                    GoTo lab_erreur
                End If
            Else ' une nouvelle coordonnée à ajouter
                If Odbc_AddNew("UtilCoordonnee", "UC_Num", "UC_Seq", False, lng, _
                               "UC_Type", "S", _
                               "UC_TypeNum", g_srvnum, _
                               "UC_ZUNum", .TextMatrix(I, GRDC_ZUNUM), _
                               "UC_Valeur", .TextMatrix(I, GRDC_VALEUR), _
                               "UC_Comm", .TextMatrix(I, GRDC_COMMENTAIRE), _
                               "UC_Niveau", .TextMatrix(I, GRDC_NIVEAU), _
                               "UC_Principal", uc_principal_val, _
                               "UC_LstPoste", "") = P_ERREUR Then
                    GoTo lab_erreur
                End If
            End If
            uc_principal_val = False
        Next I
    End With
    
    With grdCoordLiees
        For I = 1 To .Rows - 1
            .Row = I
            .col = GRDCL_PRINCIPAL
            ' La coordonnée est-elle principale ?
            If .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture Then uc_principal_val = True
            If Odbc_AddNew("Coordonnee_Associee", "CA_Num", "CA_Seq", False, lng, _
                           "CA_UCType", "S", _
                           "CA_UCTypeNum", g_srvnum, _
                           "CA_UCNum", .TextMatrix(I, GRDCL_CODE), _
                           "CA_Principal", uc_principal_val) = P_ERREUR Then
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
    Dim I As Integer
    
    ' initialiser le tableau des coordonnées
    Erase g_coord_supprimees
    g_nbr_coord_supp = 0

    cmd(CMD_MOINS_TYPE).Enabled = False
    g_position_txt_cache = ""
    grdCoord.ScrollTrack = True

    With grdCoord
        txt(TXT_CACHE).left = .CellLeft
        txt(TXT_CACHE).Top = .CellTop

        Me.txt(TXT_NOM).SetFocus

        .Rows = 1
        .FormatString = "|Type de coordonnée|Valeur|Niveau de confidentialité|Principal dans son type|Commentaire"
        .RowHeight(0) = 750
        .ColWidth(GRDC_ZUNUM) = 0
        .ColWidth(GRDC_TYPE) = 2250
        .ColWidth(GRDC_VALEUR) = 2300
        .ColWidth(GRDC_NIVEAU) = 1500
        .ColWidth(GRDC_PRINCIPAL) = 1650
        .ColWidth(GRDC_COMMENTAIRE) = 2580
        .ColWidth(GRDC_UCNUM) = 0
        .Row = 0
        For I = 0 To .Cols - 1
            .col = I
            .CellFontBold = True
            .ColAlignment(I) = 4
        Next I
        .col = GRDC_ZUNUM
    End With
    
    With grdCoordLiees
        '.SelectionMode = flexSelectionByRow
        .Rows = 1
        .FormatString = "|Type de coordonnée|Valeur|Niveau de confidentialité|Principal dans son type|Propriétaire"
        .RowHeight(0) = 750
        .ColWidth(GRDCL_CODE) = 0
        .ColWidth(GRDCL_TYPE) = 2250
        .ColWidth(GRDCL_VALEUR) = 2300
        .ColWidth(GRDCL_NIVEAU) = 1500
        .ColWidth(GRDCL_PRINCIPAL) = 1650
        .ColWidth(GRDCL_IDENTITE) = 2330
        .Row = 0
        For I = 0 To .Cols - 1
            .col = I
            .CellFontBold = True
            .ColAlignment(I) = 4
        Next I
        .CellPictureAlignment = 4
        .col = GRDCL_CODE
        .Enabled = False
    End With

    ' initialiser le combobox
    cbo.Clear
    cbo.AddItem CBO_JAMAIS, 0
    cbo.AddItem CBO_VUE_DETAILLEE, 1
    cbo.AddItem CBO_TOUJOURS, 2
    cbo.ListIndex = 0

    g_changements = False
    cmd(CMD_MOINS_COORDLIEES).Enabled = False

    If afficher_service() = P_ERREUR Then
        Unload Me
        Exit Sub
    End If

End Sub

Private Sub quitter(ByVal v_bforce As Boolean)

    Dim reponse As Integer

    If v_bforce Then
        Unload Me
        Exit Sub
    End If

    If cmd(CMD_OK).Enabled Then
        If MsgBox("Des modifications ont été effectuées !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", _
                          vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If

    Unload Me

End Sub

Private Sub supprimer_type_coord()

    Dim I As Integer, row_en_cours As Integer
    Dim num_en_cours As Long

    With grdCoord
        row_en_cours = .Row
        If .TextMatrix(row_en_cours, GRDC_UCNUM) <> "" Then
        ' ne pas prendre en compte les coordonnées nouvellement ajoutées
            ReDim Preserve g_coord_supprimees(g_nbr_coord_supp)
            g_coord_supprimees(g_nbr_coord_supp) = .TextMatrix(row_en_cours, GRDC_UCNUM)
            g_nbr_coord_supp = g_nbr_coord_supp + 1
        End If
        If .Rows = 2 Then ' On a une seule ligne + ligne fixe
            .Rows = 1     ' On ne laisse que la ligne fixe
        ElseIf .Row > 0 Then ' On a plusieurs lignes
            num_en_cours = .TextMatrix(.Row, GRDC_ZUNUM)
            .col = GRDC_PRINCIPAL
            If .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture Then
                For I = 1 To .Rows - 1
                    .Row = I
                    .col = GRDC_PRINCIPAL
                    If .TextMatrix(I, GRDC_ZUNUM) = num_en_cours And I <> row_en_cours Then
                        Set .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture
                        Exit For
                    End If
                Next I
            End If
            Call .RemoveItem(row_en_cours)
            ' Remettre la taille du grid par défaut, et la disposition des boutons
            If .Rows <= NBRMAX_ROWS And .width = LARGEUR_GRID_PAR_DEFAUT + 255 Then
                .ColWidth(GRDC_COMMENTAIRE) = 2580
            End If
        End If
        ' il ne reste plus de coordonnées dans le tableau
        If .Rows - 1 = 0 Then .Enabled = False
    End With

    g_position_txt_cache = ""
    cmd(CMD_OK).Enabled = True

End Sub

Private Function set_coordLiees(ByVal v_ucnum As Long, ByVal v_principal As Boolean) As Integer
    Dim sql As String, uc_num As String, zu_libelle As String, _
        uc_valeur As String, uc_niveau As String, uc_type As String, _
        u_nom As String, u_prenom As String, libelle_entite As String
    Dim uc_typenum As Long

    sql = "SELECT UC_Num, ZU_Libelle, UC_Valeur, UC_Niveau, UC_Type, UC_TypeNum" & _
        " FROM ZoneUtil, UtilCoordonnee" & _
        " WHERE UC_ZUNum=ZU_Num AND UC_Num=" & v_ucnum
    If Odbc_RecupVal(sql, uc_num, zu_libelle, uc_valeur, uc_niveau, uc_type, uc_typenum) = P_ERREUR Then
        GoTo lab_erreur
    End If
    With grdCoordLiees
        .AddItem (v_ucnum)
        .TextMatrix(.Rows - 1, GRDCL_CODE) = uc_num
        .TextMatrix(.Rows - 1, GRDCL_TYPE) = zu_libelle
        .TextMatrix(.Rows - 1, GRDCL_VALEUR) = uc_valeur
        .TextMatrix(.Rows - 1, GRDCL_NIVEAU) = uc_niveau
        .Row = .Rows - 1
        .col = GRDCL_PRINCIPAL
        .CellPictureAlignment = 4
        If v_principal Then
            Set .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture
        Else
            Set .CellPicture = ImageListGenerale.ListImages(IMG_PASCOCHE).Picture
        End If
        .col = GRDCL_IDENTITE
        If uc_type = "U" Then ' une personne
            If Odbc_RecupVal("SELECT U_Nom, U_Prenom FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & uc_typenum, _
                            u_nom, u_prenom) = P_ERREUR Then
                GoTo lab_erreur
            End If
            libelle_entite = u_nom & " " & u_prenom
        ElseIf uc_type = "P" Then ' poste
            libelle_entite = P_get_lib_srv_poste(uc_typenum, P_POSTE)
        Else ' pièce
            libelle_entite = P_get_nom_piece(uc_typenum)
        End If
        .TextMatrix(.Row, .col) = libelle_entite
    End With

    set_coordLiees = P_OK
    Exit Function

lab_erreur:
    set_coordLiees = P_ERREUR

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

Private Sub basculer_colonne_principal(ByVal v_row As Integer, ByVal v_col As Integer)

    Dim I As Integer, type_en_cours As Integer, nbr As Integer, ma_row As Integer

    With grdCoord
        ' ne pas tenter de basculer si on est sur la ligne fixe
        If v_row = 0 Then Exit Sub

        type_en_cours = .TextMatrix(v_row, GRDC_ZUNUM)
        If .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture Then
        ' la cellule elle est cochée
            .Row = v_row
            .col = v_col
            Set .CellPicture = ImageListGenerale.ListImages(IMG_PASCOCHE).Picture ' on décoche cette ligne
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
                Set .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture
            End If
            cmd(CMD_OK).Enabled = True
        Else ' elle n'est pas cochée
            ' vérifier s'il n'y pas d'autre principale pour le même type de coordonnée
            For I = 1 To .Rows - 1
                If .TextMatrix(I, GRDC_ZUNUM) = type_en_cours And I <> v_row Then
                    .col = GRDC_PRINCIPAL
                    .Row = I
                    If .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture Then
                        Set .CellPicture = ImageListGenerale.ListImages(IMG_PASCOCHE).Picture
                    End If
                End If
            Next I
            .Row = v_row
            .col = v_col
            Set .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture
            cmd(CMD_OK).Enabled = True
        End If
    End With

End Sub

Private Sub supprimer_coordLiee()
    Dim row_en_cours As Integer

    With grdCoordLiees
        If .Rows = 1 Then Exit Sub
        row_en_cours = .Row
        If .Rows = 2 Then
            .Rows = 1
        Else
            Call .RemoveItem(row_en_cours)
        End If
        If .Rows = 1 Then
            .Enabled = False
            cmd(CMD_MOINS_COORDLIEES).Enabled = False
        End If
    End With
    cmd(CMD_OK).Enabled = True

End Sub

Private Function valider() As Boolean
    Dim lng As Long
    Dim num_site As Integer

    num_site = 1 ' le numéro du site
    g_srvnom = txt(TXT_NOM).Text
    ' ************************ VÉRIFICATIONS ************************
    If verifier_tous_les_champs = P_NON Then
        GoTo lab_erreur
    End If
    
    ' Enregistrer le nom
    If (g_srvnum > 0) Then
        ' MAJ nom du service s'il le faut
        If (g_old_nom <> g_srvnom Or g_old_visibilite <> cbo.ListIndex) Then
            If (Odbc_Update("Service", "Srv_Num", _
                            "WHERE Srv_Num=" & g_srvnum, _
                            "Srv_Nom", g_srvnom, _
                            "Srv_Visible", cbo.ListIndex)) = P_ERREUR Then
                            
                GoTo lab_erreur
            End If
        End If
    Else ' une création
        If (Odbc_AddNew("Service", "Srv_Num", "Srv_Seq", True, lng, _
                        "Srv_Nom", txt(TXT_NOM).Text, _
                        "Srv_NumPere", g_srvnumpere, _
                        "Srv_LNum", num_site, _
                        "Srv_Code", "", _
                        "Srv_Libcourt", "", _
                        "Srv_Visible", cbo.ListIndex)) = P_ERREUR Then
            GoTo lab_erreur
        End If
        g_srvnum_creation = lng
        
    End If
    
    ' enregistrer les coordonnées liées du service
    If enregistrer_coordonnees = P_ERREUR Then
        GoTo lab_erreur
    End If
    
    g_changements = True
    valider = True
    Exit Function

lab_erreur:
    valider = False

End Function

Private Function verifier_tous_les_champs() As Integer

    Dim iligne As Integer

    With grdCoord ' vérifier la validité du NIVEAU de confidentialité et la VALEUR du coordonnée
        For iligne = 1 To .Rows - 1
            If .TextMatrix(iligne, GRDC_VALEUR) = "" Then ' la VALEUR du coordonnée est vide
                Call MsgBox("La VALEUR du coordonnée doit être renseignée.", vbCritical + vbOKOnly, "Erreur de validation")
                .col = GRDC_VALEUR
                GoTo lab_erreur
            End If
            If .TextMatrix(iligne, GRDC_NIVEAU) = "" Then ' Le NIVEAU
                Call MsgBox("Le NIVEAU de confidentialité est obligatoire." & "Il doit être compris entre 0 (niveau bas) et 9 (niveau haut).", _
                            vbCritical + vbOKOnly, "Erreur de validation")
                .col = GRDC_NIVEAU
                GoTo lab_erreur
            ElseIf Not STR_EstEntierPos(.TextMatrix(iligne, GRDC_NIVEAU)) Then ' NIVEAU de confidentialité est invalide
                Call MsgBox("Le NIVEAU de confidentialité doit être un entier positif.", vbCritical + vbOKOnly, "Erreur de validation")
                .TextMatrix(iligne, GRDC_NIVEAU) = ""
                .col = GRDC_NIVEAU
                GoTo lab_erreur
            ElseIf .TextMatrix(iligne, GRDC_NIVEAU) > 9 Then ' Le NIVEAU doit € [0; 9]
                Call MsgBox("Le NIVEAU de confidentialité doit être compris entre 0 (niveau bas) et 9 (niveau haut).", _
                            vbCritical + vbOKOnly, "Erreur de validation")
                .TextMatrix(iligne, GRDC_NIVEAU) = ""
                .col = GRDC_NIVEAU
                GoTo lab_erreur
            End If
        Next iligne
    End With

    verifier_tous_les_champs = P_OUI
    Exit Function

lab_erreur:
    verifier_tous_les_champs = P_NON

End Function

Private Sub cbo_Change()
    cmd(CMD_OK).Enabled = True
End Sub

Private Sub cbo_DropDown()
    'If g_old_visibilite <> CInt(cbo.List(cbo.ListIndex)) Then
    If g_old_visibilite <> cbo.List(cbo.ListIndex) Then
        cmd(CMD_OK).Enabled = True
    End If
End Sub

Private Sub cbo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txt(TXT_NOM).SetFocus
    End If
End Sub

Private Sub cbo_LostFocus()
    'If g_old_visibilite <> CInt(cbo.List(cbo.ListIndex)) Then
    If g_old_visibilite <> cbo.List(cbo.ListIndex) Then
        cmd(CMD_OK).Enabled = True
    End If
End Sub

Private Sub cbo_Validate(Cancel As Boolean)
    'If g_old_visibilite <> CInt(cbo.List(cbo.ListIndex)) Then
    If g_old_visibilite <> cbo.List(cbo.ListIndex) Then
        cmd(CMD_OK).Enabled = True
    End If
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case CMD_PLUS_TYPE
            Call ajouter_type_coord
        Case CMD_MOINS_TYPE ' [s'il est active]
            Call supprimer_type_coord
            cmd(CMD_MOINS_TYPE).Enabled = False
        Case CMD_OK
            If valider Then Unload Me
            Exit Sub
        Case CMD_QUITTER
            Call quitter(True)
        Case CMD_PLUS_COORDLIEES
            Call ajouter_coordLiee
        Case CMD_MOINS_COORDLIEES
            Call supprimer_coordLiee
    End Select

End Sub

Private Sub Form_Activate()

    If g_form_active Then Exit Sub

    g_form_active = True
    Call initialiser

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        If valider Then Unload Me
        Exit Sub
    ElseIf KeyCode = vbKeyH And Shift = vbAltMask Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_e_spm.htm")
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        Call quitter(False)
        Exit Sub
    End If
    
End Sub

Private Sub Form_Load()

    g_form_active = False

    g_form_width = Me.width
    g_form_height = Me.Height

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Call quitter(False)
    End If

End Sub

Private Sub grdcoord_Click()
    
    Dim mouse_row As Integer

    With grdCoord
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
        cmd(CMD_MOINS_TYPE).Enabled = True
        If .Row = 0 Then Exit Sub
        If .col = GRDC_PRINCIPAL Then
            txt(TXT_CACHE).Visible = False
        End If
    End With
    
End Sub

Private Sub grdCoord_KeyPress(KeyAscii As Integer)
    
    Dim frm As Form
    Dim ucnum As Long
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
                If .TextMatrix(.Row, GRDC_UCNUM) <> "" Then
                    ucnum = .TextMatrix(.Row, GRDC_UCNUM)
                Else
                    ucnum = 0
                End If
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

Private Sub grdCoord_LostFocus()
    
    grdCoord.tag = "focus_non"
    cmd(CMD_MOINS_TYPE).Enabled = True
    
End Sub

Private Sub grdCoord_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim old_row As Integer, old_col As Integer, mouse_row As Integer, mouse_col As Integer
    Dim frm As Form
    Dim ucnum As Long
    If Button = MouseButtonConstants.vbRightButton Then Exit Sub
    With grdCoord
        mouse_row = .MouseRow
        mouse_col = .MouseCol
        If .TextMatrix(mouse_row, GRDC_UCNUM) <> "" Then
            ucnum = .TextMatrix(mouse_row, GRDC_UCNUM)
        Else
            ucnum = 0
        End If
        .tag = "focus_oui"
        Select Case mouse_col
            ' -------------------------------------------------------
            Case GRDC_PRINCIPAL
                Call basculer_colonne_principal(mouse_row, mouse_col)
                txt(TXT_CACHE).Visible = False
                g_position_txt_cache = mouse_row & ";" & mouse_col
            ' -------------------------------------------------------
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
        If .col < GRDC_VALEUR Then
            .col = GRDC_VALEUR
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

Private Sub grdCoordLiees_GotFocus()

    With grdCoordLiees
        If .Rows = 1 Then Exit Sub
'MsgBox "gotfocus"
        cmd(CMD_MOINS_COORDLIEES).Enabled = True
    End With

End Sub

Private Sub grdCoordLiees_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        If valider Then Unload Me
        Exit Sub
    End If

End Sub

Private Sub grdCoordLiees_KeyPress(KeyAscii As Integer)

    With grdCoordLiees
        Select Case KeyAscii
            Case vbKeySpace
                KeyAscii = 0
                If .col = GRDCL_PRINCIPAL Then
                    If .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture Then
                        Set .CellPicture = ImageListGenerale.ListImages(IMG_PASCOCHE).Picture
                    Else
                        Set .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture
                    End If
                End If
                cmd(CMD_OK).Enabled = True
        End Select
    End With

End Sub

Private Sub grdCoordLiees_LostFocus()
'MsgBox "lostfocus"
    cmd(CMD_MOINS_COORDLIEES).Enabled = True

End Sub

Private Sub grdCoordLiees_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then Exit Sub

    With grdCoordLiees
        Select Case .MouseCol
           Case GRDCL_PRINCIPAL
                If .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture Then
                    Set .CellPicture = ImageListGenerale.ListImages(IMG_PASCOCHE).Picture
                Else
                    Set .CellPicture = ImageListGenerale.ListImages(IMG_COCHE).Picture
                End If
                cmd(CMD_OK).Enabled = True
        End Select
    End With

End Sub

Private Sub txt_Change(Index As Integer)

    cmd(CMD_OK).Enabled = True

End Sub

Private Sub txt_GotFocus(Index As Integer)

    cmd(CMD_MOINS_COORDLIEES).Enabled = False

End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        If valider Then Unload Me
        Exit Sub
    End If

End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Index = TXT_CACHE Then
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
    
    Dim I As Integer

    If Index = TXT_CACHE Then
            grdCoord.tag = "focus_non"
            txt(TXT_CACHE).Visible = False
    End If

    If g_mode_saisie Then
        If txt(Index).Text <> g_txt_avant Then
            cmd(CMD_OK).Enabled = True
        End If
    End If
    
End Sub
