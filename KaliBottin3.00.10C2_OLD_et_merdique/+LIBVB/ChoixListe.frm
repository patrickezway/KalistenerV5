VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form ChoixListe 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmListe 
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
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton cmdHelp 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   4080
         Picture         =   "ChoixListe.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Aucun"
         Top             =   360
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox TxtRecherche 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdTR 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1170
         Picture         =   "ChoixListe.frx":0359
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Aucun"
         Top             =   330
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton cmdTR 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   750
         Picture         =   "ChoixListe.frx":069B
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Tous"
         Top             =   330
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   855
         Left            =   720
         TabIndex        =   3
         Top             =   1020
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1508
         _Version        =   393216
         ForeColor       =   8388608
         BackColorFixed  =   12632256
         GridColorFixed  =   16777215
         GridLines       =   0
         GridLinesFixed  =   0
         SelectionMode   =   1
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
      Begin VB.Label LblRecherche 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rechercher"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   525
         TabIndex        =   7
         Top             =   390
         Visible         =   0   'False
         Width           =   1095
      End
      Begin ComctlLib.ImageList imglst 
         Left            =   3750
         Top             =   1110
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   15
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   3
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixListe.frx":09DD
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixListe.frx":0D2F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixListe.frx":1081
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frmFct 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   4575
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
         Height          =   510
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "ChoixListe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IMG_SANS_COCHE = 1
Private Const IMG_COCHE = 2
Private Const IMG_COCHE_GRIS = 3

Private Const CMD_TOUS = 0
Private Const CMD_AUCUN = 1

Private g_nblig_visible As Integer
Private g_nbbouton As Integer

Private g_form_active As Boolean
Private g_mode_saisie As Boolean


Private Sub afficher_ligne(ByVal v_row As Integer, _
                           ByVal v_col As Integer, _
                           ByVal v_str As String)
    
    Dim lg_col As Long
    
    If grd.Cols < v_col + 1 Then
        grd.Cols = v_col + 1
    End If
    grd.TextMatrix(v_row, v_col) = v_str
    lg_col = FRM_LargeurTexte(Me, grd, v_str + "     ")
    If grd.ColWidth(v_col) < lg_col Then
        grd.ColWidth(v_col) = lg_col
    End If
    
End Sub

Private Sub changer_etat_selection()
    Dim I As Integer
    
    grd.col = 0
    If grd.CellPicture = imglst.ListImages(IMG_COCHE_GRIS).Picture Then
        Exit Sub
    End If
    
    grd.col = 0
    For I = 0 To UBound(CL_liste.lignes)
        If CL_liste.lignes(I).iGrd = grd.Row - grd.FixedRows Then
            GoTo ok
        End If
    Next I
    MsgBox "changer_etat_selection : Ligne non trouvée"
    Exit Sub
    'If CL_liste.lignes(grd.Row - grd.FixedRows).selected Then
    '    CL_liste.lignes(grd.Row - grd.FixedRows).selected = False
    '    Set grd.CellPicture = imglst.ListImages(IMG_SANS_COCHE).Picture
    'Else
    '    CL_liste.lignes(grd.Row - grd.FixedRows).selected = True
    '    Set grd.CellPicture = imglst.ListImages(IMG_COCHE).Picture
    'End If
ok:
    If CL_liste.lignes(I).selected Then
        CL_liste.lignes(I).selected = False
        Set grd.CellPicture = imglst.ListImages(IMG_SANS_COCHE).Picture
    Else
        CL_liste.lignes(I).selected = True
        Set grd.CellPicture = imglst.ListImages(IMG_COCHE).Picture
    End If
    grd.ColSel = grd.Cols - 1
    
End Sub

Private Sub chercher_ligne_avec_first_car(ByVal v_car As Integer)

    Dim s As String, s_car As String
    Dim I As Integer, lig As Integer, first_col As Integer, ipasse As Integer
    Dim lig_deb As Integer, lig_fin As Integer
    
    If CL_liste.prmfrm.multi_select Then
        first_col = 1
    Else
        first_col = 0
    End If
    s_car = Chr(v_car)
    For ipasse = 0 To 1
        If ipasse = 0 Then
            lig_deb = grd.Row + 1
            lig_fin = grd.Rows - 1
        Else
            lig_deb = 0
            lig_fin = grd.Row - 1
        End If
        For I = lig_deb To lig_fin
            s = left$(Trim$(grd.TextMatrix(I, first_col)), 1)
            If s = s_car Or UCase(s) = UCase(s_car) Then
                grd.Row = I
                grd.ColSel = grd.Cols - 1
                If grd.Row >= grd.TopRow + g_nblig_visible - grd.FixedRows Then
                    lig = grd.Row - g_nblig_visible + 1
                    If lig < 0 Then
                        grd.TopRow = grd.FixedRows
                    Else
                        grd.TopRow = lig
                    End If
                End If
                Exit Sub
            End If
        Next I
    Next ipasse
    
End Sub

Private Sub init_ligne_courante()

    If grd.Rows = 0 Then
        Exit Sub
    End If
    
    If CL_liste.pointeur > 0 Then
        If CL_liste.pointeur >= grd.Rows Then
            Call MsgBox("La ligne indiquée n'existe pas.", vbExclamation + vbOKOnly, "")
            Call quitter(0)
            Exit Sub
        End If
        grd.Row = CL_liste.pointeur
        grd.TopRow = grd.Row
    Else
        grd.Row = grd.FixedRows
    End If
    grd.col = 0
    grd.ColSel = grd.Cols - 1
    
End Sub

Private Sub initialiser()

    Dim tabstr() As String, s As String
    Dim nb_label As Integer, nb_ligne As Integer, first_col As Integer
    Dim icol As Integer, ilig As Integer, I As Integer, n As Integer
    Dim ajout As Integer
    Dim h_grd As Long, h_tot As Long, left As Long, intervalle As Long
    Dim lg_grd As Long, lg_tot As Long, lg_max As Long, lg As Long
    Dim lg_bouton As Long, hauteur As Long
    Dim lg_titre As Long, lg_bouton1 As Long
    Dim OkAfficher As Boolean
    Dim nb_newligne As Integer
    Dim ajout_rech As Integer
    Dim imot As Integer, smot As String
    
    g_nbbouton = 0
    On Error Resume Next
    g_nbbouton = UBound(CL_liste.boutons) + 1
    On Error GoTo 0
    
    Call FRM_ResizeForm(Me, 0, 0)
    
    ' Titre
    frmListe.Caption = CL_liste.prmfrm.titre
    lg_titre = FRM_LargeurTexte(Me, frmListe, CL_liste.prmfrm.titre) + 255
    
    ' Lignes
    nb_ligne = 0
    On Error Resume Next
    nb_ligne = UBound(CL_liste.lignes) + 1
    On Error GoTo 0
    If nb_ligne = 0 And g_nbbouton < 2 Then
        Call MsgBox("Aucune ligne à afficher.", vbExclamation + vbOKOnly, "")
        Call quitter(g_nbbouton - 1)
        Exit Sub
    End If
    
    ' Initialisation du grd
    grd.SelectionMode = flexSelectionByRow
    grd.FocusRect = flexFocusNone
    grd.ScrollBars = flexScrollBarNone
    grd.FixedCols = 0
    grd.BackColorBkg = grd.BackColor
    grd.Rows = nb_ligne
    If nb_ligne > 0 Then
        grd.FixedRows = 0
    End If
    
    If CL_liste.prmfrm.multi_select Then
        first_col = 1
    Else
        first_col = 0
    End If
    
    ' Labels
    nb_label = 0
    On Error Resume Next
    nb_label = UBound(CL_liste.prmfrm.labels) + 1
    On Error GoTo 0
    If nb_label > 0 Then
        grd.Cols = nb_label
        grd.Rows = grd.Rows + 1
        grd.FixedRows = 1
    End If
    For icol = 0 To nb_label - 1
        Call afficher_ligne(0, icol + first_col, CL_liste.prmfrm.labels(icol))
    Next icol
        
    ' Zone de recherche
    If CL_liste.prmfrm.b_txtRecherche Then
        Me.TxtRecherche.Visible = True
        Me.TxtRecherche.Text = CL_liste.prmfrm.s_txtRecherche
        TxtRecherche.SelStart = Len(TxtRecherche.Text)
        Me.LblRecherche.Visible = True
        Me.cmdHelp.Visible = True
    End If
    
    ' Affichage des lignes
    nb_newligne = 0
    For ilig = 0 To nb_ligne - 1
        ' Voir si la condition de recherche est remplie
        If Not CL_liste.prmfrm.b_txtRecherche Then
            OkAfficher = True
        ElseIf CL_liste.lignes(ilig).texte = "<Nouvelle>" Then
            OkAfficher = True
        ElseIf CL_liste.lignes(ilig).texte = "<Nouveau>" Then
            OkAfficher = True
        ElseIf CL_liste.prmfrm.multi_select Then
            ' on affiche celle déjà sélectionnée
            If CL_liste.lignes(ilig).selected Then
                OkAfficher = True
            Else
                GoTo Suite
            End If
        Else
Suite:
            If Me.TxtRecherche.Text = "" Then
                OkAfficher = True
            Else
                OkAfficher = True
                For imot = 0 To STR_GetNbchamp(Me.TxtRecherche.Text, " ")
                    smot = STR_GetChamp(Me.TxtRecherche.Text, " ", imot)
                    smot = Trim(STR_Phonet(smot))
                    If smot <> "" Then
                        If InStr(STR_Phonet(CL_liste.lignes(ilig).texte), smot) = 0 Then
                            OkAfficher = False
                            Exit For
                        End If
                    End If
                Next imot
            End If
        End If
        CL_liste.lignes(ilig).iGrd = -1
        If OkAfficher Then
            Call STR_Decouper(CL_liste.lignes(ilig).texte, tabstr)
            For icol = 0 To UBound(tabstr)
                Call afficher_ligne(nb_newligne + grd.FixedRows, icol + first_col, tabstr(icol))
                CL_liste.lignes(ilig).iGrd = nb_newligne + grd.FixedRows
            Next icol
            If first_col = 1 Then
                grd.Row = grd.FixedRows + nb_newligne
                grd.col = 0
                If CL_liste.lignes(ilig).selected Then
                    If Not CL_liste.lignes(ilig).fmodif Then
                        Set grd.CellPicture = imglst.ListImages(IMG_COCHE_GRIS).Picture
                        grd.col = 1
                        grd.CellForeColor = P_GRIS_FONCE
                    Else
                        Set grd.CellPicture = imglst.ListImages(IMG_COCHE).Picture
                    End If
                Else
                    Set grd.CellPicture = imglst.ListImages(IMG_SANS_COCHE).Picture
                End If
            End If
            nb_newligne = nb_newligne + 1
        End If
    Next ilig
    
    grd.Rows = nb_newligne
' A VOIR AVEC RV
    If nb_ligne > 0 Then
        grd.FixedRows = 0
    End If

    ' Zone de recherche : si > 10 => zone de recherche en auto
    If Not CL_liste.prmfrm.b_txtRecherche Then
        If nb_newligne > 10 Then
            CL_liste.prmfrm.b_txtRecherche = True
            Me.TxtRecherche.Visible = True
            Me.TxtRecherche.Text = ""
            CL_liste.prmfrm.s_txtRecherche = ""
            Me.LblRecherche.Visible = True
            Me.cmdHelp.Visible = True
            Me.TxtRecherche.SetFocus
        End If
    End If
    
    ' Boutons
    lg_bouton = 0
    For I = 0 To g_nbbouton - 1
        If I > 0 Then Load cmd(I)
        cmd(I).Visible = True
        If CL_liste.boutons(I).image <> "" Then
            cmd(I).Picture = CM_LoadPicture(CL_liste.boutons(I).image)
            cmd(I).Caption = ""
            cmd(I).ToolTipText = CL_liste.boutons(I).libelle
        Else
            cmd(I).Picture = LoadPicture("")
            cmd(I).Caption = CL_liste.boutons(I).libelle
            cmd(I).ToolTipText = ""
        End If
        If CL_liste.boutons(I).largeur > 0 Then
            cmd(I).width = CL_liste.boutons(I).largeur
        End If
        lg_bouton = lg_bouton + cmd(I).width
    Next I
    lg_bouton1 = lg_bouton
    If lg_bouton > 0 Then
        lg_bouton = 255 + lg_bouton + 255 + (g_nbbouton - 1) * 510
    End If
    If nb_ligne = 0 Then
        cmd(0).Visible = False
    End If
    
    ' Reglage hauteur
    If CL_liste.prmfrm.hauteur_max > 0 Then
        h_grd = CL_liste.prmfrm.hauteur_max - 300 - 255 - 255 - cmd(0).Height - 255 - 255
        g_nblig_visible = (h_grd - 68) / grd.RowHeight(grd.FixedRows)
    Else
        g_nblig_visible = -CL_liste.prmfrm.hauteur_max
        If nb_label > 0 Then
            g_nblig_visible = g_nblig_visible + 1
        End If
    End If
    If nb_ligne > 0 Then
        h_grd = 68 + (g_nblig_visible * grd.RowHeight(grd.FixedRows))
    Else
        h_grd = 0
    End If
    
    grd.Height = h_grd
    ajout = 0
    If CL_liste.prmfrm.b_txtRecherche Then
        ajout = Me.LblRecherche.Height + 20
        ajout_rech = Me.LblRecherche.Height + 20
        If CL_liste.prmfrm.gerer_tous_rien Then
            ajout = ajout + cmdTR(0).Height + 100
        End If
    Else
        If CL_liste.prmfrm.gerer_tous_rien Then
            ajout = cmdTR(0).Height + 100
        End If
    End If
    
    grd.left = 255
    If CL_liste.prmfrm.gerer_tous_rien Then
        cmdTR(0).Top = 512 + ajout_rech
        cmdTR(0).left = grd.left
        cmdTR(0).Visible = True
        cmdTR(1).Top = cmdTR(0).Top
        cmdTR(1).left = cmdTR(0).left + cmdTR(0).width + 60
        cmdTR(1).Visible = True
    Else
        cmdTR(0).Visible = False
        cmdTR(1).Visible = False
    End If
    
    
    grd.Top = 512 + ajout
    h_tot = 255 + 255 + h_grd + 255 + ajout
    frmListe.Height = h_tot
    
    If CL_liste.prmfrm.b_txtRecherche Then
        Me.LblRecherche.left = grd.left
        Me.TxtRecherche.left = Me.LblRecherche.left + Me.LblRecherche.width
    End If
    
    ' Reglage largeur
    lg = lg_titre
    If lg < lg_bouton Then
        lg = lg_bouton
    End If
    If nb_ligne > 0 Then
        grd.Row = grd.FixedRows
        grd.col = grd.Cols - 1
    End If
    lg_grd = grd.CellLeft + grd.ColWidth(grd.Cols - 1) + 30
    If nb_ligne > g_nblig_visible Then
        grd.ScrollBars = flexScrollBarVertical
        lg_grd = lg_grd + 240
    End If
    grd.width = lg_grd
    If lg < lg_grd Then
        lg = lg_grd
    Else
        grd.ColWidth(grd.Cols - 1) = lg - grd.CellLeft - 30
        grd.width = lg
    End If
    
    If CL_liste.prmfrm.b_txtRecherche Then
        Me.TxtRecherche.width = (grd.left + grd.width) - Me.TxtRecherche.left
        Me.TxtRecherche.width = Me.TxtRecherche.width - Me.cmdHelp.width - 50
        Me.TxtRecherche.left = Me.LblRecherche.left + Me.LblRecherche.width
        Me.cmdHelp.left = Me.TxtRecherche.left + Me.TxtRecherche.width + 50
        Me.cmdHelp.Top = Me.TxtRecherche.Top - 25
    End If
    
    lg_bouton = lg + 512
    lg_tot = lg + 512
    frmListe.width = lg_tot
    frmFct.width = lg_tot
    
    
    ' Position de la form
    If CL_liste.prmfrm.x <> 0 Then
        Me.left = CL_liste.prmfrm.x
    Else
        Me.left = (Screen.width - frmListe.width) / 2
    End If
    If Me.left + frmListe.width > Screen.width Then
        Me.left = Screen.width - frmListe.width
    End If
    If CL_liste.prmfrm.y <> 0 Then
        If CL_liste.prmfrm.y < 0 Then
            Me.Top = -CL_liste.prmfrm.y - frmListe.Height
        Else
            Me.Top = CL_liste.prmfrm.y
        End If
    Else
        Me.Top = (Screen.Height - frmListe.Height) / 2
    End If
    If Me.Top + frmListe.Height > Screen.Height Then
        Me.Top = Screen.Height - frmListe.Height
    End If
    frmListe.Top = 0
    frmListe.left = 0
    frmListe.ZOrder 0
    frmFct.Top = frmListe.Height - 150
    
    ' Positionnement des boutons
    If g_nbbouton = 1 Then
        cmd(0).left = (frmFct.width - 510 - cmd(0).width) / 2
    Else
        intervalle = (frmFct.width - 510 - lg_bouton1) / (g_nbbouton - 1)
        left = 255
        For I = 0 To g_nbbouton - 1
            cmd(I).left = left
            left = left + cmd(I).width + intervalle
        Next I
    End If
    
    
Fin:
    cmd(0).Default = True
    Me.MousePointer = 0
    Me.width = lg_tot + 100
    Me.Height = frmFct.Top + frmFct.Height + 300
    For I = 0 To grd.Cols - 1
        grd.ColAlignment(I) = flexAlignLeftCenter
    Next I
    
    ' Ligne de départ
    Call init_ligne_courante
    grd.col = 0
    grd.ColSel = grd.Cols - 1

    grd.SetFocus
    g_mode_saisie = True
    Exit Sub
    
End Sub

Private Sub quitter(ByVal v_index As Integer)

    Dim Bsel As Boolean
    Dim I As Integer, TailleTab As Integer
    
    CL_liste.retour = v_index
    
    ' Liste à choix multiples
    If CL_liste.prmfrm.multi_select Then
        ' Validation
        If v_index = 0 Then
            ' Si aucune sélection -> renvoie la ligne courante
            If CL_liste.prmfrm.renvoyer_courant Then
                Bsel = False
                On Error Resume Next
                TailleTab = UBound(CL_liste.lignes)
                For I = 0 To TailleTab
                    If CL_liste.lignes(I).selected Then
                        Bsel = True
                        Exit For
                    End If
                Next I
                If Not Bsel Then
                    On Error Resume Next
                    TailleTab = UBound(CL_liste.lignes)
                    For I = 0 To TailleTab
                        If CL_liste.lignes(I).iGrd = grd.Row Then
                            CL_liste.lignes(I).selected = True
                            Exit For
                        End If
                    Next I
                End If
            End If
        Else
            CL_liste.pointeur = grd.Row - grd.FixedRows
            On Error Resume Next
            TailleTab = UBound(CL_liste.lignes)
            For I = 0 To TailleTab
                If CL_liste.lignes(I).iGrd = grd.Row Then
                    CL_liste.pointeur = I
                    Exit For
                End If
            Next I
        End If
    Else
        ' Rechercher le bon pointeur
        CL_liste.pointeur = grd.Row - grd.FixedRows
        On Error Resume Next
        TailleTab = UBound(CL_liste.lignes)
        For I = 0 To TailleTab
            If CL_liste.lignes(I).iGrd = grd.Row Then
                CL_liste.pointeur = I
                Exit For
            End If
        Next I
    End If
    
    If CL_liste.prmfrm.reste_cachée Then
        Me.Hide
    Else
        Unload Me
    End If
    
    
End Sub

Private Sub reinitialiser()

    Dim tabstr() As String, smot As String
    Dim first_col As Integer, icol As Integer, nb_ligne As Integer
    Dim n As Integer, ilig As Integer, imot As Integer
    Dim OkAfficher As Boolean
    Dim nb_newligne As Integer
    
'    If grd.Rows = 0 Then
'        Exit Sub
'    End If
    
    nb_ligne = 0
    On Error Resume Next
    nb_ligne = UBound(CL_liste.lignes) + 1
    On Error GoTo 0
    nb_newligne = grd.Rows - grd.FixedRows
    If grd.Rows - grd.FixedRows <> nb_ligne Then
        n = grd.Rows - grd.FixedRows
        For ilig = n To UBound(CL_liste.lignes)
            If Not CL_liste.prmfrm.b_txtRecherche Then
                OkAfficher = True
            ElseIf CL_liste.lignes(ilig).texte = "<Nouvelle>" Then
                OkAfficher = True
            ElseIf CL_liste.lignes(ilig).texte = "<Nouveau>" Then
                OkAfficher = True
            ElseIf CL_liste.prmfrm.multi_select Then
                ' on affiche celle déjà sélectionnée
                If CL_liste.lignes(ilig).selected Then
                    OkAfficher = True
                End If
            End If
            If Me.TxtRecherche.Text = "" Then
                OkAfficher = True
            Else
                OkAfficher = True
                For imot = 0 To STR_GetNbchamp(Me.TxtRecherche.Text, " ")
                    smot = STR_GetChamp(Me.TxtRecherche.Text, " ", imot)
                    smot = Trim(STR_Phonet(smot))
                    If smot <> "" Then
                        If InStr(STR_Phonet(CL_liste.lignes(ilig).texte), smot) = 0 Then
                            OkAfficher = False
                            Exit For
                        End If
                    End If
                Next imot
            End If
            CL_liste.lignes(ilig).iGrd = -1
            If OkAfficher Then
                grd.Rows = grd.Rows + 1
                If CL_liste.prmfrm.multi_select Then
                    first_col = 1
                Else
                    first_col = 0
                End If
                Call STR_Decouper(CL_liste.lignes(ilig).texte, tabstr)
                For icol = 0 To UBound(tabstr)
                    Call afficher_ligne(nb_newligne + grd.FixedRows, icol + first_col, tabstr(icol))
                    CL_liste.lignes(ilig).iGrd = nb_newligne + grd.FixedRows
'                    Call afficher_ligne(ilig + grd.FixedRows, icol + first_col, tabstr(icol))
'                    CL_liste.lignes(ilig).iGrd = ilig + grd.FixedRows
                Next icol
                If first_col = 1 Then
                    grd.Row = grd.FixedRows + nb_newligne
'                    grd.Row = grd.FixedRows + ilig
                    grd.col = 0
                    If CL_liste.lignes(ilig).selected Then
                        Set grd.CellPicture = imglst.ListImages(IMG_COCHE).Picture
                    Else
                        Set grd.CellPicture = imglst.ListImages(IMG_SANS_COCHE).Picture
                    End If
                End If
                nb_newligne = nb_newligne + 1
            End If
        Next ilig
        If grd.Rows - grd.FixedRows > g_nblig_visible Then
            grd.ScrollBars = flexScrollBarVertical
            grd.width = grd.width + 240
        End If
        grd.Row = grd.Rows - 1
        If grd.Row >= grd.TopRow + g_nblig_visible Then
            grd.TopRow = grd.Row - g_nblig_visible + 1
        End If
        grd.col = 0
        grd.ColSel = grd.Cols - 1
    Else
        Call init_ligne_courante
    End If
    grd.SetFocus

End Sub

Private Sub cmd_Click(Index As Integer)

    If Screen.ActiveControl.Name = "TxtRecherche" Then
        Call AfficherCacher
    Else
        Call quitter(Index)
    End If
End Sub

Private Function AfficherCacher() As Integer
    Dim tabstr() As String, s As String
    Dim nb_label As Integer, nb_ligne As Integer, first_col As Integer
    Dim icol As Integer, ilig As Integer, I As Integer, n As Integer
    Dim ajout As Integer
    Dim h_grd As Long, h_tot As Long, left As Long, intervalle As Long
    Dim lg_grd As Long, lg_tot As Long, lg_max As Long, lg As Long
    Dim lg_bouton As Long, hauteur As Long
    Dim lg_titre As Long, lg_bouton1 As Long
    Dim OkAfficher As Boolean
    Dim nb_newligne As Integer
    Dim imot As Integer, smot As String
    
    ' Lignes
    nb_ligne = 0
    On Error Resume Next
    nb_ligne = UBound(CL_liste.lignes) + 1
    On Error GoTo 0
    If nb_ligne = 0 And g_nbbouton < 2 Then
        Call MsgBox("Aucune ligne à afficher.", vbExclamation + vbOKOnly, "")
        Call quitter(g_nbbouton - 1)
        Exit Function
    End If
    
    nb_newligne = 0
    grd.Rows = nb_ligne
    If nb_ligne > 0 Then
        grd.FixedRows = 0
    End If
    
    If CL_liste.prmfrm.multi_select Then
        first_col = 1
    Else
        first_col = 0
    End If
    
    ' Affichage des lignes
    For ilig = 0 To nb_ligne - 1
        ' Voir si la condition de recherche est remplie
        If Not CL_liste.prmfrm.b_txtRecherche Then
            OkAfficher = True
        ElseIf CL_liste.lignes(ilig).texte = "<Nouvelle>" Then
            OkAfficher = True
        ElseIf CL_liste.lignes(ilig).texte = "<Nouveau>" Then
            OkAfficher = True
        ElseIf CL_liste.prmfrm.multi_select Then
            ' on affiche celle déjà sélectionnée
            If CL_liste.lignes(ilig).selected Then
                OkAfficher = True
            Else
                GoTo Suite
            End If
        Else
Suite:
            If Me.TxtRecherche.Text = "" Then
                OkAfficher = True
            Else
                OkAfficher = True
                For imot = 0 To STR_GetNbchamp(Me.TxtRecherche.Text, " ")
                    smot = STR_GetChamp(Me.TxtRecherche.Text, " ", imot)
                    smot = Trim(STR_Phonet(smot))
                    If smot <> "" Then
                        If InStr(STR_Phonet(CL_liste.lignes(ilig).texte), smot) = 0 Then
                            OkAfficher = False
                            Exit For
                        End If
                    End If
                Next imot
            End If
        End If
        CL_liste.lignes(ilig).iGrd = -1
        If OkAfficher Then
            Call STR_Decouper(CL_liste.lignes(ilig).texte, tabstr)
            For icol = 0 To UBound(tabstr)
                Call afficher_ligne(nb_newligne + grd.FixedRows, icol + first_col, tabstr(icol))
                CL_liste.lignes(ilig).iGrd = nb_newligne + grd.FixedRows
            Next icol
            If first_col = 1 Then
                grd.Row = grd.FixedRows + nb_newligne
                grd.col = 0
                If CL_liste.lignes(ilig).selected Then
                    If Not CL_liste.lignes(ilig).fmodif Then
                        Set grd.CellPicture = imglst.ListImages(IMG_COCHE_GRIS).Picture
                        grd.col = 1
                        grd.CellForeColor = P_GRIS_FONCE
                    Else
                        Set grd.CellPicture = imglst.ListImages(IMG_COCHE).Picture
                    End If
                Else
                    Set grd.CellPicture = imglst.ListImages(IMG_SANS_COCHE).Picture
                End If
            End If
            nb_newligne = nb_newligne + 1
        End If
    Next ilig
    
    ' Initialisation du grd
    grd.Rows = nb_newligne
    If nb_newligne > 0 Then
        grd.FixedRows = 0
    End If
    ' Ligne de départ
    Call init_ligne_courante
    grd.col = 0
    grd.ColSel = grd.Cols - 1

    On Error Resume Next
    grd.SetFocus
    
    AfficherCacher = nb_newligne
    
End Function

Private Sub cmdHelp_Click()
    Dim message As String
    
    message = "Vous pouvez rechercher sur plusieurs mots en les séparant par des espaces"
    message = message & Chr(13) & Chr(10)
    message = message & "ex : adjoint cadre (recherche les lignes qui contiennent le mot 'adjoint' ET le mot 'cadre'"
    MsgBox message

End Sub

Private Sub cmdTR_Click(Index As Integer)

    Dim bcoche As Boolean
    Dim i_img As Integer, lig As Integer
    
    Select Case Index
    Case CMD_TOUS
        bcoche = True
        i_img = IMG_COCHE
    Case CMD_AUCUN
        bcoche = False
        i_img = IMG_SANS_COCHE
    End Select
    
    For lig = grd.FixedRows To grd.Rows - 1
        CL_liste.lignes(lig - grd.FixedRows).selected = bcoche
        grd.Row = lig
        grd.col = 0
        Set grd.CellPicture = imglst.ListImages(i_img).Picture
    Next lig
    
    grd.SetFocus

End Sub

Private Sub Form_Activate()

    If g_form_active Then
        Call reinitialiser
        Exit Sub
    End If
    
    g_form_active = True
    Call initialiser
    
    If CL_liste.prmfrm.b_txtRecherche Then
        Me.TxtRecherche.SetFocus
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim nomchm As String, nomtopic As String
    Dim I As Integer
    
    If KeyCode = 13 And Screen.ActiveControl.Name = "TxtRecherche" Then
    End If
        
    If Shift = vbAltMask Then
        For I = 0 To g_nbbouton - 1
            If KeyCode = CL_liste.boutons(I).raccourci_alt Then
                KeyCode = 0
                Call quitter(I)
                Exit Sub
            End If
        Next I
        If KeyCode = vbKeyH Then
            KeyCode = 0
            If CL_liste.prmfrm.nomhelp <> "" Then
                If STR_GetNbchamp(CL_liste.prmfrm.nomhelp, ";") = 1 Then
                    nomchm = CL_liste.prmfrm.nomhelp
                    nomtopic = ""
                Else
                    nomchm = STR_GetChamp(CL_liste.prmfrm.nomhelp, ";", 0)
                    nomtopic = STR_GetChamp(CL_liste.prmfrm.nomhelp, ";", 1)
                End If
                Call HtmlHelp(0, nomchm, HH_DISPLAY_TOPIC, nomtopic)
            End If
        End If
    Else
        For I = 0 To g_nbbouton - 1
            If KeyCode = CL_liste.boutons(I).raccourci_touche Then
                KeyCode = 0
                Call quitter(I)
                Exit Sub
            End If
        Next I
    End If

End Sub

Private Sub Form_Load()

    g_form_active = False
    g_mode_saisie = False
    Call FRM_ResizeForm(Me, 0, 0)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Call quitter(cmd.Count - 1)
    End If
    
End Sub

Private Sub grd_Click()

    If Not g_mode_saisie Then Exit Sub
    
    If CL_liste.prmfrm.multi_select Then
        Call changer_etat_selection
    End If

End Sub

Private Sub grd_DblClick()
    
    Dim lig As Integer, ilig As Integer
    
    If Not CL_liste.prmfrm.multi_select Then
        lig = grd.Row - grd.FixedRows
        For ilig = 0 To UBound(CL_liste.lignes)
            'If ilig = lig Then
            If CL_liste.lignes(ilig).iGrd = lig Then
                CL_liste.lignes(ilig).selected = True
            Else
                CL_liste.lignes(ilig).selected = False
            End If
        Next ilig
        Call quitter(0)
    End If
    
End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim I As Integer

    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        Call quitter(0)
    ElseIf KeyCode = vbKeySpace Then
        KeyCode = 0
        If CL_liste.prmfrm.multi_select Then
            Call changer_etat_selection
        End If
    End If

End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)

    Call chercher_ligne_avec_first_car(KeyAscii)
    
End Sub



Private Sub TxtRecherche_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim nb As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = 13 Then
    ElseIf KeyCode = vbKeyPageDown Then
        grd.SetFocus
        SendKeys "{PGDN}", True
        TxtRecherche.SetFocus
        TxtRecherche.SelStart = Len(TxtRecherche.Text)
        KeyCode = 0
    ElseIf KeyCode = vbKeyPageUp Then
        grd.SetFocus
        SendKeys "{PGUP}", True
        TxtRecherche.SetFocus
        TxtRecherche.SelStart = Len(TxtRecherche.Text)
        KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
        On Error Resume Next
        grd.Row = grd.Row + 1
        grd.ColSel = grd.Cols - 1
        TxtRecherche.SetFocus
        TxtRecherche.SelStart = Len(TxtRecherche.Text)
    ElseIf KeyCode = vbKeyUp Then
        On Error Resume Next
        grd.Row = grd.Row - 1
        grd.ColSel = grd.Cols - 1
        TxtRecherche.SetFocus
        TxtRecherche.SelStart = Len(TxtRecherche.Text)
    ElseIf KeyCode = 116 Then
        KeyCode = 0
    Else
        nb = AfficherCacher()
        If nb > 1 Then
            If Len(Me.TxtRecherche.Text) = 1 Then
                ' se mettre sur le premier
                chercher_ligne_avec_first_car (Asc(Me.TxtRecherche.Text))
            End If
            TxtRecherche.SetFocus
            TxtRecherche.SelStart = Len(TxtRecherche.Text)
        End If
    End If
    TxtRecherche.SetFocus
    TxtRecherche.SelStart = Len(TxtRecherche.Text)


End Sub
