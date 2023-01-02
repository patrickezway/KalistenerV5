VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form PrmService 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Services"
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
      Height          =   7995
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8265
      Begin ComctlLib.ProgressBar pgBar 
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   7560
         Visible         =   0   'False
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
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
         Left            =   7560
         Picture         =   "PrmService.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Aucun"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   380
      End
      Begin VB.ComboBox CmbNiveau 
         Height          =   315
         Left            =   4800
         TabIndex        =   9
         Text            =   "Tous"
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox TxtRecherche 
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   480
         Width           =   2775
      End
      Begin ComctlLib.TreeView tv 
         Height          =   6885
         Left            =   195
         TabIndex        =   5
         Top             =   1035
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   12144
         _Version        =   327682
         Indentation     =   2
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "img"
         Appearance      =   1
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
         Caption         =   "Rechercher sur"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin ComctlLib.ImageList img 
         Left            =   6120
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   22
         ImageHeight     =   22
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   8
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":0359
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":0BAB
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":147D
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":1D4F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":2621
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":2EF3
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":3745
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":3D77
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblDepl 
         BackColor       =   &H000080FF&
         Caption         =   "Cliquez sur le nouveau service ou cliquez ici pour Annuler"
         Height          =   495
         Left            =   2370
         TabIndex        =   6
         Top             =   30
         Visible         =   0   'False
         Width           =   3585
      End
   End
   Begin VB.Frame frmFct 
      BackColor       =   &H00808080&
      Height          =   800
      Left            =   0
      TabIndex        =   0
      Top             =   7920
      Width           =   8265
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
         Left            =   7410
         Picture         =   "PrmService.frx":4359
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Quitter"
         Top             =   200
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
         Index           =   0
         Left            =   240
         Picture         =   "PrmService.frx":4912
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Sélectionner"
         Top             =   200
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
         Index           =   2
         Left            =   1020
         Picture         =   "PrmService.frx":4D6B
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprimer"
         Top             =   200
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.Label LbldetailSRV 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   5655
      End
   End
   Begin VB.Menu mnuFct 
      Caption         =   "mnuFct"
      Visible         =   0   'False
      Begin VB.Menu mnuCreerS 
         Caption         =   "&Créer un service"
      End
      Begin VB.Menu mnuSepCrS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModS 
         Caption         =   "&Modifier le service"
      End
      Begin VB.Menu mnuSuppS 
         Caption         =   "&Supprimer le service"
      End
      Begin VB.Menu mnuSepMSS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreerP 
         Caption         =   "C&réer un poste"
      End
      Begin VB.Menu mnuCreerPi 
         Caption         =   "Créer une p&ièce"
      End
      Begin VB.Menu mnuSepCrP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModPi 
         Caption         =   "&Modifier la pièce"
      End
      Begin VB.Menu mnuSuppPi 
         Caption         =   "&Supprimer la pièce"
      End
      Begin VB.Menu mnuPosteResp 
         Caption         =   "Poste responsable"
      End
      Begin VB.Menu mnuLibPoste 
         Caption         =   "Libellé du poste"
      End
      Begin VB.Menu mnuPosteRemplace 
         Caption         =   "Niveau de remplacement"
      End
      Begin VB.Menu mnuSuppP 
         Caption         =   "&Supprimer le poste"
      End
      Begin VB.Menu mnuTrsP 
         Caption         =   "&Transférer sur un autre poste"
      End
      Begin VB.Menu mnuSepSuppP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDepl 
         Caption         =   "&Déplacer dans un autre service"
      End
      Begin VB.Menu mnuSepDepl 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVoirPers 
         Caption         =   "&Voir les personnes"
      End
      Begin VB.Menu mnuSepVoirPers 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCrPers 
         Caption         =   "&Créer une personne"
      End
      Begin VB.Menu mnuAjPers 
         Caption         =   "&Ajouter une personne existante"
      End
      Begin VB.Menu mnuSepCrPers 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModPers 
         Caption         =   "&Modifier les caractéristiques de la personne"
      End
      Begin VB.Menu mnuSepModPers 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quitter"
      End
   End
End
Attribute VB_Name = "PrmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CMD_OK = 0
Private Const CMD_IMPRIMER = 2
Private Const CMD_QUITTER = 1

Private Const IMG_SRV = 1
Private Const IMG_POSTE = 2
Private Const IMG_SRV_SEL = 3
Private Const IMG_POSTE_SEL = 4
Private Const IMG_SRV_SEL_NOMOD = 5
Private Const IMG_POSTE_SEL_NOMOD = 6
Private Const IMG_UTIL = 7
Private Const IMG_PIECE = 8

Private Const MODE_PARAM = 0
Private Const MODE_SELECT = 1
Private Const MODE_PARAM_PERS = 2

Private g_mode_acces As Integer
Private g_smode As String

Private g_ya_niveau As Boolean

Private g_plusieurs As Boolean
Private g_stype As String
Private g_prmpers As Boolean
Private g_numserv As Long
Private g_ouvrir As String
Private g_sret As String

Private g_crfct_autor As Boolean
Private g_crutil_autor As Boolean
Private g_modutil_autor As Boolean
Private g_crspm_autor As Boolean
Private g_modspm_autor As Boolean
Private g_supspm_autor As Boolean
Private g_crpiece_autor As Boolean
Private g_modpiece_autor As Boolean
Private g_suppiece_autor As Boolean

Private g_lignes() As CL_SLIGNE

Private g_node_crt As Long
Private g_pos_depl As Long

Private g_node As Integer
Private g_expand As Boolean
Private g_button As Integer
Private g_mode_saisie As Boolean
Private g_form_active As Boolean


'V_smode :      "C" --> Quand on vient de prmClasseur

Public Function AppelFrm(ByVal v_stitre As String, _
                         ByVal v_smode As String, _
                         ByVal v_bplusieurs As Boolean, _
                         ByVal v_ssite As String, _
                         ByVal v_stype As String, _
                         ByVal v_prmpers As Boolean, _
                         Optional v_ouvrir As String = "") As String

    If v_smode = "M" Then
        g_smode = v_smode
        g_mode_acces = MODE_PARAM
        g_stype = v_stype
        g_numserv = 0
    ElseIf v_smode = "S" Or v_smode = "C" Then
        g_smode = v_smode
        g_mode_acces = MODE_SELECT
        g_stype = v_stype
        g_numserv = 0
    ElseIf v_smode = "P" Then
        g_smode = v_smode
        g_mode_acces = MODE_PARAM_PERS
        If left$(v_stype, 1) = "S" Then
            g_numserv = Mid$(v_stype, 2)
        Else
            g_numserv = 0
        End If
        g_stype = "SP"
    End If
    g_plusieurs = v_bplusieurs
'    g_ssite = v_ssite
    g_prmpers = v_prmpers
    g_ouvrir = v_ouvrir
    
    frm.Caption = v_stitre
    
    Me.Show 1
    
    AppelFrm = g_sret
    
End Function

Private Sub activer_depl()

    g_node_crt = tv.SelectedItem.Index
    g_pos_depl = g_node_crt
    lblDepl.Caption = "Cliquez sur le service destination ou cliquez ici pour ANNULER l'opération"
    lblDepl.BackColor = P_ORANGE
    lblDepl.Visible = True
    lblDepl.tag = "D"
    
End Sub

Private Sub activer_trs_poste()

    g_node_crt = tv.SelectedItem.Index
    g_pos_depl = g_node_crt
    lblDepl.Caption = "Cliquez sur le poste de remplacement ou cliquez ici pour ANNULER l'opération"
    lblDepl.BackColor = P_JAUNE
    lblDepl.Visible = True
    lblDepl.tag = "T"
    
End Sub

Private Function afficher_liste() As Integer

    Dim sql As String, s As String, sfct As String, lib As String
    Dim stag As String, libNiveau As String
    Dim fmodif As Boolean, afficher As Boolean, trouve As Boolean
    Dim img As Integer, I As Integer, nsel As Integer, n As Integer, j As Integer
    Dim numsrv As Long, lnb As Long, num As Long
    Dim rs As rdoResultset
    Dim nd As Node, ndp As Node, ndp_sav As Node
    
    g_mode_saisie = False
    
    nsel = -1
    On Error Resume Next
    nsel = UBound(g_lignes)
    On Error GoTo 0
    
    tv.Nodes.Clear
    
    sql = "select L_Num, L_Nom from Laboratoire"
    If Odbc_RecupVal(sql, num, lib) = P_ERREUR Then
        afficher_liste = P_ERREUR
        Exit Function
    End If
    Set ndp = tv.Nodes.Add(, , "L" & num, lib)
    ndp.Expanded = True
    
    ' Les services
    sql = "select SRV_Num, SRV_LNum, SRV_NumPere, SRV_Nom, SRV_NivsNum from Service " _
        & " where SRV_Actif=true" _
        & " and SRV_Numpere=0" _
        & " order by SRV_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        afficher_liste = P_ERREUR
        Exit Function
    End If
    
    If Not rs.EOF Then
        rs.MoveLast
        pgBar.Visible = True
        pgBar.Max = rs.RowCount
        pgBar.Value = 0
        rs.MoveFirst
    End If
    
    While Not rs.EOF
        pgBar.Value = pgBar.Value + 1

        lib = rs("SRV_Nom").Value
        libNiveau = recup_lib_niveau(rs("SRV_NivsNum").Value)
        If libNiveau <> "" Then
            lib = lib & " (" & libNiveau & ")"
        End If
        Set nd = tv.Nodes.Add(ndp, _
                               tvwChild, _
                               "S" & rs("SRV_Num").Value, _
                               lib, _
                               IMG_SRV, _
                               IMG_SRV)
        sql = "select count(*) from Service " _
            & " where SRV_Actif=true" _
            & " and SRV_Numpere=" & rs("SRV_Num").Value
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            afficher_liste = P_ERREUR
            Exit Function
        End If
        If lnb = 0 And g_smode <> "C" Then
            sql = "select count(*) from Poste " _
                & " where PO_Actif=true" _
                & " and PO_SRVNum=" & rs("SRV_Num").Value
            If Odbc_Count(sql, lnb) = P_ERREUR Then
                afficher_liste = P_ERREUR
                Exit Function
            End If
        End If
        If lnb = 0 Then
            nd.tag = True & "|" & True
        Else
            nd.tag = True & "|" & False
            Set nd = tv.Nodes.Add(nd, _
                               tvwChild, _
                               , _
                               "A charger")
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    pgBar.Visible = False
    
    ' Met en évidence les noeuds 'retenus'
    If g_mode_acces = MODE_SELECT And g_plusieurs Then
        For I = 0 To nsel
            n = STR_GetNbchamp(g_lignes(I).texte, ";")
            s = STR_GetChamp(g_lignes(I).texte, ";", n - 1)
            If left$(s, 1) = "S" Then
                fmodif = g_lignes(I).fmodif
                If fmodif Then
                    img = IMG_SRV_SEL
                Else
                    img = IMG_SRV_SEL_NOMOD
                End If
            Else
                fmodif = g_lignes(I).fmodif
                If fmodif Then
                    img = IMG_POSTE_SEL
                Else
                    img = IMG_POSTE_SEL_NOMOD
                End If
            End If
            If TV_NodeExiste(tv, s, nd) = P_NON Then
                Call charger_arbor(s)
            End If
            If TV_NodeExiste(tv, s, nd) = P_OUI Then
                Set nd = tv.Nodes(s)
                nd.image = img
                stag = nd.tag
                Call STR_PutChamp(stag, "|", 0, fmodif)
                nd.tag = stag
                nd.SelectedImage = img
                Set ndp = nd.Parent
                While left$(ndp.key, 1) <> "L"
                    ndp.Expanded = True
                    Set ndp = ndp.Parent
                Wend
            End If
        Next I
    End If
    
    Call ouvrir_serv_poste
    
    tv.SetFocus
    g_mode_saisie = True
    
    Set nd = Nothing
    Set ndp = Nothing
    Set ndp_sav = Nothing
    
    afficher_liste = P_OK
    
End Function

Private Function afficher_liste2() As Integer

    Dim sql As String, s As String, codsite As String, nomsrv As String, sfct As String
    Dim trouve As Boolean, faff As Boolean, afficher As Boolean
    Dim img As Integer, I As Integer, nsel As Integer, n As Integer, j As Integer
    Dim mode As Integer
    Dim numsrv As Long, lnb As Long, num As Long
    Dim rs As rdoResultset
    Dim nd As Node, ndp As Node, ndp_sav As Node
    
    g_mode_saisie = False
    
    tv.Nodes.Clear
    
    If g_numserv = 0 Then
        sql = "select L_Num, L_Nom from Laboratoire"
        If Odbc_RecupVal(sql, num, codsite) = P_ERREUR Then
            afficher_liste2 = P_ERREUR
            Exit Function
        End If
        Set ndp = tv.Nodes.Add(, , "L" & num, codsite)
        ndp.Expanded = True
        ndp.Sorted = True
    Else
        sql = "select SRV_Nom from Service" _
            & " where SRV_Num=" & g_numserv
        If Odbc_RecupVal(sql, codsite) = P_ERREUR Then
            afficher_liste2 = P_ERREUR
            Exit Function
        End If
        Set ndp = tv.Nodes.Add(, , "S" & g_numserv, codsite)
        ndp.Expanded = True
    End If
    
    Call charger_service(g_numserv)
GoTo ici

    ' Les services
    sql = "select SRV_Num, SRV_LNum, SRV_NumPere, SRV_Nom from Service" _
        & " where SRV_Actif=true" _
        & " and SRV_Numpere=" & g_numserv _
        & " order by SRV_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        afficher_liste2 = P_ERREUR
        Exit Function
    End If
    If rs.EOF And g_numserv > 0 Then
        sql = "select count(*) from Poste " _
            & " where PO_Actif=true" _
            & " and PO_SRVNum=" & g_numserv
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            afficher_liste2 = P_ERREUR
            Exit Function
        End If
        If lnb = 0 Then
            ndp.tag = True & "|" & True
        Else
            ndp.tag = True & "|" & False
            Set nd = tv.Nodes.Add(ndp, _
                               tvwChild, _
                               , _
                               "A charger")
        End If
        rs.Close
    Else
        While Not rs.EOF
            Set nd = tv.Nodes.Add(ndp, _
                                   tvwChild, _
                                   "S" & rs("SRV_Num").Value, _
                                   rs("SRV_Nom").Value, _
                                   IMG_SRV, _
                                   IMG_SRV)
            sql = "select count(*) from Service " _
                & " where SRV_Actif=true" _
                & " and SRV_Numpere=" & rs("SRV_Num").Value
            If Odbc_Count(sql, lnb) = P_ERREUR Then
                afficher_liste2 = P_ERREUR
                Exit Function
            End If
            If lnb = 0 Then
                sql = "select count(*) from Poste " _
                    & " where PO_Actif=true" _
                    & " and PO_SRVNum=" & rs("SRV_Num").Value
                If Odbc_Count(sql, lnb) = P_ERREUR Then
                    afficher_liste2 = P_ERREUR
                    Exit Function
                End If
            End If
            If lnb = 0 Then
                nd.tag = True & "|" & True
            Else
                nd.tag = True & "|" & False
                Set nd = tv.Nodes.Add(nd, _
                                   tvwChild, _
                                   , _
                                   "A charger")
            End If
            rs.MoveNext
        Wend
        rs.Close
    End If
    
ici:
    Call ouvrir_serv_poste
    
    tv.SetFocus
    g_mode_saisie = True
    
    afficher_liste2 = P_OK
    
End Function

Private Function afficher_liste3(v_rech As String) As Integer

    Dim sql As String, s As String, sfct As String, lib As String
    Dim libNiveau As String, condRech As String, op As String
    Dim smot As String, stag As String
    Dim fmodif As Boolean, afficher As Boolean, trouve As Boolean
    Dim img As Integer, I As Integer, nsel As Integer, n As Integer, j As Integer
    Dim mode As Integer, imot As Integer
    Dim numsrv As Long, lnb As Long, num As Long
    Dim rs As rdoResultset, rs2 As rdoResultset
    Dim nd As Node, ndp As Node, ndp_sav As Node
    Dim tb()
    
    g_mode_saisie = False
    
    ' Les services
    sql = "select SRV_Num, SRV_LNum, SRV_NumPere, SRV_Nom, SRV_NivsNum from Service " _
        & " where SRV_Actif=true [CONDITION_NIVEAU] [CONDITION_RECHERCHE]" _
        & " order by SRV_Nom"
    
    If Me.CmbNiveau.ListIndex <= 0 Then
        sql = Replace(sql, "[CONDITION_NIVEAU]", "")
    Else
        If Odbc_SelectV("select Nivs_num from niveau_structure" _
                        & " Where Nivs_Num=" & Me.CmbNiveau.ItemData(Me.CmbNiveau.ListIndex), rs2) = P_ERREUR Then
            Call quitter
            Exit Function
        Else
            If rs2.EOF Then
                Call quitter
                Exit Function
            Else
                sql = Replace(sql, "[CONDITION_NIVEAU]", " and SRV_NivsNum=" & rs2("Nivs_Num") & " ")
            End If
        End If
    End If
    
    op = ""
    condRech = ""
    op = ""
    For imot = 0 To STR_GetNbchamp(Me.TxtRecherche.Text, " ")
        smot = Trim(STR_GetChamp(Me.TxtRecherche.Text, " ", imot))
        smot = LCase(STR_Phonet(smot))
        If smot <> "" Then
            condRech = condRech & op & " ( translate(lower(SRV_Nom),'éèàçù', 'eeacu') like '%" & smot & "%' or translate(lower(SRV_code),'éèàçù', 'eeacu') like '%" & smot & "%' )"
            op = " And "
        End If
        'Debug.Print condRech
    Next imot
    sql = Replace(sql, "[CONDITION_RECHERCHE]", " and (" & condRech & ")")
    
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        afficher_liste3 = P_ERREUR
        Exit Function
    End If
    
    If rs.EOF Then
        If Me.CmbNiveau.ListIndex <= 0 Then
            MsgBox "Aucun  trouvé"
        Else
            MsgBox "Aucun '" & Me.CmbNiveau.Text & "' trouvé"
        End If
        Exit Function
    End If
            
    nsel = -1
    On Error Resume Next
    nsel = UBound(g_lignes)
    On Error GoTo 0
    
    tv.Nodes.Clear
    
    sql = "select L_Num, L_Nom from Laboratoire"
    If Odbc_RecupVal(sql, num, lib) = P_ERREUR Then
        afficher_liste3 = P_ERREUR
        Exit Function
    End If
    p_NumLabo = num
    Set ndp = tv.Nodes.Add(, , "L" & p_NumLabo, lib)
    ndp.Expanded = True
    
    If Not rs.EOF Then
        rs.MoveLast
        pgBar.Visible = True
        pgBar.Max = rs.RowCount
        pgBar.Value = 0
        rs.MoveFirst
    End If
    
    While Not rs.EOF
        pgBar.Value = pgBar.Value + 1
        If TV_NodeExiste(tv, "S" & rs("SRV_Num").Value, nd) = P_OUI Then
            GoTo lab_suivant
        End If
        If rs("SRV_NumPere").Value = 0 Then
            Set ndp = tv.Nodes("L" & rs("SRV_LNum").Value)
        Else
            If TV_NodeExiste(tv, "S" & rs("SRV_NumPere").Value, nd) = P_NON Then
                Call ajouter_service(rs("SRV_NumPere").Value)
            End If
            Set ndp = tv.Nodes("S" & rs("SRV_NumPere").Value)
        End If
        
        lib = rs("SRV_Nom").Value
        libNiveau = recup_lib_niveau(rs("SRV_NivsNum").Value)
        If libNiveau <> "" Then
            lib = lib & " (" & libNiveau & ")"
        End If
        Set nd = tv.Nodes.Add(ndp, _
                               tvwChild, _
                               "S" & rs("SRV_Num").Value, _
                               lib, _
                               IMG_SRV, _
                               IMG_SRV)
        nd.tag = True & "|" & True
lab_suivant:
        rs.MoveNext
    Wend
    rs.Close
    pgBar.Visible = False
    
    ' on les ouvre tous
    For n = 1 To tv.Nodes.Count
        Set ndp = tv.Nodes(n)
        ' Les postes
        If g_smode <> "C" Then
            If left$(ndp.key, 1) = "S" Then
                If ndp.Children > 0 Then
                    sql = "select PO_Num, PO_Libelle, FT_Libelle" _
                        & " from Poste, FctTrav" _
                        & " where FT_Num=PO_FTNum" _
                        & " and PO_Actif=true" _
                        & " and PO_SRVNum=" & Mid$(ndp.key, 2) _
                        & " order by PO_Libelle desc"
                    If Odbc_SelectV(sql, rs) = P_ERREUR Then
                        afficher_liste3 = P_ERREUR
                        Exit Function
                    End If
                    While Not rs.EOF
                        sfct = rs("FT_Libelle").Value
                        If rs("FT_Libelle").Value <> rs("PO_Libelle").Value Then
                            sfct = sfct & " *"
                        End If
                        Set nd = tv.Nodes.Add(ndp.Child, _
                                               tvwPrevious, _
                                               "P" & rs("PO_Num").Value, _
                                               sfct, _
                                               IMG_POSTE, _
                                               IMG_POSTE)
                        nd.tag = True
                        rs.MoveNext
                    Wend
                    ndp.Expanded = True
                    rs.Close
                Else
                    sql = "select count(*) from Poste" _
                        & " where PO_Actif=true" _
                        & " and PO_SRVNum=" & Mid$(ndp.key, 2)
                    If Odbc_Count(sql, lnb) = P_ERREUR Then
                        afficher_liste3 = P_ERREUR
                        Exit Function
                    End If
                    If lnb > 0 Then
                        Set nd = tv.Nodes.Add(ndp, _
                                           tvwChild, _
                                           , _
                                           "A charger")
                        stag = ndp.tag
                        Call STR_PutChamp(stag, "|", 1, False)
                        ndp.tag = stag
                    End If
                End If
            End If
        End If
    Next n
    
    ' Met en évidence les noeuds 'retenus'
    If g_mode_acces = MODE_SELECT And g_plusieurs Then
        For I = 0 To nsel
            n = STR_GetNbchamp(g_lignes(I).texte, ";")
            s = STR_GetChamp(g_lignes(I).texte, ";", n - 1)
            If left$(s, 1) = "S" Then
                fmodif = g_lignes(I).fmodif
                If fmodif Then
                    img = IMG_SRV_SEL
                Else
                    img = IMG_SRV_SEL_NOMOD
                End If
            Else
                fmodif = g_lignes(I).fmodif
                If fmodif Then
                    img = IMG_POSTE_SEL
                Else
                    img = IMG_POSTE_SEL_NOMOD
                End If
            End If
            If TV_NodeExiste(tv, s, nd) = P_OUI Then
                nd.image = img
                nd.tag = fmodif
                nd.SelectedImage = img
                Set ndp = nd.Parent
                While left$(ndp.key, 1) <> "L"
                    ndp.Expanded = True
                    Set ndp = ndp.Parent
                Wend
            End If
        Next I
    End If
    
    Call ouvrir_serv_poste
    
    tv.SetFocus
    g_mode_saisie = True
    
    Set nd = Nothing
    Set ndp = Nothing
    Set ndp_sav = Nothing
    
    pgBar.Visible = False
    
    afficher_liste3 = P_OK

End Function

Private Function recup_lib_niveau(ByVal v_nivsnum As Long) As String
    
    Dim sql As String, nNiv As Long
    Dim libNiveau As String
    
    If Not g_ya_niveau Then
        recup_lib_niveau = ""
        Exit Function
    End If
    
    sql = "select count(*) from niveau_structure where Nivs_Num=" & v_nivsnum
    If Odbc_Count(sql, nNiv) = P_ERREUR Then
        recup_lib_niveau = ""
        Exit Function
    Else
        If v_nivsnum = 0 Then
            recup_lib_niveau = ""
        Else
            sql = "select Nivs_Nom from niveau_structure where Nivs_Num=" & v_nivsnum
            If Odbc_RecupVal(sql, libNiveau) = P_ERREUR Then
                recup_lib_niveau = ""
                Exit Function
            Else
                recup_lib_niveau = libNiveau
            End If
        End If
    End If
    
End Function

Private Sub afficher_menu(ByVal v_bclavier As Boolean)

    Dim key As String, tag As String, libresp As String, sql As String, libposte As String
    Dim numposte As Long, numresp As Long
    Dim FT_NivRemplace As Integer, srv_num As Integer
    Dim strRemplace As String
    
    key = tv.SelectedItem.key
    Select Case left$(key, 1)
    Case "L"
        mnuCreerS.Visible = g_crspm_autor
        mnuCreerS.Caption = "&Créer un service " & IIf(g_ya_niveau, propose_niveau(key, "Lib"), "")
        mnuSepCrS.Visible = g_crspm_autor
        mnuModS.Visible = False
        mnuSuppS.Visible = False
        mnuSepMSS.Visible = False
        mnuCreerP.Visible = False
        If g_mode_acces = MODE_PARAM And p_appli_kalibottin > 0 Then
            mnuCreerPi.Visible = g_crpiece_autor
            mnuSepCrP.Visible = g_crpiece_autor
        Else
            mnuCreerPi.Visible = False
            mnuSepCrP.Visible = False
        End If
        mnuPosteResp.Visible = False
        mnuLibPoste.Visible = False
        mnuSuppP.Visible = False
        mnuTrsP.Visible = False
        mnuModPi.Visible = False
        mnuSuppPi.Visible = False
        mnuSepSuppP.Visible = False
        mnuDepl.Visible = False
        mnuSepDepl.Visible = False
        mnuVoirPers.Visible = False
        mnuSepVoirPers.Visible = False
        mnuCrPers.Visible = False
        mnuAjPers.Visible = False
        mnuSepCrPers.Visible = False
        mnuModPers.Visible = False
        mnuSepModPers.Visible = False
        mnuPosteRemplace.Visible = False
    Case "S"
        mnuCreerS.Visible = g_crspm_autor
        mnuCreerS.Caption = "&Créer un service " & IIf(g_ya_niveau, propose_niveau(key, "Lib"), "")
        mnuSepCrS.Visible = g_crspm_autor
        mnuModS.Visible = g_modspm_autor
        mnuSuppS.Visible = g_supspm_autor
        mnuSepMSS.Visible = IIf(g_modspm_autor Or g_supspm_autor, True, False)
        If g_mode_acces = MODE_PARAM And p_appli_kalibottin > 0 Then
            mnuCreerPi.Visible = g_crpiece_autor
        Else
            mnuCreerPi.Visible = False
        End If
        mnuCreerP.Visible = g_crspm_autor
        mnuSepCrP.Visible = g_crspm_autor
        mnuPosteResp.Visible = False
        mnuLibPoste.Visible = False
        mnuSuppP.Visible = False
        mnuTrsP.Visible = False
        mnuModPi.Visible = False
        mnuSuppPi.Visible = False
        mnuSepSuppP.Visible = False
        If tv.SelectedItem.Index = tv.SelectedItem.Root.Index Then
            mnuDepl.Visible = False
            mnuSepDepl.Visible = False
        Else
            mnuDepl.Visible = g_modspm_autor
            mnuSepDepl.Visible = g_modspm_autor
        End If
        If InStr(g_stype, "P") > 0 Or g_stype = "" Then
            mnuVoirPers.Visible = True
            mnuSepVoirPers.Visible = True
        Else
            mnuVoirPers.Visible = False
            mnuSepVoirPers.Visible = False
        End If
        mnuCrPers.Visible = False
        mnuAjPers.Visible = False
        mnuSepCrPers.Visible = False
        mnuModPers.Visible = False
        mnuSepModPers.Visible = False
        mnuPosteRemplace.Visible = False
    Case "P"
        mnuCreerS.Visible = False
        mnuSepCrS.Visible = False
        mnuModS.Visible = False
        mnuSuppS.Visible = False
        mnuSepMSS.Visible = False
        mnuCreerP.Visible = False
        mnuCreerPi.Visible = False
        mnuSepCrP.Visible = False
        mnuPosteResp.Visible = True
        numposte = Mid$(key, 2)
        sql = "select PO_NumResp, PO_Libelle from Poste" _
            & " where PO_Num=" & numposte
        If Odbc_RecupVal(sql, numposte, libposte) = P_ERREUR Then
            libresp = "???"
        ElseIf numposte > 0 Then
            sql = "select FT_Libelle from Poste, FctTrav" _
                & " where PO_Num=" & numposte _
                & " and FT_Num=PO_FTNum"
            If Odbc_RecupVal(sql, libresp) = P_ERREUR Then
                libresp = "???"
            End If
        ElseIf numposte = 0 Then
            libresp = "Est responsable"
        Else
            libresp = "NON RENSEIGNE"
        End If
        libresp = "--- Poste responsable : " & libresp & " ---"
        
        ' Niveau de remplacement
        sql = "select FT_NivRemplace, PO_SrvNum from Poste, FctTrav" _
            & " where PO_Num=" & Mid$(key, 2) _
            & " and FT_Num=PO_FTNum"
        If Odbc_RecupVal(sql, FT_NivRemplace, srv_num) <> P_ERREUR Then
            If FT_NivRemplace = 0 Then
                mnuPosteRemplace.Visible = False
            Else
                mnuPosteRemplace.Visible = True
                mnuPosteRemplace.Caption = ""
                If FctNivRemplace(FT_NivRemplace, srv_num, strRemplace) < 0 Then
                    strRemplace = "!!! Attention : " & strRemplace
                End If
                mnuPosteRemplace.Caption = strRemplace
            End If
        End If
        
        mnuPosteResp.Caption = libresp
        mnuLibPoste.Visible = True
        mnuLibPoste.Caption = "Poste : " & libposte
        mnuSuppP.Visible = g_supspm_autor
        mnuTrsP.Visible = g_supspm_autor
        mnuModPi.Visible = False
        mnuSuppPi.Visible = False
        mnuSepSuppP.Visible = True
        mnuDepl.Visible = g_modspm_autor
        mnuSepDepl.Visible = g_modspm_autor
        If g_mode_acces = MODE_PARAM_PERS Then
            mnuVoirPers.Visible = False
            mnuSepVoirPers.Visible = False
        Else
            If tv.SelectedItem.Children > 0 Then
                mnuVoirPers.Visible = False
                mnuSepVoirPers.Visible = False
            Else
                mnuVoirPers.Visible = True
                mnuSepVoirPers.Visible = True
            End If
        End If
        If g_mode_acces = MODE_PARAM_PERS Or g_prmpers Then
            mnuCrPers.Visible = g_crutil_autor
            mnuAjPers.Visible = g_crutil_autor
            mnuSepCrPers.Visible = g_crutil_autor
        Else
            mnuCrPers.Visible = False
            mnuAjPers.Visible = False
            mnuSepCrPers.Visible = False
        End If
        mnuModPers.Visible = False
        mnuSepModPers.Visible = False
    Case "C"
        mnuCreerS.Visible = False
        mnuSepCrS.Visible = False
        mnuModS.Visible = False
        mnuSuppS.Visible = False
        mnuSepMSS.Visible = False
        mnuCreerP.Visible = False
        mnuCreerPi.Visible = False
        mnuSepCrP.Visible = False
        mnuPosteResp.Visible = False
        mnuLibPoste.Visible = False
        mnuSuppP.Visible = False
        mnuTrsP.Visible = False
        mnuSepSuppP.Visible = True
        mnuModPi.Visible = g_modpiece_autor
        mnuSuppPi.Visible = g_modpiece_autor
        mnuDepl.Visible = g_modpiece_autor
        mnuSepDepl.Visible = g_modpiece_autor
        mnuCrPers.Visible = False
        mnuAjPers.Visible = False
        mnuSepCrPers.Visible = False
        mnuVoirPers.Visible = False
        mnuSepVoirPers.Visible = False
        mnuModPers.Visible = False
        mnuSepModPers.Visible = False
        mnuPosteRemplace.Visible = False
    Case Else
        mnuCreerS.Visible = False
        mnuSepCrS.Visible = False
        mnuModS.Visible = False
        mnuSuppS.Visible = False
        mnuSepMSS.Visible = False
        mnuCreerP.Visible = False
        mnuCreerPi.Visible = False
        mnuSepCrP.Visible = False
        mnuPosteResp.Visible = False
        mnuLibPoste.Visible = False
        mnuSuppP.Visible = False
        mnuModPi.Visible = False
        mnuSuppPi.Visible = False
        mnuTrsP.Visible = False
        mnuSepSuppP.Visible = False
        mnuDepl.Visible = False
        mnuSepDepl.Visible = False
        mnuVoirPers.Visible = False
        mnuSepVoirPers.Visible = False
        mnuCrPers.Visible = False
        mnuSepCrPers.Visible = False
        mnuAjPers.Visible = False
        If g_mode_acces = MODE_PARAM_PERS Or g_prmpers Then
            mnuModPers.Visible = g_modutil_autor
            mnuSepModPers.Visible = g_modutil_autor
        Else
            mnuModPers.Visible = False
            mnuSepModPers.Visible = False
        End If
        mnuPosteRemplace.Visible = False
    End Select
    
    If v_bclavier Then
        Call PopupMenu(mnuFct, , tv.left, tv.Top)
    Else
        Call PopupMenu(mnuFct)
    End If
    
End Sub

Private Function propose_niveau(ByVal keyService As String, _
                                ByVal v_Trait As String)
    Dim sql As String
    Dim rs As rdoResultset
    Dim nNiv As Long
    Dim libNiveau As String
    Dim op As String
    
    If Not g_ya_niveau Then Exit Function
    
    If Mid(keyService, 1, 1) = "L" Then
        ' seule possibilité : un Pôle
        sql = "select * from niveau_structure where Nivs_NivPere=0"
        GoTo Suite
    ElseIf Mid(keyService, 1, 1) = "S" Then
    Else
        MsgBox "Clé " & keyService & " non conforme"
    End If
    sql = "select * from Service where SRV_Num = " & Replace(keyService, "S", "")
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        propose_niveau = P_ERREUR
        Exit Function
    ElseIf rs.EOF Then
        propose_niveau = P_ERREUR
        Exit Function
    Else
        If rs("srv_nivsnum").Value = 0 Then
            propose_niveau = ""
        Else
            sql = "select count(*) from niveau_structure where Nivs_Num=" & rs("srv_nivsnum").Value
            If Odbc_Count(sql, nNiv) = P_ERREUR Then
                propose_niveau = ""
                Exit Function
            Else
                sql = "select * from niveau_structure where Nivs_NivPere=" & rs("srv_nivsnum")
Suite:
                If Odbc_SelectV(sql, rs) = P_ERREUR Then
                    MsgBox "Erreur SQL " & sql
                    propose_niveau = ""
                    Exit Function
                Else
                    op = ""
                    If v_Trait = "Lib" Then propose_niveau = "( "
                    While Not rs.EOF
                        If v_Trait = "Lib" Then
                            propose_niveau = propose_niveau & op & rs("Nivs_Nom")
                            op = " ou "
                        Else
                            propose_niveau = propose_niveau & op & rs("Nivs_Code")
                            op = "|"
                        End If
                        rs.MoveNext
                    Wend
                    If v_Trait = "Lib" Then propose_niveau = propose_niveau & " )"
                End If
            End If
        End If
    End If
    rs.Close
End Function

Private Function ajouter_fils(ByVal v_numsrv As Long) As Integer

    Dim sql As String
    Dim rs As rdoResultset
    Dim nd As Node, ndp As Node
    
    sql = "select SRV_Num, SRV_Nom from Service" _
        & " where SRV_NumPere=" & v_numsrv _
        & " order by SRV_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        ajouter_fils = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        Set ndp = tv.Nodes("S" & v_numsrv)
        Set nd = tv.Nodes.Add(ndp, _
                               tvwChild, _
                               "S" & rs("SRV_Num").Value, _
                               rs("SRV_Nom").Value, _
                               IMG_SRV, _
                               IMG_SRV)
'        nd.Sorted = True
        nd.tag = True
        Call ajouter_fils(rs("SRV_Num").Value)
        rs.MoveNext
    Wend
    rs.Close
    
    ajouter_fils = P_OK
    
End Function

Private Sub ajouter_pers_tv()

    Dim sql As String, sposte As String, s As String
    Dim I As Integer, n   As Integer
    Dim spm As Variant
    Dim nd As Node, ndp As Node
    Dim rs As rdoResultset
    Dim ndf As Node
    
    sql = "select U_Num, U_Nom, U_Prenom, U_SPM from Utilisateur" _
        & " where U_SPM like '%" & tv.SelectedItem.key & ";%' and U_Actif=true" _
        & " order by U_Nom, U_Prenom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    If Not rs.EOF Then
        While Not rs.EOF
            spm = rs("U_SPM").Value
            n = STR_GetNbchamp(spm, "|")
            For I = 0 To n - 1
                s = STR_GetChamp(spm, "|", I)
                If InStr(s, tv.SelectedItem.key + ";") > 0 Then
                    sposte = STR_GetChamp(s, ";", STR_GetNbchamp(s, ";") - 1)
                    If TV_NodeExiste(tv, sposte, ndp) = P_NON Then
                        Call charger_arbor(sposte)
                        Set ndp = tv.Nodes(sposte)
                    End If
                    ' Vérifier s'il faut ajouter la personne ou si elle existe déjà
                    If ndp.Children > 0 Then
                        Set ndf = ndp.Child
continue:
                        If ndf.tag = "U" & rs("U_Num").Value Then
                            GoTo affich_suiv
                        End If
                        If TV_NodeNext(ndf) Then
                            GoTo continue
                        End If
                    End If
                    
                    Set nd = tv.Nodes.Add(ndp, _
                                           tvwChild, _
                                           "", _
                                           rs("U_Nom").Value & " " & rs("U_Prenom").Value, _
                                           IMG_UTIL, _
                                           IMG_UTIL)
                    nd.tag = "U" & rs("U_Num").Value
                    ndp.Expanded = True
                    Set ndp = ndp.Parent
                    While left$(ndp.key, 1) <> "L"
                        ndp.Expanded = True
                        Set ndp = ndp.Parent
                    Wend
affich_suiv:
                End If
            Next I
            rs.MoveNext
        Wend
    End If
    rs.Close

End Sub

Private Sub ajouter_personne()

    Dim sret As String, sfct_old As String, spm_old As String, le_spm As String
    Dim spm As String, sfct As String
    Dim far As Boolean
    Dim numposte As Long, numutil As Long, numfct As Long
    Dim frm As Form
    
    p_siz_tblu = -1
    Set frm = ChoixUtilisateur
    sret = ChoixUtilisateur.AppelFrm("Choix d'une personne", "", False, False, "", True, True)
    Set frm = Nothing
    
    If sret = "" Then
        Exit Sub
    End If
    
    numutil = p_tblu_sel(0)
    
    numposte = Mid$(tv.SelectedItem.key, 2)
    If Odbc_RecupVal("select u_fcttrav, u_spm, u_ar from utilisateur" _
                        & " where u_num=" & numutil, _
                       sfct_old, spm_old, far) = P_ERREUR Then
        Exit Sub
    End If
    
    ' Vérifier que la personne ne possède pas déjà ce poste
    If InStr(spm_old, "P" & numposte & ";") > 0 Then
        Call MsgBox("Cette personne est déjà associée à ce poste!", vbInformation + vbOKOnly, "")
        Exit Sub
    End If
    
    ' Maj u_fcttrav et u_spm
    If build_services(numposte, le_spm) = P_ERREUR Then
        Exit Sub
    End If
    spm = spm_old + le_spm + "|"
    If Odbc_RecupVal("select po_ftnum from poste where po_num=" & numposte, _
                     numfct) = P_ERREUR Then
        Exit Sub
    End If
    sfct = sfct_old
    If InStr(sfct_old, "F" & numfct & ";") = 0 Then
        sfct = sfct & "F" & numfct & ";"
    End If
    If Odbc_Update("Utilisateur", "U_Num", "where u_num=" & numutil, _
                     "u_fcttrav", sfct, _
                     "u_spm", spm) = P_ERREUR Then
        Exit Sub
    End If
    
    ' Gestion des diffusions ...
    Call KS_PrmPersonne.gerer_chgt_prmutil(numutil, far, "", "", sfct_old, sfct, spm_old, spm)
    
End Sub

Private Function ajouter_service(ByVal v_numsrv As Long) As Integer

    Dim sql As String, lib As String
    Dim trouve As Boolean
    Dim mode As Integer, I As Integer, n As Integer
    Dim rs As rdoResultset
    Dim nd As Node, ndp As Node, ndp_sav As Node
    Dim libNiveau As String
    
    If v_numsrv = 0 Then
        ajouter_service = P_OK
        Exit Function
    End If
    
    If TV_NodeExiste(tv, "S" & v_numsrv, nd) = P_OUI Then
        ajouter_service = P_OK
        Exit Function
    End If
    
    sql = "select SRV_LNum, SRV_NumPere, SRV_Nom, SRV_NivsNum from Service" _
        & " where SRV_Num=" & v_numsrv
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        ajouter_service = P_ERREUR
        Exit Function
    End If
    If rs("SRV_NumPere").Value > 0 Then
        If TV_NodeExiste(tv, "S" & rs("SRV_NumPere").Value, nd) = P_NON Then
            Call ajouter_service(rs("SRV_NumPere").Value)
        End If
        Set ndp = tv.Nodes("S" & rs("SRV_NumPere").Value)
    Else
        Set ndp = tv.Nodes(1).Root
    End If
    lib = rs("SRV_Nom").Value
    libNiveau = recup_lib_niveau(rs("SRV_NivsNum").Value)
    If libNiveau <> "" Then
        lib = lib & " (" & libNiveau & ")"
    End If
    Set nd = tv.Nodes.Add(ndp, _
                           tvwChild, _
                           "S" & v_numsrv, _
                           lib, _
                           IMG_SRV, _
                           IMG_SRV)
    nd.tag = True & "|" & True
    
    ajouter_service = P_OK
    
End Function

Private Sub basculer_selection()

    Dim img As Long
    
    If tv.SelectedItem.tag = "" Or left$(tv.SelectedItem.tag, 1) = "U" Then
        Exit Sub
    End If
    If STR_GetChamp(tv.SelectedItem.tag, "|", 0) = False Then
        Exit Sub
    End If
    
    Select Case tv.SelectedItem.SelectedImage
    Case IMG_SRV
        If InStr(g_stype, left$(tv.SelectedItem.key, 1)) = 0 Then
            Call MsgBox("Vous ne pouvez pas sélectionner un service.", vbInformation + vbOKOnly, "")
            Exit Sub
        End If
        img = IMG_SRV_SEL
    Case IMG_POSTE
        If InStr(g_stype, left$(tv.SelectedItem.key, 1)) = 0 Then
            Call MsgBox("Vous ne pouvez pas sélectionner un poste.", vbInformation + vbOKOnly, "")
            Exit Sub
        End If
        img = IMG_POSTE_SEL
    Case IMG_SRV_SEL
        img = IMG_SRV
    Case IMG_POSTE_SEL
        img = IMG_POSTE
    End Select
    
    tv.SelectedItem.SelectedImage = img
    tv.SelectedItem.image = img
    
End Sub

Private Function build_services(ByVal v_numposte As Long, _
                               ByRef r_sp As String) As Integer
                               
    Dim s As String, sql As String
    Dim numsrv As Long
    
    s = "P" & v_numposte & ";"
    sql = "select PO_SRVNum from Poste where PO_Num=" & v_numposte
    If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
        build_services = P_ERREUR
        Exit Function
    End If
    s = "S" & numsrv & ";" & s
    Do
        sql = "select SRV_NumPere from Service where SRV_Num=" & numsrv
        If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
            build_services = P_ERREUR
            Exit Function
        End If
        If numsrv > 0 Then
            s = "S" & numsrv & ";" & s
        End If
    Loop Until numsrv = 0
    
    r_sp = s
    
    build_services = P_OK
    
End Function

Private Function charger_arbor(ByVal v_ssrv As String) As Integer

    Dim sql As String, s_srv As String, s As String
    Dim I As Integer, n As Integer
    Dim numsrv As Long, numposte As Long
    Dim nd As Node, ndp As Node
    
    If left$(v_ssrv, 1) = "P" Then
        numposte = Mid$(v_ssrv, 2)
        sql = "select PO_SRVNum from Poste where PO_Num=" & numposte
        If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
            charger_arbor = P_ERREUR
            Exit Function
        End If
    Else
        numsrv = Mid$(v_ssrv, 2)
        sql = "select SRV_NumPere from Service where SRV_Num=" & numsrv
        If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
            charger_arbor = P_ERREUR
            Exit Function
        End If
    End If
    s_srv = numsrv & ";"
    While numsrv > 0
        sql = "select SRV_NumPere from Service where SRV_Num=" & numsrv
        If Odbc_RecupVal(sql, numsrv) = P_ERREUR Then
            charger_arbor = P_ERREUR
            Exit Function
        End If
        If numsrv > 0 Then
            s_srv = numsrv & ";" & s_srv
        End If
    Wend
    
    n = STR_GetNbchamp(s_srv, ";")
    For I = 0 To n - 1
        numsrv = STR_GetChamp(s_srv, ";", I)
        If TV_NodeExiste(tv, "S" & numsrv, nd) = P_OUI Then
            If STR_GetChamp(nd.tag, "|", 1) = False Then
                tv.Nodes.Remove (nd.Child.Index)
                If charger_service(numsrv) = P_ERREUR Then
                    charger_arbor = P_ERREUR
                    Exit Function
                End If
            End If
        End If
    Next I
    
    charger_arbor = P_OK

End Function

Private Function charger_service(ByVal v_numsrv As Long) As Integer

    Dim sql As String, sfct As String, lib As String, stag As String
    Dim libNiveau As String
    Dim img As Integer, I As Integer
    Dim lnb As Long
    Dim rs As rdoResultset, rsu As rdoResultset
    Dim nd As Node, ndp As Node, ndu As Node
    Dim strRemplace As String
    
    If v_numsrv = 0 Then
        Set ndp = tv.Nodes(1)
    Else
        Set ndp = tv.Nodes("S" & v_numsrv)
        ndp.tag = True & "|" & False
    End If

    ' Les pièces
    If g_mode_acces = MODE_PARAM And p_appli_kalibottin > 0 Then
        sql = "select PC_Num, PC_Nom" _
            & " from Piece" _
            & " where PC_SRVNum=" & v_numsrv _
            & " order by PC_Nom"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            charger_service = P_ERREUR
            Exit Function
        End If
        While Not rs.EOF
            Set nd = tv.Nodes.Add(ndp, _
                                   tvwChild, _
                                   "C" & rs("PC_Num").Value, _
                                   rs("PC_Nom").Value, _
                                   IMG_PIECE, _
                                   IMG_PIECE)
            nd.tag = True
            rs.MoveNext
        Wend
        rs.Close
    End If

    ' Les postes
    If g_smode <> "C" Then
        sql = "select PO_Num, PO_Libelle, FT_Libelle, FT_NivRemplace" _
            & " from Poste, FctTrav" _
            & " where FT_Num=PO_FTNum" _
            & " and PO_Actif=true" _
            & " and PO_SRVNum=" & v_numsrv _
            & " order by PO_Libelle"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            charger_service = P_ERREUR
            Exit Function
        End If
        
        If Not rs.EOF Then
            rs.MoveLast
            pgBar.Visible = True
            pgBar.Max = rs.RowCount
            pgBar.Value = 0
            rs.MoveFirst
        End If
        
        While Not rs.EOF
            pgBar.Value = pgBar.Value + 1
            sfct = rs("FT_Libelle").Value
            If rs("FT_Libelle").Value <> rs("PO_Libelle").Value Then
                sfct = sfct & " *"
            End If
            ' Vérification du niveau de remplacement
            If FctNivRemplace(rs("FT_NivRemplace"), v_numsrv, strRemplace) < 0 Then
                MsgBox strRemplace
                sfct = sfct & " (" & strRemplace & ")"
            End If
            Set nd = tv.Nodes.Add(ndp, _
                                   tvwChild, _
                                   "P" & rs("PO_Num").Value, _
                                   sfct, _
                                   IMG_POSTE, _
                                   IMG_POSTE)
            nd.tag = True
            If g_mode_acces = MODE_PARAM_PERS Then
                ' Les personnes associées
                sql = "select U_Num, U_Nom, U_Prenom from Utilisateur" _
                    & " where U_SPM like '%P" & rs("PO_Num").Value & ";%'" _
                    & " and U_Actif=true"
                If Odbc_SelectV(sql, rsu) = P_ERREUR Then
                    charger_service = P_ERREUR
                    Exit Function
                End If
                While Not rsu.EOF
                    Set ndu = tv.Nodes.Add(nd, _
                                           tvwChild, _
                                           "", _
                                           rsu("U_Nom").Value & " " & rsu("U_Prenom").Value, _
                                           IMG_UTIL, _
                                           IMG_UTIL)
                    ndu.tag = "U" & rsu("U_Num").Value
                    rsu.MoveNext
                Wend
                rsu.Close
            End If
            rs.MoveNext
        Wend
        rs.Close
        pgBar.Visible = False
    End If
    
    If TxtRecherche.Text <> "" Then
        GoTo lab_fin
    End If
    
    ' Les services
    sql = "select SRV_Num, SRV_LNum, SRV_NumPere, SRV_Nom, SRV_NivsNum from Service " _
        & " where SRV_Actif=true" _
        & " and SRV_Numpere=" & v_numsrv _
        & " order by SRV_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        charger_service = P_ERREUR
        Exit Function
    End If
    
    If Not rs.EOF Then
        rs.MoveLast
        pgBar.Visible = True
        pgBar.Max = rs.RowCount
        pgBar.Value = 0
        rs.MoveFirst
    End If
    
    While Not rs.EOF
        pgBar.Value = pgBar.Value + 1

        lib = rs("SRV_Nom").Value
        libNiveau = recup_lib_niveau(rs("SRV_NivsNum").Value)
        If libNiveau <> "" Then
            lib = lib & " (" & libNiveau & ")"
        End If
        Set nd = tv.Nodes.Add(ndp, _
                               tvwChild, _
                               "S" & rs("SRV_Num").Value, _
                               lib, _
                               IMG_SRV, _
                               IMG_SRV)
        sql = "select count(*) from Service " _
            & " where SRV_Actif=true" _
            & " and SRV_Numpere=" & rs("SRV_Num").Value
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            charger_service = P_ERREUR
            Exit Function
        End If
        If lnb = 0 And g_smode <> "C" Then
            sql = "select count(*) from Poste " _
                & " where PO_Actif=true" _
                & " and PO_SRVNum=" & rs("SRV_Num").Value
            If Odbc_Count(sql, lnb) = P_ERREUR Then
                charger_service = P_ERREUR
                Exit Function
            End If
        End If
        If lnb = 0 Then
            nd.tag = True & "|" & True
        Else
            nd.tag = True & "|" & False
            Set nd = tv.Nodes.Add(nd, _
                               tvwChild, _
                               , _
                               "A charger")
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    pgBar.Visible = False
    
lab_fin:
    
    stag = ndp.tag
    Call STR_PutChamp(stag, "|", 1, True)
    ndp.tag = stag
    
    charger_service = P_OK
    
End Function

Private Sub creer_fonction(ByRef r_num As Long, _
                           ByRef r_lib As String)

    Dim sret As String
    Dim frm As Form
    
    Set frm = KS_PrmFonction
    sret = KS_PrmFonction.AppelFrm("CREATE")
    Set frm = Nothing
    If sret = "" Then
        r_num = 0
        Exit Sub
    End If
    r_num = STR_GetChamp(sret, "|", 0)
    r_lib = STR_GetChamp(sret, "|", 1)
    
End Sub

Private Function creer_poste() As Integer
    
    Dim sql As String, lib As String
    Dim trouve As Boolean
    Dim I As Integer, nbenf As Integer, n As Integer, btn_sortie As Integer
    Dim nfct As Integer, ie As Integer
    Dim numsrv As Long, num As Long, tbl_fct() As Long, numlabo As Long, lnb As Long
    Dim numfct
    Dim rs As rdoResultset
    Dim nd As Node, nde As Node, ndp As Node
    
    Set nd = tv.SelectedItem
    numsrv = Mid$(nd.key, 2)
    nbenf = nd.Children
    nfct = -1
    Set nde = nd.Child
    For I = 1 To nbenf
        If left$(nde.key, 1) = "P" Then
            num = Mid$(nde.key, 2)
            sql = "select PO_FTNum from Poste" _
                & " where PO_Num=" & num
            If Odbc_RecupVal(sql, numfct) = P_ERREUR Then
                creer_poste = P_ERREUR
                Exit Function
            End If
            nfct = nfct + 1
            ReDim Preserve tbl_fct(nfct) As Long
            tbl_fct(nfct) = numfct
        End If
        Set nde = nde.Next
    Next I
    
    Call CL_Init
    
    sql = "select * from FctTrav" _
        & " where FT_Actif=true" _
        & " order by FT_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        creer_poste = P_ERREUR
        Exit Function
    End If
    n = 0
    While Not rs.EOF
        trouve = False
        For I = 0 To nfct
            If tbl_fct(I) = rs("FT_Num").Value Then
                trouve = True
                Exit For
            End If
        Next I
        If Not trouve Then
            Call CL_AddLigne(rs("FT_Libelle").Value, rs("FT_Num").Value, "", False)
            n = n + 1
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    If n = 0 And Not g_crfct_autor Then
        Call MsgBox("Aucune fonction ne peut être ajoutée à ce service.", vbInformation + vbOKOnly, "")
        creer_poste = P_OK
        Exit Function
    End If
    
    Call CL_InitTitreHelp("Fonctions du personnel", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    If g_crfct_autor Then
        Call CL_AddBouton("&Créer une fonction", "", 0, 0, 1800)
        btn_sortie = 2
    Else
        btn_sortie = 1
    End If
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -15)
    Call CL_InitMultiSelect(True, True)
    Call CL_AffiSelFirst
    Call CL_InitResteCachée(True)
lab_choix:
    ChoixListe.Show 1
    ' Sortie
    If CL_liste.retour = btn_sortie Then
        GoTo lab_fin
    End If
    
    ' Création
    If CL_liste.retour = 1 Then
        Call creer_fonction(num, lib)
        If num > 0 Then
            Call CL_AddLigne(lib, num, "", True)
            n = n + 1
        End If
        GoTo lab_choix
    End If
    
    ' Ajout des sélectionnés
'    Call TV_FirstParent(nd, ndp)
'    numlabo = Mid$(ndp.key, 2)
    ' Ca ne marchait pas si on avait sélectionné un service et non tout le site
    ' -> numlabo se retrouvait avec le numserv
    numlabo = 1
    For I = 0 To n - 1
        If CL_liste.lignes(I).selected Then
            Call Odbc_AddNew("Poste", _
                             "PO_Num", _
                             "po_seq", _
                             True, _
                             num, _
                             "PO_SRVNum", numsrv, _
                             "PO_FTNum", CL_liste.lignes(I).num, _
                             "PO_Libelle", CL_liste.lignes(I).texte, _
                             "PO_NumResp", -1, _
                             "PO_LNum", numlabo, _
                             "PO_Actif", True)
            Set nde = tv.Nodes.Add(nd, _
                                   tvwChild, _
                                   "P" & num, _
                                   CL_liste.lignes(I).texte, _
                                   IMG_POSTE, _
                                   IMG_POSTE)
            nde.tag = True
        End If
    Next I
    
lab_fin:
    Unload ChoixListe
    creer_poste = P_OK

End Function

Private Function creer_service() As Integer

    Dim liste_multiselect As Boolean, b_sel As Boolean
    Dim code As String, lib As String, libcourt As String, smasque As String
    Dim liste_nomtable As String, liste_lsttypchp As String, liste_chpretour As String, liste_chpnum   As String
    Dim nivs_code As String, sql As String
    Dim first_chp As Integer, srv_visible As Integer, n As Integer
    Dim num As Long, numpere As Long, numlabo As Long, lnb As Long
    Dim nivs_num As Long
    Dim nd As Node, ndp As Node
    Dim rs As rdoResultset
    
    Set nd = tv.SelectedItem
    
    code = ""
    lib = ""
    libcourt = ""
    
lab_saisie:
    Call SAIS_Init
    Call SAIS_InitTitreHelp("Service", "dico_d_spm.htm")
    Call SAIS_AddBouton("", p_chemin_appli + "\btnEnregistrer.gif", vbKeyE, vbKeyF1, 0)
    If left$(nd.key, 1) = "S" Then
        Call SAIS_AddChamp("Rattaché à", -50, 0, True, nd.Text)
        first_chp = 1
        numpere = Mid$(nd.key, 2)
    Else
        first_chp = 0
        numpere = 0
    End If
    Call SAIS_AddChamp("Code", 8, SAIS_TYP_TOUT_CAR, True, code)
    Call SAIS_AddChamp("Nom", 120, SAIS_TYP_TOUT_CAR, False, lib)
    Call SAIS_AddChamp("Nom court", 30, SAIS_TYP_TOUT_CAR, True, libcourt)
    Call SAIS_AddChamp("Masque", 30, SAIS_TYP_TOUT_CAR, True, smasque)

    sql = "select count(*) from niveau_structure"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
    End If
    If lnb > 0 Then
        g_ya_niveau = True
        nivs_code = propose_niveau(nd.key, "Code")
        liste_nomtable = "select * from niveau_structure"
        liste_multiselect = False
        liste_chpretour = "nivs_code"
        liste_chpnum = "nivs_num"
        Call SAIS_AddListe("Niveau", liste_nomtable, liste_multiselect, liste_chpretour, liste_chpnum, SAIS_TYP_CHOIXLISTE, True, nivs_code)
        ' Ajouter les champs listes
        sql = "select * from niveau_structure"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            MsgBox "Erreur SQL " & sql
            Exit Function
        End If
        While Not rs.EOF
            Call SAIS_AddItemListe(UBound(SAIS_Saisie.champs), rs("nivs_num"), rs("nivs_code"), rs("nivs_nom"), (rs("nivs_code") = nivs_code))
            rs.MoveNext
        Wend
        rs.Close
    Else
        g_ya_niveau = False
    End If
    
    If p_appli_kalibottin > 0 Then
        Call SAIS_AddListe("Visible dans l'annuaire", "", False, "", "", SAIS_TYP_CHOIXLISTE, True, "")
        Call SAIS_AddItemListe(UBound(SAIS_Saisie.champs), 0, "0", "Jamais", False)
        Call SAIS_AddItemListe(UBound(SAIS_Saisie.champs), 1, "1", "Vue détaillée seulement", False)
        Call SAIS_AddItemListe(UBound(SAIS_Saisie.champs), 2, "2", "Toujours", True)
        Call SAIS_AddBouton("Coordonnées", "", 0, 0, 1500)
    End If
    Call SAIS_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    
    Saisie.Show 1
    
    If SAIS_Saisie.retour <> SAIS_RET_MODIF Then
        creer_service = P_NON
        Exit Function
    End If
    code = SAIS_Saisie.champs(first_chp).sval
    If code <> "" Then
        Call Odbc_Count("select count(*) from service where srv_code=" & Odbc_String(SAIS_Saisie.champs(first_chp).sval), lnb)
        If lnb > 0 Then
            Call MsgBox("Le code '" & code & "' est déjà attribué." & vbCrLf & vbCrLf & "Veuillez choisir un autre code.", vbInformation + vbOKOnly, "")
            GoTo lab_saisie
        End If
    End If
    lib = SAIS_Saisie.champs(first_chp + 1).sval
    libcourt = SAIS_Saisie.champs(first_chp + 2).sval
    smasque = SAIS_Saisie.champs(first_chp + 3).sval
    nivs_num = 0
    If g_ya_niveau Then
        n = 1
        nivs_code = SAIS_Saisie.champs(first_chp + 4).sval
        Call Odbc_RecupVal("select nivs_num from niveau_structure where  nivs_code='" & nivs_code & "'", _
                            nivs_num)
    Else
        n = 0
    End If
    If p_appli_kalibottin > 0 Then
        srv_visible = CInt(SAIS_Saisie.champs(first_chp + 4 + n).sval)
    Else
        srv_visible = 2
    End If
    
    Call TV_FirstParent(nd, ndp)
    numlabo = Mid$(ndp.key, 2)
    Call Odbc_AddNew("Service", _
                     "SRV_Num", _
                     "srv_seq", _
                     True, _
                     num, _
                     "SRV_Code", code, _
                     "SRV_Code_Masque", smasque, _
                     "SRV_Nom", lib, _
                     "SRV_libcourt", libcourt, _
                     "SRV_NivsNum", nivs_num, _
                     "SRV_NumPere", numpere, _
                     "SRV_LNum", numlabo, _
                     "SRV_Visible", srv_visible, _
                     "SRV_Actif", True)
    
'    nd.Sorted = True
    nd.Expanded = True
    If nivs_code <> "" Then
        lib = lib & " (" & nivs_code & ")"
    End If
    Set nd = tv.Nodes.Add(nd, tvwChild, "S" & num, lib, IMG_SRV, IMG_SRV)
    nd.tag = True & "|" & True
    Set tv.SelectedItem = nd
    SendKeys "{DOWN}"
    SendKeys "{UP}"
    DoEvents
    Set tv.SelectedItem = nd
    
    creer_service = P_OUI
    
End Function

Private Function deplacer_piece() As Integer

    Dim key As String, sql As String, stype_dest As String
    Dim key_depl As String
    Dim I As Integer
    Dim numsrv_dest As Long, lnb As Long, num_piece As Long
    Dim nd_src As Node, nd_dest As Node, nd As Node, ndp As Node
    Dim rs As rdoResultset
    
    lblDepl.Visible = False
    
    Set nd_dest = tv.SelectedItem
    Set nd_src = tv.Nodes(g_pos_depl)
    g_pos_depl = 0
    
    stype_dest = left$(nd_dest.key, 1)
    If stype_dest <> "L" And stype_dest <> "S" Then
        Call MsgBox("Vous ne pouvez déplacer une pièce que dans un service !", vbExclamation + vbOKOnly, "")
        deplacer_piece = P_OK
        Exit Function
    End If
    
    num_piece = CLng(Mid$(nd_src.key, 2))
    numsrv_dest = CLng(Mid$(nd_dest.key, 2))
    
    key_depl = nd_src.key
    
    ' Mise à jour Piece : PC_SRVNum
    If Odbc_Update("Piece", _
                    "PC_Num", _
                    "where PC_Num=" & num_piece, _
                    "PC_SRVNum", numsrv_dest) = P_ERREUR Then
        GoTo err_enreg
    End If
    
    Set nd_src.Parent = nd_dest
    Set tv.SelectedItem = tv.Nodes(key_depl)
    tv.Nodes(key_depl).EnsureVisible
    
    ' LN - Permet de supprimer la ligne "A charger" : je n'ai pas réussi autrement ...
    If nd_dest.Children > 0 Then
        Set nd = nd_dest.Child
        For I = 1 To nd_dest.Children
            If nd.Text = "A charger" Then
                tv.Nodes.Remove (nd.Index)
                Exit For
            End If
            If I < nd_dest.Children Then
                Set nd = nd.Next
            End If
        Next I
    End If
    
    deplacer_piece = P_OK
    Exit Function

err_enreg:
    deplacer_piece = P_ERREUR
    
End Function

Private Function deplacer_sp() As Integer

    Dim key As String, sql As String, stype_src As String, stype_dest As String
    Dim s_sp_src As String, s_sp_dest As String, s_sp As String, s As String
    Dim key_depl As String
    Dim encore As Boolean
    Dim I As Integer, nbch As Integer, nutil As Integer, iutil As Integer
    Dim iu As Integer
    Dim numsrv_src As Long, numsrv_dest As Long, lnb As Long
    Dim tbl_util() As Long
    Dim nd_src As Node, nd_dest As Node, nd As Node, ndp As Node
    Dim rs As rdoResultset, rsu As rdoResultset
    
    lblDepl.Visible = False
    
    Set nd_dest = tv.SelectedItem
    Set nd_src = tv.Nodes(g_pos_depl)
    g_pos_depl = 0
    
    If nd_src.key = nd_dest.key Then
        Call MsgBox("Vous ne pouvez pas rattacher l'objet à lui-même !", vbExclamation + vbOKOnly, "")
        deplacer_sp = P_OK
        Exit Function
    End If
    
    stype_src = left$(nd_src.key, 1)
    stype_dest = left$(nd_dest.key, 1)
    If stype_dest = "P" Then
        Call MsgBox("Vous ne pouvez pas sélectionner un poste !", vbExclamation + vbOKOnly, "")
        deplacer_sp = P_OK
        Exit Function
    End If
    If stype_src = "P" And stype_dest = "L" Then
        Call MsgBox("Vous ne pouvez déplacer un poste que dans un service !", vbExclamation + vbOKOnly, "")
        deplacer_sp = P_OK
        Exit Function
    End If
    
    ' On regarde si nd_dest n'est pas enfant de nd_src !
    If left$(nd_dest.key, 1) <> "L" Then
        Set nd = nd_dest.Parent
        While left$(nd.key, 1) <> "L"
            If nd.Index = nd_src.Index Then
                Call MsgBox("Vous ne pouvez déplacer le service dans un des ses fils !", vbExclamation + vbOKOnly, "")
                deplacer_sp = P_OK
                Exit Function
            End If
            Set nd = nd.Parent
        Wend
        numsrv_dest = CLng(Mid$(nd_dest.key, 2))
    Else
        numsrv_dest = 0
    End If
    numsrv_src = CLng(Mid$(nd_src.key, 2))
    
    key_depl = nd_src.key
    s_sp_src = ""
    Set nd = nd_src
    While left$(nd.key, 1) <> "L"
        s_sp_src = nd.key & ";" & s_sp_src
        Set nd = nd.Parent
    Wend
        
    s_sp_dest = nd_src.key & ";"
    Set nd = nd_dest
    While left$(nd.key, 1) <> "L"
        s_sp_dest = nd.key & ";" & s_sp_dest
        Set nd = nd.Parent
    Wend
    
    If Odbc_BeginTrans() = P_ERREUR Then
        deplacer_sp = P_ERREUR
        Exit Function
    End If
    ' Mise à jour Service : SRV_NumPere
    If stype_src = "P" Then
        If Odbc_Update("Poste", _
                        "PO_Num", _
                        "where PO_Num=" & numsrv_src, _
                        "PO_SRVNum", numsrv_dest) = P_ERREUR Then
            GoTo err_enreg
        End If
    Else
        If Odbc_Update("Service", _
                        "SRV_Num", _
                        "where SRV_Num=" & numsrv_src, _
                        "SRV_NumPere", numsrv_dest) = P_ERREUR Then
            GoTo err_enreg
        End If
    End If
    ' Mise à jour Documentation
    sql = "select DO_Num, DO_Dest from Documentation" _
        & " where DO_Dest like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        s_sp = STR_Remplacer(rs("DO_Dest").Value & "", s_sp_src, s_sp_dest)
        If Odbc_Update("Documentation", _
                        "DO_Num", _
                        "where DO_Num=" & rs("DO_Num").Value, _
                        "DO_Dest", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close
    ' Mise à jour Dossier
    sql = "select DS_Num, DS_Dest from Dossier" _
        & " where DS_Dest like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        s_sp = STR_Remplacer(rs("DS_Dest").Value, s_sp_src, s_sp_dest)
        If Odbc_Update("Dossier", _
                        "DS_Num", _
                        "where DS_Num=" & rs("DS_Num").Value, _
                        "DS_Dest", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close
    ' Mise à jour Document
    sql = "select D_Num, D_Dest from Document" _
        & " where D_Dest like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        s_sp = STR_Remplacer(rs("D_Dest").Value & "", s_sp_src, s_sp_dest)
        If Odbc_Update("Document", _
                        "D_Num", _
                        "where D_Num=" & rs("D_Num").Value, _
                        "D_Dest", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close
    ' Mise à jour GroupeUtil
    sql = "select GU_Num, GU_Lst from GroupeUtil" _
        & " where GU_Lst like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        s_sp = STR_Remplacer(rs("GU_Lst").Value, s_sp_src, s_sp_dest)
        If Odbc_Update("GroupeUtil", _
                        "GU_Num", _
                        "where GU_Num=" & rs("GU_Num").Value, _
                        "GU_Lst", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close
    ' Mise à jour Utilisateur
    sql = "select U_Num, U_SPM from Utilisateur" _
        & " where U_Spm like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        nbch = STR_GetNbchamp(rs("U_SPM").Value, "|")
        s_sp = ""
        For I = 1 To nbch
            s = STR_GetChamp(rs("U_SPM").Value, "|", I - 1)
            s = Replace(s, s_sp_src, s_sp_dest)
            s_sp = s_sp + s + "|"
        Next I
        If Odbc_Update("Utilisateur", _
                        "U_Num", _
                        "where U_Num=" & rs("U_Num").Value, _
                        "U_SPM", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close
    If stype_src = "S" Then
        ' Mise à jour DocPrmDiffusion
        sql = "select D_Num from Document" _
            & " where D_Dest like '%S" & numsrv_dest & ";|%'"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            GoTo err_enreg
        End If
        If Not rs.EOF Then
            sql = "select u_num from utilisateur where u_spm like '%S" & numsrv_src & ";%'"
            If Odbc_SelectV(sql, rsu) = P_ERREUR Then
                GoTo err_enreg
            End If
            nutil = -1
            While Not rsu.EOF
                nutil = nutil + 1
                ReDim Preserve tbl_util(nutil) As Long
                tbl_util(nutil) = rsu("u_num").Value
                rsu.MoveNext
            Wend
            rsu.Close
            If nutil >= 0 Then
                While Not rs.EOF
                    For iu = 0 To nutil
                        If P_AjouterDocUtil_Dest(rs("d_num").Value, tbl_util(iu), -1) = P_ERREUR Then
                            GoTo err_enreg
                        End If
                    Next iu
                    rs.MoveNext
                Wend
            End If
        End If
        rs.Close
    End If
    Call Odbc_CommitTrans
    
    Set nd_src.Parent = nd_dest
    Set tv.SelectedItem = tv.Nodes(key_depl)
    tv.Nodes(key_depl).EnsureVisible
    
    ' LN - Permet de supprimer la ligne "A charger" : je n'ai pas réussi autrement ...
    If nd_dest.Children > 0 Then
        Set nd = nd_dest.Child
        For I = 1 To nd_dest.Children
            If nd.Text = "A charger" Then
                tv.Nodes.Remove (nd.Index)
                Exit For
            End If
            If I < nd_dest.Children Then
                Set nd = nd.Next
            End If
        Next I
    End If
    
    deplacer_sp = P_OK
    Exit Function

err_enreg:
    Call Odbc_RollbackTrans
    deplacer_sp = P_ERREUR
    
End Function

Private Function fct_dans_dest_do(ByVal v_numfct As Long) As Integer

    Dim sql As String
    Dim lnb As Long
    
    sql = "select count(*) from Documentation" _
        & " where DO_Dest like '%F" & v_numfct & "|%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        fct_dans_dest_do = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        fct_dans_dest_do = 1
        Exit Function
    End If
    
    sql = "select count(*) from Dossier" _
        & " where DS_Dest like '%F" & v_numfct & "|%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        fct_dans_dest_do = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        fct_dans_dest_do = 2
        Exit Function
    End If
    
    sql = "select count(*) from Document" _
        & " where D_Dest like '%F" & v_numfct & "|%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        fct_dans_dest_do = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        fct_dans_dest_do = 3
        Exit Function
    End If
    
    fct_dans_dest_do = 0
    
End Function

Private Function fct_dans_donnees_form(ByVal v_numfct As Long) As Integer

    Dim sql As String
    Dim lnb As Long
    Dim rs As rdoResultset
    
    sql = "select distinct forec_fornum, forec_nom from formetapechp" _
        & " where forec_fctvalid='%NUMFCT'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        fct_dans_donnees_form = P_OUI
        Exit Function
    End If
    While Not rs.EOF
        sql = "select count(*) from donnees_" & rs("forec_fornum").Value _
            & " where " & rs("forec_nom").Value & " like '" & v_numfct & "#%'"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            fct_dans_donnees_form = P_OUI
            Exit Function
        End If
        If lnb > 0 Then
            fct_dans_donnees_form = P_OUI
            Exit Function
        End If
        rs.MoveNext
    Wend
    rs.Close
            
    fct_dans_donnees_form = P_NON
    
End Function

Private Function fct_dans_util(ByVal v_numfct As Long) As Integer

    Dim sql As String
    Dim lnb As Long
    
    ' Utilisateurs actifs associés à cette fonction
    sql = "select count(*) from Utilisateur" _
        & " where U_FctTrav like '%F" & v_numfct & ";%'" _
        & " and U_Actif=true"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        fct_dans_util = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        fct_dans_util = 1
        Exit Function
    End If
    
    ' Utilisateurs inactifs associés à cette fonction
    sql = "select count(*) from Utilisateur" _
        & " where U_FctTrav like '%F" & v_numfct & ";%'" _
        & " and U_Actif=false"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        fct_dans_util = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        fct_dans_util = 2
        Exit Function
    End If
    
    fct_dans_util = 0
    
End Function

Private Function FctNivRemplace(ByVal v_FT_NivRemplace, ByVal v_numsrv, ByRef r_strRemplace As String) As Integer
    ' Voir si pour ce service il a un père du niveau de remplacement indiqué
    Dim sql As String, rs As rdoResultset
    Dim encore As Boolean
    Dim ilya As Boolean
    Dim srvnum As Long
    Dim srvnom_prem As String
    Dim strNiveau As String
    
    If v_FT_NivRemplace = 0 Then
        FctNivRemplace = 0
    Else
        srvnum = v_numsrv
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
                FctNivRemplace = -1
                rs.Close
                Exit Function
            ElseIf rs("SRV_NivsNum") = v_FT_NivRemplace Then
                Call Odbc_RecupVal("select nivs_nom from niveau_structure where nivs_num='" & v_FT_NivRemplace & "'", strNiveau)
                r_strRemplace = "Niveau de remplacement : " & strNiveau & " => " & rs("SRV_Nom")
                FctNivRemplace = 1
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

Private Sub imprimer()

    Dim s As String, stexte(0) As String
    Dim fl_fax As Boolean, fl_bid As Boolean, encore As Boolean
    Dim decal As Integer
    Dim nd As Node, ndp As Node, nds As Node, ndf As Node
    
    If tv.SelectedItem.Children = 0 Then
        Call MsgBox("Rien à imprimer pour " & tv.SelectedItem.Text, vbOKOnly + vbInformation, "")
        Exit Sub
    End If
    
    ' Choix de l'imprimante
    If PR_ChoixImp(False, False, fl_fax, fl_bid) = False Then Exit Sub
        
    On Error Resume Next
    Printer.ScaleMode = vbTwips
    Printer.PaperSize = vbPRPSA4
    On Error GoTo err_printer
    
    stexte(0) = ""
    Set ndf = tv.SelectedItem
    If left$(ndf.key, 1) = "L" Then
        s = "Services - " & ndf.Text
    Else
        s = ndf.Text
    End If
    Call PR_InitFormat(10, _
                       True, _
                       s, _
                       False, _
                       "g", _
                       stexte())
    Set nd = ndf
    If left$(nd.key, 1) = "S" Then
        If STR_GetChamp(nd.tag, "|", 1) = False Then
            tv.Nodes.Remove (nd.Child.Index)
            charger_service (Mid$(nd.key, 2))
        End If
    End If
    Set ndf = ndf.Child
    encore = True
    decal = 0
    While encore
        If left$(ndf.key, 1) = "S" Then
            If STR_GetChamp(ndf.tag, "|", 1) = False Then
                tv.Nodes.Remove (ndf.Child.Index)
                charger_service (Mid$(ndf.key, 2))
            End If
        End If
        If decal > 0 Then
            s = String(decal * 5, " ")
        Else
            s = ""
        End If
        s = s & left$(ndf.key, 1) & " - "
        stexte(0) = s & ndf.Text
        Call PR_ImpLigne(stexte())
        If ndf.Children > 0 Then
            Set ndf = ndf.Child
            decal = decal + 1
        Else
            encore = False
            While ndf.Index <> nd.Index And Not encore
                encore = True
                Set nds = ndf
                Set ndf = ndf.Next
                On Error GoTo pas_de_suivant
                s = ndf.key
                On Error GoTo 0
                If Not encore Then
                    Set ndf = nds.Parent
                    decal = decal - 1
                End If
            Wend
            If ndf.Index = nd.Index Then
                encore = False
            End If
        End If
    Wend
    Printer.EndDoc
    Call PR_RestoreImp
    Exit Sub
    
pas_de_suivant:
    encore = False
    Resume Next
    
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

    Dim I As Integer, nbchp As Integer, nsel As Integer
    Dim sql As String, rs As rdoResultset
    Dim lnb As Long
    
    g_crfct_autor = P_UtilEstAutorFct("CR_FCTTRAV")
    g_crutil_autor = P_UtilEstAutorFct("CR_UTIL")
    g_modutil_autor = P_UtilEstAutorFct("MOD_UTIL")
    g_crspm_autor = P_UtilEstAutorFct("CR_SPM")
    g_modspm_autor = P_UtilEstAutorFct("MOD_SPM")
    g_supspm_autor = P_UtilEstAutorFct("SUPP_SPM")
    If p_appli_kalibottin > 0 Then
        g_crpiece_autor = P_UtilEstAutorFct("CR_SPM")
        g_modpiece_autor = P_UtilEstAutorFct("MOD_SPM")
        g_suppiece_autor = P_UtilEstAutorFct("SUPP_SPM")
    End If
    g_node_crt = 1
    
    If g_mode_acces = MODE_SELECT Then
        nsel = -1
        On Error Resume Next
        nsel = UBound(CL_liste.lignes)
        On Error GoTo 0
        If nsel >= 0 Then
            ReDim g_lignes(nsel) As CL_SLIGNE
            For I = 0 To nsel
                g_lignes(I) = CL_liste.lignes(I)
            Next I
        Else
            Erase g_lignes()
        End If
        cmd(CMD_IMPRIMER).Visible = False
    Else
        p_NumLabo = p_NumLaboDefaut
        p_CodeLabo = p_CodeLaboDefaut
        cmd(CMD_OK).Visible = False
        cmd(CMD_IMPRIMER).left = cmd(CMD_OK).left
    End If
    
    ' charger le combo des niveaux
    g_ya_niveau = False
    sql = "select count(*) from niveau_structure"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
    End If
    
    If lnb > 0 Then
        g_ya_niveau = True
        Me.CmbNiveau.AddItem "Tous"
        Me.CmbNiveau.ItemData(Me.CmbNiveau.ListCount - 1) = 0
        sql = "select Nivs_Nom, Nivs_Num from niveau_structure Order by Nivs_Num"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            Call quitter
            Exit Sub
        Else
            While Not rs.EOF
                Me.CmbNiveau.AddItem rs("Nivs_Nom")
                Me.CmbNiveau.ItemData(Me.CmbNiveau.ListCount - 1) = rs("Nivs_Num")
                rs.MoveNext
            Wend
        End If
        Me.CmbNiveau.ListIndex = 0
    Else
        Me.CmbNiveau.Visible = False
    End If
    
    If g_mode_acces <> MODE_PARAM_PERS Then
        If afficher_liste() = P_ERREUR Then
            Call quitter
            Exit Sub
        End If
    Else
        If afficher_liste2() = P_ERREUR Then
            Call quitter
            Exit Sub
        End If
    End If
    
'    Set tv.SelectedItem = tv.Nodes(1)
    tv.SetFocus
'    SendKeys "{PGDN}"
'    SendKeys "{HOME}"
'    DoEvents
    
End Sub

Private Sub majLibDetailSRV(ByVal v_tv_key As String, ByVal v_tv_tag As String)
    Dim sql As String
    Dim srv_code As String, srv_stru_import As String
    Dim srv_num As Long
    Dim u_matricule As String
    
    LbldetailSRV.Visible = False
    If Mid(v_tv_key, 1, 1) = "S" Then
        sql = "select srv_code,srv_stru_import,srv_num from service where srv_num=" & Replace(v_tv_key, "S", "")
        Call Odbc_RecupVal(sql, srv_code, srv_stru_import, srv_num)
        LbldetailSRV.Visible = True
        LbldetailSRV.Caption = "Num=" & srv_num
        LbldetailSRV.Caption = LbldetailSRV.Caption & "   " & IIf(srv_code <> "", "Code " & srv_code, "")
        LbldetailSRV.Caption = LbldetailSRV.Caption & "   " & IIf(srv_stru_import <> "", "Import " & srv_stru_import, "")
        LbldetailSRV.Caption = Trim(LbldetailSRV.Caption)
    ElseIf Mid(v_tv_tag, 1, 1) = "U" Then
        sql = "select u_matricule from utilisateur where u_num=" & Replace(v_tv_tag, "U", "")
        Call Odbc_RecupVal(sql, u_matricule)
        LbldetailSRV.Visible = True
        LbldetailSRV.Caption = IIf(u_matricule <> "", "Matricule " & u_matricule, "")
        LbldetailSRV.Visible = IIf(Trim(LbldetailSRV.Caption) <> "", True, False)
    End If
End Sub

Private Sub maj_fct_util(ByVal v_numutil As Long, _
                         ByVal v_sp As Variant)
    
    Dim sfct As String, s As String, spo As String, sql As String
    Dim n As Integer, n2 As Integer, I As Integer
    Dim numfct As Long
    
    sfct = ""
    n = STR_GetNbchamp(v_sp, "|")
    For I = 0 To n - 1
        s = STR_GetChamp(v_sp, "|", I)
        n2 = STR_GetNbchamp(s, ";")
        spo = STR_GetChamp(s, ";", n2 - 1)
        sql = "select po_ftnum from poste where po_num=" & Mid$(spo, 2)
        Call Odbc_RecupVal(sql, numfct)
        If InStr(sfct, "F" & numfct & ";") = 0 Then
            sfct = sfct & "F" & numfct & ";"
        End If
    Next I
    Call Odbc_Update("Utilisateur", "U_Num", "where U_Num=" & v_numutil, _
                     "U_Fcttrav", sfct)
                     
End Sub

Private Sub modifier_coordonnees(ByVal v_stype As String, _
                                 ByVal v_num As Long)

    Dim sutil As String, url As String
    
    If p_chemin_webbrowser = "" Then
        Call MsgBox("Aucun navigateur n'a été trouvé.", vbInformation + vbOKOnly, "")
        Exit Sub
    End If
    
    sutil = STR_CrypterNombre(format(p_NumUtil, "#0000000"))
    url = p_CheminPHP & "/pident.php?in=kb_coord&V_util=" & sutil _
        & "&V_niveau=3&V_nument=" & v_num & "&V_type=" & v_stype
    If p_sversconf <> "" Then
        url = url & "&s_vers_conf=" & p_sversconf
    End If
    ' Chargement de la page
    Shell p_chemin_webbrowser & " " & url, vbMaximizedFocus

End Sub

Private Function modifier_creer_piece() As Integer

    Dim lib As String, sql As String, s As String
    Dim nch As Integer
    Dim num As Long, numsrv As Long
    Dim nd As Node, ndpi As Node
    Dim frm As Form
    
    Set nd = tv.SelectedItem
    If left$(nd.key, 1) = "L" Then
        num = 0
        numsrv = 0
    ElseIf left$(nd.key, 1) = "S" Then
        num = 0
        numsrv = CLng(Mid$(nd.key, 2))
    Else
        num = CLng(Mid$(nd.key, 2))
    End If
    
lab_debut:
    If num > 0 Then
        sql = "select PC_Nom from Piece" _
            & " where PC_Num=" & num
        If Odbc_RecupVal(sql, lib) = P_ERREUR Then
            modifier_creer_piece = P_ERREUR
            Exit Function
        End If
    Else
        lib = ""
    End If
    
    Call SAIS_Init
    Call SAIS_InitTitreHelp("Pièce", "dico_d_spm.htm")
    Call SAIS_AddBouton("", p_chemin_appli + "\btnEnregistrer.gif", vbKeyE, vbKeyF1, 0)
    nch = 0
    If num > 0 Then
        If left$(nd.Parent.key, 1) <> "L" Then
            Call SAIS_AddChamp("Service", -50, 0, True, nd.Parent.Text)
            nch = 1
        End If
    Else
        If left$(nd.key, 1) <> "L" Then
            Call SAIS_AddChamp("Service", -50, 0, True, nd.Text)
            nch = 1
        End If
    End If
    Call SAIS_AddChamp("Nom", 80, SAIS_TYP_TOUT_CAR, False, lib)
    If p_appli_kalibottin > 0 And num > 0 Then
        Call SAIS_AddBouton("Coordonnées", "", 0, 0, 1500)
    End If
    Call SAIS_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Saisie.Show 1
    If SAIS_Saisie.retour = 1 And num > 0 And p_appli_kalibottin > 0 Then
        Call modifier_coordonnees("C", num)
        GoTo lab_debut
    End If
    If SAIS_Saisie.retour <> SAIS_RET_MODIF Then
        modifier_creer_piece = P_NON
        Exit Function
    End If
    lib = SAIS_Saisie.champs(nch).sval
    
    If num > 0 Then
        Call Odbc_Update("Piece", _
                         "PC_Num", _
                         "where PC_Num=" & num, _
                         "PC_Nom", lib)
        nd.Text = lib
    Else
        Call Odbc_AddNew("Piece", "PC_Num", "pc_seq", True, num, _
                         "PC_SRVNum", numsrv, _
                         "PC_Nom", lib)
        Set ndpi = tv.Nodes.Add(nd, _
                               tvwChild, _
                               "C" & num, _
                               lib, _
                               IMG_PIECE, _
                               IMG_PIECE)
        ndpi.tag = True
    End If
    
    modifier_creer_piece = P_OK
    
End Function

Private Function modifier_poste() As Integer

    Dim libposte As String, libfct As String, sql As String, s As String
    Dim nch As Integer, ft_visible As Integer
    Dim num As Long, numfct As Long, ft_niveau As Long
    Dim nd As Node
    Dim frm As Form
    
    Set nd = tv.SelectedItem
    num = CLng(Mid$(nd.key, 2))
    
lab_debut:
    sql = "select PO_Libelle, FT_Num, FT_Libelle, FT_Visible, FT_Niveau from Poste, FctTrav" _
        & " where PO_Num=" & num _
        & " and FT_Num=PO_FTNum"
    If Odbc_RecupVal(sql, libposte, numfct, libfct, ft_visible, ft_niveau) = P_ERREUR Then
        modifier_poste = P_ERREUR
        Exit Function
    End If
    
    Call SAIS_Init
    Call SAIS_InitTitreHelp("Poste", "dico_d_spm.htm")
    Call SAIS_AddBouton("", p_chemin_appli + "\btnEnregistrer.gif", vbKeyE, vbKeyF1, 0)
    Call SAIS_AddBouton("&Accès Fonction", "", 0, 0, 1000)
    Call SAIS_AddChamp("Fonction", -80, 0, True, libfct)
    If left$(nd.Parent.key, 1) <> "L" Then
        Call SAIS_AddChamp("Service", -50, 0, True, nd.Parent.Text)
        nch = 2
    Else
        nch = 1
    End If
    Call SAIS_AddChamp("Nom", 80, SAIS_TYP_TOUT_CAR, False, libposte)
    If p_appli_kalibottin > 0 Then
        sql = "SELECT FNIV_Libelle FROM FCT_Niveau WHERE FNIV_Num=" & ft_niveau
        Call Odbc_RecupVal(sql, s)
        Call SAIS_AddChamp("Niveau de coordonnées accessibles", -30, 0, True, ft_niveau & " - " & s)
        If ft_visible = 0 Then
            s = "Jamais"
        ElseIf ft_visible = 1 Then
            s = "Vue détaillée seulement"
        Else
            s = "Toujours"
        End If
        Call SAIS_AddChamp("Visible dans l'annuaire", -25, 0, True, s)
        Call SAIS_AddBouton("Coordonnées", "", 0, 0, 1500)
    End If
    Call SAIS_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Saisie.Show 1
    If SAIS_Saisie.retour = 1 Then
        Set frm = KS_PrmFonction
        Call KS_PrmFonction.AppelFrm("DIRECT", numfct)
        Set frm = Nothing
        GoTo lab_debut
    End If
    If SAIS_Saisie.retour = 2 And p_appli_kalibottin > 0 Then
        Call modifier_coordonnees("P", num)
        GoTo lab_debut
    End If
    If SAIS_Saisie.retour <> SAIS_RET_MODIF Then
        modifier_poste = P_NON
        Exit Function
    End If
    libposte = SAIS_Saisie.champs(nch).sval
    
    Call Odbc_Update("Poste", _
                     "PO_Num", _
                     "where PO_Num=" & num, _
                     "PO_Libelle", libposte)
    If libposte <> libfct Then
        libfct = libfct & " *"
    End If
    nd.Text = libfct
    
    modifier_poste = P_OK
    
End Function

Private Function modifier_service() As Integer

    Dim code As String, lib As String, libcourt As String, smasque As String
    Dim gerer_niveau As Boolean
    Dim first_chp As Integer, srv_visible As Integer, n As Integer
    Dim num As Long, lnb As Long, nivs_num As Long
    Dim nd As Node
    Dim liste_nomtable As String, liste_lsttypchp As String, liste_chpretour As String, liste_chpnum   As String
    Dim liste_multiselect As Boolean
    Dim nivs_code As String, s As String
    Dim b_sel As Boolean
    Dim sql As String, rs As rdoResultset
    
    Set nd = tv.SelectedItem
    num = CLng(Mid$(nd.key, 2))
    Call Odbc_RecupVal("select srv_code, srv_nom, srv_libcourt, srv_nivsnum, srv_code_masque, srv_visible from service where srv_num=" & num, _
                        code, lib, libcourt, nivs_num, smasque, srv_visible)
    
lab_saisie:
    Call SAIS_Init
    Call SAIS_InitTitreHelp("Service", "dico_d_spm.htm")
    Call SAIS_AddBouton("", p_chemin_appli + "\btnEnregistrer.gif", vbKeyE, vbKeyF1, 0)
    If left$(nd.Parent.key, 1) <> "L" Then
        Call SAIS_AddChamp("Rattaché à", -50, 0, True, nd.Parent.Text)
        first_chp = 1
    Else
        first_chp = 0
    End If
    Call SAIS_AddChamp("Code", 15, SAIS_TYP_TOUT_CAR, True, code)
    Call SAIS_AddChamp("Nom", 120, SAIS_TYP_TOUT_CAR, False, lib)
    Call SAIS_AddChamp("Nom court", 30, SAIS_TYP_TOUT_CAR, True, libcourt)
    Call SAIS_AddChamp("Masque", 30, SAIS_TYP_TOUT_CAR, True, smasque)

    gerer_niveau = False
    sql = "select count(*) from niveau_structure"
    If Odbc_Count(sql, lnb) = P_OK Then
        If lnb > 0 Then
            gerer_niveau = True
        End If
    End If
    If gerer_niveau Then
        liste_nomtable = "select * from niveau_structure"
        liste_multiselect = False
        liste_chpretour = "nivs_code"
        liste_chpnum = "nivs_num"
        Call SAIS_AddListe("Niveau", liste_nomtable, liste_multiselect, liste_chpretour, liste_chpnum, SAIS_TYP_CHOIXLISTE, True, nivs_code)
        ' Ajouter les champs listes
        sql = "select * from niveau_structure"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            MsgBox "Erreur SQL " & sql
            Exit Function
        End If
        While Not rs.EOF
            Call SAIS_AddItemListe(UBound(SAIS_Saisie.champs), rs("nivs_num"), rs("nivs_code"), rs("nivs_nom"), (rs("nivs_num") = nivs_num))
            rs.MoveNext
        Wend
        rs.Close
    End If
    If p_appli_kalibottin > 0 Then
        Call SAIS_AddListe("Visible dans l'annuaire", "", False, "", "", SAIS_TYP_CHOIXLISTE, True, "")
        Call SAIS_AddItemListe(UBound(SAIS_Saisie.champs), 0, "0", "Jamais", IIf(srv_visible = 0, True, False))
        Call SAIS_AddItemListe(UBound(SAIS_Saisie.champs), 1, "1", "Vue détaillée seulement", IIf(srv_visible = 1, True, False))
        Call SAIS_AddItemListe(UBound(SAIS_Saisie.champs), 2, "2", "Toujours", IIf(srv_visible = 2, True, False))
        Call SAIS_AddBouton("Coordonnées", "", 0, 0, 1500)
    End If
    
    Call SAIS_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Saisie.Show 1
    
    If SAIS_Saisie.retour = 1 And p_appli_kalibottin > 0 Then
        Call modifier_coordonnees("S", num)
        GoTo lab_saisie
    End If
    
    If SAIS_Saisie.retour <> SAIS_RET_MODIF Then
        modifier_service = P_NON
        Exit Function
    End If
    code = SAIS_Saisie.champs(first_chp).sval
    If code <> "" Then
        Call Odbc_Count("select count(*) from service where srv_code=" & Odbc_String(SAIS_Saisie.champs(first_chp).sval) & " and srv_num<>" & num, lnb)
        If lnb > 0 Then
            Call MsgBox("Le code '" & code & "' est déjà attribué." & vbCrLf & vbCrLf & "Veuillez choisir un autre code.", vbInformation + vbOKOnly, "")
            GoTo lab_saisie
        End If
    End If
    lib = SAIS_Saisie.champs(first_chp + 1).sval
    libcourt = SAIS_Saisie.champs(first_chp + 2).sval
    smasque = SAIS_Saisie.champs(first_chp + 3).sval
    
    nivs_num = 0
    If gerer_niveau Then
        nivs_code = SAIS_Saisie.champs(first_chp + 4).sval
        If nivs_code <> "" Then
            Call Odbc_RecupVal("select nivs_num from niveau_structure where nivs_code='" & nivs_code & "'", _
                                nivs_num)
        End If
        n = 1
    Else
        n = 0
    End If
    
    If p_appli_kalibottin > 0 Then
        srv_visible = CInt(SAIS_Saisie.champs(first_chp + 4 + n).sval)
    Else
        srv_visible = 2
    End If
    
    Call Odbc_Update("Service", _
                     "SRV_Num", _
                     "where SRV_Num=" & num, _
                     "SRV_code", code, _
                     "SRV_code_masque", smasque, _
                     "SRV_Nom", lib, _
                     "SRV_NivsNum", nivs_num, _
                     "SRV_visible", srv_visible, _
                     "SRV_libcourt", libcourt)
    
    s = recup_lib_niveau(nivs_num)
    If s <> "" Then
        lib = lib & " (" & s & ")"
    End If
    nd.Text = lib
    
    modifier_service = P_OK
    
End Function

Private Sub ouvrir_serv_poste()

    Dim sql As String
    Dim encore As Boolean
    Dim numposte As Long
    Dim nd As Node
    Dim rs As rdoResultset
    
    If g_ouvrir <> "" Then
        numposte = 0
        If left$(g_ouvrir, 1) = "P" Then
            numposte = Mid$(g_ouvrir, 2)
        ElseIf left$(g_ouvrir, 1) = "U" Then
            sql = "select U_Po_Princ from Utilisateur" _
                & " where U_Num=" & Mid$(g_ouvrir, 2)
            If Odbc_SelectV(sql, rs) = P_OK Then
                If Not rs.EOF Then
                    numposte = rs("U_Po_Princ").Value
                End If
                rs.Close
            End If
        End If
        If numposte > 0 Then
            If TV_NodeExiste(tv, "P" & numposte, nd) = P_NON Then
                Call charger_arbor("P" & numposte)
            End If
            Set nd = tv.Nodes("P" & numposte)
            encore = True
            While encore
                If nd.Index = nd.Root.Index Then
                    encore = False
                Else
                    Set nd = nd.Parent
                    nd.Expanded = True
                End If
            Wend
            tv.SetFocus
            Set tv.SelectedItem = tv.Nodes("P" & numposte)
            SendKeys "{DOWN}"
            SendKeys "{UP}"
            DoEvents
            Set tv.SelectedItem = tv.Nodes("P" & numposte)
            If left$(g_ouvrir, 1) = "U" Then
                Call ajouter_pers_tv
            End If
        End If
    Else
        Set tv.SelectedItem = tv.Nodes(1).Root
    End If
    
End Sub

Private Function po_dans_histordoc(ByVal v_num As Long) As Boolean

    Dim sql As String
    Dim lnb As Long
    
    sql = "select count(*) from DocEtapeVersion" _
        & " where DEV_PONum=" & v_num
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        po_dans_histordoc = True
        Exit Function
    End If
    If lnb = 0 Then
        po_dans_histordoc = False
    Else
        po_dans_histordoc = True
    End If

End Function

Private Function poste_est_posteresp(ByVal v_numposte As Long) As Boolean

    Dim sql As String
    Dim lnb As Long
    
    sql = "select count(*) from poste where po_numresp=" & v_numposte
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        poste_est_posteresp = False
        Exit Function
    End If
    If lnb > 0 Then
        Call MsgBox("Attention : Ce poste est un poste RESPONSABLE." & vbCrLf & "Vous devrez aller indiquer le nouveau poste dans le dictionnaire des responsables.", vbInformation + vbOKOnly, "")
        poste_est_posteresp = True
    Else
        poste_est_posteresp = False
    End If
    
End Function

Private Sub quitter()
    
    g_sret = ""
    
    Unload Me
    
End Sub

Private Sub selectionner()

    Dim sp As String, sm As String, s As String
    Dim encore As Boolean
    Dim n As Integer, I As Integer, j As Integer, nbch As Integer, img As Integer
    Dim nd As Node, ndp As Node
    
    If g_plusieurs Then
        n = 0
        For I = 2 To tv.Nodes.Count
            img = tv.Nodes(I).image
            If img = IMG_POSTE_SEL Or img = IMG_SRV_SEL Or img = IMG_POSTE_SEL_NOMOD Or img = IMG_SRV_SEL_NOMOD Then
                sp = tv.Nodes(I).key & ";"
                Set nd = tv.Nodes(I)
                encore = True
                Do
                    Set ndp = nd.Parent
                    s = ndp.key
                    If left$(s, 1) = "L" Then
                        encore = False
                    Else
                        sp = sp + s + ";"
                        Set nd = ndp
                    End If
                Loop Until Not encore
                s = ""
                nbch = STR_GetNbchamp(sp, ";")
                For j = nbch To 1 Step -1
                    s = s + STR_GetChamp(sp, ";", j - 1) + ";"
                Next j
                ReDim Preserve CL_liste.lignes(n)
                CL_liste.lignes(n).texte = s
                If img = IMG_POSTE_SEL Or img = IMG_SRV_SEL Then
                    CL_liste.lignes(n).tag = True
                Else
                    CL_liste.lignes(n).tag = False
                End If
                n = n + 1
            End If
        Next I
        
        ' Cas ou l'on souhaite pouvoir selectionner tous les services
        If g_smode = "C" Then
            If n = 0 Then
                If tv.Nodes.Item(1) = tv.SelectedItem Then
                    ReDim Preserve CL_liste.lignes(n)
                    CL_liste.lignes(n).texte = tv.SelectedItem
                    CL_liste.lignes(n).tag = True
                End If
            End If
        End If
        
        g_sret = "N" & n
    Else
        Set nd = tv.SelectedItem
        If InStr(g_stype, left$(nd.key, 1)) = 0 Or nd.key = "" Then
            If nd.key = "" Then
                Call MsgBox("Vous ne pouvez pas sélectionner une personne.", vbInformation + vbOKOnly, "")
            ElseIf left$(nd.key, 1) = "S" Then
                Call MsgBox("Vous ne pouvez pas sélectionner un service.", vbInformation + vbOKOnly, "")
            Else
                Call MsgBox("Vous ne pouvez pas sélectionner un poste.", vbInformation + vbOKOnly, "")
            End If
            Exit Sub
        End If
        sp = nd.key & ";"
        If left$(sp, 1) = "L" Then
            encore = False
        Else
            encore = True
        End If
        While encore
            Set ndp = nd.Parent
            s = ndp.key
            If left$(s, 1) = "L" Then
                encore = False
            Else
                sp = sp + s + ";"
                Set nd = ndp
            End If
        Wend
        s = ""
        nbch = STR_GetNbchamp(sp, ";")
        For j = nbch To 1 Step -1
            s = s + STR_GetChamp(sp, ";", j - 1) + ";"
        Next j
        g_sret = s
    End If
    
    Unload Me
    Exit Sub
    
lab_no_prev:
    encore = False
    Resume Next
    
End Sub

Private Function serv_dans_donnees_form(ByVal v_numsrv As Long) As Integer

    Dim sql As String
    Dim lnb As Long
    Dim rs As rdoResultset
    
    sql = "select distinct forec_fornum, forec_nom from formetapechp" _
        & " where forec_fctvalid='%NUMSERVICE'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        serv_dans_donnees_form = P_OUI
        Exit Function
    End If
    While Not rs.EOF
        sql = "select count(*) from donnees_" & rs("forec_fornum").Value _
            & " where " & rs("forec_nom").Value & " like '" & v_numsrv & "#%'"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            serv_dans_donnees_form = P_OUI
            Exit Function
        End If
        If lnb > 0 Then
            serv_dans_donnees_form = P_OUI
            Exit Function
        End If
        rs.MoveNext
    Wend
    rs.Close
            
    serv_dans_donnees_form = P_NON
    
End Function

Private Function srvpo_dans_do(ByVal v_stype As String, _
                               ByVal v_num As Long, _
                               ByRef r_cas As Integer) As Boolean

    Dim sql As String
    Dim lnb As Long
    
    r_cas = 0
    sql = "select count(*) from Documentation" _
        & " where DO_Dest like '%" & v_stype & v_num & ";%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        srvpo_dans_do = True
        Exit Function
    End If
    If lnb > 0 Then
        srvpo_dans_do = True
        Exit Function
    End If

    r_cas = 1
    sql = "select count(*) from Dossier" _
        & " where DS_Dest like '%" & v_stype & v_num & ";%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        srvpo_dans_do = True
        Exit Function
    End If
    If lnb > 0 Then
        srvpo_dans_do = True
        Exit Function
    End If

    r_cas = 2
    sql = "select count(*) from Document" _
        & " where D_Dest like '%" & v_stype & v_num & ";%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        srvpo_dans_do = True
        Exit Function
    End If
    If lnb > 0 Then
        srvpo_dans_do = True
        Exit Function
    End If

    If v_stype = "S" Then
        r_cas = 3
        sql = "select count(*) from Document" _
            & " where D_srvnum_emet=" & v_num
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            srvpo_dans_do = True
            Exit Function
        End If
        If lnb > 0 Then
            srvpo_dans_do = True
            Exit Function
        End If
        sql = "select count(*) from Dossier" _
            & " where DS_srvnum_emet=" & v_num
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            srvpo_dans_do = True
            Exit Function
        End If
        If lnb > 0 Then
            srvpo_dans_do = True
            Exit Function
        End If
    End If

    srvpo_dans_do = False
    
End Function

Private Function srvpo_dans_util(ByVal v_stype As String, _
                                 ByVal v_num As Long, _
                                 ByRef r_yaactif As Boolean) As Boolean

    Dim sql As String
    Dim lnb As Long
    
    sql = "select count(*) from Utilisateur" _
        & " where U_SPM like '%" & v_stype & v_num & ";%'" _
        & " and U_Actif=true"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        r_yaactif = False
        srvpo_dans_util = True
        Exit Function
    End If
    If lnb > 0 Then
        r_yaactif = True
        srvpo_dans_util = True
        Exit Function
    End If
    
    sql = "select count(*) from Utilisateur" _
        & " where U_SPM like '%" & v_stype & v_num & ";%'" _
        & " and U_Actif=false"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        r_yaactif = False
        srvpo_dans_util = True
        Exit Function
    End If
    If lnb > 0 Then
        r_yaactif = False
        srvpo_dans_util = True
        Exit Function
    End If
    
    srvpo_dans_util = False

End Function

Private Function supprimer_piece() As Integer

    Dim reponse As Integer
    Dim num As Long, lnb As Long
    
    num = CLng(Mid$(tv.SelectedItem.key, 2))
    
    ' Il y a des réunions / types de réunions asscociés à cette pièce
    'sql = "select count(*)
    'If Odbc_Count(sql, lnb) = P_ERREUR Then
    '    supprimer = P_ERREUR
    '    Exit Function
    'End If
    'If lnb > 0 Then
    '    Call MsgBox("Veuillez d'abord supprimer les services rattachés à ce service.", vbExclamation + vbOKOnly, "")
    '    supprimer = P_ERREUR
    '    Exit Function
    'End If

    reponse = MsgBox("Confirmez-vous la suppression de la pièce '" & tv.SelectedItem.Text & "' ?", vbQuestion + vbYesNo + vbDefaultButton2, "")
    If reponse = vbNo Then
        supprimer_piece = P_OK
        Exit Function
    End If
    
    Call Odbc_Delete("Piece", "PC_Num", "where PC_Num=" & num, lnb)
    
    tv.Nodes.Remove (tv.SelectedItem.Index)
    tv.Refresh
    
End Function

Private Function supprimer_po_srv() As Integer

    Dim lib As String, sql As String, sobj As String, stype As String
    Dim s As String
    Dim ya_actif As Boolean, ya_resp As Boolean, inhiber As Boolean
    Dim reponse As Integer, cas As Integer, n As Integer, I As Integer
    Dim num As Long, lnb As Long, numsrv As Long, numfct As Long
    Dim lnb_actif As Long, lnb_inactif As Long
    Dim slst As Variant
    Dim nd As Node, ndr As Node
    Dim rs As rdoResultset
    
    num = CLng(Mid$(tv.SelectedItem.key, 2))
    stype = left$(tv.SelectedItem.key, 1)
    
    If stype = "S" Then
        ' Il y a des services fils
        sql = "select count(*) from service where srv_numpere=" & num _
            & " and srv_actif=true"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            supprimer_po_srv = P_ERREUR
            Exit Function
        End If
        If lnb > 0 Then
            Call MsgBox("Veuillez d'abord supprimer les services rattachés à ce service.", vbExclamation + vbOKOnly, "")
            supprimer_po_srv = P_ERREUR
            Exit Function
        End If
        ' Il y a des postes rattachés à ce service
        sql = "select count(*) from poste where po_srvnum=" & num _
            & " and po_actif=true"
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            supprimer_po_srv = P_ERREUR
            Exit Function
        End If
        If lnb > 0 Then
            Call MsgBox("Veuillez d'abord supprimer les postes rattachés à ce service.", vbExclamation + vbOKOnly, "")
            supprimer_po_srv = P_ERREUR
            Exit Function
        End If
        ' Il y a des documents rattachés à ce service (service émetteur)
        sql = "select count(*) from document where d_srvnum_emet=" & num
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            supprimer_po_srv = P_ERREUR
            Exit Function
        End If
        If lnb > 0 Then
            Call MsgBox("Il y a des documents avec ce service comme service émetteur.", vbExclamation + vbOKOnly, "")
            supprimer_po_srv = P_ERREUR
            Exit Function
        End If
        ' Il y a des dossiers rattachés à ce service (service émetteur)
        sql = "select count(*) from dossier where ds_srvnum_emet=" & num
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            supprimer_po_srv = P_ERREUR
            Exit Function
        End If
        If lnb > 0 Then
            Call MsgBox("Il y a des dossiers avec ce service comme service émetteur.", vbExclamation + vbOKOnly, "")
            supprimer_po_srv = P_ERREUR
            Exit Function
        End If
        sobj = "ce service"
    Else
        sobj = "ce poste"
    End If
    
    ' Poste/Service associé à documentation, dossier, document
    If p_appli_kalidoc > 0 Then
        If srvpo_dans_do(stype, num, cas) Then
            Select Case cas
            Case 0
                Call MsgBox("Une (ou plusieurs) documentation est associée à " & sobj & "." & vbCrLf & "Il ne peut donc pas être supprimé.", vbExclamation + vbOKOnly, "")
            Case 1
                Call MsgBox("Un (ou plusieurs) dossier est associé à " & sobj & "." & vbCrLf & "Il ne peut donc pas être supprimé.", vbExclamation + vbOKOnly, "")
            Case 2
                Call MsgBox("Un (ou plusieurs) document est associé à " & sobj & "." & vbCrLf & "Il ne peut donc pas être supprimé.", vbExclamation + vbOKOnly, "")
            Case 3
                Call MsgBox("Un (ou plusieurs) dossier/document possède " & sobj & " comme service émetteur." & vbCrLf & "Il ne peut donc pas être supprimé.", vbExclamation + vbOKOnly, "")
            End Select
            supprimer_po_srv = P_OK
            Exit Function
        End If
    End If
    
    inhiber = False
    ' Utilisateurs associés à ce Poste/Service
    If srvpo_dans_util(stype, num, ya_actif) Then
        ' Utilisateurs actifs associés à ce Poste/Service
        If ya_actif Then
            Call MsgBox("Une (ou plusieurs) personne est associée à " & sobj & "." & vbCrLf & "Il ne peut donc pas être supprimé.", vbExclamation + vbOKOnly, "")
            supprimer_po_srv = P_OK
            Exit Function
        End If
        ' Utilisateurs inactifs associés à ce Poste/Service
        inhiber = True
    End If
    
    If stype = "P" Then
        ya_resp = poste_est_posteresp(num)
        ' Poste dans l'historique des versions -> on l'inhibe
        If po_dans_histordoc(num) Then
            inhiber = True
        End If
    Else
        ' Formulaires avec ce service dans un champ NUMSERVICE
        If serv_dans_donnees_form(num) = P_OUI Then
            inhiber = True
        End If
    End If
    
    Select Case stype
    Case "S"
        If P_RecupSrvNom(num, lib) = P_ERREUR Then
            supprimer_po_srv = P_ERREUR
            Exit Function
        End If
        sobj = "du service " & lib
    Case "P"
        sql = "select PO_SRVNum, PO_FTNum, FT_Libelle from Poste, FctTrav" _
            & " where PO_Num=" & num _
            & " and FT_Num=PO_FTNum"
        If Odbc_RecupVal(sql, numsrv, numfct, lib) = P_ERREUR Then
            supprimer_po_srv = P_ERREUR
            Exit Function
        End If
        sobj = "du poste " & lib
        If P_RecupSrvNom(numsrv, lib) = P_ERREUR Then
            supprimer_po_srv = P_ERREUR
            Exit Function
        End If
        sobj = sobj & " dans le service " & lib
    End Select
    
    If inhiber Then
        reponse = MsgBox(lib & " ne peut être supprimé mais seulement rendu inactif." & vbCrLf & vbCrLf & "Confirmez-vous cette opération ?", vbQuestion + vbYesNo + vbDefaultButton2, "")
    Else
        reponse = MsgBox("Confirmez-vous la suppression " & sobj & " ?", vbQuestion + vbYesNo + vbDefaultButton2, "")
    End If
    If reponse = vbNo Then
        supprimer_po_srv = P_OK
        Exit Function
    End If
    
    If Odbc_BeginTrans() = P_ERREUR Then
        supprimer_po_srv = P_ERREUR
        Exit Function
    End If
    
    If stype = "S" Then
        If inhiber Then
            If Odbc_Update("Service", "SRV_Num", "where SRV_Num=" & num, _
                           "SRV_Actif", False) = P_ERREUR Then
                GoTo err_enreg
            End If
        Else
            If Odbc_Delete("Service", "SRV_Num", "where SRV_Num=" & num, lnb) = P_ERREUR Then
                GoTo err_enreg
            End If
        End If
    Else
        If inhiber Then
            If Odbc_Update("Poste", "PO_Num", "where PO_Num=" & num, _
                           "PO_Actif", False) = P_ERREUR Then
                GoTo err_enreg
            End If
            If ya_resp Then
                Call Odbc_UpdateP("Poste", "po_num", "where po_numresp=" & num, lnb, _
                                  "po_numresp", -1)
            End If
        Else
            If Odbc_Delete("Poste", "PO_Num", "where PO_Num=" & num, lnb) = P_ERREUR Then
                GoTo err_enreg
            End If
        End If
' Le poste doit être supprimé de toutes les références :
'  formetape.fore_dest

        ' On propose de supprimer la fonction associée si plus de poste rattaché
        sql = "select count(*) from poste" _
            & " where po_ftnum=" & numfct _
            & " and po_actif=true"
        Call Odbc_Count(sql, lnb_actif)
        sql = "select count(*) from poste" _
            & " where po_ftnum=" & numfct _
            & " and po_actif=false"
        Call Odbc_Count(sql, lnb_inactif)
        If lnb_actif = 0 Then
            sql = "select FT_Libelle from FctTrav" _
                & " where FT_Num=" & numfct
            If Odbc_RecupVal(sql, lib) = P_ERREUR Then
                supprimer_po_srv = P_ERREUR
                Exit Function
            End If
            reponse = MsgBox("Plus aucun poste n'est rattaché à la fonction '" & lib & "'." & vbCrLf & vbCrLf _
                            & "Voulez-vous supprimer cette fonction ?", vbQuestion + vbYesNo + vbDefaultButton2, "")
            If reponse = vbYes Then
                Call supprimer_fonction(numfct, IIf(lnb_inactif > 0, True, False))
            End If
        End If
    End If
    
    ' On supprime le poste/service qui est ds les créateurs possibles
    ' dossiers/documentations
    sql = "select do_num, do_lstcrdoc from documentation where do_lstcrdoc like '%" & stype & num & ";%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        n = STR_GetNbchamp(rs("do_lstcrdoc").Value, "|")
        slst = ""
        For I = 0 To n - 1
            s = STR_GetChamp(rs("do_lstcrdoc").Value, "|", I)
            If InStr(s, stype & num & ";") = 0 Then
                slst = slst + s + "|"
            End If
        Next I
        If slst <> rs("do_lstcrdoc").Value Then
            Call Odbc_Update("documentation", "do_num", "where do_num=" & rs("do_num").Value, _
                             "do_lstcrdoc", slst)
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    sql = "select ds_num, ds_lstcrdoc from dossier where ds_lstcrdoc like '%" & stype & num & ";%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        n = STR_GetNbchamp(rs("ds_lstcrdoc").Value, "|")
        slst = ""
        For I = 0 To n - 1
            s = STR_GetChamp(rs("ds_lstcrdoc").Value, "|", I)
            If InStr(s, stype & num & ";") = 0 Then
                slst = slst + s + "|"
            End If
        Next I
        If slst <> rs("ds_lstcrdoc").Value Then
            Call Odbc_Update("dossier", "ds_num", "where ds_num=" & rs("ds_num").Value, _
                             "ds_lstcrdoc", slst)
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    If Odbc_CommitTrans() = P_ERREUR Then
        supprimer_po_srv = P_ERREUR
        Exit Function
    End If
    
    tv.Nodes.Remove (tv.SelectedItem.Index)
    tv.Refresh
    
    supprimer_po_srv = P_OK
    Exit Function
    
err_enreg:
    Call Odbc_RollbackTrans
    supprimer_po_srv = P_ERREUR
    
End Function

Private Sub supprimer_fonction(ByVal v_numfct As Long, _
                               ByVal v_inhiber As Boolean)

    Dim sobjet As String
    Dim inhiber As Boolean
    Dim reponse As Integer, cr As Integer
    Dim lnb As Long
    
    If p_appli_kalidoc > 0 Then
        ' Fonction définie ds les destinataires de documentations, dossiers ou documents
        cr = fct_dans_dest_do(v_numfct)
        If cr = P_ERREUR Then
            Exit Sub
        End If
        If cr > 0 Then
            If cr = 1 Then
                sobjet = "documentations"
            ElseIf cr = 2 Then
                sobjet = "dossiers"
            Else
                sobjet = "documents"
            End If
            Call MsgBox("Des " & sobjet & " ont cette fonction comme destinataires." & vbLf & vbCr & "Cette fonction ne peut donc pas être supprimée.", vbExclamation + vbOKOnly, "")
            Exit Sub
        End If
    End If
    
    ' Formulaires avec cette fonction dans un champ NUMFCT
    If fct_dans_donnees_form(v_numfct) = P_OUI Then
        v_inhiber = True
    End If
    
    
    If v_inhiber Then
        If Odbc_Update("FctTrav", "FT_Num", "where FT_Num=" & v_numfct, _
                       "FT_Actif", False) = P_ERREUR Then
            Exit Sub
        End If
        Call MsgBox("La fonction a été désactivée.", vbInformation + vbOKOnly, "")
    Else
        ' Maj table
        If Odbc_Delete("FctTrav", _
                       "FT_Num", _
                        "where FT_Num=" & v_numfct, _
                        lnb) = P_ERREUR Then
            Exit Sub
        End If
        Call MsgBox("La fonction a été supprimée.", vbInformation + vbOKOnly, "")
    End If

End Sub

Private Function transferer_poste() As Integer

    Dim key As String, sql As String, stype_src As String, stype_dest As String
    Dim s_sp_src As String, s_sp_dest As String, s_sp As String, s As String
    Dim key_depl As String
    Dim encore As Boolean
    Dim I As Integer, nbch As Integer, n As Integer
    Dim numposte_src As Long, numposte_dest As Long, lnb As Long, numposte As Long
    Dim nd_src As Node, nd_dest As Node, nd As Node, ndp As Node
    Dim rs As rdoResultset
    
    lblDepl.Visible = False
    
    Set nd_dest = tv.SelectedItem
    Set nd_src = tv.Nodes(g_pos_depl)
    g_pos_depl = 0
    
    If nd_src.key = nd_dest.key Then
        Call MsgBox("Vous ne pouvez pas transférer le poste vers lui-même !", vbExclamation + vbOKOnly, "")
        transferer_poste = P_OK
        Exit Function
    End If
    
    stype_dest = left$(nd_dest.key, 1)
    If stype_dest <> "P" Then
        Call MsgBox("Vous devez sélectionner un poste !", vbExclamation + vbOKOnly, "")
        transferer_poste = P_OK
        Exit Function
    End If
    If nd_dest.Parent.key <> nd_src.Parent.key Then
        Call MsgBox("Vous devez sélectionner un poste du même service !", vbExclamation + vbOKOnly, "")
        transferer_poste = P_OK
        Exit Function
    End If
    
    key_depl = nd_src.key
    s_sp_src = ""
    Set nd = nd_src
    numposte_src = Mid$(nd.key, 2)
    While left$(nd.key, 1) <> "L"
        s_sp_src = nd.key & ";" & s_sp_src
        Set nd = nd.Parent
    Wend
        
    s_sp_dest = ""
    Set nd = nd_dest
    numposte_dest = Mid$(nd.key, 2)
    While left$(nd.key, 1) <> "L"
        s_sp_dest = nd.key & ";" & s_sp_dest
        Set nd = nd.Parent
    Wend
    
    If Odbc_BeginTrans() = P_ERREUR Then
        transferer_poste = P_ERREUR
        Exit Function
    End If
    ' Mise à jour Documentation
    sql = "select DO_Num, DO_Dest from Documentation" _
        & " where DO_Dest like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        s_sp = Replace(rs("DO_Dest").Value & "", s_sp_src, s_sp_dest)
        If Odbc_Update("Documentation", _
                        "DO_Num", _
                        "where DO_Num=" & rs("DO_Num").Value, _
                        "DO_Dest", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close
    ' Mise à jour Dossier
    sql = "select DS_Num, DS_Dest from Dossier" _
        & " where DS_Dest like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        s_sp = Replace(rs("DS_Dest").Value, s_sp_src, s_sp_dest)
        If Odbc_Update("Dossier", _
                        "DS_Num", _
                        "where DS_Num=" & rs("DS_Num").Value, _
                        "DS_Dest", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close
    ' Mise à jour Document
    sql = "select D_Num, D_Dest from Document" _
        & " where D_Dest like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        s_sp = Replace(rs("D_Dest").Value & "", s_sp_src, s_sp_dest)
        If Odbc_Update("Document", _
                        "D_Num", _
                        "where D_Num=" & rs("D_Num").Value, _
                        "D_Dest", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close
    ' Mise à jour GroupeUtil
    sql = "select GU_Num, GU_Lst from GroupeUtil" _
        & " where GU_Lst like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        s_sp = Replace(rs("GU_Lst").Value, s_sp_src, s_sp_dest)
        If Odbc_Update("GroupeUtil", _
                        "GU_Num", _
                        "where GU_Num=" & rs("GU_Num").Value, _
                        "GU_Lst", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close
    ' Mise à jour Utilisateur
    sql = "select U_Num, U_PO_Princ, U_SPM from Utilisateur" _
        & " where U_Spm like '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        If rs("U_PO_Princ").Value = numposte_src Then
            numposte = numposte_dest
        Else
            numposte = rs("U_PO_Princ").Value
        End If
        nbch = STR_GetNbchamp(rs("U_SPM").Value, "|")
        s_sp = ""
        For I = 1 To nbch
            s = STR_GetChamp(rs("U_SPM").Value, "|", I - 1)
            s = Replace(s, s_sp_src, s_sp_dest)
            If InStr(s_sp, s) = 0 Then
                s_sp = s_sp + s + "|"
            End If
        Next I
        If Odbc_Update("Utilisateur", _
                        "U_Num", _
                        "where U_Num=" & rs("U_Num").Value, _
                        "U_SPM", s_sp, _
                        "U_PO_Princ", numposte) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        Call maj_fct_util(rs("U_Num").Value, s_sp)
        Call KS_PrmPersonne.gerer_chgt_poste_act(rs("U_Num").Value, s_sp)
        rs.MoveNext
    Wend
    rs.Close

    Call Odbc_CommitTrans
    
    ' Rattacher les personnes à l'autre poste
    If nd_src.Children > 0 Then
        n = nd_src.Children
        For I = 1 To n
            Set nd = nd_src.Child
            Set nd.Parent = nd_dest
            Set nd = nd.Next
        Next I
    End If
    
    Set tv.SelectedItem = tv.Nodes(key_depl)
    tv.Nodes(key_depl).EnsureVisible
    transferer_poste = P_OK
    Exit Function

err_enreg:
    Call Odbc_RollbackTrans
    transferer_poste = P_ERREUR
    
End Function

Private Sub CmbNiveau_Click()
    If g_mode_saisie Then Call afficher_liste3(Me.TxtRecherche.Text)
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
    Case CMD_OK
        Call selectionner
    Case CMD_IMPRIMER
        Call imprimer
    Case CMD_QUITTER
        Call quitter
    End Select
    
End Sub

Private Sub cmdHelp_Click()
    Dim message As String
    
    message = "Vous pouvez rechercher sur plusieurs mots en les séparant par des espaces" & Chr(13) & Chr(10) _
        & Chr(13) & Chr(10) _
        & "Vous pouvez limiter votre recherche à un seul niveau (Pôle, CR, CA, ...)" & Chr(13) & Chr(10) _
        & Chr(13) & Chr(10) _
        & "La recherche s'effectue sur le libellé, mais aussi sur le code" & Chr(13) & Chr(10) _
        & Chr(13) & Chr(10) _
        & "ex : inform 163 recherche les services qui contiennent 'inform' ET 163 dans le libellé et le code"
    MsgBox message

End Sub

Private Sub Form_Activate()

    If g_form_active Then
        Exit Sub
    End If
    
    g_form_active = True
    Call initialiser
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyO And Shift = vbAltMask) Or KeyCode = vbKeyF1 Then
        KeyCode = 0
        If g_mode_acces = MODE_SELECT Then Call selectionner
    ElseIf (KeyCode = vbKeyI And Shift = vbAltMask) Or KeyCode = vbKeyF3 Then
        KeyCode = 0
        If g_mode_acces = MODE_PARAM Then Call imprimer
    ElseIf KeyCode = vbKeyH And Shift = vbAltMask Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_d_spm.htm")
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        Call quitter
        Exit Sub
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then KeyAscii = 0
    
End Sub

Private Sub Form_Load()

    g_mode_saisie = False
    g_form_active = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Call quitter
    End If
    
End Sub

Private Sub lblDepl_Click()

    g_pos_depl = 0
    lblDepl.Visible = False
    
End Sub

Private Sub mnuAjPers_Click()

    Call ajouter_personne
    
End Sub

Private Sub mnuCreerP_Click()

    Call creer_poste
    
End Sub

Private Sub mnuCreerPi_Click()

    Call modifier_creer_piece
    
End Sub

Private Sub mnucreerS_Click()

    Call creer_service
    
End Sub

Private Sub mnuCrPers_Click()

    g_sret = "0|" + tv.SelectedItem.key
    Unload Me

End Sub

Private Sub mnuDepl_Click()

    Call activer_depl
    
End Sub

Private Sub mnuLibPoste_Click()

    Call modifier_poste
    
End Sub


Private Sub mnuModPers_Click()

    g_sret = Mid$(tv.SelectedItem.tag, 2)
    Unload Me
    
End Sub

Private Sub mnuModPi_Click()

    Call modifier_creer_piece
    
End Sub

Private Sub mnuModS_Click()

    Call modifier_service
    
End Sub

Private Sub mnuSuppP_Click()

    Call supprimer_po_srv
    
End Sub

Private Sub mnuSuppPi_Click()

    Call supprimer_piece
    
End Sub

Private Sub mnuSuppS_Click()

    Call supprimer_po_srv
    
End Sub

Private Sub mnuTrsP_Click()

    Call activer_trs_poste
    
End Sub

Private Sub mnuVoirPers_Click()

    Call ajouter_pers_tv
    
End Sub

Private Sub tv_Click()

    If g_node = tv.SelectedItem.Index And g_expand <> tv.SelectedItem.Expanded Then
        Exit Sub
    End If
    
    If g_button = vbRightButton Then
        If tv.SelectedItem.Expanded = False And left$(tv.SelectedItem.key, 1) = "S" Then
            If STR_GetChamp(tv.SelectedItem.tag, "|", 1) = False Then
                tv.Nodes.Remove (tv.SelectedItem.Child.Index)
                charger_service (Mid$(tv.SelectedItem.key, 2))
                tv.SelectedItem.Expanded = True
            End If
        End If
        Call afficher_menu(False)
    ElseIf g_button = vbLeftButton Then
        If g_pos_depl <> 0 Then
            If lblDepl.tag = "D" Then
                If left$(tv.Nodes(g_pos_depl).key, 1) = "C" Then
                    If deplacer_piece() = P_ERREUR Then
                        Call quitter
                        Exit Sub
                    End If
                Else
                    If deplacer_sp() = P_ERREUR Then
                        Call quitter
                        Exit Sub
                    End If
                End If
            Else
                If transferer_poste() = P_ERREUR Then
                    Call quitter
                    Exit Sub
                End If
            End If
        ElseIf g_plusieurs Then
            Call basculer_selection
        End If
        Call majLibDetailSRV(tv.SelectedItem.key, tv.SelectedItem.tag)
    End If
        
End Sub

Private Sub tv_Collapse(ByVal Node As ComctlLib.Node)

    If Not g_mode_saisie Then
        Exit Sub
    End If
    
    g_button = -1

End Sub

Private Sub tv_Expand(ByVal Node As ComctlLib.Node)

    If Not g_mode_saisie Then
        Exit Sub
    End If
    
    g_button = -1
    Me.LbldetailSRV.Visible = False
    If left$(Node.key, 1) = "S" Then
        If STR_GetChamp(Node.tag, "|", 1) = False Then
            tv.Nodes.Remove (Node.Child.Index)
            charger_service (Mid$(Node.key, 2))
            Call majLibDetailSRV(Node.key, Node.tag)
        End If
    End If
    
End Sub

Private Sub tv_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        Call afficher_menu(True)
    End If
    
End Sub

Private Sub tv_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeySpace Then
        KeyAscii = 0
        If g_plusieurs Then
            Call basculer_selection
        End If
    End If
    
End Sub

Private Sub tv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    g_button = Button
    g_node = tv.SelectedItem.Index
    g_expand = tv.SelectedItem.Expanded
    Me.LbldetailSRV.Visible = False
    
End Sub

Private Sub txtRecherche_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        If TxtRecherche.Text <> "" Then
            Call afficher_liste3(TxtRecherche.Text)
        Else
            If g_mode_acces <> MODE_PARAM_PERS Then
                If afficher_liste() = P_ERREUR Then
                    Call quitter
                    Exit Sub
                End If
            Else
                If afficher_liste2() = P_ERREUR Then
                    Call quitter
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

