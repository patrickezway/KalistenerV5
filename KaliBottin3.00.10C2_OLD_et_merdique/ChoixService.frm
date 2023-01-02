VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ChoixService 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
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
      Height          =   7515
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8265
      Begin ComctlLib.TreeView tv 
         Height          =   6645
         Left            =   330
         TabIndex        =   5
         Top             =   660
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   11721
         _Version        =   327682
         Indentation     =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
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
      Begin ComctlLib.ImageList img 
         Left            =   6780
         Top             =   270
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   23
         ImageHeight     =   20
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   11
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixService.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixService.frx":08D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixService.frx":11A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixService.frx":1A76
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixService.frx":2348
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixService.frx":2C1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixService.frx":346C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixService.frx":3AFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixService.frx":412C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixService.frx":47CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ChoixService.frx":4E68
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frmFct 
      BackColor       =   &H00808080&
      Height          =   825
      Left            =   -30
      TabIndex        =   0
      Top             =   7380
      Width           =   8295
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
         Picture         =   "ChoixService.frx":5506
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   230
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
         Left            =   360
         Picture         =   "ChoixService.frx":5ABF
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   230
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
         Left            =   1620
         Picture         =   "ChoixService.frx":5F18
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
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
      Begin VB.Menu mnuCreerPiece 
         Caption         =   "Créer u&ne pièce"
      End
      Begin VB.Menu mnuSepCreerPiece 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModPiece 
         Caption         =   "Modifier &la pièce"
      End
      Begin VB.Menu mnuSuppPiece 
         Caption         =   "Sup&primer la pièce"
      End
      Begin VB.Menu mnuSepMSPiece 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreerP 
         Caption         =   "C&réer un poste"
      End
      Begin VB.Menu mnuSepCrP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModP 
         Caption         =   "Mo&difier le poste"
      End
      Begin VB.Menu mnuSuppP 
         Caption         =   "Supprimer le pos&te"
      End
      Begin VB.Menu mnuPosteResp 
         Caption         =   "Poste responsa&ble"
      End
      Begin VB.Menu mnuSepSuppP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDepl 
         Caption         =   "Dépl&acer dans un autre service"
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
Attribute VB_Name = "ChoixService"
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
Private Const IMG_SITE = 7
Private Const IMG_UTIL = 8
Private Const IMG_PIECE = 9
Private Const IMG_PIECE_SEL = 10
Private Const IMG_PIECE_SEL_NOMOD = 11

Private Const MODE_PARAM = 0
Private Const MODE_SELECT = 1
Private Const MODE_PARAM_PERS = 2
Private g_mode_acces As Integer

Private g_plusieurs As Boolean
Private g_ssite As String
Private g_stype As String
Private g_prmpers As Boolean
Private g_numserv As Long
Private g_numsite As Long
Private g_sret As String

Private g_crfct_autor As Boolean

Private g_tbl_site() As Long

Private g_lignes() As CL_SLIGNE

Private g_node_crt As Long
Private g_pos_depl As Long

Private g_node As Integer
Private g_expand As Boolean
Private g_button As Integer
Private g_mode_saisie As Boolean
Private g_form_active As Boolean

Public Function AppelFrm(ByVal v_stitre As String, _
                         ByVal v_smode As String, _
                         ByVal v_bplusieurs As Boolean, _
                         ByVal v_ssite As String, _
                         ByVal v_stype As String, _
                         ByVal v_prmpers As Boolean) As String

    If v_smode = "M" Then
        g_mode_acces = MODE_PARAM
        g_stype = v_stype
        g_numserv = 0
    ElseIf v_smode = "S" Then
        g_mode_acces = MODE_SELECT
        g_stype = v_stype
        g_numserv = 0
    ElseIf v_smode = "P" Then
        g_mode_acces = MODE_PARAM_PERS
        If left$(v_stype, 1) = "S" Then
            g_numserv = Mid$(v_stype, 2)
            g_numsite = 0
        Else
            g_numserv = 0
            g_numsite = Mid$(v_stype, 2)
        End If
        g_stype = "SP"
    End If
    g_plusieurs = v_bplusieurs
    g_ssite = v_ssite
    g_prmpers = v_prmpers
    
    frm.Caption = v_stitre

    Me.Show 1
    
    AppelFrm = g_sret
    
End Function

Private Sub activer_depl()

    g_node_crt = tv.SelectedItem.Index
    g_pos_depl = g_node_crt
    lblDepl.Caption = "Cliquez sur le nouveau service de rattachement ou cliquez ici pour ANNULER l'opération"
    lblDepl.BackColor = P_ORANGE
    lblDepl.Visible = True
    
End Sub

Private Function afficher_liste() As Integer

    Dim sql As String, s As String
    Dim fmodif As Boolean
    Dim img As Integer, I As Integer, nsel As Integer, n As Integer, j As Integer
    Dim rs As rdoResultset
    Dim nd As Node, ndp As Node
    
    g_mode_saisie = False
    
    nsel = -1
    On Error Resume Next
    nsel = UBound(g_lignes)
    On Error GoTo 0
    
    tv.Nodes.Clear
    
    sql = "SELECT L_Num, L_Code FROM Laboratoire"
    If Odbc_Select(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF
        Set nd = tv.Nodes.Add(, , "L" & rs("L_Num").Value, rs("L_Code").Value, IMG_SITE, IMG_SITE)
        If rs("L_Num").Value = p_NumLabo Then
            nd.selected = True
            nd.Expanded = True
        End If
        nd.Expanded = True
        nd.Sorted = True
        rs.MoveNext
    Wend
    rs.Close
    
    ' Les services
    sql = "SELECT SRV_Num, SRV_LNum, SRV_NumPere, SRV_Nom FROM Service" _
        & " ORDER BY SRV_LNum, SRV_NumPere"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF
        If rs("SRV_NumPere").Value = 0 Then
            Set ndp = tv.Nodes("L" & rs("SRV_LNum").Value)
        Else
            If TV_NodeExiste(tv, "S" & rs("SRV_Num").Value, nd) = P_OUI Then
                GoTo lab_suivant
            End If
            If TV_NodeExiste(tv, "S" & rs("SRV_NumPere").Value, nd) = P_NON Then
                Call ajouter_service(rs("SRV_NumPere").Value)
            End If
            Set ndp = tv.Nodes("S" & rs("SRV_NumPere").Value)
        End If
        Set nd = tv.Nodes.Add(ndp, _
                               tvwChild, _
                               "S" & rs("SRV_Num").Value, _
                               rs("SRV_Nom").Value, _
                               IMG_SRV, _
                               IMG_SRV)
        nd.Sorted = True
        nd.tag = True
lab_suivant:
        rs.MoveNext
    Wend
    rs.Close
 
    ' Les postes
    sql = "SELECT PO_Num, PO_SRVNum, FT_Libelle" _
        & " FROM Poste, FctTrav" _
        & " WHERE FT_Num=PO_FTNum"
    If g_mode_acces <> MODE_PARAM Then
        If g_tbl_site(0) > 0 Then
            For I = 0 To UBound(g_tbl_site())
                If I = 0 Then
                    sql = sql & " AND ("
                Else
                    sql = sql & " OR "
                End If
                sql = sql & " PO_LNum=" & g_tbl_site(I)
            Next I
            sql = sql + ")"
        End If
    End If
    sql = sql & " ORDER BY PO_Num"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF
        Set ndp = tv.Nodes("S" & rs("PO_SRVNum").Value)
        Set nd = tv.Nodes.Add(ndp, _
                               tvwChild, _
                               "P" & rs("PO_Num").Value, _
                               rs("FT_Libelle").Value, _
                               IMG_POSTE, _
                               IMG_POSTE)
        nd.tag = True
        rs.MoveNext
    Wend
    rs.Close

    ' Les pièces
    sql = "SELECT PC_Num, PC_Nom, PC_SrvNum FROM Piece" _
        & " ORDER BY PC_Num, PC_SrvNum"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF
        If rs("PC_SrvNum").Value = 0 Then
            Set ndp = tv.Nodes("L1")
        Else
            Set ndp = tv.Nodes("S" & rs("PC_SrvNum").Value)
        End If
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
            Set nd = tv.Nodes(s)
            nd.image = img
            nd.tag = fmodif
            nd.SelectedImage = img
            Set ndp = nd.Parent
            While left$(ndp.key, 1) <> "L"
                ndp.Expanded = True
                Set ndp = ndp.Parent
            Wend
        Next I
    End If
    
    tv.SetFocus
    g_mode_saisie = True
    
    afficher_liste = P_OK
    Exit Function

lab_erreur:
    afficher_liste = P_ERREUR

End Function

Private Function afficher_liste2() As Integer

    Dim sql As String, s As String, codsite As String, nomsrv As String, key As String
    Dim trouve As Boolean, faff As Boolean
    Dim img As Integer, I As Integer, nsel As Integer, n As Integer, j As Integer
    Dim rs As rdoResultset
    Dim nd As Node, ndp As Node
    
    g_mode_saisie = False
    
    tv.Nodes.Clear
    
    If g_numserv = 0 Then
        sql = "SELECT L_Code FROM Laboratoire" _
            & " WHERE L_Num=" & g_numsite
        If Odbc_RecupVal(sql, codsite) = P_ERREUR Then
            GoTo lab_erreur
        End If
        Set nd = tv.Nodes.Add(, , "L" & g_numsite, codsite, IMG_SITE, IMG_SITE)
        nd.Expanded = True
        nd.Sorted = True
        ' Les services
        sql = "SELECT SRV_Num, SRV_LNum, SRV_NumPere, SRV_Nom FROM Service" _
            & " ORDER BY SRV_LNum, SRV_NumPere"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            GoTo lab_erreur
        End If
        While Not rs.EOF
            If rs("SRV_NumPere").Value = 0 Then
                Set ndp = tv.Nodes("L" & rs("SRV_LNum").Value)
            Else
                If TV_NodeExiste(tv, "S" & rs("SRV_Num").Value, nd) = P_OUI Then
                    GoTo lab_suivant
                End If
                If TV_NodeExiste(tv, "S" & rs("SRV_NumPere").Value, nd) = P_NON Then
                    Call ajouter_service(rs("SRV_NumPere").Value)
                End If
                Set ndp = tv.Nodes("S" & rs("SRV_NumPere").Value)
            End If
            Set nd = tv.Nodes.Add(ndp, _
                                   tvwChild, _
                                   "S" & rs("SRV_Num").Value, _
                                   rs("SRV_Nom").Value, _
                                   IMG_SRV, _
                                   IMG_SRV)
            nd.Sorted = True
            nd.tag = True
lab_suivant:
            rs.MoveNext
        Wend
        rs.Close
    Else
        ' Les services
        sql = "SELECT SRV_Nom FROM Service" _
            & " WHERE SRV_Num=" & g_numserv
        If Odbc_RecupVal(sql, nomsrv) = P_ERREUR Then
            GoTo lab_erreur
        End If
        Set nd = tv.Nodes.Add(, _
                               tvwChild, _
                               "S" & g_numserv, _
                               nomsrv, _
                               IMG_SRV, _
                               IMG_SRV)
        nd.Expanded = True
        nd.Sorted = True
        nd.tag = True
        ajouter_fils (g_numserv)
    End If
    
    ' Les postes
    sql = "SELECT PO_Num, PO_SRVNum, FT_Libelle" _
        & " FROM Poste, FctTrav" _
        & " WHERE FT_Num=PO_FTNum"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF
        If TV_NodeExiste(tv, "S" & rs("PO_SRVNum").Value, ndp) = P_OUI Then
            Set nd = tv.Nodes.Add(ndp, _
                                   tvwChild, _
                                   "P" & rs("PO_Num").Value, _
                                   rs("FT_Libelle").Value, _
                                   IMG_POSTE, _
                                   IMG_POSTE)
            nd.tag = True
            nd.Expanded = True
            nd.Sorted = True
        End If
        rs.MoveNext
    Wend
    rs.Close

    ' Les personnes associées
    For I = 1 To tv.Nodes.Count
        Set ndp = tv.Nodes(I)
        If left$(ndp.key, 1) = "P" Then
            sql = "SELECT U_Num, U_Nom, U_Prenom FROM Utilisateur" _
                & " WHERE U_kb_actif=True AND U_SPM LIKE '%" & ndp.key & ";%'"
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                GoTo lab_erreur
            End If
            While Not rs.EOF
                Set nd = tv.Nodes.Add(ndp, _
                                       tvwChild, _
                                       "", _
                                       rs("U_Nom").Value & " " & rs("U_Prenom").Value, _
                                       IMG_UTIL, _
                                       IMG_UTIL)
                nd.tag = "U" & rs("U_Num").Value
                rs.MoveNext
            Wend
            rs.Close
        End If
    Next I

    ' Les pièces
    sql = "SELECT PC_Num, PC_Nom, PC_SrvNum FROM Piece" _
        & " ORDER BY PC_Num, PC_SrvNum"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF
        If rs("PC_SrvNum").Value = 0 Then
            key = "L1"
        Else
            key = "S" & rs("PC_SrvNum").Value
        End If
        If TV_NodeExiste(tv, key, ndp) = P_OUI Then
            Set nd = tv.Nodes.Add(ndp, _
                                  tvwChild, _
                                  "C" & rs("PC_Num").Value, _
                                  rs("PC_Nom").Value, _
                                  IMG_PIECE, _
                                  IMG_PIECE)
            nd.tag = True
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    tv.SetFocus
    g_mode_saisie = True
    
    afficher_liste2 = P_OK
    Exit Function

lab_erreur:
    afficher_liste2 = P_ERREUR

End Function

Private Sub afficher_menu(ByVal v_bclavier As Boolean)

    Dim key As String, tag As String, lib As String, sql As String
    Dim numposte As Long, numresp As Long
    
    key = tv.SelectedItem.key
    Select Case left$(key, 1)
    Case "L"
        mnuCreerS.Visible = True
        mnuSepCrS.Visible = True
        mnuModS.Visible = False
        mnuSuppS.Visible = False
        mnuSepMSS.Visible = False
        '*******************************
        mnuCreerPiece.Visible = True
        mnuSepCreerPiece.Visible = True
        mnuModPiece.Visible = False
        mnuSuppPiece.Visible = False
        mnuSepMSPiece.Visible = False
        '*******************************
        mnuCreerP.Visible = False
        mnuSepCrP.Visible = False
        mnuPosteResp.Visible = False
        mnuModP.Visible = False
        mnuSuppP.Visible = False
        mnuSepSuppP.Visible = False
        mnuDepl.Visible = False
        mnuSepDepl.Visible = False
        mnuVoirPers.Visible = False
        mnuSepVoirPers.Visible = False
        mnuCrPers.Visible = False
        mnuSepCrPers.Visible = False
        mnuModPers.Visible = False
        mnuSepModPers.Visible = False
    Case "S"
        mnuCreerS.Visible = True
        mnuSepCrS.Visible = True
        mnuModS.Visible = True
        mnuSuppS.Visible = True
        mnuSepMSS.Visible = True
        '*******************************
        mnuCreerPiece.Visible = True
        mnuSepCreerPiece.Visible = True
        mnuModPiece.Visible = False
        mnuSuppPiece.Visible = False
        mnuSepMSPiece.Visible = False
        '*******************************
        mnuCreerP.Visible = True
        mnuSepCrP.Visible = True
        mnuPosteResp.Visible = False
        mnuModP.Visible = False
        mnuSuppP.Visible = False
        mnuSepSuppP.Visible = False
        If tv.SelectedItem.Index = tv.SelectedItem.Root.Index Then
            mnuDepl.Visible = False
            mnuSepDepl.Visible = False
        Else
            mnuDepl.Visible = True
            mnuSepDepl.Visible = True
        End If
        mnuVoirPers.Visible = False
        mnuSepVoirPers.Visible = False
        mnuCrPers.Visible = False
        mnuSepCrPers.Visible = False
        mnuModPers.Visible = False
        mnuSepModPers.Visible = False
    Case "C" ' pièCe
        mnuCreerS.Visible = False
        mnuSepCrS.Visible = False
        mnuModS.Visible = False
        mnuSuppS.Visible = False
        mnuSepMSS.Visible = False
        '*******************************
        mnuCreerPiece.Visible = False
        mnuSepCreerPiece.Visible = False
        mnuModPiece.Visible = True
        mnuSuppPiece.Visible = True
        mnuSepMSPiece.Visible = True
        '*******************************
        mnuCreerP.Visible = False
        mnuSepCrP.Visible = False
        mnuPosteResp.Visible = False
        mnuModP.Visible = False
        mnuSuppP.Visible = False
        mnuSepSuppP.Visible = False
        mnuDepl.Visible = True
        mnuSepDepl.Visible = True
        mnuVoirPers.Visible = False
        mnuSepVoirPers.Visible = False
        mnuCrPers.Visible = False
        mnuSepCrPers.Visible = False
        mnuModPers.Visible = False
        mnuSepModPers.Visible = False
    Case "P"
        mnuCreerS.Visible = False
        mnuSepCrS.Visible = False
        mnuModS.Visible = False
        mnuSuppS.Visible = False
        mnuSepMSS.Visible = False
        '*******************************
        mnuCreerPiece.Visible = False
        mnuSepCreerPiece.Visible = False
        mnuModPiece.Visible = False
        mnuSuppPiece.Visible = False
        mnuSepMSPiece.Visible = False
        '*******************************
        mnuCreerP.Visible = False
        mnuSepCrP.Visible = False
        mnuPosteResp.Visible = False
        mnuModP.Visible = True
        mnuSuppP.Visible = True
        mnuSepSuppP.Visible = True
        mnuDepl.Visible = True
        mnuSepDepl.Visible = True
        If g_mode_acces = MODE_PARAM_PERS Then
            mnuVoirPers.Visible = False
            mnuSepVoirPers.Visible = False
            mnuCrPers.Visible = True
            mnuSepCrPers.Visible = True
        Else
            If tv.SelectedItem.Children > 0 Then
                mnuVoirPers.Visible = False
                mnuSepVoirPers.Visible = False
            Else
                mnuVoirPers.Visible = True
                mnuSepVoirPers.Visible = True
            End If
            mnuCrPers.Visible = g_prmpers
            mnuSepCrPers.Visible = g_prmpers
        End If
        mnuModPers.Visible = False
        mnuSepModPers.Visible = False
    Case Else
        mnuCreerS.Visible = False
        mnuSepCrS.Visible = False
        mnuModS.Visible = False
        mnuSuppS.Visible = False
        mnuSepMSS.Visible = False
        mnuCreerP.Visible = False
        '*******************************
        mnuCreerPiece.Visible = False
        mnuSepCreerPiece.Visible = False
        mnuModPiece.Visible = False
        mnuSuppPiece.Visible = False
        mnuSepMSPiece.Visible = False
        '*******************************
        mnuCreerP.Visible = False
        mnuSepCrP.Visible = False
        mnuPosteResp.Visible = False
        mnuModP.Visible = False
        mnuSuppP.Visible = False
        mnuSepSuppP.Visible = False
        '*******************************
        mnuDepl.Visible = False
        mnuSepDepl.Visible = False
        mnuVoirPers.Visible = False
        mnuSepVoirPers.Visible = False
        mnuCrPers.Visible = False
        mnuSepCrPers.Visible = False
        If g_mode_acces = MODE_PARAM_PERS Then
            mnuModPers.Visible = True
            mnuSepModPers.Visible = True
        Else
            mnuModPers.Visible = g_prmpers
            mnuSepModPers.Visible = g_prmpers
        End If
    End Select
    
    If v_bclavier Then
        Call PopupMenu(mnuFct, , tv.left, tv.Top)
    Else
        Call PopupMenu(mnuFct)
    End If
    
End Sub

Private Function ajouter_fils(ByVal v_numsrv As Long) As Integer

    Dim sql As String
    Dim rs As rdoResultset
    Dim nd As Node, ndp As Node
    
    sql = "select SRV_Num, SRV_Nom from Service" _
        & " where SRV_NumPere=" & v_numsrv
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
        nd.Sorted = True
        nd.tag = True
        Call ajouter_fils(rs("SRV_Num").Value)
        rs.MoveNext
    Wend
    rs.Close
    
    ajouter_fils = P_OK
    
End Function

Private Sub ajouter_pers_poste_tv()

    Dim sql As String
    Dim nd As Node
    Dim rs As rdoResultset
    
    sql = "select U_Num, U_Nom, U_Prenom from Utilisateur" _
        & " where U_kb_actif=True AND U_SPM like '%" & tv.SelectedItem.key & ";%' and U_Actif=true"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    If Not rs.EOF Then
        tv.SelectedItem.Expanded = True
        While Not rs.EOF
            Set nd = tv.Nodes.Add(tv.SelectedItem, _
                                   tvwChild, _
                                   "", _
                                   rs("U_Nom").Value & " " & rs("U_Prenom").Value, _
                                   IMG_UTIL, _
                                   IMG_UTIL)
            nd.tag = "U" & rs("U_Num").Value
            rs.MoveNext
        Wend
    End If
    rs.Close

End Sub

Private Function ajouter_service(ByVal v_numsrv As Long) As Integer

    Dim sql As String
    Dim rs As rdoResultset
    Dim nd As Node, ndp As Node
    
    sql = "select SRV_Num, SRV_LNum, SRV_NumPere, SRV_Nom from Service" _
        & " where SRV_Num=" & v_numsrv
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        ajouter_service = P_ERREUR
        Exit Function
    End If
    If TV_NodeExiste(tv, "S" & rs("SRV_Num").Value, nd) = P_OUI Then
        ajouter_service = P_OK
        Exit Function
    End If
    If TV_NodeExiste(tv, "S" & rs("SRV_NumPere").Value, nd) = P_NON Then
        Call ajouter_service(rs("SRV_NumPere").Value)
    End If
    Set ndp = tv.Nodes("S" & rs("SRV_NumPere").Value)
    Set nd = tv.Nodes.Add(ndp, _
                           tvwChild, _
                           "S" & rs("SRV_Num").Value, _
                           rs("SRV_Nom").Value, _
                           IMG_SRV, _
                           IMG_SRV)
    nd.Sorted = True
    nd.tag = True
    
    ajouter_service = P_OK
    
End Function

Private Sub basculer_selection()

    Dim img As Long
    
    If tv.SelectedItem.tag = False Then
        Exit Sub
    End If
    
    Select Case tv.SelectedItem.SelectedImage
    Case IMG_SRV
        If InStr(g_stype, left$(tv.SelectedItem.key, 1)) = 0 Then
' ************** ignorer le message et développer le noeud si possible
            tv.SelectedItem.Expanded = True
            'Call MsgBox("Vous ne pouvez pas sélectionner un service.", vbInformation + vbOKOnly, "")
            Exit Sub
        End If
        img = IMG_SRV_SEL
    Case IMG_POSTE
        If InStr(g_stype, left$(tv.SelectedItem.key, 1)) = 0 Then
            Call MsgBox("Vous ne pouvez pas sélectionner un poste.", vbInformation + vbOKOnly, "")
            Exit Sub
        End If
        img = IMG_POSTE_SEL
    Case IMG_PIECE
        If InStr(g_stype, left$(tv.SelectedItem.key, 1)) = 0 Then Exit Sub
        img = IMG_PIECE_SEL
    Case IMG_SRV_SEL
        img = IMG_SRV
    Case IMG_POSTE_SEL
        img = IMG_POSTE
    End Select
    
    tv.SelectedItem.SelectedImage = img
    tv.SelectedItem.image = img
    
End Sub

Private Sub creer_fonction(ByRef r_num As Long, _
                           ByRef r_lib As String)

    Dim sret As String
    Dim frm As Form

    Set frm = PrmFonction
    sret = PrmFonction.AppelFrm("F", 0)
    Set frm = Nothing
    If sret = "" Then
        r_num = 0
        Exit Sub
    End If
    r_num = STR_GetChamp(sret, "|", 0)
    r_lib = STR_GetChamp(sret, "|", 1)
    
End Sub

Private Sub creer_piece()

    Dim nom_piece As String, sret As String
    Dim num_srv As Long, num_piece As Long
    Dim frm As Form
    Dim nd As Node, nd_piece As Node

    Set nd = tv.SelectedItem
    num_srv = Mid$(nd.key, 2)
    Set frm = PrmPiece
    sret = PrmPiece.AppelFrm(0, num_srv, True)
    Set frm = Nothing
    ' Si on n'a rien ajouté
    If sret = "" Then
        Exit Sub
    End If
    num_piece = STR_GetChamp(sret, "|", 0)
    nom_piece = STR_GetChamp(sret, "|", 1)
    ' Ajout de la nouvelle pièce
    Set nd_piece = tv.Nodes.Add(nd, _
                                tvwChild, _
                                "C" & num_piece, _
                                nom_piece, _
                                IMG_PIECE, _
                                IMG_PIECE)
    tv.Nodes(nd_piece.Index).selected = True

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
            sql = "SELECT PO_FTNum FROM Poste" _
                & " WHERE PO_Num=" & num
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

    sql = "SELECT * FROM FctTrav" _
        & " ORDER BY FT_Libelle"
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

    If n = 0 Then
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
    Call TV_FirstParent(nd, ndp)
    numlabo = Mid$(ndp.key, 2)
    For I = 0 To n - 1
        If CL_liste.lignes(I).selected Then
            Call Odbc_AddNew("Poste", _
                             "PO_Num", _
                             "po_seq", _
                             True, _
                             num, _
                             "PO_SRVNum", numsrv, _
                             "PO_FTNum", CL_liste.lignes(I).num, _
                             "PO_LNum", numlabo, _
                             "PO_NumResp", -1, _
                             "PO_Libelle", CL_liste.lignes(I).texte, _
                             "PO_Actif", True)
            Set nde = tv.Nodes.Add(nd, _
                                   tvwChild, _
                                   "P" & num, _
                                   CL_liste.lignes(I).texte, _
                                   IMG_POSTE, _
                                   IMG_POSTE)
        End If
    Next I

lab_fin:
    Unload ChoixListe
    creer_poste = P_OK

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
    Call PR_InitFormat(True, _
                       s, _
                       False, _
                       "g", _
                       stexte())
    Set nd = ndf
    Set ndf = ndf.Child
    encore = True
    decal = 0
    While encore
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
    
    g_crfct_autor = True
    
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
        If g_ssite = "" Then
            ReDim g_tbl_site(0) As Long
            g_tbl_site(0) = 0
        Else
            nbchp = STR_GetNbchamp(g_ssite, ";")
            ReDim g_tbl_site(nbchp - 1) As Long
            For I = 0 To nbchp - 1
                g_tbl_site(I) = STR_GetChamp(g_ssite, ";", I)
            Next I
        End If
        cmd(CMD_IMPRIMER).Visible = False
    Else
        p_NumLabo = p_NumLaboDefaut
        p_CodeLabo = p_CodeLaboDefaut
        cmd(CMD_OK).Visible = False
        cmd(CMD_IMPRIMER).left = cmd(CMD_OK).left
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
    
    Set tv.SelectedItem = tv.Nodes(1)
    tv.SetFocus
    SendKeys "{PGDN}"
    SendKeys "{HOME}"
    DoEvents
    
End Sub

Private Sub modifier_piece()

    Dim nom_piece As String, sret As String
    Dim num_srv As Long, num_piece As Long
    Dim frm As Form
    Dim nd As Node

    Set nd = tv.SelectedItem
    num_srv = Mid$(nd.Parent.key, 2)
    num_piece = Mid$(tv.SelectedItem.key, 2)
    Set frm = PrmPiece
    sret = PrmPiece.AppelFrm(num_piece, num_srv, False)
    Set frm = Nothing
    If sret = "" Then ' on a rien modifié
        Exit Sub
    End If
    ' Il y eu une modification/suppression
    num_piece = STR_GetChamp(sret, "|", 0)
    nom_piece = STR_GetChamp(sret, "|", 1)
    If num_piece = 0 Then ' suppression
        tv.Nodes.Remove (tv.SelectedItem.Index)
    Else ' Mise-a-jour de la pièce
        ' Maj nom de la pièce
        nd.Text = nom_piece
    End If

End Sub

Private Sub modifier_poste()

    Dim nom_poste As String, sret As String
    Dim num_srv As Long, num_poste As Long, num_fct As Long
    Dim frm As Form
    Dim nd As Node

    Set nd = tv.SelectedItem
    num_srv = Mid$(nd.Parent.key, 2)
    num_poste = Mid$(tv.SelectedItem.key, 2)
    Set frm = PrmFonction
    sret = PrmFonction.AppelFrm("P", num_poste)
    Set frm = Nothing
    If sret = "" Then ' on a rien modifié
        Exit Sub
    End If
    ' Il y eu une modification/suppression
    num_poste = STR_GetChamp(sret, "|", 0)
    nom_poste = STR_GetChamp(sret, "|", 1)
    If num_poste = 0 Then ' suppression
        tv.Nodes.Remove (tv.SelectedItem.Index)
    Else ' Mise-a-jour de la pièce
        ' Maj nom de la pièce
        nd.Text = nom_poste
    End If

End Sub

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
        g_sret = "N" & n
    Else
        Set nd = tv.SelectedItem
        If InStr(g_stype, left$(nd.key, 1)) = 0 Or nd.key = "" Then
            If nd.key = "" Then
                Call MsgBox("Vous ne pouvez pas sélectionner une personne.", vbInformation + vbOKOnly, "")
            ElseIf left$(nd.key, 1) = "S" Then
                Call MsgBox("Vous ne pouvez pas sélectionner un service.", vbInformation + vbOKOnly, "")
            ElseIf left$(nd.key, 1) = "S" Then
                Call MsgBox("Vous ne pouvez pas sélectionner une pièce.", vbInformation + vbOKOnly, "")
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

Private Function srvpo_dans_util(ByVal v_stype As String, _
                                 ByVal v_num As Long) As Boolean

    Dim sql As String
    Dim lnb As Long
    
    sql = "SELECT COUNT(*) FROM Utilisateur" _
        & " WHERE U_kb_actif=True AND U_SPM LIKE '%" & v_stype & v_num & ";%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        srvpo_dans_util = True
        Exit Function
    End If
    If lnb > 0 Then
        srvpo_dans_util = True
        Exit Function
    End If
    
    srvpo_dans_util = False

End Function

Private Function supprimer() As Integer

    Dim lib As String, sql As String, sobj As String, stype As String
    Dim reponse As Integer
    Dim num As Long, lnb As Long, numsrv As Long
    Dim nd As Node, ndr As Node
    
    num = CLng(Mid$(tv.SelectedItem.key, 2))
    stype = left$(tv.SelectedItem.key, 1)
    
    If stype = "S" Then
        sobj = "ce service"
    Else
        sobj = "ce poste"
    End If
    
    If srvpo_dans_util(stype, num) Then
        Call MsgBox("Une (ou plusieurs) personne est associée à " & sobj & "." & vbCrLf & "Il ne peut donc pas être supprimé.", vbExclamation + vbOKOnly, "")
        supprimer = P_OK
        Exit Function
    End If
    
    Select Case stype
    Case "S"
        If P_RecupSrvNom(num, lib) = P_ERREUR Then
            supprimer = P_ERREUR
            Exit Function
        End If
        sobj = "du service " & lib
    Case "P"
        sql = "select PO_SRVNum, FT_Libelle from Poste, FctTrav" _
            & " where PO_Num=" & num _
            & " and FT_Num=PO_FTNum"
        If Odbc_RecupVal(sql, numsrv, lib) = P_ERREUR Then
            supprimer = P_ERREUR
            Exit Function
        End If
        sobj = "du poste " & lib
        If P_RecupSrvNom(numsrv, lib) = P_ERREUR Then
            supprimer = P_ERREUR
            Exit Function
        End If
        sobj = sobj & " dans le service " & lib
    End Select
    
    reponse = MsgBox("Confirmez-vous la suppression " & sobj & " ?", vbQuestion + vbYesNo + vbDefaultButton2, "")
    If reponse = vbNo Then
        supprimer = P_OK
        Exit Function
    End If
    
    If Odbc_BeginTrans() = P_ERREUR Then
        supprimer = P_ERREUR
        Exit Function
    End If
    
    Set ndr = tv.SelectedItem
    Set nd = ndr
    Do
        num = Mid$(nd.key, 2)
        If left$(nd.key, 1) = "S" Then
            If Odbc_Delete("Service", "SRV_Num", "WHERE SRV_Num=" & num, lnb) = P_ERREUR Then
                GoTo err_enreg
            End If
        Else
            If Odbc_Delete("Poste", "PO_Num", "WHERE PO_Num=" & num, lnb) = P_ERREUR Then
                GoTo err_enreg
            End If
        End If
    Loop Until Not TV_ChildNextParent(nd, ndr)
    
    If Odbc_CommitTrans() = P_ERREUR Then
        supprimer = P_ERREUR
        Exit Function
    End If
    
    tv.Nodes.Remove (ndr.Index)
    tv.Refresh
    
    supprimer = P_OK
    Exit Function
    
err_enreg:
    Call Odbc_RollbackTrans
    supprimer = P_ERREUR
    
End Function

Private Sub supprimer_piece()

    Dim reponse As String, num_piece As String

    reponse = MsgBox("Etes-vous sûr de vouloir suppreimer cette pièce ?", vbYesNo + vbDefaultButton2, "Attention !")
    If reponse = vbYes Then ' suppression effective
        num_piece = Mid$(tv.SelectedItem.key, 2)
        Call Odbc_Cnx.Execute("DELETE FROM Piece WHERE PC_Num = " & num_piece)
        tv.Nodes.Remove (tv.SelectedItem.Index)
        tv.Nodes(tv.SelectedItem.Parent.Index).selected = True ' selectionner le service parent
    End If

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
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_e_spm.htm")
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

Private Sub mnuCrPers_Click()

    g_sret = "0|" + tv.SelectedItem.key
    Unload Me
    
End Sub

Private Sub mnuVoirPers_Click()

    Call ajouter_pers_poste_tv

End Sub

Private Sub tv_Click()

    If g_node = tv.SelectedItem.Index And g_expand <> tv.SelectedItem.Expanded Then
        Exit Sub
    End If
    
    If g_button = vbRightButton Then
'        Call afficher_menu(False)
    ElseIf g_button = vbLeftButton Then
'        If g_pos_depl <> 0 Then
'            If deplacer_sp() = P_ERREUR Then
'                Call quitter
'                Exit Sub
'            End If
'        ElseIf g_plusieurs Then
'            Call basculer_selection
'        End If
    End If
        
End Sub

Private Sub tv_DblClick()

    If Mid$(tv.SelectedItem.key, 1, 1) = "C" Then
'        Call modifier_piece
        Call selectionner
    ElseIf Mid$(tv.SelectedItem.key, 1, 1) = "P" Then
'        Call modifier_poste
        Call selectionner
    End If

End Sub


Private Sub tv_Expand(ByVal Node As ComctlLib.Node)

    g_button = -1
'    g_expand = True
    
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
    
End Sub

