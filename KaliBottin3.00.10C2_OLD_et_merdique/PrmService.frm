VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form PrmService 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   8280
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
      Begin VB.TextBox TxtRecherche 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   480
         Width           =   2655
      End
      Begin ComctlLib.TreeView tv 
         Height          =   6405
         Left            =   330
         TabIndex        =   5
         Top             =   900
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   11298
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
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rechercher : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   480
         Width           =   1575
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
               Picture         =   "PrmService.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":08D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":11A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":1A76
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":2348
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":2C1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":346C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":3AFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":412C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":47CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "PrmService.frx":4E68
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
         Left            =   6930
         Picture         =   "PrmService.frx":5506
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Index           =   0
         Left            =   360
         Picture         =   "PrmService.frx":5ABF
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
         Index           =   2
         Left            =   1620
         Picture         =   "PrmService.frx":5F18
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
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
      Begin VB.Menu MnuRempP 
         Caption         =   ""
      End
      Begin VB.Menu mnuModP 
         Caption         =   "Mo&difier le poste"
      End
      Begin VB.Menu mnuActiverP 
         Caption         =   "Activer le poste"
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
Private Const IMG_POSTE_SEL_INACTIF = 6
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

    Dim sql As String, s As String, lib As String
    Dim fmodif As Boolean
    Dim img As Integer, I As Integer, nsel As Integer, n As Integer, j As Integer
    Dim num As Long, lnb As Long
    Dim rs As rdoResultset
    Dim nd As Node, ndP As Node
    
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
    Set ndP = tv.Nodes.Add(, , "L" & num, lib)
    ndP.Expanded = True
    
    ' Les pièces
    sql = "SELECT PC_Num, PC_Nom FROM Piece" _
        & " where PC_SRVNum=0" _
        & " ORDER BY PC_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF
        Set nd = tv.Nodes.Add(ndP, _
                              tvwChild, _
                              "C" & rs("PC_Num").Value, _
                              rs("PC_Nom").Value, _
                              IMG_PIECE, _
                              IMG_PIECE)
        nd.tag = True
        rs.MoveNext
    Wend
    rs.Close
    
    ' Les services
    sql = "select SRV_Num, SRV_NumPere, SRV_Nom, SRV_NivsNum from Service " _
        & " where SRV_Actif=true" _
        & " and SRV_Numpere=0" _
        & " order by SRV_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        afficher_liste = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        'lib = rs("SRV_Nom").Value
        'libNiveau = recup_lib_niveau(rs("SRV_NivsNum").Value)
        'If libNiveau <> "" Then
        '    lib = lib & " (" & libNiveau & ")"
        'End If
        Set nd = tv.Nodes.Add(ndP, _
                               tvwChild, _
                               "S" & rs("SRV_Num").Value, _
                               rs("SRV_Nom").Value, _
                               IMG_SRV, _
                               IMG_SRV)
        sql = "select count(*) from Service " _
            & " where SRV_Actif=true" _
            & " and SRV_Numpere=" & rs("SRV_Num").Value
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            afficher_liste = P_ERREUR
            Exit Function
        End If
        If lnb = 0 Then
            sql = "select count(*) from Poste " _
                & " where PO_Actif=true" _
                & " and PO_SRVNum=" & rs("SRV_Num").Value
            If Odbc_Count(sql, lnb) = P_ERREUR Then
                afficher_liste = P_ERREUR
                Exit Function
            End If
            If lnb = 0 Then
                sql = "select count(*) from Piece " _
                    & " where PC_SRVNum=" & rs("SRV_Num").Value
                If Odbc_Count(sql, lnb) = P_ERREUR Then
                    afficher_liste = P_ERREUR
                    Exit Function
                End If
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
            Set nd = tv.Nodes(s)
            nd.image = img
            nd.tag = fmodif
            nd.SelectedImage = img
            Set ndP = nd.Parent
            While left$(ndP.key, 1) <> "L"
                ndP.Expanded = True
                Set ndP = ndP.Parent
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
    Dim num As Long, lnb As Long
    Dim rs As rdoResultset
    Dim nd As Node, ndP As Node
    
    g_mode_saisie = False
    
    tv.Nodes.Clear
    
    If g_numserv = 0 Then
        sql = "select L_Num, L_Nom from Laboratoire"
        If Odbc_RecupVal(sql, num, codsite) = P_ERREUR Then
            afficher_liste2 = P_ERREUR
            Exit Function
        End If
        Set ndP = tv.Nodes.Add(, , "L" & num, codsite)
        ndP.Expanded = True
        ndP.Sorted = True
    Else
        sql = "select SRV_Nom from Service" _
            & " where SRV_Num=" & g_numserv
        If Odbc_RecupVal(sql, codsite) = P_ERREUR Then
            afficher_liste2 = P_ERREUR
            Exit Function
        End If
        Set ndP = tv.Nodes.Add(, , , codsite)
        ndP.Expanded = True
    End If
    
    ' Les services
    sql = "select SRV_Num, SRV_LNum, SRV_NumPere, SRV_Nom from Service" _
        & " where SRV_Actif=true" _
        & " and SRV_Numpere=" & g_numserv _
        & " order by SRV_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        afficher_liste2 = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        Set nd = tv.Nodes.Add(ndP, _
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
            If lnb = 0 Then
                sql = "select count(*) from Piece " _
                    & " where PC_SRVNum=" & rs("SRV_Num").Value
                If Odbc_Count(sql, lnb) = P_ERREUR Then
                    afficher_liste2 = P_ERREUR
                    Exit Function
                End If
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
    Dim FT_NivRemplace As Integer
    Dim srv_num As Long
    Dim strRemplace As String
    
    key = tv.SelectedItem.key
    Select Case left$(key, 1)
    Case "L"
        mnuCreerS.Visible = False
        mnuSepCrS.Visible = False
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
        mnuCreerS.Visible = False
        mnuSepCrS.Visible = False
        mnuModS.Visible = True
        mnuSuppS.Visible = False
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
' ON INTERDIT LE DEPLACEMENT D'UN SERVICE DANS TOUS LES CAS
        If tv.SelectedItem.Index = tv.SelectedItem.Root.Index Then
            mnuDepl.Visible = False
            mnuSepDepl.Visible = False
        Else
            mnuDepl.Visible = True
            mnuSepDepl.Visible = True
        End If
            mnuDepl.Visible = False
            mnuSepDepl.Visible = False
' ***********************************
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
        If tv.Nodes(key).image = IMG_POSTE_SEL_INACTIF Then
            mnuActiverP.Visible = True
        End If
        mnuSuppP.Visible = True
        mnuSepSuppP.Visible = True
' ON INTERDIT LE DEPLACEMENT D'UN POSTE DANS TOUS LES CAS
        mnuDepl.Visible = True
        mnuSepDepl.Visible = True
        mnuDepl.Visible = False
        mnuSepDepl.Visible = False
' ********************
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
        
        ' Niveau de remplacement
        sql = "select FT_NivRemplace, PO_SrvNum from Poste, FctTrav" _
            & " where PO_Num=" & Mid$(key, 2) _
            & " and FT_Num=PO_FTNum"
        If Odbc_RecupVal(sql, FT_NivRemplace, srv_num) <> P_ERREUR Then
            If FT_NivRemplace = 0 Then
                MnuRempP.Visible = False
            Else
                MnuRempP.Visible = True
                MnuRempP.Caption = ""
                If FctNivRemplace(FT_NivRemplace, srv_num, strRemplace) < 0 Then
                    strRemplace = "!!! Attention : " & strRemplace
                End If
                MnuRempP.Caption = strRemplace
            End If
        End If
    
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
    Dim nd As Node, ndP As Node
    
    sql = "select SRV_Num, SRV_Nom from Service" _
        & " where SRV_NumPere=" & v_numsrv
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        ajouter_fils = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        Set ndP = tv.Nodes("S" & v_numsrv)
        Set nd = tv.Nodes.Add(ndP, _
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
    Dim nd As Node, ndP As Node
    
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
    Set ndP = tv.Nodes("S" & rs("SRV_NumPere").Value)
    Set nd = tv.Nodes.Add(ndP, _
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

Private Function charger_arbor(ByVal v_ssrv As String) As Integer

    Dim sql As String, s_srv As String, s As String
    Dim I As Integer, n As Integer
    Dim numsrv As Long, numposte As Long
    Dim nd As Node, ndP As Node
    
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
    Dim nd As Node, ndP As Node, ndu As Node
    Dim strRemplace As String
    
    Set ndP = tv.Nodes("S" & v_numsrv)
    
    ' Les pièces
    sql = "SELECT PC_Num, PC_Nom FROM Piece" _
        & " where PC_SrvNum=" & v_numsrv _
        & " ORDER BY PC_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        charger_service = P_ERREUR
        Exit Function
    End If
    While Not rs.EOF
        If TV_NodeExiste(tv, "S" & v_numsrv, ndP) = P_OUI Then
            Set nd = tv.Nodes.Add(ndP, _
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
    
    ' Les postes actifs
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
    
    While Not rs.EOF
        sfct = rs("FT_Libelle").Value
        If rs("FT_Libelle").Value <> rs("PO_Libelle").Value Then
            sfct = sfct & " *"
        End If
        ' Vérification du niveau de remplacement
        If FctNivRemplace(rs("FT_NivRemplace"), v_numsrv, strRemplace) < 0 Then
            MsgBox strRemplace
            sfct = sfct & " (" & strRemplace & ")"
        End If
        Set nd = tv.Nodes.Add(ndP, _
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
    
    ' Les postes non Actifs
    sql = "select PO_Num, PO_Libelle, FT_Libelle" _
        & " from Poste, FctTrav" _
        & " where FT_Num=PO_FTNum" _
        & " and PO_Actif=false" _
        & " and PO_SRVNum=" & v_numsrv _
        & " order by PO_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        charger_service = P_ERREUR
        Exit Function
    End If
    
    While Not rs.EOF
        sfct = rs("FT_Libelle").Value
        If rs("FT_Libelle").Value <> rs("PO_Libelle").Value Then
            sfct = sfct & " *"
        End If
        sfct = sfct & " (poste inactivé)"
        Set nd = tv.Nodes.Add(ndP, _
                               tvwChild, _
                               "P" & rs("PO_Num").Value, _
                               sfct, _
                               IMG_POSTE_SEL_INACTIF, _
                               IMG_POSTE_SEL_INACTIF)
        nd.tag = True
        rs.MoveNext
    Wend
    rs.Close
    
    ' Les services
    sql = "select SRV_Num, SRV_LNum, SRV_NumPere, SRV_Nom, SRV_NivsNum from Service " _
        & " where SRV_Actif=true" _
        & " and SRV_Numpere=" & v_numsrv _
        & " order by SRV_Nom"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        charger_service = P_ERREUR
        Exit Function
    End If
    
    While Not rs.EOF
        'lib = rs("SRV_Nom").Value
        'libNiveau = recup_lib_niveau(rs("SRV_NivsNum").Value)
        'If libNiveau <> "" Then
        '    lib = lib & " (" & libNiveau & ")"
        'End If
        Set nd = tv.Nodes.Add(ndP, _
                               tvwChild, _
                               "S" & rs("SRV_Num").Value, _
                               rs("SRV_Nom").Value, _
                               IMG_SRV, _
                               IMG_SRV)
        sql = "select count(*) from Service " _
            & " where SRV_Actif=true" _
            & " and SRV_Numpere=" & rs("SRV_Num").Value
        If Odbc_Count(sql, lnb) = P_ERREUR Then
            charger_service = P_ERREUR
            Exit Function
        End If
        If lnb = 0 Then
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
    
lab_fin:
    
    stag = ndP.tag
    Call STR_PutChamp(stag, "|", 1, True)
    ndP.tag = stag
    
    charger_service = P_OK
    
End Function

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
    Dim nd As Node, nde As Node, ndP As Node

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
    Call TV_FirstParent(nd, ndP)
    numlabo = Mid$(ndP.key, 2)
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

Private Function creer_service() As Integer

    Dim lib As String
    Dim nch As Integer
    Dim num As Long, numpere As Long, numlabo As Long
    Dim creation As Boolean
    Dim nd As Node, ndP As Node
    Dim frm As Form
    
    creation = False

    Set nd = tv.SelectedItem
    numpere = 0

    ' c'est un service
'    Call TV_FirstParent(nd, ndp)
    If left$(nd.key, 1) = "S" Then
        numpere = Mid$(nd.key, 2)
    End If
    ' vérifier si création est ok
    Set frm = PrmServiceModif
    If PrmServiceModif.AppelFrm(0, numpere, lib, num) Then
        creation = True
    End If
    Set frm = Nothing
    
    If Not creation Then ' on a annulé la création
        creer_service = P_NON
        Exit Function
    End If

    nd.Sorted = True
    nd.Expanded = True
    Set nd = tv.Nodes.Add(nd, tvwChild, "S" & num, lib, IMG_SRV, IMG_SRV)
    Set tv.SelectedItem = nd
    SendKeys "{DOWN}"
    SendKeys "{UP}"
    DoEvents
    Set tv.SelectedItem = nd
    
    creer_service = P_OUI
    
End Function

Private Function deplacer_sp() As Integer

    Dim key As String, sql As String, stype_src As String, stype_dest As String
    Dim s_sp_src As String, s_sp_dest As String, s_sp As String, s As String
    Dim key_depl As String
    Dim encore As Boolean
    Dim I As Integer, nbch As Integer
    Dim numsrv_src As Long, numsrv_dest As Long, lnb As Long
    Dim nd_src As Node, nd_dest As Node, nd As Node, ndP As Node
    Dim rs As rdoResultset
    
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
    If (stype_src = "P" Or stype_src = "C") And stype_dest = "L" Then
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
                        "WHERE PO_Num=" & numsrv_src, _
                        "PO_SRVNum", numsrv_dest) = P_ERREUR Then
            GoTo err_enreg
        End If
    ElseIf stype_src = "S" Then
        If Odbc_Update("Service", _
                        "SRV_Num", _
                        "WHERE SRV_Num=" & numsrv_src, _
                        "SRV_NumPere", numsrv_dest) = P_ERREUR Then
            GoTo err_enreg
        End If
    ElseIf stype_src = "C" Then
        If Odbc_Update("Piece", _
                        "PC_Num", _
                        "WHERE PC_Num=" & numsrv_src, _
                        "PC_SRVNum", numsrv_dest) = P_ERREUR Then
            GoTo err_enreg
        End If
        GoTo lab_commit
    End If
    
    ' Mise à jour Utilisateur
    sql = "SELECT U_Num, U_SPM FROM Utilisateur" _
        & " WHERE U_kb_actif=True AND U_SPM LIKE '%" & s_sp_src & "%'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo err_enreg
    End If
    While Not rs.EOF
        nbch = STR_GetNbchamp(rs("U_SPM").Value, "|")
        s_sp = ""
        For I = 1 To nbch
            s = STR_GetChamp(rs("U_SPM").Value, "|", I - 1)
            s = STR_Remplacer(s, s_sp_src, s_sp_dest)
            s_sp = s_sp + s
        Next I
        If Odbc_Update("Utilisateur", _
                       "U_Num", _
                       "WHERE U_kb_actif=True AND U_Num=" & rs("U_Num").Value, _
                       "U_SPM", s_sp) = P_ERREUR Then
            rs.Close
            GoTo err_enreg
        End If
        rs.MoveNext
    Wend
    rs.Close

lab_commit:
    Call Odbc_CommitTrans
    
    Call afficher_liste
    
    tv.SetFocus
    ' Se repositionne sur le "SP" déplacé
    For I = 1 To tv.Nodes.Count
        If tv.Nodes(I).key = key_depl Then
            Set ndP = tv.Nodes(I).Parent
            While left$(ndP.key, 1) <> "L"
                ndP.Expanded = True
                Set ndP = ndP.Parent
            Wend
        End If
    Next I
    
    Set tv.SelectedItem = tv.Nodes(key_depl)
    tv.SetFocus
    SendKeys "{DOWN}"
    SendKeys "{UP}"
    
    deplacer_sp = P_OK
    Exit Function
    
err_enreg:
    Call Odbc_RollbackTrans
    deplacer_sp = P_ERREUR
    
End Function

Private Function FctNivRemplace(ByVal v_FT_NivRemplace, ByVal v_numsrv As Long, ByRef r_strRemplace As String) As Integer
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
    Dim nd As Node, ndP As Node, nds As Node, ndF As Node
    
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
    Set ndF = tv.SelectedItem
    If left$(ndF.key, 1) = "L" Then
        s = "Services - " & ndF.Text
    Else
        s = ndF.Text
    End If
    Call PR_InitFormat(True, _
                       s, _
                       False, _
                       "g", _
                       stexte())
    Set nd = ndF
    Set ndF = ndF.Child
    encore = True
    decal = 0
    While encore
        If decal > 0 Then
            s = String(decal * 5, " ")
        Else
            s = ""
        End If
        s = s & left$(ndF.key, 1) & " - "
        stexte(0) = s & ndF.Text
        Call PR_ImpLigne(stexte())
        If ndF.Children > 0 Then
            Set ndF = ndF.Child
            decal = decal + 1
        Else
            encore = False
            While ndF.Index <> nd.Index And Not encore
                encore = True
                Set nds = ndF
                Set ndF = ndF.Next
                On Error GoTo pas_de_suivant
                s = ndF.key
                On Error GoTo 0
                If Not encore Then
                    Set ndF = nds.Parent
                    decal = decal - 1
                End If
            Wend
            If ndF.Index = nd.Index Then
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
    
    'MODIF KD
    g_crfct_autor = False
    
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

Private Sub activer_poste()

    Dim nom_poste As String, sret As String
    Dim sql As String, rs As rdoResultset
    Dim num_poste As Long, num_fct As Long
    Dim frm As Form
    Dim nd As Node
    

    Set nd = tv.SelectedItem
    num_poste = Mid$(tv.SelectedItem.key, 2)
    If num_poste > 0 Then
        sql = "Update Poste set PO_Actif='t' where PO_Num=" & num_poste
        Call Odbc_Cnx.Execute(sql)
        tv.SelectedItem.Text = Replace(tv.SelectedItem.Text, "(poste inactivé)", "")
        tv.SelectedItem.image = IMG_POSTE
        tv.SelectedItem.SelectedImage = IMG_POSTE
    End If

End Sub

Private Function modifier_service() As Integer

    Dim lib As String
    Dim nch As Integer
    Dim num As Long, numpere As Long
    Dim modification As Boolean
    Dim nd As Node, ndP As Node
    
    Set nd = tv.SelectedItem
    numpere = 0
    num = CLng(Mid$(nd.key, 2))
    If P_RecupSrvNom(num, lib) = P_ERREUR Then
        modifier_service = P_ERREUR
        Exit Function
    End If
    Call TV_FirstParent(nd, ndP)
    numpere = Mid$(ndP.key, 2)
    ' vérifier si création est ok
    If PrmServiceModif.AppelFrm(num, numpere, lib, num) Then
        modification = True
    End If

    If Not modification Then ' on a annulé la modification ou pas de modif pour le nom
        modifier_service = P_NON
        Exit Function
    End If

    nd.Sorted = True
    Set tv.SelectedItem = nd
    SendKeys "{DOWN}"
    SendKeys "{UP}"
    DoEvents
    Set tv.SelectedItem = nd
    nd.Text = lib
    
    modifier_service = P_OK
    
End Function

Private Sub quitter()
    
    g_sret = ""
    
    Unload Me
    
End Sub

Private Sub selectionner()

    Dim sp As String, sm As String, s As String
    Dim encore As Boolean
    Dim n As Integer, I As Integer, j As Integer, nbch As Integer, img As Integer
    Dim nd As Node, ndP As Node
    
    If g_plusieurs Then
        n = 0
        For I = 2 To tv.Nodes.Count
            img = tv.Nodes(I).image
            If img = IMG_POSTE_SEL Or img = IMG_SRV_SEL Or img = IMG_POSTE_SEL_NOMOD Or img = IMG_SRV_SEL_NOMOD Then
                sp = tv.Nodes(I).key & ";"
                Set nd = tv.Nodes(I)
                encore = True
                Do
                    Set ndP = nd.Parent
                    s = ndP.key
                    If left$(s, 1) = "L" Then
                        encore = False
                    Else
                        sp = sp + s + ";"
                        Set nd = ndP
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
            Set ndP = nd.Parent
            s = ndP.key
            If left$(s, 1) = "L" Then
                encore = False
            Else
                sp = sp + s + ";"
                Set nd = ndP
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
            If Odbc_Delete("Coordonnee_Associee", _
                            "CA_Num", _
                            "WHERE CA_UCTypeNum=" & num & " AND CA_UCType='S'", _
                            lnb) = P_ERREUR Then
                GoTo err_enreg
            End If
        Else
            If Odbc_Delete("Poste", "PO_Num", "WHERE PO_Num=" & num, lnb) = P_ERREUR Then
                GoTo err_enreg
            End If
            If Odbc_Delete("UtilCoordonnee", _
                           "UC_Num", _
                           "WHERE UC_TypeNum=" & num & " AND UC_Type='P'", _
                           lnb) = P_ERREUR Then
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

    reponse = MsgBox("Etes-vous sûr de vouloir supprimer cette pièce ?", vbQuestion + vbYesNo + vbDefaultButton2, "Attention !")
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

Private Sub mnuActiverP_Click()
    
    Call activer_poste

End Sub

Private Sub mnuCreerP_Click()

    Call creer_poste
    
End Sub

Private Sub mnuCreerPiece_Click()

    Call creer_piece

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

Private Sub mnuModP_Click()

    Call modifier_poste

End Sub

Private Sub mnuModPers_Click()

    g_sret = Mid$(tv.SelectedItem.tag, 2)
    Unload Me
    
End Sub

Private Sub mnuModPiece_Click()

    Call modifier_piece

End Sub


Private Sub mnuModS_Click()

    Call modifier_service
    
End Sub

Private Sub mnuSuppP_Click()

    Call supprimer
    
End Sub

Private Sub mnuSuppPiece_Click()

    Call supprimer_piece

End Sub

Private Sub mnuSuppS_Click()

    Call supprimer
    
End Sub

Private Sub mnuVoirPers_Click()

    Call ajouter_pers_poste_tv

End Sub

Private Sub tv_Click()

    If g_node = tv.SelectedItem.Index And g_expand <> tv.SelectedItem.Expanded Then
        Exit Sub
    End If
    
    If g_button = vbRightButton Then
        If left$(tv.SelectedItem.key, 1) = "S" Then
            If STR_GetChamp(tv.SelectedItem.tag, "|", 1) = False Then
                tv.Nodes.Remove (tv.SelectedItem.Child.Index)
                charger_service (Mid$(tv.SelectedItem.key, 2))
                tv.SelectedItem.Expanded = True
            End If
        End If
        Call afficher_menu(False)
    ElseIf g_button = vbLeftButton Then
        If g_pos_depl <> 0 Then
            If deplacer_sp() = P_ERREUR Then
                Call quitter
                Exit Sub
            End If
        ElseIf g_plusieurs Then
            Call basculer_selection
        End If
    End If
        
End Sub

Private Sub tv_DblClick()

    If g_mode_acces = MODE_SELECT Then Exit Sub
    
    If Mid$(tv.SelectedItem.key, 1, 1) = "C" Then
        Call modifier_piece
    ElseIf Mid$(tv.SelectedItem.key, 1, 1) = "P" Then
        Call modifier_poste
    End If

End Sub

Private Sub tv_Expand(ByVal Node As ComctlLib.Node)

    If Not g_mode_saisie Then
        Exit Sub
    End If
    
    g_button = -1
    If left$(Node.key, 1) = "S" Then
        If STR_GetChamp(Node.tag, "|", 1) = False Then
            tv.Nodes.Remove (Node.Child.Index)
            charger_service (Mid$(Node.key, 2))
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
    
End Sub

Private Sub TxtRecherche_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim I As Integer
    
    If KeyCode = 13 Then
        For I = 1 To tv.Nodes.Count
            If left$(tv.Nodes(I).key, 1) = "S" Then
                If STR_GetChamp(tv.Nodes(I).tag, "|", 1) = False Then
                    tv.Nodes.Remove (tv.Nodes(I).Child.Index)
                    charger_service (Mid$(tv.Nodes(I).key, 2))
                End If
            End If
        Next I
        If TxtRecherche.Text = "" Then
            For I = 1 To tv.Nodes.Count - 1
                tv.Nodes(I).Text = Replace(tv.Nodes(I).Text, "     <===", "")
            Next I
        Else
            For I = 1 To tv.Nodes.Count - 1
                'MsgBox tv.Nodes(I).Text
                If left$(tv.Nodes(I).key, 1) = "S" Then
                    If InStr(UCase(tv.Nodes(I).Text), UCase(TxtRecherche.Text)) = 0 Then
                        tv.Nodes(I).Expanded = False
                        tv.Nodes(I).Text = Replace(tv.Nodes(I).Text, "     <===", "")
                    Else
                        tv.Nodes(I).Expanded = True
                        If InStr(UCase(tv.Nodes(I).Text), "<===") = 0 Then
                            tv.Nodes(I).Text = tv.Nodes(I).Text & "     <==="
                        End If
                    End If
                End If
                If left$(tv.Nodes(I).key, 1) = "P" Then
                    Call Filtrer_Postes(tv.Nodes(I).Parent.key, UCase(TxtRecherche.Text))
                End If
            Next I
        End If
    End If

End Sub

Private Sub Filtrer_Postes(v_key, v_texte)
    Dim nd As Node
    Dim ndF As Node
    Dim ndP As Node
    
    Dim I As Integer
    Dim j As Integer
    
    Set nd = tv.Nodes(v_key)
    If nd.Children > 0 Then
        Set ndF = nd.Child
        For I = 1 To nd.Children
            If left$(ndF.key, 1) = "P" Then
                If InStr(UCase(ndF.Text), UCase(TxtRecherche.Text)) > 0 Then
                    If InStr(UCase(ndF.Text), "<===") = 0 Then
                        ndF.Text = ndF.Text & "     <==="
                    End If
                    ' Ouvrir ses pères
                    Set ndP = ndF.Parent
                    For j = 1 To 10
                        If Mid(ndP.key, 1, 1) = "L" Then
                            Exit For
                        Else
                            ndP.Expanded = True
                            Set ndP = ndP.Parent
                        End If
                    Next j
                Else
                    ndF.Text = Replace(ndF.Text, "     <===", "")
                End If
            End If
        Next I
    End If
End Sub
