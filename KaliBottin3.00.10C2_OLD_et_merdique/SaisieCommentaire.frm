VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form SaisieCommentaire 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   600
      Picture         =   "SaisieCommentaire.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Aucun"
      Top             =   4440
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
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
      Left            =   120
      Picture         =   "SaisieCommentaire.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Tous"
      Top             =   4440
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0C0C0&
      Height          =   510
      Index           =   0
      Left            =   2400
      Picture         =   "SaisieCommentaire.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   550
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0C0C0&
      Height          =   510
      Index           =   1
      Left            =   6960
      Picture         =   "SaisieCommentaire.frx":0ADD
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   550
   End
   Begin VB.TextBox txt 
      Height          =   615
      Left            =   120
      MaxLength       =   255
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   600
      Width           =   10215
   End
   Begin MSFlexGridLib.MSFlexGrid grdPoste 
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   3836
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
   Begin VB.Label Lblposte 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Postes liés à cette coordonnée"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   10215
   End
   Begin ComctlLib.ImageList imglst 
      Left            =   9600
      Top             =   4200
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
            Picture         =   "SaisieCommentaire.frx":1096
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SaisieCommentaire.frx":13E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SaisieCommentaire.frx":173A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Saisissez le commentaire ici"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "SaisieCommentaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'images du grid
Private Const IMG_SANS_COCHE = 1
Private Const IMG_COCHE = 2
Private Const IMG_COCHE_GRIS = 3

' Index des CMD
Private Const CMD_OK = 0
Private Const CMD_ANNULER = 1

Private g_return As Boolean
Private g_ucnum As Long
Private g_affichage As Boolean
Private g_lstoldposte As String
Private g_grdCoord As MSFlexGrid

Public Function AppelFrm(ByVal v_grdCoord As MSFlexGrid, Optional v_ucnum As Long) As Boolean
' **************************************
' ENTRÉE: le grid à manipuler
' SORTIE: est-ce que le texte a changé ?
' **************************************
    Set g_grdCoord = v_grdCoord
    g_ucnum = v_ucnum
    g_return = False
    g_lstoldposte = ""

    With g_grdCoord
        txt.Text = .TextMatrix(.Row, .col)
        lbl = "Saisissez le commentaire concernant : " & .TextMatrix(.Row, 1)
    End With

    Me.Show 1

    AppelFrm = g_return

End Function

Private Sub quitter()

    g_return = False
    Unload Me

End Sub

Private Sub valider()
    
    Dim lst_newposte As String
    Dim nbr_ligne As Integer, I As Integer
    
    With g_grdCoord
        If .TextMatrix(.Row, .col) <> txt.Text Then
            g_return = True
        Else
            g_return = False
        End If
        If g_ucnum > 0 Then
            On Error Resume Next
            nbr_ligne = UBound(CL_liste.lignes) + 1
            On Error GoTo 0
            lst_newposte = ""
            For I = 0 To nbr_ligne - 1
                If CL_liste.lignes(I).selected Then
                    lst_newposte = lst_newposte + CL_liste.lignes(I).tag
                End If
            Next I
            'Faire un update de la liste si different
            If LCase$(lst_newposte) <> LCase$(g_lstoldposte) Then
                Call update_liste_poste(lst_newposte)
                g_return = True
            End If
        End If
        .TextMatrix(.Row, .col) = txt.Text
        Unload Me
    End With

End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case CMD_OK
            Call valider
        Case CMD_ANNULER
            Call quitter
    End Select

End Sub

Private Sub ajouter_ligne(ByVal v_str As String)
    
    Dim nom As String, numservice As Long, numposte As Long
    Dim nomservice As String, nomposte As String, couple As String
    Dim nbchp As Integer
    Dim checked As Boolean
    
    checked = False
    nbchp = STR_GetNbchamp(v_str, ";")
    numservice = CLng(Mid$(STR_GetChamp(v_str, ";", nbchp - 2), 2))
    numposte = CLng(Mid$(STR_GetChamp(v_str, ";", nbchp - 1), 2))
    couple = STR_GetChamp(v_str, ";", nbchp - 2) & ";" & STR_GetChamp(v_str, ";", nbchp - 1) & ";" & "|"
    If P_RecupSrvNom(numservice, nomservice) = P_ERREUR Then
        nomservice = "Inconnu"
    End If
    If P_RecupPosteNom(numposte, nomposte) = P_ERREUR Then
        nomposte = "Inconnu"
    End If
    nom = nomservice & " - " & nomposte
    If InStr(g_lstoldposte, couple) > 0 Then
        checked = True
    End If
    Call CL_AddLigne(nom, g_ucnum, couple, checked)
    
End Sub

Private Sub afficher_ligne(ByVal v_row As Integer, _
                           ByVal v_col As Integer, _
                           ByVal v_str As String)
    
    Dim lg_col As Long
    
    If grdPoste.Cols < v_col + 1 Then
        grdPoste.Cols = v_col + 1
    End If
    grdPoste.TextMatrix(v_row, v_col) = v_str
    lg_col = FRM_LargeurTexte(Me, grdPoste, v_str + "     ")
    If grdPoste.ColWidth(v_col) < lg_col Then
        grdPoste.ColWidth(v_col) = lg_col
    End If
    
End Sub

Private Sub update_liste_poste(ByVal v_lstposte As String)
    
    Dim rs As rdoResultset
    Dim sql As String
    
    If Odbc_Update("UtilCoordonnee", "UC_Num", "WHERE UC_Num=" & g_ucnum, _
                                "UC_Lstposte", v_lstposte) = P_ERREUR Then
        MsgBox " Erreur inscription liste "
    End If
    
End Sub

Private Sub changer_etat_selection()
    
    grdPoste.col = 0
    If CL_liste.lignes(grdPoste.Row - grdPoste.FixedRows).selected Then
        CL_liste.lignes(grdPoste.Row - grdPoste.FixedRows).selected = False
        Set grdPoste.CellPicture = imglst.ListImages(IMG_SANS_COCHE).Picture
    Else
        CL_liste.lignes(grdPoste.Row - grdPoste.FixedRows).selected = True
        Set grdPoste.CellPicture = imglst.ListImages(IMG_COCHE).Picture
    End If
    grdPoste.ColSel = grdPoste.Cols - 1
    
End Sub

Private Sub Form_Activate()
    
    Call initialiser
    
End Sub

Private Sub initialiser()

    Dim tabstr() As String
    Dim nb_ligne As Integer, first_col As Integer
    Dim icol As Integer, ilig As Integer, I As Integer, n As Integer
    Dim sql As String
    Dim rs As rdoResultset
    
    'Si pas affichage
    If g_ucnum = 0 Then
        GoTo Fin
    End If
    'Initialisation du tableau
    Call CL_Init
    Call CL_InitMultiSelect(True, False)
    ' Lignes
    nb_ligne = 0
    ' liste des postes
    sql = "SELECT U_Spm, UC_lstposte FROM UtilCoordonnee, Utilisateur " _
        & " WHERE UC_Num = " & g_ucnum _
        & " AND UC_Typenum = U_num "
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo Fin
    End If
    If Not rs.EOF Then
        If IsNull(rs("U_Spm").Value) Then
            GoTo Fin
        Else
            nb_ligne = STR_GetNbchamp(rs("U_Spm").Value, "|")
        End If
        g_lstoldposte = rs("UC_Lstposte").Value & ""
    End If
    
    ' Ajout des lignes
    For ilig = 0 To nb_ligne - 1
        Call ajouter_ligne(STR_GetChamp(rs("U_Spm").Value, "|", ilig))
    Next ilig
    
    On Error Resume Next
    nb_ligne = UBound(CL_liste.lignes) + 1
    On Error GoTo 0
    ' Initialisation du grd
    grdPoste.SelectionMode = flexSelectionByRow
    grdPoste.FocusRect = flexFocusNone
    grdPoste.ScrollBars = flexScrollBarNone
    grdPoste.FixedCols = 0
    grdPoste.BackColorBkg = grdPoste.BackColor
    grdPoste.Rows = nb_ligne
    If nb_ligne > 0 Then
        grdPoste.FixedRows = 0
    End If
    first_col = 1
    ' Affichage des lignes
    For ilig = 0 To nb_ligne - 1
        Call STR_Decouper(CL_liste.lignes(ilig).texte, tabstr)
        For icol = 0 To UBound(tabstr)
            Call afficher_ligne(ilig + grdPoste.FixedRows, icol + first_col, tabstr(icol))
        Next icol
        grdPoste.Row = grdPoste.FixedRows + ilig
        grdPoste.col = 0
        If CL_liste.lignes(ilig).selected Then
            Set grdPoste.CellPicture = imglst.ListImages(IMG_COCHE).Picture
        Else
            Set grdPoste.CellPicture = imglst.ListImages(IMG_SANS_COCHE).Picture
        End If
    Next ilig
    
    grdPoste.col = 0
    grdPoste.ColSel = 1
    grdPoste.SetFocus
    
    Exit Sub
    
Fin:
    
    Lblposte.Visible = False
    grdPoste.Visible = False
    txt.Height = 3130

    
End Sub

Private Sub grdPoste_Click()
    
    Call changer_etat_selection
    
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape
            Call quitter
        Case vbKeyReturn
            Call valider
    End Select

End Sub
