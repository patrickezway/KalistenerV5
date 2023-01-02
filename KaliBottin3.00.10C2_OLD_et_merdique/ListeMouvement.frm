VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ListeMouvement 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm 
      Caption         =   "Liste des personnes modifiées dans KaliDoc"
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
      Height          =   7425
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11895
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   6435
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   11351
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedCols       =   0
         BackColorFixed  =   12648447
         BackColorSel    =   8388608
         SelectionMode   =   1
      End
      Begin ComctlLib.ImageList ImgList 
         Left            =   9360
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   17
         ImageHeight     =   17
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   3
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ListeMouvement.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ListeMouvement.frx":03C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ListeMouvement.frx":078C
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
      TabIndex        =   0
      Top             =   7300
      Width           =   11895
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
         Left            =   5640
         Picture         =   "ListeMouvement.frx":0B52
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Quitter"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "ListeMouvement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Index des colonnes du grid
Private Const GRDC_U_NUM = 0
Private Const GRDC_ACTION = 1
Private Const GRDC_MATRICULE = 2
Private Const GRDC_NOM = 3
Private Const GRDC_PRENOM = 4
Private Const GRDC_POSTE = 5
Private Const GRDC_DETAIL = 6
Private Const GRDC_DESACTIVEE = 7

' Index des IMG
Private Const IMG_DETAIL = 1

Private Const COLOR_DESACTIVEE = &HE0E0E0 ' GRIS

Private Const CMD_QUITTER = 0

Private g_form_active As Boolean
Private g_col_tri As Boolean
Private g_sens_tri As Boolean

Public Sub AppelFrm()
'***********************************
' Le point d'entrée pour cette forme
'***********************************
    Call FRM_ResizeForm(Me, Me.width, Me.Height)

    Me.Show 1

End Sub

Private Sub desactiver_ligne(ByVal v_row As Integer)

    Dim i As Integer

    If grd.Rows = 2 Then
        Call quitter
        Exit Sub
    End If
    grd.RemoveItem v_row
    
    Exit Sub
    
' NE SERT PLUS
    With grd
        .Row = v_row
        For i = 0 To .Cols - 1
            .col = i
            .CellBackColor = COLOR_DESACTIVEE
            .col = GRDC_DETAIL
            Set .CellPicture = LoadPicture("")
        Next i
        .col = GRDC_ACTION
        Set .CellPicture = LoadPicture("")
        .TextMatrix(v_row, GRDC_DESACTIVEE) = ""
    End With

End Sub

Private Function get_poste(ByVal v_u_num As Long) As String
'**************************************************************
' Retourne le poste à afficher dans le grid pour cette personne
'**************************************************************
    Dim sql As String, spm As String

    get_poste = ""
    sql = "SELECT U_SPM FROM Utilisateur WHERE U_Num=" & v_u_num
    If Odbc_RecupVal(sql, spm) = P_ERREUR Then
        Exit Function
    End If

    If spm <> "" Then
        get_poste = P_get_lib_srv_poste(P_get_num_srv_poste(spm, P_POSTE), P_POSTE)
        get_poste = get_poste & " - " & P_get_lib_srv_poste(P_get_num_srv_poste(spm, P_SERVICE), P_SERVICE)
    End If

End Function

Private Sub initialiser()

    Dim i As Integer

    cmd(CMD_QUITTER).left = (Me.width / 2) - (cmd(CMD_QUITTER).width / 2)
    With grd
        .FormatString = "u_num||MATRICULE|NOM|PRENOM|POSTE KaliBottin|"
        .ColWidth(GRDC_U_NUM) = 0
        .ColWidth(GRDC_ACTION) = 255
        .ColWidth(GRDC_MATRICULE) = 1300
        .ColWidth(GRDC_NOM) = 1700
        .ColWidth(GRDC_PRENOM) = 1700
        .ColWidth(GRDC_POSTE) = 5680
        .ColWidth(GRDC_DETAIL) = 260
        .ColWidth(GRDC_DESACTIVEE) = 0
        .Row = 0
        For i = 0 To .Cols - 1
            .col = i
            .CellFontBold = True
            .CellAlignment = flexAlignCenterCenter
        Next i
    End With

    Call remplir_grid

End Sub

Private Sub quitter()

    Unload Me

End Sub

Private Sub remplir_action(ByVal v_unum As Long)
'*********************************************************
' Determiner la pastille à afficher dans la cellule ACTION
'*********************************************************
    Dim u_spm As String

    If Odbc_RecupVal("SELECT U_SPM FROM Utilisateur" _
                   & " WHERE U_Num=" & v_unum, u_spm) = P_ERREUR Then
        Exit Sub
    End If

    grd.TextMatrix(grd.Rows - 1, GRDC_ACTION) = IIf(u_spm = "", _
                                                    "C", _
                                                    "M")
    grd.TextMatrix(grd.Rows - 1, GRDC_DESACTIVEE) = IIf(u_spm = "", _
                                                        "Création", _
                                                        "Modification")

End Sub

Private Sub remplir_grid()

    Dim sql As String
    Dim rs As rdoResultset

    sql = "SELECT U_Num, U_Nom, U_Prenom, U_Matricule" _
        & " FROM UtilMouvement, Utilisateur" _
        & " WHERE UM_UNum=U_Num AND UM_APPNum=" & p_appli_kalidoc _
        & " AND UM_DatePEC_KD_KB IS NULL" _
        & " GROUP BY U_Num, U_Nom, U_Prenom, U_Matricule"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    With grd
        While Not rs.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, GRDC_U_NUM) = rs("U_Num").Value
            Call remplir_action(rs("U_Num").Value)
            .TextMatrix(.Rows - 1, GRDC_MATRICULE) = rs("U_Matricule").Value
            .TextMatrix(.Rows - 1, GRDC_NOM) = rs("U_Nom").Value
            .TextMatrix(.Rows - 1, GRDC_PRENOM) = rs("U_Prenom").Value
            .TextMatrix(.Rows - 1, GRDC_POSTE) = get_poste(rs("U_Num").Value)
            .Row = .Rows - 1
            .col = GRDC_DETAIL
            .CellPictureAlignment = flexAlignCenterCenter
            Set .CellPicture = ImgList.ListImages(IMG_DETAIL).Picture
            rs.MoveNext
        Wend
        rs.Close
        ' trier le grid
        .Row = 0
        .col = GRDC_NOM
        .Sort = 1
        g_sens_tri = 1
        g_col_tri = GRDC_NOM
        ' redimension pour la bare de défilement vertical
        If .Rows - 1 > 18 Then
            .ColWidth(GRDC_NOM) = .ColWidth(GRDC_NOM) - 50
            .ColWidth(GRDC_PRENOM) = .ColWidth(GRDC_PRENOM) - 55
            .ColWidth(GRDC_MATRICULE) = .ColWidth(GRDC_MATRICULE) - 50
            .ColWidth(GRDC_POSTE) = .ColWidth(GRDC_POSTE) - 60
            .ColWidth(GRDC_DETAIL) = .ColWidth(GRDC_DETAIL) - 50
        End If
        .Visible = True
    End With

End Sub

Private Sub cmd_Click(Index As Integer)

    Call quitter

End Sub

Private Sub Form_Activate()

    If g_form_active Then Exit Sub

    g_form_active = True
    Call initialiser

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Call quitter
    End If
    
End Sub

Private Sub Form_Load()

    g_form_active = False

End Sub

Private Sub grd_Click()

    Dim mouse_row As Integer

    With grd
        mouse_row = .MouseRow
        If mouse_row = 0 Then Exit Sub
        If HistoriqueMouvement.AppelFrm(.TextMatrix(mouse_row, GRDC_U_NUM), _
                                        " AND UM_APPNum=" & p_appli_kalidoc _
                                        & " and UM_datepec_kd_kb is null", _
                                        IIf(.TextMatrix(mouse_row, GRDC_DESACTIVEE) = "", False, True)) Then
            Call desactiver_ligne(mouse_row)
        End If
    End With

End Sub

Private Sub grd_DblClick()

    Dim mouse_row As Integer, mouse_col As Integer, i As Integer

    With grd
        If .Rows - 1 = 0 Then Exit Sub
        mouse_row = .MouseRow
        mouse_col = .MouseCol
        ' Trier selon la colonne choisie dans la ligne fixe uniquement
        If mouse_row = 0 Then
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
            .TopRow = 1
            .col = 0
            .Row = 1
            .RowSel = 1
            .ColSel = .Cols - 1
            g_col_tri = mouse_col
        ElseIf mouse_row <> 0 Then ' n'importe où sur le GRID sauf la 1° ligne
        End If
    End With

End Sub

Private Sub grd_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

    Dim mouse_row As Integer, mouse_col As Integer

    mouse_row = grd.MouseRow
    mouse_col = grd.MouseCol
    If mouse_col <> GRDC_ACTION Then Exit Sub
    grd.ToolTipText = grd.TextMatrix(mouse_row, GRDC_DESACTIVEE)

End Sub


