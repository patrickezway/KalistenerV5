VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form HistoriqueMouvement 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame frm 
      Caption         =   "Veuillez patienter s'il vous plait"
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
      Height          =   7455
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11895
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   6090
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   10742
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
         ForeColor       =   0
         BackColorFixed  =   12648447
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.Label lbl 
         Caption         =   "* signe (+) pour les ajouts / signe (-) pour les suppressions."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   6960
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Label lbl 
         Caption         =   "** ces valeurs ne concernent que les coordonnées principales."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   4
         Top             =   6960
         Visible         =   0   'False
         Width           =   5415
      End
   End
   Begin VB.Frame frm 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   7320
      Width           =   11895
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Prendre en compte"
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
         Index           =   1
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1815
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
         Left            =   5520
         Picture         =   "HistoriqueMouvement.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Quitter"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "HistoriqueMouvement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Index des boutons
Private Const CMD_QUITTER = 0
Private Const CMD_PRENDE_EN_COMPTE = 1

' Index des libellés
Private Const LBL_ASTERIX_POSTES = 0
Private Const LBL_ASTERIX_COORDONNEES = 1

' Index des colonnes
Private Const GRDC_INVISIBLE = 0 ' pour une selection correcte des lignes
Private Const GRDC_ACTION = 1
Private Const GRDC_DATE = 2
Private Const GRDC_APPLICATION = 3
Private Const GRDC_VALEUR_AVANT = 4
Private Const GRDC_VALEUR_APRES = 5

' Pour les cellules VALEUR AVANT/APRES dans CREATION, INACTIVE et ACTIVE
' &H80000018& tooltip
Private Const COLOR_DESACTIVE = &HE0E0E0        ' GRIS
Private Const COLOR_VALEUR_AVANT = &H80000013   ' BLEU CLAIR
Private Const COLOR_VALEUR_APRES = &H80000016
Private Const COLOR_VALEUR_AVANT_2 = &H80000003 ' BLEU FONCÉ
Private Const COLOR_VALEUR_APRES_2 = &H8000000F

Private g_numutil As Long
Private g_form_active As Boolean
Private g_complement_sql As String
Private g_asterix_coordonnees_visible As Boolean
Private g_asterix_postes_visible As Boolean
Private g_PrendreEnCompte As Boolean
Private g_desactiver_ligne As Boolean

Public Function AppelFrm(ByVal v_numutil As Long, ByVal v_complement_sql As String, _
                         ByVal v_btnPrendreEnCompte As Boolean) As Boolean
' ****************************************************************************************
' Retourne a ListeMouvement s'il faut désactiver la ligne après une prise en compte ou non
' ****************************************************************************************
    g_complement_sql = v_complement_sql
    g_numutil = v_numutil
    g_PrendreEnCompte = v_btnPrendreEnCompte

    Call FRM_ResizeForm(Me, Me.width, Me.Height)

    Me.Show 1

    AppelFrm = g_desactiver_ligne

End Function

Private Sub initialiser()

    Dim I As Integer

    If g_PrendreEnCompte Then
        cmd(CMD_QUITTER).left = (2 * (frm(1).width / 3)) - (cmd(CMD_QUITTER).width / 2)
        cmd(CMD_PRENDE_EN_COMPTE).left = (frm(1).width / 3) - (cmd(CMD_PRENDE_EN_COMPTE).width / 2)
        g_desactiver_ligne = False
    Else
        cmd(CMD_QUITTER).left = (frm(1).width / 2) - (cmd(CMD_QUITTER).width / 2)
        cmd(CMD_PRENDE_EN_COMPTE).Visible = False
    End If
    With grd
        .ScrollTrack = True
        .FormatString = "invisible|ACTION|DATE|APPLI.|VALEUR AVANT|VALEUR APRÈS"
        .ColWidth(GRDC_INVISIBLE) = 0
        .ColWidth(GRDC_ACTION) = 2000
        .ColWidth(GRDC_DATE) = 1000
        .ColWidth(GRDC_APPLICATION) = 1200
        .ColWidth(GRDC_VALEUR_AVANT) = 3440
        .ColWidth(GRDC_VALEUR_APRES) = 3440
        .ColAlignment(GRDC_APPLICATION) = flexAlignCenterCenter
        .ColAlignment(GRDC_VALEUR_AVANT) = flexAlignLeftCenter
        .ColAlignment(GRDC_VALEUR_APRES) = flexAlignLeftCenter
        .Row = 0
        For I = 0 To .Cols - 1
            .col = I
            .CellFontBold = True
            .CellAlignment = flexAlignCenterCenter
        Next I
    End With

    g_asterix_postes_visible = False
    g_asterix_coordonnees_visible = False

    Call remplir_grid

    Me.Visible = True

End Sub

Private Function maj_um_datepec_kd_kb() As Integer
'********************************************************************
' MAJ UtilMouvement pour toutes les modifs faites dans KaliDoc et
' ayant UM_DatePEC_KD_KB = Null
'********************************************************************
    Dim sql As String
    Dim um_num As Long
    Dim rs As rdoResultset

    sql = "SELECT UM_Num FROM UtilMouvement" _
        & " WHERE UM_UNum=" & g_numutil _
        & " AND UM_APPNum=" & p_appli_kalidoc _
        & " AND UM_DatePEC_KD_KB IS NULL"
    If Odbc_Select(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    While Not rs.EOF
        um_num = rs("UM_Num").Value
        If Odbc_Update("UtilMouvement", "UM_Num", _
                       "WHERE UM_Num=" & um_num, _
                       "UM_DatePEC_KD_KB", Date) = P_ERREUR Then
            rs.Close
            GoTo lab_erreur
        End If
        rs.MoveNext
    Wend
    rs.Close

    maj_um_datepec_kd_kb = P_OK
    Exit Function

lab_erreur:
    maj_um_datepec_kd_kb = P_ERREUR

End Function

'********************************************************************
' 1° MAJ UM_DatePEC_KD_KB
' 2° Selectionner un poste KaliBottin pour les CREATIONS dans KaliDoc
'    ou MODIFICATION DU POSTE dans KalDoc
'********************************************************************
Private Sub prendre_en_compte()
    
    Dim sql As String, str As String, spm As String, ssite As String
    Dim nbr As Integer, I As Integer, nbr_postes_kd As Integer
    Dim creation As Boolean
    Dim rs As rdoResultset
    Dim frm As Form

    Call Odbc_BeginTrans

    ' maj [UtilMouvement] de tous les cas
    If maj_um_datepec_kd_kb() = P_ERREUR Then
        GoTo lab_erreur
    End If

    g_desactiver_ligne = True
    Call Odbc_CommitTrans
    Exit Sub

lab_erreur:
    g_desactiver_ligne = False

lab_annuler:
    Call Odbc_RollbackTrans
    Call quitter

End Sub

Private Sub quitter()

    Unload Me
    Exit Sub

End Sub

Private Function recup_PSLib(ByVal v_numposte As Long) As String
    
    Dim s As String, lib As String, sql As String
    Dim numsrv As Long
    Dim rs As rdoResultset
    
    lib = ""

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
    lib = rs("FT_Libelle").Value
    numsrv = rs("PO_SRVNum").Value
    rs.Close
    sql = "select SRV_Nom from Service" _
        & " where SRV_Num=" & numsrv
    If Odbc_SelectV(sql, rs) = P_OK Then
        If Not rs.EOF Then
            lib = lib & " - " & rs("SRV_Nom").Value
        End If
        rs.Close
    End If
    
    recup_PSLib = lib

End Function

Private Sub remplir_grid()

    Dim sql As String, str As String, str_modifiee As String, val_avant As String, val_apres As String, _
        poste_kalidoc As String, service_kalidoc As String, s As String
        Dim n As Integer
    Dim spm As Variant
    Dim rs As rdoResultset
' ARRET renommer  postekalidoc et servicekalidoc en kalibottin
    ' ***************** NOM, PRENOM et MATRICULE de cet personne *****************
    sql = "SELECT * FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & g_numutil
    If Odbc_Select(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    str = "Historique des mouvements: " & rs("U_Nom").Value & " " & rs("U_Prenom").Value _
        & " (Matricule " & rs("U_Matricule").Value & ")"
    If g_PrendreEnCompte Then ' récupérer le poste et le service KaliDoc
        service_kalidoc = P_get_lib_srv_poste(P_get_num_srv_poste(rs("U_SPM").Value, P_SERVICE), P_SERVICE)
        poste_kalidoc = P_get_lib_srv_poste(P_get_num_srv_poste(rs("U_SPM").Value, P_POSTE), P_POSTE)
    End If
    rs.Close
    ' *************** les MOUVEMENTS triés par ordre chronologique ***************
    sql = "SELECT * FROM UtilMouvement WHERE UM_UNum=" & g_numutil & g_complement_sql & " ORDER BY UM_Num"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Sub
    End If
    ' **************************** pas de mouvements ****************************
    ' LE TEST SE FAIT AUSSI DANS PrmPersonne.affich_utilisateur()
    If rs.EOF Then
        'frm(0).Caption = "Il n'y a pas de mouvements enregistrés pour cette personne."
        Call MsgBox("Il n'y a pas de mouvements enregistrés pour cette personne.", vbInformation + vbOKOnly, "")
        Call quitter
        Exit Sub
    End If
    ' ***************************** remplir le grid *****************************
    With grd
        While Not rs.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, GRDC_DATE) = rs("UM_Date").Value
            .TextMatrix(.Rows - 1, GRDC_APPLICATION) = P_get_nom_appli(rs("UM_APPNum").Value)
            .Row = .Rows - 1
            .col = GRDC_VALEUR_AVANT
            .CellBackColor = COLOR_DESACTIVE
            .col = GRDC_VALEUR_APRES
            .CellBackColor = COLOR_DESACTIVE
            If rs("UM_TypeMvt").Value = "C" Then ' ------------------------------- CREATION
                .TextMatrix(.Rows - 1, GRDC_ACTION) = "CREATION"
                If g_PrendreEnCompte Then ' lors de la prise en compte après une modification dans KaliDoc
                    If Odbc_RecupVal("select U_SPM from Utilisateur where U_kb_actif=True AND U_Num=" & g_numutil, _
                                      spm) = P_ERREUR Then
                        Exit Sub
                    End If
                    str_modifiee = "|A:"
                    For n = 0 To STR_GetNbchamp(spm, "|") - 1
                        s = STR_GetChamp(spm, "|", n)
                        str_modifiee = str_modifiee + STR_GetChamp(s, ";", STR_GetNbchamp(s, ";") - 1) + ";"
                    Next n
                    Call remplir_poste(str_modifiee)
                End If
            ElseIf rs("UM_TypeMvt").Value = "A" Then ' --------------------------- ACTIVE
                .TextMatrix(.Rows - 1, GRDC_ACTION) = "ETAT DU COMPTE"
                .TextMatrix(.Rows - 1, GRDC_VALEUR_AVANT) = "INACTIF"
                .TextMatrix(.Rows - 1, GRDC_VALEUR_APRES) = "ACTIF"
            ElseIf rs("UM_TypeMvt").Value = "I" Then ' --------------------------- INACTIVE
                .TextMatrix(.Rows - 1, GRDC_ACTION) = "ETAT DU COMPTE"
                .TextMatrix(.Rows - 1, GRDC_VALEUR_AVANT) = "ACTIF"
                .TextMatrix(.Rows - 1, GRDC_VALEUR_APRES) = "INACTIF"
            ElseIf rs("UM_TypeMvt").Value = "M" Then ' --------------------------- MODIFICATION
                .TextMatrix(.Rows - 1, GRDC_ACTION) = "MODIF.: " & STR_GetChamp(rs("UM_Commentaire").Value, "=", 0)
                str_modifiee = STR_GetChamp(rs("UM_Commentaire").Value, "=", 1)
                If STR_GetChamp(rs("UM_Commentaire").Value, "=", 0) = "POSTE" Then ' POSTE
                    .TextMatrix(.Rows - 1, GRDC_ACTION) = .TextMatrix(.Rows - 1, GRDC_ACTION) & " *"
                    g_asterix_postes_visible = True
                    Call remplir_poste(str_modifiee)
                ElseIf STR_GetChamp(rs("UM_Commentaire").Value, "=", 0) = "NOM" _
                    Or STR_GetChamp(rs("UM_Commentaire").Value, "=", 0) = "PRENOM" _
                    Or STR_GetChamp(rs("UM_Commentaire").Value, "=", 0) = "MATRICULE" Then ' NOM-PRENOM-MATRICULE
                    val_avant = STR_GetChamp(str_modifiee, ";", 0)
                    val_apres = STR_GetChamp(str_modifiee, ";", 1)
                    .TextMatrix(.Rows - 1, GRDC_VALEUR_AVANT) = IIf(val_avant = "", "[inexistant ou vide]", val_avant)
                    .TextMatrix(.Rows - 1, GRDC_VALEUR_APRES) = IIf(val_apres = "", "[vide ou supprimé]", val_apres)
                    .col = GRDC_VALEUR_AVANT
                    .CellBackColor = COLOR_VALEUR_AVANT
                    .col = GRDC_VALEUR_APRES
                    .CellBackColor = COLOR_VALEUR_APRES
                Else '                              COORDONNEES
                    .TextMatrix(.Rows - 1, GRDC_ACTION) = .TextMatrix(.Rows - 1, GRDC_ACTION) & " **"
                    g_asterix_coordonnees_visible = True
                    val_avant = STR_GetChamp(str_modifiee, ";", 0)
                    val_apres = STR_GetChamp(str_modifiee, ";", 1)
                    .TextMatrix(.Rows - 1, GRDC_VALEUR_AVANT) = IIf(val_avant = "", "[inexistant(e) ou vide]", val_avant)
                    .TextMatrix(.Rows - 1, GRDC_VALEUR_APRES) = IIf(val_apres = "", "[vide ou supprimé(e)]", val_apres)
                    .col = GRDC_VALEUR_AVANT
                    .CellBackColor = COLOR_VALEUR_AVANT
                    .col = GRDC_VALEUR_APRES
                    .CellBackColor = COLOR_VALEUR_APRES
                End If
            End If
            rs.MoveNext
        Wend
        rs.Close
        If .Rows - 1 = 0 Then
            .Enabled = False
        ElseIf grd.Rows - 1 > 15 Then ' redimensionner la bare de défilement vertical
            .ColWidth(GRDC_ACTION) = .ColWidth(GRDC_ACTION) - 180
            .ColWidth(GRDC_APPLICATION) = .ColWidth(GRDC_APPLICATION) - 75
        End If
        frm(0).Caption = str
        If g_asterix_postes_visible Then
            lbl(LBL_ASTERIX_POSTES).Visible = True
            If g_asterix_coordonnees_visible Then
                lbl(LBL_ASTERIX_COORDONNEES).Visible = True
            End If
        ElseIf g_asterix_coordonnees_visible Then
            lbl(LBL_ASTERIX_COORDONNEES).left = lbl(LBL_ASTERIX_POSTES).left
            lbl(LBL_ASTERIX_COORDONNEES).Visible = True
        End If
        .Visible = True
    End With

End Sub

Private Sub remplir_poste(ByVal v_str As String)
' Remplir les valeurs avant et après pour les postes ( exp. v_str = "P4;P6;|A:P40;P1;S:P6;" )

    Dim str_avant As String, str_apres As String, str_ajout As String, str_suppression As String
    Dim lib As String
    Dim nbr_avant As Integer, I As Integer, nbr_ajout As Integer, nbr_suppression As Integer, nbr As Integer
    Dim n_avant As Integer, n_ajout As Integer, n_suppression As Integer
    
    str_avant = STR_GetChamp(v_str, "|", 0)
    str_apres = STR_GetChamp(v_str, "|", 1)
    str_ajout = STR_GetChamp(STR_GetChamp(str_apres, "A:", 1), "S:", 0)
    str_suppression = STR_GetChamp(str_apres, "S:", 1)
    nbr_avant = STR_GetNbchamp(str_avant, ";")
    n_avant = 0
    nbr_ajout = STR_GetNbchamp(str_ajout, ";")
    n_ajout = 0
    nbr_suppression = STR_GetNbchamp(str_suppression, ";")
    n_suppression = 0
    nbr = IIf(nbr_avant >= nbr_ajout + nbr_suppression, nbr_avant, nbr_ajout + nbr_suppression)
    With grd
        .Row = .Rows - 1
        .col = GRDC_ACTION
        .CellBackColor = COLOR_VALEUR_AVANT
        .col = GRDC_DATE
        .CellBackColor = COLOR_VALEUR_AVANT
        .col = GRDC_APPLICATION
        .CellBackColor = COLOR_VALEUR_AVANT
        .col = GRDC_VALEUR_AVANT
        .CellBackColor = COLOR_VALEUR_AVANT
        .col = GRDC_VALEUR_APRES
        .CellBackColor = COLOR_VALEUR_APRES
        If nbr_avant > 0 Then
            .TextMatrix(.Rows - 1, GRDC_VALEUR_AVANT) = recup_PSLib(Mid$(STR_GetChamp(str_avant, ";", 0), 2))
            n_avant = n_avant + 1
        End If
        If nbr_ajout > 0 Then
            .TextMatrix(.Rows - 1, GRDC_VALEUR_APRES) = " + " & recup_PSLib(Mid$(STR_GetChamp(str_ajout, ";", 0), 2))
            n_ajout = n_ajout + 1
        ElseIf nbr_suppression > 0 Then
            .TextMatrix(.Rows - 1, GRDC_VALEUR_APRES) = " -  " & recup_PSLib(Mid$(STR_GetChamp(str_suppression, ";", 0), 2))
            n_suppression = n_suppression + 1
        End If
        For I = 1 To nbr - 1
            .AddItem ""
            .Row = .Rows - 1
            .col = GRDC_VALEUR_AVANT
            .CellBackColor = COLOR_VALEUR_AVANT
            .col = GRDC_VALEUR_APRES
            If n_avant < nbr_avant Then
                .TextMatrix(.Rows - 1, GRDC_VALEUR_AVANT) = recup_PSLib(Mid$(STR_GetChamp(str_avant, ";", n_avant), 2))
                n_avant = n_avant + 1
            End If
            If n_ajout < nbr_ajout Then
                .TextMatrix(.Rows - 1, GRDC_VALEUR_APRES) = " + " & recup_PSLib(Mid$(STR_GetChamp(str_ajout, ";", n_ajout), 2))
                .CellBackColor = COLOR_VALEUR_APRES
                n_ajout = n_ajout + 1
            ElseIf n_suppression < nbr_suppression Then
                .TextMatrix(.Rows - 1, GRDC_VALEUR_APRES) = " -  " & recup_PSLib(Mid$(STR_GetChamp(str_suppression, ";", n_suppression), 2))
                .CellBackColor = COLOR_VALEUR_APRES
                n_suppression = n_suppression + 1
            End If
        Next I
    End With

End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = CMD_PRENDE_EN_COMPTE Then
        Call prendre_en_compte
        Call quitter
    Else ' Index = CMD_QUITTER
        Call quitter
    End If

End Sub

Private Sub cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape
            Call quitter
    End Select

End Sub


Private Sub Form_Activate()

    If g_form_active Then Exit Sub

    g_form_active = True
    Call initialiser

End Sub

Private Sub Form_Load()

    g_form_active = False

End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape
            Call quitter
    End Select

End Sub

