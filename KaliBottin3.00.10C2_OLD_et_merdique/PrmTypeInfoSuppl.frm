VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PrmTypeInfoSuppl 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Type d'information supplémentaire"
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
      Height          =   4215
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6735
      Begin VB.CheckBox chk 
         Caption         =   "Check1"
         Height          =   195
         Left            =   3360
         TabIndex        =   13
         Top             =   1800
         Width           =   210
      End
      Begin VB.CommandButton cmd 
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
         Index           =   4
         Left            =   6150
         Picture         =   "PrmTypeInfoSuppl.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton cmd 
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
         Index           =   5
         Left            =   6150
         Picture         =   "PrmTypeInfoSuppl.frx":0457
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3660
         Width           =   360
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   1215
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2143
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         ForeColor       =   -2147483647
         BackColorBkg    =   16777215
         GridColorFixed  =   16777215
         GridLines       =   0
         GridLinesFixed  =   0
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   960
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1200
         Width           =   5055
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   960
         MaxLength       =   10
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label LblNum 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6120
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Géré dans le fichiers d'importation"
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
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Qui peut créer/modifier"
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
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Libellé"
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
         TabIndex        =   5
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Code"
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
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   1005
      Left            =   0
      TabIndex        =   0
      Top             =   4080
      Width           =   6740
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Height          =   510
         Index           =   3
         Left            =   3090
         Picture         =   "PrmTypeInfoSuppl.frx":089E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Supprimer ce type d'information"
         Top             =   290
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmTypeInfoSuppl.frx":0E33
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
         Left            =   360
         Picture         =   "PrmTypeInfoSuppl.frx":138F
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Enregistrer"
         Top             =   290
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
         Left            =   5850
         Picture         =   "PrmTypeInfoSuppl.frx":18F8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Quitter sans enregistrer"
         Top             =   290
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "PrmTypeInfoSuppl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' l'index des boutons
Private Const CMD_OK = 0
Private Const CMD_QUITTER = 1
Private Const CMD_IMPRIMER = 2
Private Const CMD_SUPPRIMER = 3
Private Const CMD_PLUS_REPS = 4
Private Const CMD_MOINS_RESP = 5

' L'index des textbox
Private Const TXT_CODE = 0
Private Const TXT_LIBELLE = 1

' Index des objets LBL
Private Const LBL_CODE = 0
Private Const LBL_LIBELLE = 1
Private Const LBL_CREER_MODIFIER = 2

' Index des colonnes du grid
Private Const GRDC_U_NUM = 0
Private Const GRDC_U_NOM = 1

' indiquer si la forme a déjà été activée
Private g_form_active As Boolean

Private g_tis_num As Long
Private g_sret As String
Private g_mode_direct As Boolean

Private Function afficher_tis(ByVal v_tisnum As Long) As Integer

    Dim sql As String
    Dim rs As rdoResultset

    g_tis_num = v_tisnum

    grd.Rows = 0

    If (v_tisnum > 0) Then ' mode modif
        sql = "SELECT * FROM KB_TypeInfoSuppl WHERE KB_TisNum=" & v_tisnum
        If Odbc_Select(sql, rs) = P_ERREUR Then
            GoTo lab_erreur
        End If
        If Not rs.EOF Then
            txt(TXT_CODE).Text = rs("KB_TisCode").Value
            txt(TXT_LIBELLE).Text = rs("KB_TisLibelle").Value
            LblNum.Caption = v_tisnum
            chk.Value = IIf(rs("KB_TisImport").Value, vbChecked, vbUnchecked)
            If remplir_grid(rs("KB_TisLstUtilModif").Value) = P_ERREUR Then
                GoTo lab_erreur
            End If
        End If
    Else ' mode création
        txt(TXT_CODE).Text = ""
        txt(TXT_LIBELLE).Text = ""
        chk.Value = vbUnchecked
    End If

    cmd(CMD_OK).Enabled = False
    txt(TXT_CODE).SetFocus
    afficher_tis = P_OK
    Exit Function

lab_erreur:
    afficher_tis = P_ERREUR
End Function

Private Sub ajouter_responsable()
    
    Dim sret As String, sql As String, nom As String, prenom As String
    Dim numutil As Long
    Dim I As Integer
    Dim frm As Form

lab_afficher:
    p_siz_tblu = -1
    Set frm = ChoixUtilisateur
    sret = ChoixUtilisateur.AppelFrm("Qui peut créer/modifier", _
                                    "", _
                                    False, _
                                    False, _
                                    "")
    Set frm = Nothing
    If sret = "" Then
        Exit Sub
    End If
    numutil = p_tblu_sel(0)
    ' vérifier si la personne n'a pas déjà été choisie
    With grd
        For I = 0 To .Rows - 1
            If .TextMatrix(I, GRDC_U_NUM) = numutil Then
                ' personne existante => demander
                If MsgBox("La personne sélectionnée existe déjà dans la liste." & vbCrLf & vbCrLf & _
                            "Voulez-vous en choisir une autre ?", vbInformation + vbYesNo, "Attention") = vbYes Then
                    GoTo lab_afficher
                Else
                    Exit Sub
                End If
            End If
        Next I
    End With
    sql = "SELECT U_Nom, U_Prenom FROM Utilisateur WHERE U_kb_actif=True AND U_Num=" & numutil
    If Odbc_RecupVal(sql, nom, prenom) = P_ERREUR Then
        Exit Sub
    End If
    ' ajouter au tableau
    With grd
        .AddItem ""
        .TextMatrix(.Rows - 1, GRDC_U_NUM) = numutil
        .TextMatrix(.Rows - 1, GRDC_U_NOM) = nom + " " + prenom
    End With

    cmd(CMD_OK).Enabled = True

End Sub

Public Function AppelFrm(ByVal v_mode_param As Integer, ByVal v_tis_num As Long) As String

    If v_tis_num >= 0 Then
        ' mode de création direct
        g_mode_direct = True
    Else
        g_mode_direct = False
    End If

    g_tis_num = v_tis_num

    Me.Show 1

    AppelFrm = g_sret

End Function

Private Function choisir_tis() As Integer

    Dim sql As String
    Dim rs As rdoResultset

    Call FRM_ResizeForm(Me, 0, 0)
    
lab_affiche:
    Call CL_Init
    'Choix du TIS
    sql = "SELECT * FROM KB_TypeInfoSuppl" _
        & " ORDER BY KB_TisLibelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_erreur
    End If
    Call CL_AddLigne("<Nouveau>", 0, "", False)

    While Not rs.EOF
        Call CL_AddLigne(rs("KB_TisLibelle").Value, rs("KB_TisNum").Value, "", False)
        rs.MoveNext
    Wend
    rs.Close

    Call CL_InitTitreHelp("Liste des types d'informations supplémentaires", "")
    Call CL_InitTaille(0, -15)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)

    ChoixListe.Show 1

    ' Quitter
    If CL_liste.retour = 1 Then
        choisir_tis = P_NON
        Exit Function
    End If

    ' => CL_liste.retour = 0
    If afficher_tis(CL_liste.lignes(CL_liste.pointeur).num) = P_ERREUR Then
        GoTo lab_erreur
    End If
    
    Call FRM_ResizeForm(Me, Me.width, Me.Height)
    
    choisir_tis = P_OUI
    Exit Function

lab_erreur:
    choisir_tis = P_NON

End Function

Private Sub enlever_responsable()

    Dim row_en_cours As Integer

    With grd
        If .Rows = 0 Then Exit Sub
        row_en_cours = .Row
        If .Rows = 1 Then
            .Rows = 0
        Else
            Call .RemoveItem(row_en_cours)
        End If
        If .Rows = 0 Then .Enabled = False
    End With

    cmd(CMD_OK).Enabled = True

End Sub

Private Sub initialiser()

    cmd(CMD_OK).Enabled = False

    With grd
        .Rows = 0
        .FormatString = "u_num|Nom"
        .col = GRDC_U_NOM
        .CellFontBold = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionByRow
        .ColWidth(GRDC_U_NUM) = 0
        .ColWidth(GRDC_U_NOM) = .width - 100
    End With

    If g_mode_direct Then
        ' accéder directement au paramétrage
        Call afficher_tis(g_tis_num)
    Else
        ' selectionner un TIS
        If choisir_tis() <> P_OUI Then
            'Call quitter(True)
            Unload Me
            Exit Sub
        End If
    End If

End Sub

Private Function get_tis_lstUtilModif(ByRef r_str) As Integer

    Dim str As String
    Dim I As Integer

    With grd
        For I = 0 To .Rows - 1
            r_str = r_str & "U" & .TextMatrix(I, GRDC_U_NUM) & ";"
        Next I
    End With

    get_tis_lstUtilModif = P_OK

End Function

Private Sub quitter(ByVal v_bforce As Boolean)

    If v_bforce Then
'        If g_mode_direct Then
'            Unload Me
'            Exit Sub
'        Else
'            If choisir_tis() <> P_OUI Then
'                Unload Me
'                Exit Sub
'            End If
'        End If
        If choisir_tis() <> P_OUI Then Unload Me
'        Unload Me
        Exit Sub
    Else
        If cmd(CMD_OK).Enabled Then ' y'a eu des changements
            If MsgBox("Des modifications ont été effectuées !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", _
                        vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                'Unload Me
                Exit Sub
            End If
        End If
    End If
    If g_mode_direct Then
        Unload Me
    Else
        If choisir_tis() <> P_OUI Then Unload Me
    End If

End Sub

Private Function remplir_grid(ByVal v_tis_lstutilmodif As String) As Integer
' remplir le grid des personnes pouvant créer/modifier ce TIS
    Dim sql As String
    Dim rs As rdoResultset
    Dim u_num As Long
    Dim nbr_tis As Integer, I As Integer

    With grd
        ' parcourir les responsables
        nbr_tis = STR_GetNbchamp(v_tis_lstutilmodif, ";")
        For I = 0 To nbr_tis - 1
            u_num = Mid$(STR_GetChamp(v_tis_lstutilmodif, ";", I), 2)
            sql = "SELECT U_Nom, U_Prenom FROM Utilisateur" & _
                    " Where U_kb_actif=True AND U_Num=" & u_num
            If Odbc_Select(sql, rs) = P_ERREUR Then
                GoTo lab_erreur
            End If
            ' ajouter cette personne
            .AddItem ""
            .TextMatrix(.Rows - 1, GRDC_U_NUM) = u_num
            .TextMatrix(.Rows - 1, GRDC_U_NOM) = rs("U_Nom").Value + " " + rs("U_Prenom").Value
            rs.Close
        Next I
    End With

    remplir_grid = P_OK
    Exit Function

lab_erreur:
    remplir_grid = P_ERREUR

End Function

Private Sub supprimer()

    Dim sql As String, s As String
    Dim I As Integer, n As Integer
    Dim lnb As Long, lnb_ise As Long, lnb_prm As Long
    Dim sval As Variant, snew_val As Variant, message As Variant
    Dim rs As rdoResultset
    
    ' poser les questions
    ' a-t-on des personnes ayant des infosuppl de ce type ?
    sql = "SELECT COUNT(*) FROM InfoSupplEntite WHERE ISE_TisNum=" & g_tis_num
    If Odbc_Count(sql, lnb_ise) = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    ' et le paramétrage général + importation ?
    sql = "SELECT COUNT(*) FROM PrmGenB WHERE PGB_LstPosInfoAutre LIKE '%I" & g_tis_num & ";%'"
    If Odbc_Count(sql, lnb_prm) = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    If lnb_ise > 0 Then ' des personnes avec des infos
        message = "- Il existe des personnes ayant des informations de ce type !"
    End If
    If lnb_prm > 0 Then ' le fichier d'import
        message = message & vbCrLf & "- Le fichier d'importation utilise ce type d'information supplémentaire !"
    End If
    ' on a une question à posser
    If Len(message) > 0 Then
        If MsgBox(message & vbCrLf & vbCrLf & _
                "Confirmez-vous quand même la suppression de ce type d'information ?", _
                vbYesNo + vbQuestion, "Attention") = vbNo Then
            Exit Sub
        End If
    Else
        ' Simple confirmation
        If MsgBox("Confirmez-vous la suppression de ce type de coordonnée ?", vbQuestion + vbYesNo + vbDefaultButton2, "") = vbNo Then
            Exit Sub
        End If
    End If
    
    Call Odbc_BeginTrans
    
    ' Suppression du type d'information
    If Odbc_Delete("KB_TypeInfoSuppl", "KB_TisNum", "WHERE KB_TisNum=" & g_tis_num, lnb) = P_ERREUR Then
        GoTo lab_erreur
    End If
    If lnb_ise > 0 Then
        ' Suppression au niveau des utilisateurs
        If Odbc_Delete("InfoSupplEntite", "ISE_TisNum", "WHERE ISE_TisNum=" & g_tis_num, lnb) = P_ERREUR Then
            GoTo lab_erreur
        End If
    End If
    If lnb_prm > 0 Then
        ' Supprime ce type d'info dans la liste du prm général
        snew_val = ""
        sql = "SELECT PGB_LstPosInfoAutre FROM PrmGenB"
        Call Odbc_RecupVal(sql, sval)
        n = STR_GetNbchamp(sval, "|")
        For I = 0 To n - 1
            s = STR_GetChamp(sval, "|", I)
            If InStr(s, "I" & g_tis_num & ";") = 0 Then
                snew_val = snew_val + s + "|"
            End If
        Next I
        Call Odbc_Update("PrmGenB", "PGB_Num", "", "PGB_LstPosInfoAutre", snew_val)
    End If
    
    Call Odbc_CommitTrans
    
    ' quitter
    Call quitter(True)

    Exit Sub
    
lab_erreur:
    Call Odbc_RollbackTrans
    
End Sub

Private Function valider() As Integer

    Dim tis_lstUtilModif As String
    Dim tis_import As Boolean
    Dim lng As Long

    ' vérifier avant d'enregister
    If verifier_champs = P_OK Then
        If get_tis_lstUtilModif(tis_lstUtilModif) <> P_OK Then
            GoTo lab_erreur
        End If

        tis_import = (chk.Value = vbChecked)

        If (g_tis_num > 0) Then ' modif
            If Odbc_Update("KB_TypeInfoSuppl", "KB_TisNum", _
                        "WHERE KB_TisNum=" & g_tis_num, _
                        "KB_TisCode", txt(TXT_CODE).Text, _
                        "KB_TisLibelle", txt(TXT_LIBELLE).Text, _
                        "KB_TisLstUtilModif", tis_lstUtilModif, _
                        "KB_TisImport", tis_import) = P_ERREUR Then
                GoTo lab_erreur
            End If
        Else ' nouveau TIS
            If Odbc_AddNew("KB_TypeInfoSuppl", "KB_TisNum", "KB_Tis_Seq", False, lng, _
                        "KB_TisCode", txt(TXT_CODE).Text, _
                        "KB_TisLibelle", txt(TXT_LIBELLE).Text, _
                        "KB_TisLstUtilModif", tis_lstUtilModif, _
                        "KB_TisImport", tis_import) = P_ERREUR Then
                GoTo lab_erreur
            End If
        End If
    Else
        GoTo lab_erreur
    End If
    ' tout est bien terminé
    valider = P_OK
    Exit Function

lab_erreur: ' rollback
    valider = P_ERREUR

End Function

Private Function verifier_champs() As Integer

    Dim str As String

    str = ""
    ' le code
    If txt(TXT_CODE).Text = "" Then
        str = "* Le CODE est une valeur obligatoire"
    End If
    ' le libellé
    If txt(TXT_LIBELLE).Text = "" Then
        str = str & vbCrLf & "* Le LIBELLE est une valeur obligatoire"
    End If
    ' Les responsables
    If grd.Rows = 0 Then
        str = str & vbCrLf & "* Vous n'avez pas spécifié de personne pouvant créer/modifier cette information."
    End If

    If str <> "" Then
        MsgBox "Les erreurs suivantes ont été enregistrées :" & vbCrLf & _
                str & vbCrLf _
                , vbCritical + vbOKOnly, "Erreur de données"
        GoTo lab_erreur
    End If

    verifier_champs = P_OK
    Exit Function

lab_erreur:
    verifier_champs = P_ERREUR

End Function

Private Sub chk_Click()

    cmd(CMD_OK).Enabled = True

End Sub

Private Sub chk_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys "{TAB}"
        Case vbKeyEscape
            Call quitter(False)
        Case vbKeyF1
            If valider = P_OK Then Call quitter(True)
    End Select

    KeyCode = 0

End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case CMD_OK
            If valider = P_OK Then Call quitter(True)
        Case CMD_QUITTER
            Call quitter(False)
        Case CMD_SUPPRIMER
            Call supprimer
        Case CMD_PLUS_REPS
            Call ajouter_responsable
        Case CMD_MOINS_RESP
            Call enlever_responsable
    End Select

End Sub

Private Sub cmd_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape
            Call quitter(False)
        Case vbKeyF1
            If valider = P_OK Then Call quitter(True)
    End Select

    KeyCode = 0

End Sub

Private Sub Form_Activate()

    If Not g_form_active Then
        Call initialiser
        g_form_active = True
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyEscape
            KeyAscii = 0
            Call quitter(False)
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{TAB}"
        'Case vbKeyF1
        '    KeyAscii = 0
        '    If cmd(CMD_OK).Enabled Then
        '        If valider = P_OK Then Call quitter(True)
        '    End If
    End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF1
            KeyCode = 0
            If cmd(CMD_OK).Enabled Then
                If valider = P_OK Then Call quitter(True)
            End If
    End Select

End Sub

Private Sub Form_Load()

    g_form_active = False

End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyEscape
            KeyAscii = 0
            Call quitter(False)
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys "{TAB}"
        Case vbKeyF1
            KeyAscii = 0
            If valider() = P_OK Then Call quitter(True)
    End Select

End Sub

Private Sub txt_Change(Index As Integer)

    cmd(CMD_OK).Enabled = True

End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)

'    Select Case KeyAscii
'        Case vbKeyEscape
'            KeyAscii = 0
'            Call quitter(False)
'        Case vbKeyReturn
'            KeyAscii = 0
'            SendKeys "{TAB}"
'        Case vbKeyF1
'            KeyAscii = 0
'            If cmd(CMD_OK).Enabled Then
'                If valider = P_OK Then Call quitter(True)
'            End If
'    End Select

End Sub

Private Sub txt_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

'    Select Case KeyCode
'        Case vbKeyF1
'            KeyCode = 0
'            If valider = P_OK Then Call quitter(True)
'    End Select

End Sub
