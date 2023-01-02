VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PrmTypeInfSupPersonne 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Informations supplémentaires"
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
      Height          =   4335
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton cmd 
         Height          =   330
         Index           =   3
         Left            =   7695
         Picture         =   "PrmTypeInfSupPersonne.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3765
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         Height          =   315
         Index           =   2
         Left            =   7695
         Picture         =   "PrmTypeInfSupPersonne.frx":0447
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   3375
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5953
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   8454143
         BackColorBkg    =   16777215
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   885
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   8295
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmTypeInfSupPersonne.frx":089E
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
         Left            =   720
         Picture         =   "PrmTypeInfSupPersonne.frx":0DFA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
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
         Left            =   6960
         Picture         =   "PrmTypeInfSupPersonne.frx":1363
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "PrmTypeInfSupPersonne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Index des boutons
Private Const CMD_OK = 0
Private Const CMD_QUITTER = 1
Private Const CMD_PLUS_INFO_SUPPL = 2
Private Const CMD_MOINS_INFO_SUPPL = 3

' Index des colonnes du Grid
Private Const GRDC_TIS_NUM = 0
Private Const GRDC_TIS_LIBELLE = 1
Private Const GRDC_TIS_VALEUR = 2

' indiquer si la forme a déjà été activée
Private g_form_active As Boolean
' Numéro de la personne
Private g_numutil As Long

Public Sub AppelFrm(ByVal v_numutil As Long)

    g_numutil = v_numutil

    Me.Show 1

End Sub

Private Function ajouter_info_suppl() As Integer

    Dim sql As String
    Dim i As Integer, nbr_infosuppl As Integer
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

    nbr_infosuppl = 0

    While Not rs.EOF
        ' chercher si on a pas déjà ce type
        With grd
            For i = 1 To .Rows - 1
                If rs("KB_TisNum").Value = .TextMatrix(i, GRDC_TIS_NUM) Then
                    GoTo lab_suivant
                End If
            Next i
        End With
        ' TIS n'existe pas
        Call CL_AddLigne(rs("KB_TisLibelle").Value, rs("KB_TisNum").Value, "", False)
        nbr_infosuppl = nbr_infosuppl + 1
lab_suivant:
        rs.MoveNext
    Wend
    rs.Close

    ' verification si on peut ajouter
'    If grd.Rows - 1 = nbr_infosuppl Then
    If nbr_infosuppl = 0 Then
        MsgBox "Vous ne pouvez pas ajouter d'informations supplémentaires pour cette personne." & vbInformation + vbOKOnly
        GoTo lab_erreur
    End If

    Call CL_InitTitreHelp("Liste des types de coordonnée", p_chemin_appli + "\help\kalidoc.chm" & ";" & "dico_d_fonction.htm")
    Call CL_InitTaille(0, -15)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)

    ChoixListe.Show 1

    ' Quitter
    If CL_liste.retour = 1 Then
        GoTo lab_erreur
    End If

    ' => CL_liste.retour = 0
    ' ajouter
    With grd
        .AddItem ""
        .TextMatrix(.Rows - 1, GRDC_TIS_NUM) = CL_liste.lignes(CL_liste.pointeur).num
        .TextMatrix(.Rows - 1, GRDC_TIS_LIBELLE) = CL_liste.lignes(CL_liste.pointeur).texte
        .TextMatrix(.Rows - 1, GRDC_TIS_VALEUR) = ""
        cmd(CMD_OK).Enabled = True
    End With

lab_enreg:
    Call FRM_ResizeForm(Me, Me.width, Me.Height)
    ajouter_info_suppl = P_OUI
    Exit Function

lab_erreur:
    Call FRM_ResizeForm(Me, Me.width, Me.Height)
    ajouter_info_suppl = P_NON

End Function

Private Sub initialiser()

    cmd(CMD_OK).Enabled = False

    With grd
        .Rows = 1
        .FormatString = "tis_num|Libellé|Valeur"
        .col = GRDC_TIS_NUM
        .CellFontBold = True
        .col = GRDC_TIS_LIBELLE
        .CellFontBold = True
        .col = GRDC_TIS_VALEUR
        .CellFontBold = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionByRow
        .ColWidth(GRDC_TIS_NUM) = 0
        .ColWidth(GRDC_TIS_LIBELLE) = (1 / 2) * .width
        .ColWidth(GRDC_TIS_VALEUR) = (1 / 2) * .width - 110
        .ColAlignment(GRDC_TIS_LIBELLE) = flexAlignCenterCenter
        .ColAlignment(GRDC_TIS_VALEUR) = flexAlignCenterCenter
    End With

    ' selectionner un TIS
    If remplir_grid() <> P_OK Then
        Call quitter(True)
    End If

End Sub

Private Sub enlever_info_suppl()

    With grd
        If .Rows = 1 Then Exit Sub
        If .Rows = 2 Then
            .Rows = 1
            Exit Sub
        End If
        Call .RemoveItem(.Row)
    End With

    cmd(CMD_OK).Enabled = True

End Sub

Private Sub imprimer()
MsgBox "imprimer"
End Sub

Private Sub quitter(ByVal v_bforce As Boolean)

    If v_bforce Then
        Unload Me
        Exit Sub
    Else
        If cmd(CMD_OK).Enabled Then ' y'a eu des changements
            If MsgBox("Etes vous sûr de vouloir quitter le paramétrage des " & _
                    " informations supplémentaires ?", vbQuestion + vbYesNo, _
                    "Quitter") = vbYes Then
                Unload Me
                Exit Sub
            End If
        Else
            Unload Me
            Exit Sub
        End If
    End If

End Sub

Private Function remplir_grid() As Integer
' remplir le grid des personnes pouvant créer/modifier ce TIS
    Dim sql As String
    Dim rs As rdoResultset
    Dim u_num As Long
    Dim nbr_tis As Integer, i As Integer

    With grd
        .Rows = 1
        sql = "SELECT * FROM InfoSupplEntite, KB_TypeInfoSuppl" _
            & " WHERE ISE_TisNum=KB_TisNum AND ISE_Type='U' AND ISE_TypeNum=" & g_numutil _
            & " ORDER BY KB_TisLibelle"
        If Odbc_SelectV(sql, rs) = P_ERREUR Then
            GoTo lab_erreur
        End If
        ' parcourir les info suppl de cette personne
        While Not rs.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, GRDC_TIS_NUM) = rs("KB_TisNum").Value
            .TextMatrix(.Rows - 1, GRDC_TIS_LIBELLE) = rs("KB_TisLibelle").Value
            .TextMatrix(.Rows - 1, GRDC_TIS_VALEUR) = rs("ISE_Valeur").Value
            rs.MoveNext
        Wend
        rs.Close
    End With

    remplir_grid = P_OK
    Exit Function

lab_erreur:
    remplir_grid = P_ERREUR

End Function

Private Function valider() As Integer

    Dim rs As rdoResultset
    Dim lng As Long
    Dim i As Integer

    ' vérifier avant d'enregister
    If verifier_champs = P_OK Then
        ' supprimer les anciennes infos suppl
        If Odbc_Delete("InfoSupplEntite", _
                    "ISE_Num", _
                    "WHERE ISE_Type='U' AND ISE_TypeNum=" & g_numutil, _
                    lng) = P_ERREUR Then
            GoTo lab_erreur
        End If
        ' Enregistrer les nouvelles infos
        With grd
            For i = 1 To .Rows - 1
                If Odbc_AddNew("InfoSupplEntite", "ISE_Num", "ISE_Seq", False, lng, _
                                "ISE_TisNum", .TextMatrix(i, GRDC_TIS_NUM), _
                                "ISE_Type", "U", _
                                "ISE_TypeNum", g_numutil, _
                                "ISE_Valeur", .TextMatrix(i, GRDC_TIS_VALEUR)) = P_ERREUR Then
                    GoTo lab_erreur
                End If
            Next i
        End With
    Else
        GoTo lab_erreur
    End If

    ' tout est bien terminé
    valider = P_OK
    Exit Function

lab_erreur: ' rollback ?
    valider = P_ERREUR

End Function

Private Function verifier_champs() As Integer

    Dim i As Integer

    With grd
        ' a-t-on des valeurs valides ?
        For i = 1 To .Rows - 1
            If .TextMatrix(i, GRDC_TIS_VALEUR) = "" Then
                MsgBox "Veuillez saisir une valeur valide pour l'information:" _
                        & vbCrLf & vbCrLf & .TextMatrix(i, GRDC_TIS_LIBELLE), _
                        vbOKOnly + vbCritical, "Valeur non valide"
                GoTo lab_erreur
            End If
        Next i
    End With
    ' tout est OK !
    verifier_champs = P_OK
    Exit Function

lab_erreur:
    verifier_champs = P_ERREUR

End Function

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case CMD_OK
            If valider = P_OK Then Call quitter(True)
        Case CMD_QUITTER
            Call quitter(False)
        Case CMD_PLUS_INFO_SUPPL
            Call ajouter_info_suppl
        Case CMD_MOINS_INFO_SUPPL
            Call enlever_info_suppl
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
        Case vbKeyF1
            KeyAscii = 0
            If cmd(CMD_OK).Enabled Then
                If valider = P_OK Then Call quitter(True)
            End If
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

Private Sub grd_Click()

    Dim valeur_retournee As String
    Dim i As Integer, mouse_row As Integer, mouse_col As Integer, _
        mon_left As Integer, mon_top As Integer

    With grd
        If .Rows = 1 Then Exit Sub

        mouse_row = .MouseRow
        mouse_col = .MouseCol

        ' vérifications
        If mouse_row = 0 Then Exit Sub
        If mouse_col <> GRDC_TIS_VALEUR Then Exit Sub
        ' demander la valeur
        mon_top = ((Me.Height - 2000) / 2) + Me.Top
        mon_left = ((Me.width - 5000) / 2) + Me.left
        valeur_retournee = InputBox("Veuillez saisir la valeur de : " & vbCrLf & vbCrLf _
                        & .TextMatrix(mouse_row, GRDC_TIS_LIBELLE), _
                        "Valeur de l'information", _
                        .TextMatrix(mouse_row, mouse_col), _
                        mon_left, mon_top)
        If StrPtr(valeur_retournee) = 0 Then
            ' on a annulé
        Else ' on a validé (même pour les valeurs vides)
            .col = mouse_col
            .Row = mouse_row
            .Text = valeur_retournee
            cmd(CMD_OK).Enabled = True
        End If
    End With

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
            If valider = P_OK Then Call quitter(True)
    End Select

End Sub
