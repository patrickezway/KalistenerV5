VERSION 5.00
Begin VB.Form KS_Saisie 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSaisie 
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
      ForeColor       =   &H00008000&
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4575
      Begin VB.TextBox txtSaisie 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblSaisie 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame frmFct 
      BackColor       =   &H00808080&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   2340
      Width           =   5535
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "KS_Saisie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private g_nbbouton As Integer

Private g_form_active As Boolean
Private g_ind_txt_enabled As Integer

Private Function ctrl_saisie_ok(ByVal v_index As Integer, _
                                ByVal v_lafin As Boolean)

    Dim hh As Integer, mm As Integer, pos As Integer
    Dim stmp As String, texte As String, s As String, s2 As String
    
    If Len(txtSaisie(v_index)) = 0 Then
        If Not v_lafin Then
            ctrl_saisie_ok = True
        Else
            If SAIS_Saisie.champs(v_index).facu Then
                ctrl_saisie_ok = True
            Else
                MsgBox "La saisie de cette rubrique est obligatoire", vbOKOnly + vbExclamation, "Saisie Erronnée"
                ctrl_saisie_ok = False
            End If
        End If
        Exit Function
    End If
        
    texte = txtSaisie(v_index).Text
    If Not SAIS_CtrlChamp(texte, SAIS_Saisie.champs(v_index).type) Then
        txtSaisie(v_index).Text = ""
        ctrl_saisie_ok = False
    Else
        txtSaisie(v_index).Text = texte
        ctrl_saisie_ok = True
    End If
    
End Function

Private Sub initialiser()

    Dim hauteur As Integer, nb_champ As Integer, Index As Integer, i As Integer
    Dim max_nbcar_txt As Integer, max_size_lbl As Integer, nb_car As Integer
    Dim max_size_txt As Integer, marge As Integer, intervalle As Integer
    Dim size_txt As Integer
    Dim lg_titre As Long, lg_texte As Long, lg_bouton As Long, lg_bouton1 As Long
    Dim lg As Long, lg_tot As Long
    
    frmSaisie.Caption = SAIS_Saisie.prmfrm.titre
    lg_titre = FRM_LargeurTexte(Me, frmSaisie, SAIS_Saisie.prmfrm.titre) + 255

    'Hauteur de chaque intervalle
    hauteur = 325
    marge = 255
    
    'Nbre de caractères du plus long libellé
    max_size_lbl = 0
    'Nbre de caractères du plus long texte
    max_nbcar_txt = 0
    
    ' Textbox + labels
    nb_champ = 0
    On Error Resume Next
    nb_champ = UBound(SAIS_Saisie.champs) + 1
    On Error GoTo 0
    Index = 1
    g_ind_txt_enabled = -1
    For i = 0 To nb_champ - 1
        If i > 0 Then
            Load lblSaisie(i)
            Load txtSaisie(i)
        End If
        If SAIS_Saisie.champs(i).conversion = SAIS_CONV_SECRET Then
            txtSaisie(i).PasswordChar = "*"
        End If
        lblSaisie(i).visible = True
        lblSaisie(i).Top = 325 + hauteur + ((i * 2) * hauteur)
        lblSaisie(i).Caption = SAIS_Saisie.champs(i).libelle
        lblSaisie(i).TabIndex = Index
        Index = Index + 1
        If FRM_LargeurTexte(Me, lblSaisie(i), lblSaisie(i).Caption) > max_size_lbl Then
            max_size_lbl = FRM_LargeurTexte(Me, lblSaisie(i), lblSaisie(i).Caption)
        End If
        txtSaisie(i).visible = True
        txtSaisie(i).Top = lblSaisie(i).Top
        nb_car = SAIS_Saisie.champs(i).len
        If nb_car > 0 Then
            If g_ind_txt_enabled = -1 Then
                g_ind_txt_enabled = i
            End If
'            txtSaisie(i).BackColor = &HFFFFFF
            txtSaisie(i).Enabled = True
        ElseIf nb_car < 0 Then
            'Texte non modifiable
'            txtSaisie(i).BackColor = &HC0C0C0
            txtSaisie(i).Enabled = False
            nb_car = -nb_car
        Else
            txtSaisie(i).visible = False
        End If
        txtSaisie(i).MaxLength = nb_car
        txtSaisie(i).Text = SAIS_Saisie.champs(i).sval
        If nb_car > max_nbcar_txt Then
            max_nbcar_txt = nb_car
        End If
        txtSaisie(i).TabIndex = Index
        Index = Index + 1
    Next i
    
    'Nbre de caractères max du texte
    If SAIS_Saisie.prmfrm.max_nbcar_visible > 0 And max_nbcar_txt > SAIS_Saisie.prmfrm.max_nbcar_visible Then
        max_nbcar_txt = SAIS_Saisie.prmfrm.max_nbcar_visible
    End If
    'Conversion de caractères en pixels
    max_size_txt = FRM_LargeurTexte(Me, txtSaisie(0), String$(max_nbcar_txt, "M"))
    lg_texte = 255 + max_size_lbl + 255 + max_size_txt + 255
    
    ' Boutons
    On Error Resume Next
    g_nbbouton = UBound(SAIS_Saisie.boutons) + 1
    On Error GoTo 0
    lg_bouton = 0
    For i = 0 To g_nbbouton - 1
        If i > 0 Then Load cmd(i)
        cmd(i).visible = True
        If SAIS_Saisie.boutons(i).image <> "" Then
            cmd(i).Picture = CM_LoadPicture(SAIS_Saisie.boutons(i).image)
            cmd(i).Caption = ""
            cmd(i).ToolTipText = SAIS_Saisie.boutons(i).libelle
        Else
            cmd(i).Picture = LoadPicture("")
            cmd(i).Caption = SAIS_Saisie.boutons(i).libelle
        End If
        If SAIS_Saisie.boutons(i).largeur > 0 Then
            cmd(i).width = SAIS_Saisie.boutons(i).largeur
        End If
        lg_bouton = lg_bouton + cmd(i).width
    Next i
    lg_bouton1 = lg_bouton
    If lg_bouton > 0 Then
        lg_bouton = 255 + lg_bouton + 255 + (g_nbbouton - 1) * 510
    End If
        
    ' Labels et textes alignés
    For i = 0 To nb_champ - 1
        lblSaisie(i).width = max_size_lbl
        lblSaisie(i).left = marge
        txtSaisie(i).left = lblSaisie(i).left + max_size_lbl + 255
        If txtSaisie(i).MaxLength > max_nbcar_txt Then
            size_txt = FRM_LargeurTexte(Me, txtSaisie(0), String$(max_nbcar_txt, "M"))
        Else
            size_txt = FRM_LargeurTexte(Me, txtSaisie(0), String$(txtSaisie(i).MaxLength, "M"))
        End If
        txtSaisie(i).width = size_txt
    Next i

    ' Reglage largeur
    lg = lg_titre
    If lg < lg_bouton Then
        lg = lg_bouton
    End If
    If lg < lg_texte Then
        lg = lg_texte
    End If
    lg_bouton = lg + 512
    lg_tot = lg + 512
    frmSaisie.width = lg_tot
    frmFct.width = lg_tot
    Me.width = lg_tot
    
    ' Positionnement des boutons
    If g_nbbouton = 1 Then
        cmd(0).left = (frmFct.width - 510 - cmd(0).width) / 2
    Else
        intervalle = (frmFct.width - 510 - lg_bouton1) / (g_nbbouton - 1)
        left = 255
        For i = 0 To g_nbbouton - 1
            cmd(i).left = left
            left = left + cmd(i).width + intervalle
        Next i
    End If
        
    ' Calcul de la hauteur
    frmSaisie.Height = 255 + (2 * nb_champ * hauteur) + 300
    frmSaisie.Top = 0
    frmSaisie.ZOrder 0
    frmFct.Top = frmSaisie.Top + frmSaisie.Height - 150
    Me.Height = frmFct.Top + frmFct.Height + 300
    
    If SAIS_Saisie.prmfrm.x > 0 Then
        Me.left = SAIS_Saisie.prmfrm.x
    Else
        Me.left = (Screen.width - Me.width) / 2
    End If
    If SAIS_Saisie.prmfrm.Y > 0 Then
        Me.Top = SAIS_Saisie.prmfrm.Y
    Else
        Me.Top = (Screen.Height - Me.Height) / 2
    End If
    frmSaisie.left = 0
    
End Sub

Private Sub quitter(ByVal Index As Integer)

    Dim modif As Boolean
    Dim i As Integer
    
    Select Case Index
    Case 0
        modif = False
        For i = 0 To UBound(SAIS_Saisie.champs())
            If Not ctrl_saisie_ok(i, True) Then
                txtSaisie(i).SetFocus
                Exit Sub
            End If
            If SAIS_Saisie.champs(i).sval <> txtSaisie(i).Text Then modif = True
            SAIS_Saisie.champs(i).sval = txtSaisie(i).Text
        Next i
        If modif Then
            SAIS_Saisie.retour = SAIS_RET_MODIF
        Else
            SAIS_Saisie.retour = SAIS_RET_NOMODIF
        End If
    Case Else
        SAIS_Saisie.retour = Index
    End Select
    
    If SAIS_Saisie.prmfrm.reste_chargée Then
        Me.Hide
    Else
        Unload Me
    End If
 
End Sub

Private Sub cmd_Click(Index As Integer)

    Call quitter(Index)
    
End Sub

Private Sub Form_Activate()
    
    Dim i As Integer
    
    If Not g_form_active Then
        g_form_active = True
        Call FRM_ResizeForm(Me, 0, 0)
        Call initialiser
    Else
        For i = 0 To UBound(SAIS_Saisie.champs())
            txtSaisie(i).Text = SAIS_Saisie.champs(i).sval
        Next i
    End If
    If Len(txtSaisie(g_ind_txt_enabled)) > 0 Then
        txtSaisie(g_ind_txt_enabled).SelLength = Len(txtSaisie(g_ind_txt_enabled).Text)
    End If
    txtSaisie(g_ind_txt_enabled).SetFocus
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim nomchm As String, nomtopic As String
    Dim i As Integer
    
    If Shift = vbAltMask Then
        For i = 0 To g_nbbouton - 1
            If KeyCode = SAIS_Saisie.boutons(i).raccourci_alt Then
                KeyCode = 0
                Call quitter(i)
                Exit Sub
            End If
        Next i
        If KeyCode = vbKeyH Then
            KeyCode = 0
            If SAIS_Saisie.prmfrm.nomhelp <> "" Then
                If STR_GetNbchamp(SAIS_Saisie.prmfrm.nomhelp, ";") = 1 Then
                    nomchm = SAIS_Saisie.prmfrm.nomhelp
                    nomtopic = ""
                Else
                    nomchm = STR_GetChamp(SAIS_Saisie.prmfrm.nomhelp, ";", 0)
                    nomtopic = STR_GetChamp(SAIS_Saisie.prmfrm.nomhelp, ";", 1)
                End If
                Call HtmlHelp(0, nomchm, HH_DISPLAY_TOPIC, nomtopic)
            End If
        End If
    Else
        For i = 0 To g_nbbouton - 1
            If KeyCode = SAIS_Saisie.boutons(i).raccourci_touche Then
                KeyCode = 0
                Call quitter(i)
                Exit Sub
            End If
        Next i
    End If

End Sub

Private Sub Form_Load()

    g_form_active = False
    Call FRM_ResizeForm(Me, 0, 0)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Call quitter(cmd.Count - 1)
    End If
    
End Sub

Private Sub txtSaisie_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If UBound(SAIS_Saisie.champs()) = 0 Or _
            (SAIS_Saisie.champs(Index).validationdirecte And _
            txtSaisie(Index).Text <> "") Then
            Call quitter(0)
        Else
            SendKeys "{TAB}"
        End If
        Exit Sub
    End If
    
    If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub
    
    Select Case SAIS_Saisie.champs(Index).type
    Case SAIS_TYP_JOUR_SEMAINE
        If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
            If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    Case SAIS_TYP_HEURE
        If KeyAscii = Asc(":") Then Exit Sub
        If KeyAscii = Asc("h") Then Exit Sub
        If KeyAscii = Asc("H") Then Exit Sub
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Case SAIS_TYP_DATE
        If KeyAscii = Asc("/") Then Exit Sub
        If KeyAscii = Asc("+") Then Exit Sub
        If KeyAscii = Asc("-") Then Exit Sub
        If KeyAscii = Asc("j") Then Exit Sub
        If KeyAscii = Asc("J") Then Exit Sub
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Case SAIS_TYP_ENTIER
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Case SAIS_TYP_ENTIER_NEG
        If KeyAscii = Asc("-") Then Exit Sub
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Case SAIS_TYP_LETTRE
        If Not STR_EstAlpha(Chr(KeyAscii)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Case SAIS_TYP_LETTRE_PONCT
        If Not STR_EstAlpha(Chr(KeyAscii)) Then
            If Not STR_EstPonctuation(Chr(KeyAscii)) Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    Case SAIS_TYP_PERIODE
        If KeyAscii = Asc("j") Then Exit Sub
        If KeyAscii = Asc("J") Then Exit Sub
        If KeyAscii = Asc("s") Then Exit Sub
        If KeyAscii = Asc("S") Then Exit Sub
        If KeyAscii = Asc("m") Then Exit Sub
        If KeyAscii = Asc("M") Then Exit Sub
        If KeyAscii = Asc("a") Then Exit Sub
        If KeyAscii = Asc("A") Then Exit Sub
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    Case SAIS_TYP_CAR_PARTICULIER
        If InStr(SAIS_Saisie.champs(Index).chaine_type, UCase(Chr(KeyAscii))) <= 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    End Select
    
    If SAIS_Saisie.champs(Index).conversion = SAIS_CONV_MAJUSCULE Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    ElseIf SAIS_Saisie.champs(Index).conversion = SAIS_CONV_MINUSCULE Then
        KeyAscii = Asc(LCase(Chr(KeyAscii)))
    End If

End Sub

Private Sub txtSaisie_LostFocus(Index As Integer)

    Dim i As Integer, n As Integer
    
    If g_nbbouton = 1 Then
        n = 0
    Else
        n = g_nbbouton - 1
    End If
    For i = 0 To n
        If cmd(i).tag = "" Then
            If Not ctrl_saisie_ok(Index, False) Then
                txtSaisie(Index).SetFocus
            End If
        End If
    Next i
    
End Sub


