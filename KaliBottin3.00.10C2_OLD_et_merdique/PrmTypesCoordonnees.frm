VERSION 5.00
Begin VB.Form PrmTypeCoordonnees 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4740
   ClientLeft      =   4350
   ClientTop       =   3315
   ClientWidth     =   6825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Type de coordonnée"
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
      Height          =   4065
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6825
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "alimentée par"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   360
         MaskColor       =   &H0080C0FF&
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Quitter sans enregistrer"
         Top             =   3600
         UseMaskColor    =   -1  'True
         Width           =   1755
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Les valeurs sont une liste de choix"
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
         Left            =   315
         TabIndex        =   14
         Top             =   3255
         Width           =   3570
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   1
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1035
         Width           =   4305
      End
      Begin VB.TextBox txt 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   2
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1575
         Width           =   675
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   0
         Left            =   1830
         MaxLength       =   10
         TabIndex        =   0
         Top             =   540
         Width           =   1995
      End
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
         Height          =   435
         Index           =   3
         Left            =   3420
         Picture         =   "PrmTypesCoordonnees.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Choisir une image"
         Top             =   2280
         Width           =   405
      End
      Begin VB.PictureBox pct 
         Height          =   735
         Left            =   2280
         ScaleHeight     =   675
         ScaleWidth      =   825
         TabIndex        =   9
         ToolTipText     =   "Image associée à ce type de coordonnée"
         Top             =   2130
         Width           =   885
      End
      Begin VB.Label lblInfSup 
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
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Top             =   3600
         Width           =   4335
      End
      Begin VB.Label LblNum 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6240
         TabIndex        =   15
         Top             =   240
         Width           =   495
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
         Height          =   285
         Index           =   1
         Left            =   390
         TabIndex        =   13
         Top             =   1125
         Width           =   735
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre max par personne"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   390
         TabIndex        =   12
         Top             =   1590
         Width           =   2265
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
         Height          =   285
         Index           =   0
         Left            =   390
         TabIndex        =   11
         Top             =   660
         Width           =   615
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Image associée"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   390
         TabIndex        =   10
         Top             =   2340
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   880
      Left            =   0
      TabIndex        =   7
      Top             =   3860
      Width           =   6825
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmTypesCoordonnees.frx":0672
         Height          =   510
         Index           =   2
         Left            =   3090
         Picture         =   "PrmTypesCoordonnees.frx":0C01
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Supprimer ce type de coordonnée"
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
         Picture         =   "PrmTypesCoordonnees.frx":1196
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Quitter sans enregistrer"
         Top             =   290
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "PrmTypesCoordonnees.frx":174F
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
         Picture         =   "PrmTypesCoordonnees.frx":1CAB
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Enregistrer"
         Top             =   290
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "PrmTypeCoordonnees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Entrée: Code|Libellé|Nombre max par personne|Image associée
' Sortie: Insertion dans les tables adéquates

Private Const CMD_OK = 0
Private Const CMD_QUITTER = 1
Private Const CMD_DETRUIRE = 2
Private Const CMD_PARCOURIR = 3
Private Const CMD_ALIMENTE = 4

' Index des objets TXT
Private Const TXT_CODE = 0
Private Const TXT_LIBELLE = 1
Private Const TXT_NBR_MAX = 2
' Index des objets LBL
Private Const LBL_CODE = 0
Private Const LBL_LIBELLE = 1
Private Const LBL_NBR_MAX = 2
' Index des objets CHL
Private Const CHK_LISTE = 0

Private Const CODE_MAIL = "ADRMAIL"

Private Const COUL_GRIS = &HC0C0C0
Private Const COUL_ORANGE = &H80C0FF

Private Const zones_excluts = " AND ZU_Code <> 'ALIASMAIL' "
' Liste des fichiers graphiques supportés par VisualBasic:
' Graphics formats recognized by Visual Basic include:
' bitmap (.bmp) files, icon (.ico) files, cursor (.cur) files,
' run-length encoded (.rle) files, metafile (.wmf) files,
' enhanced metafiles (.emf), GIF (.gif) files, and JPEG (.jpg) files.
Private Const LISTE_FICHIERS_IMAGES = "*.gif;*.bmp;*.jpg;*.jpeg;*.ico;*.cur;*.rle;*.wmf;*.emf"

' ZoneUtil.ZU_Type: 'C' pour coordonnée, 'X' pour autre chose
Private Const TYPE_PAR_DEFAUT = "C"
Private Const AUTRE_TYPE = "X"

' g_nomimage pour le chemin de l'image associée
Private g_nomimage As String
' g_img_update pour indiquer que l'image a été modifier
Private g_img_update As Boolean

' No fction en saisie (0 si nouveau)
Private g_numzoneutilisateur As Long

Private g_sret As String
Private g_mode_direct As Boolean

Private g_crzoneutilisateur_autor As Boolean

' Indique si la forme a déjà été activée
Private g_form_active As Boolean

' Indique si la saisie est en-cours
Private g_mode_saisie As Boolean

' Stocke le texte avant modif pour gérer le changement
Private g_txt_avant As String

' Le préfix à donner aux images associées aux types de coordonnée
Private Const g_prefix_image = "img_type_coord_"

Private Function afficher_zoneutilisateur( _
        ByVal v_numzoneutilisateur As Long) As Integer

    Dim sql As String, nomimg As String, nomimg_loc As String
    Dim rs As rdoResultset

    Call FRM_ResizeForm(Me, Me.width, Me.Height)
    
    g_img_update = False
    
    If v_numzoneutilisateur > 0 Then
        ' libellé existant
        cmd(CMD_DETRUIRE).Visible = True
        sql = "SELECT * FROM ZoneUtil" _
            & " WHERE ZU_Num=" & v_numzoneutilisateur

        If Odbc_Select(sql, rs) = P_ERREUR Then
            afficher_zoneutilisateur = P_ERREUR
            Exit Function
        End If
        g_numzoneutilisateur = v_numzoneutilisateur
        LblNum.Caption = v_numzoneutilisateur
        txt(TXT_CODE).Text = rs("ZU_Code").Value
        txt(TXT_LIBELLE).Text = rs("ZU_Libelle").Value
        txt(TXT_NBR_MAX).Text = rs("ZU_NbreMax").Value
        If rs("ZU_alimente") > 0 Then
            Me.lblInfSup.Visible = True
            Me.lblInfSup.Caption = Get_InfoSup(rs("ZU_alimente"))
            Me.lblInfSup.tag = rs("ZU_alimente")
            cmd(CMD_ALIMENTE).Caption = "alimentée par"
            cmd(CMD_ALIMENTE).BackColor = COUL_ORANGE
        Else
            Me.lblInfSup.Visible = True
            Me.lblInfSup.tag = 0
            cmd(CMD_ALIMENTE).Caption = "non alimentée"
            cmd(CMD_ALIMENTE).BackColor = COUL_GRIS
        End If
        ' On a pas trouvé l'image associée
        pct.Picture = LoadPicture("")
        If rs("ZU_Image").Value <> "" Then
            'initialiser la variable qui contient le nom de l'image
            g_nomimage = rs("ZU_Image").Value
            nomimg = p_CheminKW & "/kalibottin/grafx/TypeCoord/" & rs("ZU_Image").Value
            If KF_FichierExiste(nomimg) Then
                nomimg_loc = p_chemin_appli & "/tmp/" & rs("ZU_Image").Value
                If KF_GetFichier(nomimg, nomimg_loc) = P_OK Then
                    pct.Picture = LoadPicture(nomimg_loc)
                End If
                pct.tag = nomimg
            Else
                pct.tag = ""
            End If
        End If
        If rs("ZU_Liste").Value Then
            chk(CHK_LISTE).Value = 1
        Else
            chk(CHK_LISTE).Value = 0
        End If
        rs.Close
        cmd(CMD_DETRUIRE).Enabled = True
        ' si ADRMAIL, autoriser uniquement la modification de l'image associée
        If txt(TXT_CODE).Text = CODE_MAIL Then
            txt(TXT_CODE).Locked = True
            txt(TXT_LIBELLE).Locked = True
            txt(TXT_NBR_MAX).Locked = True
        Else
            txt(TXT_CODE).Locked = False
            txt(TXT_LIBELLE).Locked = False
            txt(TXT_NBR_MAX).Locked = False
        End If
        ' gestion de l'evenement clavier F2 pour supprimer
    Else
        ' nouveau libellé
        txt(TXT_CODE).Text = ""
        txt(TXT_LIBELLE).Text = ""
        txt(TXT_NBR_MAX).Text = ""
        Me.lblInfSup.tag = 0
        Me.lblInfSup.Visible = True
        Me.lblInfSup.tag = 0
        cmd(CMD_ALIMENTE).Caption = "non alimentée"
        cmd(CMD_ALIMENTE).BackColor = COUL_GRIS
        pct.Picture = LoadPicture("")
        chk(CHK_LISTE).Value = 0
        g_numzoneutilisateur = 0
        g_nomimage = ""
        ' cmd(CMD_DETRUIRE).Enabled = False
        cmd(CMD_DETRUIRE).Visible = False
    End If

    cmd(CMD_OK).Enabled = False
    txt(TXT_CODE).SetFocus
    Me.MousePointer = 0
    g_mode_saisie = True

    afficher_zoneutilisateur = P_OK

End Function

Public Function AppelFrm(ByVal v_mode_param As Long) As String

    If v_mode_param >= 0 Then
        ' mode de création direct
        g_mode_direct = True
    Else
        g_mode_direct = False
    End If

    Me.Show 1

    AppelFrm = g_sret

End Function
Private Function choisir_tis() As Long

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

    Call CL_AddLigne("=> non alimentée", 0, "", False)
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
        Call FRM_ResizeForm(Me, Me.width, Me.Height)
        choisir_tis = 0
        Exit Function
    End If

    If CL_liste.retour = 0 Then
        If CL_liste.lignes(CL_liste.pointeur).num = 0 Then
            cmd(CMD_ALIMENTE).BackColor = COUL_GRIS
            cmd(CMD_ALIMENTE).Caption = "non alimentée"
            choisir_tis = CL_liste.lignes(CL_liste.pointeur).num
            Me.lblInfSup.Visible = False
            Me.lblInfSup.Caption = ""
            Me.lblInfSup.tag = 0
            Call FRM_ResizeForm(Me, Me.width, Me.Height)
            cmd(CMD_OK).Enabled = True
        Else
            cmd(CMD_ALIMENTE).BackColor = COUL_ORANGE
            cmd(CMD_ALIMENTE).Caption = "alimentée par"
            choisir_tis = CL_liste.lignes(CL_liste.pointeur).num
            Me.lblInfSup.Visible = True
            Me.lblInfSup.Caption = Get_InfoSup(CL_liste.lignes(CL_liste.pointeur).num)
            Me.lblInfSup.tag = CL_liste.lignes(CL_liste.pointeur).num
            Call FRM_ResizeForm(Me, Me.width, Me.Height)
            cmd(CMD_OK).Enabled = True
        End If
        Exit Function
    End If
    GoTo lab_affiche
    
lab_erreur:
    choisir_tis = 0

End Function

Private Function choisir_zoneutilisateur() As Integer

    Dim sret As String, sql As String
    Dim n As Integer
    Dim nofct As Long
    Dim rs As rdoResultset

    Call FRM_ResizeForm(Me, 0, 0)

lab_affiche:
    Call CL_Init
    'Choix du libelle
    sql = "SELECT * FROM ZoneUtil" _
        & " WHERE ZU_Type='" & TYPE_PAR_DEFAUT & "'" _
        & zones_excluts _
        & " ORDER BY ZU_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        choisir_zoneutilisateur = P_ERREUR
        Exit Function
    End If
    n = 1
    Call CL_AddLigne("<Nouvelle>", 0, "", False)

    While Not rs.EOF
        Call CL_AddLigne(rs("ZU_Libelle").Value, rs("ZU_Num").Value, "", False)
        n = n + 1
        rs.MoveNext
    Wend
    rs.Close

    If n = 0 Then
        Call MsgBox("Aucun type de coordonnée n'a été trouvé.", vbOKOnly + vbInformation, "")
        choisir_zoneutilisateur = P_NON
        Exit Function
    End If

    Call CL_InitTitreHelp("Liste des types de coordonnées", "")
    Call CL_InitTaille(0, -15)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)

    ChoixListe.Show 1

    ' Quitter
    If CL_liste.retour = 1 Then
        choisir_zoneutilisateur = P_NON
        Exit Function
    End If

    If afficher_zoneutilisateur(CL_liste.lignes(CL_liste.pointeur).num) = P_ERREUR Then
        choisir_zoneutilisateur = P_ERREUR
        Exit Function
    End If

    choisir_zoneutilisateur = P_OUI

End Function

Private Sub copier_image(ByVal v_nomfichier As String, ByVal v_fDest As String)
' ******************************************************************
' Copier l'image associée à ce type de coordonnée dans le répértoire:
' P_CHEMIN_SMB_KW/kalibottin/grafx/TypeCoord
' ******************************************************************

    ' Effacer l'ancienne image si elle existe
    If pct.tag <> "" Then
        If Dir(pct.tag) <> "" Then
            Call FICH_EffacerFichier(pct.tag, False)
        End If
    End If
    ' Copier la nouvelle image
    Call KF_PutFichier(v_fDest, v_nomfichier)

End Sub
Private Function enregistrer_zoneutilisateur() As Integer

    If g_numzoneutilisateur = 0 Then
        ' Ajout d'un nouveau enregistrement
        If Odbc_AddNew("ZoneUtil", _
                        "ZU_Num", _
                        "ZU_Seq", _
                        True, _
                        g_numzoneutilisateur, _
                        "ZU_Libelle", txt(TXT_CODE).Text) = P_ERREUR Then
            enregistrer_zoneutilisateur = P_ERREUR
            Exit Function
        End If
    Else
        ' Maj la table ZoneUtil avec les nouvelles données
        If Odbc_Update("ZoneUtil", _
                        "ZU_Num", _
                        "WHERE ZU_Num=" & g_numzoneutilisateur, _
                        "ZU_Libelle", txt(TXT_CODE).Text) = P_ERREUR Then
            enregistrer_zoneutilisateur = P_ERREUR
            Exit Function
        End If
    End If

    enregistrer_zoneutilisateur = P_OK

End Function

Private Function est_entier_positif(ByVal valeur As String) As Boolean

' voir la version MCommon.STR_EstEntierPos()
    On Error GoTo erreur

    ' si la valeur est numérique
    If IsNumeric(valeur) Then
        ' si la valeur est un entier
        If val(valeur) = CInt(valeur) And CDbl(valeur) = CInt(valeur) Then
            'si cet entier est positif
            If val(valeur) >= 0 Then
                est_entier_positif = True
            Else
                est_entier_positif = False
            End If
        Else
            est_entier_positif = False
        End If
    Else
        est_entier_positif = False
    End If

    Exit Function

erreur:
    est_entier_positif = False

End Function

Private Function Get_InfoSup(v_num)
    Dim sql As String, rs As rdoResultset
    
    sql = "SELECT * FROM KB_TypeInfoSuppl where KB_TisNum=" & v_num
    If Odbc_Select(sql, rs) = P_ERREUR Then
        Get_InfoSup = ""
    Else
        Get_InfoSup = rs("KB_TisLibelle")
        rs.Close
    End If
End Function

Private Sub initialiser()

    g_crzoneutilisateur_autor = True

    Call FRM_ResizeForm(Me, 0, 0)

    g_mode_saisie = False

    cmd(CMD_OK).Visible = True
    cmd(CMD_DETRUIRE).Visible = True

    If g_mode_direct Then
        Call afficher_zoneutilisateur(0)
    Else
        If choisir_zoneutilisateur() <> P_OUI Then
            Call quitter(True)
            Exit Sub
        End If
    End If

End Sub

Private Sub parcourir()

    Dim nom_fich As String
    Dim type_coordonnee_num As Integer

    ' La recherche du fichier image
    nom_fich = Com_ChoixFichier.AppelFrm("Selectionnez l'image associée à ce type de coordonnée:", "", "", _
                    LISTE_FICHIERS_IMAGES, False)
    ' On n'active le bouton Enregistrer que si on a bien choisi une image
    If nom_fich <> "" Then
        g_img_update = True
        g_nomimage = nom_fich
        pct.Picture = LoadPicture(g_nomimage)
        cmd(CMD_OK).Enabled = True
    End If

End Sub

Private Sub quitter(ByVal v_bforce As Boolean)

    Dim sql As String
    Dim reponse As Integer
    Dim rs As rdoResultset

    If v_bforce Then
        
        If g_mode_direct Then
'            sql = "SELECT ZU_Num FROM ZoneUtil WHERE ZU_Code=" & txt(TXT_CODE).Text
'            If Odbc_Select(sql, rs) = P_ERREUR Then
'                Exit Sub
'            End If
'            g_sret = rs("ZU_Num").Value & "|" & "z"
'            rs.Close
            Unload Me
            Exit Sub
        End If
            
        'g_sret = ""
        g_sret = txt(TXT_CODE).Text & "|" & "ee"
        Unload Me
        Exit Sub
    End If

    If cmd(CMD_OK).Visible And cmd(CMD_OK).Enabled Then
        reponse = MsgBox("Des modifications ont été effectuées !" & vbLf & vbLf & "Confirmez-vous l'abandon ?", _
                          vbYesNo + vbDefaultButton2 + vbQuestion)
        If reponse = vbNo Then Exit Sub
    End If

    If choisir_zoneutilisateur() <> P_OUI Then Unload Me

End Sub

Private Sub supprimer()

    ' CRA, CRU : Code Retour Application ou Utlisiateur
    Dim reponse As Integer, CRA As Integer, CRU As Integer
    Dim lnb As Long

    ' Utilisateur associé à ce type de coordonnée ?
    CRU = zone_dans_coordonnee()
    ' personne n'est associé à ce type de coordonnée
    If CRU = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    ' message de prevention sans modification(s)
    If CRU = P_OUI Then
        Call MsgBox("Des utilisateurs ont des coordonnées de ce type." & vbLf & vbCr _
                    & "Ce type de coordonnée ne peut donc être supprimé.", _
                    vbExclamation + vbOKOnly, "")
        Exit Sub
    End If

    ' Application associée à ce type de coordonnée ?
    CRA = zone_dans_application()

    ' aucune application n'est associée à ce type de coordonnée
    If CRA = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If
    ' message de prevention sans modification(s)
    If CRA = P_OUI Then
        Call MsgBox("Des applications sont associées à ce type de coordonnée." & vbLf & vbCr _
                    & "Ce type de coordonnée ne peut donc être supprimé.", _
                    vbExclamation + vbOKOnly, "")
        Exit Sub
    End If

    reponse = MsgBox("Confirmez-vous la suppression de ce type de coordonnée ?", vbQuestion + vbYesNo + vbDefaultButton2, "")
    If reponse = vbNo Then
        Exit Sub
    End If

    ' Maj table zoneutil, en supprimant le type de coordonnée en cours
    If Odbc_Delete("ZoneUtil", _
                    "ZU_Num", _
                    "WHERE ZU_Num=" & g_numzoneutilisateur, _
                    lnb) = P_ERREUR Then
        Call quitter(True)
        Exit Sub
    End If

    If g_mode_direct Then
        g_sret = g_numzoneutilisateur & "|" & txt(TXT_CODE).Text
        Unload Me
        Exit Sub
    End If

    If choisir_zoneutilisateur() <> P_OUI Then Call quitter(True)

End Sub

Private Sub valider()
' ***************************************************************************
' La vérification de chaque champ se fait lors du txt_LostFocus()
' Copier l'image (si elle existe): p_chemin_smb_kw\kalibottin\grafx\TypeCoord
' ***************************************************************************
    Dim extension As String, sql As String, file_dest As String
    Dim I As Integer
    Dim lbid As Long
    Dim rs As rdoResultset

    ' maj la table: kalibottin.zoneutil
    If g_numzoneutilisateur = 0 Then ' mode creation
        Call Odbc_AddNew("ZoneUtil", "ZU_Num", "ZU_Seq", True, g_numzoneutilisateur, _
                    "ZU_Code", txt(TXT_CODE).Text, _
                    "ZU_Libelle", txt(TXT_LIBELLE).Text, _
                    "ZU_Type", TYPE_PAR_DEFAUT, _
                    "ZU_NbreMax", txt(TXT_NBR_MAX).Text, _
                    "ZU_Image", g_nomimage, _
                    "ZU_alimente", Me.lblInfSup.tag, _
                    "ZU_Liste", IIf(chk(CHK_LISTE) = 1, True, False))
    Else ' mode modification
        Call Odbc_Update("ZoneUtil", "ZU_Num", "WHERE ZU_Num=" & g_numzoneutilisateur, _
                    "ZU_Code", txt(TXT_CODE).Text, _
                    "ZU_Libelle", txt(TXT_LIBELLE).Text, _
                    "ZU_Type", TYPE_PAR_DEFAUT, _
                    "ZU_NbreMax", txt(TXT_NBR_MAX).Text, _
                    "ZU_Image", g_nomimage, _
                    "ZU_alimente", Me.lblInfSup.tag, _
                    "ZU_Liste", IIf(chk(CHK_LISTE) = 1, True, False))
    End If

    ' Récupérer l'extension de l'image
    If g_nomimage <> "" And g_img_update Then
        extension = Right$(g_nomimage, Len(g_nomimage) - InStrRev(g_nomimage, ".") + 1)
        ' Copier l'image
        If extension <> "" Then
            file_dest = p_CheminKW & "/kalibottin/grafx/TypeCoord/" & g_prefix_image & g_numzoneutilisateur & extension
            Call copier_image(g_nomimage, file_dest)
            ' Donner le nouveau nom de l'image
            g_nomimage = g_prefix_image & g_numzoneutilisateur & extension
        Else
            g_nomimage = ""
        End If
        Call Odbc_Update("ZoneUtil", "ZU_Num", "where zu_num=" & g_numzoneutilisateur, "ZU_Image", g_nomimage)
    End If
    
    If g_mode_direct Then
        sql = "SELECT ZU_Num FROM ZoneUtil WHERE ZU_Code='" & txt(TXT_CODE).Text & "'"
        If Odbc_Select(sql, rs) = P_ERREUR Then
            Exit Sub
        End If
        g_sret = rs("ZU_Num").Value & "|" & "z"
        rs.Close
        Unload Me
        Exit Sub
    End If

    Call initialiser

End Sub

Private Function verif_un_champ(ByVal v_index As Integer) As Integer
' ******************************************************************
' La validation des champs se fait lors de la perte du focus.
' Pas besoin de faire Valider_Tous_Champs() lors de l'enregistrement
' ******************************************************************
    Dim lib As String, sql As String, message As String
    Dim rs As rdoResultset

    lib = txt(v_index).Text
    Select Case v_index
    Case TXT_CODE ' ***************** CODE ***********************
        If lib <> "" Then
            sql = "SELECT ZU_Num FROM ZoneUtil" _
                & " WHERE ZU_Code=" & Odbc_String(lib) _
                & " AND ZU_Num<>" & g_numzoneutilisateur
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                ' il n'y a pas de doublon, on peut quitter
                GoTo lab_erreur
            End If
            If Not rs.EOF Then
                If rs("ZU_Num").Value <> g_numzoneutilisateur Then
                    rs.Close
                    lbl(LBL_CODE).ForeColor = vbRed
                    Call MsgBox("Code déjà existant.", vbOKOnly + vbExclamation, "")
                    ' code existant, on ne quitte pas
                    GoTo lab_erreur
                End If
            End If
            rs.Close
        Else ' If txt(TXT_CODE).Text = ""
            MsgBox ("Le CODE est une rubrique obligatoire !")
            lbl(LBL_CODE).ForeColor = vbRed
            GoTo lab_erreur
        End If
        lbl(LBL_CODE).ForeColor = vbBlack
    Case TXT_LIBELLE ' ************** LIBELLE ********************
        If lib <> "" Then
            sql = "SELECT ZU_Num FROM ZoneUtil" _
                & " WHERE ZU_Libelle=" & Odbc_String(lib) _
                & " AND ZU_Num<>" & g_numzoneutilisateur
            If Odbc_SelectV(sql, rs) = P_ERREUR Then
                ' il n'y a pas de doublon, on peut quitter
                GoTo lab_erreur
            End If
            If Not rs.EOF Then
                If rs("ZU_Num").Value <> g_numzoneutilisateur Then
                    rs.Close
                    lbl(LBL_LIBELLE).ForeColor = vbRed
                    Call MsgBox("Libellé déjà existant.", vbOKOnly + vbExclamation, "")
                    ' code existant, on ne quitte pas
                    GoTo lab_erreur
                End If
            End If
            rs.Close
        Else ' If txt(TXT_LIBELLE).Text = ""
            lbl(LBL_LIBELLE).ForeColor = vbRed
            MsgBox ("Le LIBELLE est une rubrique obligatoire !")
            GoTo lab_erreur
        End If
        lbl(LBL_LIBELLE).ForeColor = vbBlack
    Case TXT_NBR_MAX ' *************** NBR_MAX *******************
        If lib <> "" Then
            If Not est_entier_positif(txt(TXT_NBR_MAX).Text) Then
                lbl(LBL_NBR_MAX).ForeColor = vbRed
                MsgBox ("Le nombre maximum par personne doit être un entier positif")
                GoTo lab_erreur
            End If
        Else ' If txt(TXT_NBR_MAX).Text = ""
            lbl(LBL_NBR_MAX).ForeColor = vbRed
            MsgBox ("Le NOMBRE MAX PAR PERSONNE est une rubrique obligatoire !")
            GoTo lab_erreur
        End If
        lbl(LBL_NBR_MAX).ForeColor = vbBlack
    End Select ' ************************************************

    verif_un_champ = P_OUI
    Exit Function

lab_erreur:
    verif_un_champ = P_NON

End Function

Private Function zone_dans_application() As Integer

    Dim sql As String
    Dim lnb As Long

    sql = "SELECT COUNT(*) FROM Application" _
        & " WHERE APP_ZoneRens LIKE '%Z" & g_numzoneutilisateur & ";%'" _
        & " OR APP_ZonePrev LIKE '%Z" & g_numzoneutilisateur & ";%'"
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        zone_dans_application = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        zone_dans_application = P_OUI
        Exit Function
    End If

    zone_dans_application = P_NON

End Function

Private Function zone_dans_coordonnee() As Integer
' ***********************************************************************
' Déterminer si des utilisateurs sont associés avec ce type de coordonnée
' ***********************************************************************
    Dim sql As String
    Dim lnb As Long

    sql = "SELECT COUNT(*) FROM UtilCoordonnee" _
        & " WHERE UC_ZUNum=" & g_numzoneutilisateur
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        zone_dans_coordonnee = P_ERREUR
        Exit Function
    End If
    If lnb > 0 Then
        zone_dans_coordonnee = P_OUI
        Exit Function
    End If

    zone_dans_coordonnee = P_NON

End Function

Private Sub chk_Click(Index As Integer)

    If g_mode_saisie Then
        cmd(CMD_OK).Enabled = True
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case CMD_OK
            If verif_un_champ(TXT_CODE) And verif_un_champ(TXT_LIBELLE) And verif_un_champ(TXT_NBR_MAX) Then
                Call valider
            End If

        Case CMD_QUITTER
            If g_mode_direct Then
                Call quitter(True)
            Else
                Call quitter(False)
            End If
        Case CMD_DETRUIRE
            Call supprimer
        Case CMD_PARCOURIR
            Call parcourir
        Case CMD_ALIMENTE
            Call choisir_tis
    End Select

End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Index = CMD_DETRUIRE Or Index = CMD_QUITTER Then
        g_mode_saisie = False
    End If

End Sub

Private Sub Form_Activate()

    If Not g_form_active Then
        Call initialiser
        g_form_active = True
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyE And Shift = vbAltMask) Or KeyCode = vbKeyF1 Then
        KeyCode = 0
        If verif_un_champ(TXT_CODE) And verif_un_champ(TXT_LIBELLE) And verif_un_champ(TXT_NBR_MAX) Then
            If cmd(CMD_OK).Enabled Then Call valider
        End If
    ElseIf (KeyCode = vbKeyS And Shift = vbAltMask) Or KeyCode = vbKeyF2 Then
        KeyCode = 0
        If cmd(CMD_DETRUIRE).Enabled Then
            Call supprimer
        End If
    ElseIf KeyCode = vbKeyH And Shift = vbAltMask Then
        KeyCode = 0
        Call HtmlHelp(0, p_chemin_appli + "\help\kalidoc.chm", HH_DISPLAY_TOPIC, "dico_d_fonction.htm")
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Call quitter(False)
    End If

End Sub

Private Sub Form_Load()

    g_form_active = False

End Sub

Private Sub txt_Change(Index As Integer)

    cmd(CMD_OK).Enabled = True

End Sub

Private Sub txt_GotFocus(Index As Integer)

    ' on récupère l'ancienne valeur de l'objet
    g_txt_avant = txt(Index).Text

End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case KeyAscii
    Case vbKeyReturn
        ' Enter = TAB
        SendKeys "{TAB}"
    Case vbKeyEscape
        ' Echape = Quitter
        Call quitter(True)
    End Select

End Sub

Private Sub txt_LostFocus(Index As Integer)

    Dim cr As Integer

    ' tester si le libellé est unique
    If g_mode_saisie Then
        If txt(Index).Text <> g_txt_avant Then
            ' la valeur du champs à changé
            ' on gere la valeur du champs
            cr = verif_un_champ(Index)
            ' s'il n'y a pas de redondance
            If cr = P_ERREUR Then
                Call quitter(True)
                Exit Sub
            End If
            ' si le code existe déjà
            If cr = P_NON Then
                ' on remet l'ancienne valeur sans quitter
                txt(Index).Text = g_txt_avant
                txt(Index).SetFocus
                Exit Sub
            End If
            ' On peut enregistrer et quitter
            cmd(CMD_OK).Enabled = True
        Else ' le TXT est bon !
            Select Case Index
            Case TXT_CODE
                lbl(LBL_CODE).ForeColor = vbBlack
            Case TXT_LIBELLE
                lbl(LBL_LIBELLE).ForeColor = vbBlack
            Case TXT_NBR_MAX
                lbl(LBL_NBR_MAX).ForeColor = vbBlack
            End Select
        End If
    End If

End Sub
