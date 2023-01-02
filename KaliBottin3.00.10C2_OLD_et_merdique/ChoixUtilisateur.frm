VERSION 5.00
Begin VB.Form ChoixUtilisateur 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Personne"
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
      Height          =   3105
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5985
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1110
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rechercher par"
         ForeColor       =   &H00800080&
         Height          =   1155
         Left            =   480
         TabIndex        =   7
         Top             =   1800
         Width           =   4965
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
            Height          =   540
            Index           =   4
            Left            =   3420
            Picture         =   "ChoixUtilisateur.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Rechercher par services"
            Top             =   420
            UseMaskColor    =   -1  'True
            Width           =   660
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
            Height          =   540
            Index           =   5
            Left            =   630
            Picture         =   "ChoixUtilisateur.frx":058F
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Rechercher par fonctions"
            Top             =   420
            Width           =   660
         End
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   0
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ou"
         Height          =   255
         Index           =   2
         Left            =   1170
         TabIndex        =   11
         Top             =   870
         Width           =   315
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
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nom"
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
         Left            =   1080
         TabIndex        =   6
         Top             =   630
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   825
      Left            =   0
      TabIndex        =   4
      Top             =   2940
      Width           =   5985
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
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
         Left            =   480
         Picture         =   "ChoixUtilisateur.frx":09F2
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
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
         Index           =   1
         Left            =   5010
         Picture         =   "ChoixUtilisateur.frx":0E4B
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "ChoixUtilisateur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Entrée : Titre du choix | un seul utilisateur autorisé |
'          Plusieurs utilisateurs autorisés | Dans la liste
' Sortie : U si utilisateur   p_choixlistem contient les utilisateurs selectionnés
'   ou     Fx si fonction
'   ou     Sx si service
'   ou     Rien si Abandon

'Index sur les objets cmd
Private Const CMD_OK = 0
Private Const CMD_CHOIX_FONCTION = 5
Private Const CMD_CHOIX_SERVICE = 4
Private Const CMD_FERMER = 1

Private Const TXT_NOM = 0
Private Const TXT_CODE = 1

Private g_titre As String
Private g_titre_doc As String
Private g_plusieurs_util_autor As Boolean
Private g_choix_dans_liste As Boolean
Private g_ssite As String
Private g_scr As String

Private g_sql_exclu As String ' la personne à exclur du choix

Private g_mode_saisie As Boolean

Private g_txt_avant As String

Private g_form_active As Boolean

Public Function AppelFrm(ByVal v_titre As String, _
                          ByVal v_titre_doc As String, _
                          ByVal v_plusieurs_util_autor As Boolean, _
                          ByVal v_choix_dans_liste As Boolean, _
                          ByVal v_ssite As String, _
                          Optional ByVal v_unum_exclu As Long) As String

    g_titre = v_titre
    g_titre_doc = v_titre_doc
    g_plusieurs_util_autor = v_plusieurs_util_autor
    g_choix_dans_liste = v_choix_dans_liste
    g_ssite = v_ssite
    g_sql_exclu = IIf(v_unum_exclu > 0, " AND U_Num<>" & v_unum_exclu, "")

    ChoixUtilisateur.Show 1
    
    AppelFrm = g_scr

End Function

Private Sub build_sql_fsu(ByVal v_sfct As String, _
                          ByVal v_sspm As String, _
                          ByRef r_sql As String)

    Dim clause_labo As String, clause As String, sDest As String
    Dim s As String, s2 As String, slst As String
    Dim n As Integer, i As Integer
    Dim rs As rdoResultset

    clause_labo = ""
    If g_ssite <> "" Then
        clause = " AND ("
        n = STR_GetNbchamp(g_ssite, ";")
        For i = 0 To n - 1
            clause_labo = clause_labo & clause & "U_Labo LIKE '%" & STR_GetChamp(g_ssite, ";", i) & ";%'"
            clause = " OR "
        Next i
        If clause_labo <> "" Then
            clause_labo = clause_labo + ")"
        End If
    End If

    sDest = ""
    If v_sfct <> "" Then
        sDest = sDest + v_sfct
    ElseIf v_sspm <> "" Then
        sDest = sDest + v_sspm
    End If

    r_sql = "SELECT U_Num, U_Nom, U_Prenom, U_SPM, U_FctTrav" _
            & " FROM Utilisateur" _
            & " WHERE U_kb_actif=True AND U_Actif=True" _
            & clause_labo
    If sDest <> "" Then
        n = STR_GetNbchamp(sDest, "|")
        For i = 0 To n - 1
            s = STR_GetChamp(sDest, "|", i)
            If i = 0 Then
                r_sql = r_sql & " AND ("
            Else
                r_sql = r_sql & " OR"
            End If
            Select Case left$(s, 1)
                Case "F"
                    r_sql = r_sql & " U_FctTrav LIKE '%" & s & ";%'"
                Case "S"
                    r_sql = r_sql & " U_SPM LIKE '%" & s & "%'"
                Case "U"
                    r_sql = r_sql & " U_Num=" & Mid$(s, 2)
            End Select
        Next i
        r_sql = r_sql & ")"
    End If
    r_sql = r_sql & " ORDER BY U_Nom, U_Prenom"

End Sub

Private Sub choisir_dans_la_liste()

    Dim nomutil As String, libspm As String, libfct As String
    Dim i As Integer

    Call FRM_ResizeForm(Me, 0, 0)

    p_siz_tblu_sel = -1

    Call CL_Init
    Call CL_InitTitreHelp(g_titre + " " + g_titre_doc, "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    If g_plusieurs_util_autor Then
        Call CL_AddBouton("&Tous", "", 0, 0, 0)
    End If
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    For i = 0 To p_siz_tblu
        If recup_fsp(p_tblu(i), nomutil, libfct, libspm) = P_ERREUR Then
            Call quitter
            Exit Sub
        End If
        Call CL_AddLigne(nomutil & vbTab & libfct & vbTab & libspm, _
                         p_tblu(i), _
                         "", _
                         False)
    Next i
    Call CL_InitTaille(0, -20)
    If g_plusieurs_util_autor Then
        Call CL_InitMultiSelect(True, True)
        ChoixListe.Show 1
        If CL_liste.retour = 2 Then
            Call quitter
            Exit Sub
        End If
        If CL_liste.retour = 1 Then
            p_siz_tblu_sel = p_siz_tblu
            ReDim p_tblu_sel(p_siz_tblu_sel) As Long
            For i = 0 To p_siz_tblu
                p_tblu_sel(i) = p_tblu(i)
            Next i
        Else
            For i = 0 To p_siz_tblu
                If CL_liste.lignes(i).selected = True Then
                    p_siz_tblu_sel = p_siz_tblu_sel + 1
                    ReDim Preserve p_tblu_sel(p_siz_tblu_sel) As Long
                    p_tblu_sel(p_siz_tblu_sel) = CL_liste.lignes(i).num
                End If
            Next i
        End If
    Else
        ChoixListe.Show 1
        If CL_liste.retour = 1 Then
            Call quitter
            Exit Sub
        End If
        p_siz_tblu_sel = 0
        ReDim p_tblu_sel(0) As Long
        p_tblu_sel(0) = CL_liste.lignes(CL_liste.pointeur).num
    End If

    g_scr = "U"
    Unload Me

End Sub

Private Sub choisir_fonction()

    Dim sql As String, sret As String, sfct As String
    Dim n As Integer, i As Integer
    Dim rs As rdoResultset

    Call FRM_ResizeForm(Me, 0, 0)
    
    Call CL_Init
    n = 0
    sql = "SELECT * FROM FctTrav" _
        & " ORDER BY FT_Libelle"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Call FRM_ResizeForm(Me, Me.width, Me.Height)
        Exit Sub
    End If
    While Not rs.EOF
        Call CL_AddLigne(rs("FT_Libelle").Value, rs("FT_Num").Value, "", False)
        n = n + 1
lab_suivant:
        rs.MoveNext
    Wend
    rs.Close
    If n = 0 Then
        Exit Sub
    End If

    Call CL_InitTitreHelp("Fonctions du personnel", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    Call CL_InitTaille(0, -15)
    Call CL_InitMultiSelect(True, True)
    ChoixListe.Show 1
    If CL_liste.retour = 1 Then
        Call FRM_ResizeForm(Me, Me.width, Me.Height)
        Exit Sub
    End If
    sfct = ""
    For i = 0 To n - 1
        If CL_liste.lignes(i).selected Then
            sfct = sfct & "F" & CL_liste.lignes(i).num & "|"
        End If
    Next i
    Call choisir_utilisateur_fsu(sfct, "")

End Sub

Private Sub choisir_service()

    Dim sret As String, ssite As String, s_srv As String, sprm As String
    Dim encore As Boolean
    Dim i As Integer, n As Integer
    Dim numlabo As Long, numutil As Long
    Dim frm As Form

    numlabo = p_NumLabo

    Call FRM_ResizeForm(Me, 0, 0)
    
    If g_ssite <> "" Then
        ssite = STR_Supprimer(g_ssite, "L")
    End If
    encore = True
    Do
        Call CL_Init
'        Set frm = PrmService
'        sret = PrmService.AppelFrm("Choix d'un service / poste", "S", g_plusieurs_util_autor, ssite, "SP", True)
        Set frm = ChoixService
        sret = ChoixService.AppelFrm("Choix d'un service / poste", "S", g_plusieurs_util_autor, ssite, "S", True)
        Set frm = Nothing
        If sret = "" Then
            Call FRM_ResizeForm(Me, Me.width, Me.Height)
            Exit Sub
        End If
        If g_plusieurs_util_autor And left$(sret, 1) = "N" Then
            encore = False
        ElseIf Not g_plusieurs_util_autor And left$(sret, 1) = "S" Then
            encore = False
        Else
            Set frm = PrmPersonne
            numutil = STR_GetChamp(sret, "|", 0)
            If numutil = 0 Then
                sprm = "POSTE=" & Mid$(STR_GetChamp(sret, "|", 1), 2)
            Else
                sprm = ""
            End If
            Call PrmPersonne.AppelFrm(numutil, sprm)
            Set frm = Nothing
        End If
    Loop Until encore = False
    
    p_NumLabo = numlabo

    If g_plusieurs_util_autor Then
        s_srv = ""
        n = CLng(Mid$(sret, 2))
        If n = 0 Then
            Exit Sub
        End If
        For i = 0 To n - 1
            s_srv = s_srv + CL_liste.lignes(i).texte + "|"
        Next i
    Else
        s_srv = sret
    End If
    Call choisir_utilisateur_fsu("", s_srv)
    
End Sub

Private Sub choisir_utilisateur_fsu(ByVal v_sfct As String, _
                                     ByVal v_sp As String)

    Dim sql As String, libsp As String, libfct As String, sret As String
    Dim s As String
    Dim n As Integer, i As Integer
    Dim rs As rdoResultset
    
    Call build_sql_fsu(v_sfct, v_sp, sql)
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        GoTo lab_fin
    End If

    Me.MousePointer = 11

    Call CL_Init
    Call CL_InitTitreHelp(g_titre, "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    n = 1
    If g_plusieurs_util_autor Then
        Call CL_AddBouton("&Tous", "", 0, 0, 0)
        n = n + 1
    End If
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    n = 0
    While Not rs.EOF
        s = rs("U_Nom").Value & " " & rs("U_Prenom").Value & vbTab
        If P_RecupSPLib(rs("U_SPM").Value, libsp) = P_ERREUR Then
            GoTo lab_fin
        End If
        s = s & libsp & vbTab
        Call CL_AddLigne(s, rs("U_Num").Value, "", False)
        n = n + 1
lab_suiv1:
        rs.MoveNext
    Wend
    rs.Close
    
    If n = 0 Then
        GoTo lab_fin
    End If

    Call CL_InitTaille(0, -20)
    If g_plusieurs_util_autor Then
        Call CL_InitMultiSelect(True, True)
        ChoixListe.Show 1
        If CL_liste.retour = 2 Then
            GoTo lab_fin
        End If
        If CL_liste.retour = 1 Then
            p_siz_tblu_sel = n - 1
            ReDim p_tblu_sel(n - 1) As Long
            For i = 0 To n - 1
                p_tblu_sel(i) = CL_liste.lignes(i).num
            Next i
        Else
            For i = 0 To n - 1
                If CL_liste.lignes(i).selected = True Then
                    p_siz_tblu_sel = p_siz_tblu_sel + 1
                    ReDim Preserve p_tblu_sel(p_siz_tblu_sel) As Long
                    p_tblu_sel(p_siz_tblu_sel) = CL_liste.lignes(i).num
                End If
            Next i
        End If
    Else
        ChoixListe.Show 1
        If CL_liste.retour = 1 Then
            GoTo lab_fin
        End If
        p_siz_tblu_sel = 0
        ReDim p_tblu_sel(0) As Long
        p_tblu_sel(0) = CL_liste.lignes(CL_liste.pointeur).num
    End If

    Me.MousePointer = 0

    g_scr = "U"
    Unload Me
    Exit Sub
    
lab_fin:
    Me.MousePointer = 0
    Call FRM_ResizeForm(Me, Me.width, Me.Height)

End Sub

Private Sub initialiser()

    p_siz_tblu_sel = -1

    g_mode_saisie = False

    If g_choix_dans_liste And p_siz_tblu <> -1 Then
        Call choisir_dans_la_liste
        Exit Sub
    End If

    frm.Caption = g_titre

    g_mode_saisie = True
    txt(TXT_NOM).SetFocus

End Sub

Private Sub quitter()

    g_scr = ""
    Unload Me

End Sub

Private Function recup_fcttrav(ByVal v_sfct As String, _
                               ByRef r_lib As String) As Integer

    Dim s As String
    Dim num As Long

    s = STR_GetChamp(v_sfct, ";", 0)
    If s <> "" Then
        num = CLng(Mid$(s, 2))
        If Odbc_RecupVal("SELECT FT_Libelle FROM FctTrav WHERE FT_Num=" & num, _
                         r_lib) = P_ERREUR Then
            recup_fcttrav = P_ERREUR
            Exit Function
        End If
    Else
        r_lib = ""
    End If

    recup_fcttrav = P_OK

End Function

Private Function recup_fsp(ByVal v_numutil As Long, _
                           ByRef r_nomutil As String, _
                           ByRef r_libfct As String, _
                           ByRef r_libsp As String) As Integer

    Dim sfct As String, prenom As String
    Dim n As Integer
    Dim s_sp As Variant
    Dim rs As rdoResultset

    If Odbc_RecupVal("SELECT U_Nom, U_Prenom, U_FctTrav, U_SPM FROM Utilisateur" & _
                     " WHERE U_kb_actif=True AND U_Num=" & v_numutil, _
                     r_nomutil, _
                     prenom, _
                     sfct, _
                     s_sp) = P_ERREUR Then
        recup_fsp = P_ERREUR
        Exit Function
    End If

    r_nomutil = r_nomutil + " " + prenom
    If recup_fcttrav(sfct, r_libfct) = P_ERREUR Then
        recup_fsp = P_ERREUR
        Exit Function
    End If

    s_sp = STR_GetChamp(s_sp, "|", 0)
    n = STR_GetNbchamp(s_sp, ";")
    s_sp = STR_GetChamp(s_sp, ";", n - 2)
    If P_RecupSPLib(s_sp, r_libsp) = P_ERREUR Then
        recup_fsp = P_ERREUR
        Exit Function
    End If

    recup_fsp = P_OK

End Function

Private Function verif_util() As Integer

    Dim sql As String, nomutil As String, libfct As String, libspm As String
    Dim codutil As String
    Dim est_ok As Boolean
    Dim n As Integer, i As Integer, j As Integer, nbchp_u As Integer, nbchp_p As Integer
    Dim nbtot As Integer
    Dim numutil As Long, numlabo As Long, lnb As Long
    Dim rs As rdoResultset

    If txt(TXT_NOM).Text = "" And txt(TXT_CODE).Text = "" Then
        verif_util = P_NON
        Exit Function
    End If

    If txt(TXT_NOM).Text <> "" Then
        nomutil = UCase(txt(TXT_NOM).Text)
        sql = "SELECT U_Num, U_Nom, U_Prenom, U_Labo" _
            & " FROM Utilisateur" _
            & " WHERE U_kb_actif=True AND U_Num>1" _
            & " AND " & Odbc_upper() & "(U_Nom)=" & Odbc_String(nomutil) _
            & " AND U_Actif=True" _
            & g_sql_exclu
    Else
        codutil = UCase(txt(TXT_CODE).Text)
        sql = "SELECT U_Num, U_Nom, U_Prenom, U_Labo" _
            & " FROM Utilisateur, UtilAppli" _
            & " WHERE U_kb_actif=True AND U_Num>1" _
            & " AND UAPP_UNum=U_Num" _
            & " AND UAPP_APPNum=" & p_appli_kalidoc _
            & " AND UAPP_Code=" & Odbc_String(codutil) _
            & " AND U_Actif=True" _
            & g_sql_exclu
    End If
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        verif_util = P_ERREUR
        Exit Function
    End If
    If Not rs.EOF Then
        rs.MoveNext
        If rs.EOF Then
            rs.MovePrevious
            p_siz_tblu_sel = 0
            ReDim p_tblu_sel(0) As Long
            p_tblu_sel(0) = rs("U_Num").Value
            rs.Close
            GoTo lab_fin
        Else
            rs.MovePrevious
            GoTo lab_affiche
        End If
    End If
    rs.Close
    If txt(TXT_NOM).Text <> "" Then
        nomutil = UCase(txt(TXT_NOM).Text)
'        If left$(nomutil, 1) <> "*" Then
'            nomutil = "*" + nomutil
'        End If
        If Right$(nomutil, 1) <> "*" Then
            nomutil = nomutil + "*"
        End If
        sql = "SELECT U_Num, U_Nom, U_Prenom, U_Labo" _
            & " FROM Utilisateur" _
            & " WHERE U_kb_actif=True AND U_Num>1" _
            & " AND " & Odbc_upper & "(U_Nom) LIKE " & Odbc_String(nomutil) _
            & " AND U_Actif=True" _
            & g_sql_exclu _
            & " ORDER BY U_Nom"
    Else
        codutil = UCase(txt(TXT_CODE).Text)
'        If left$(codutil, 1) <> "*" Then
'            codutil = "*" + codutil
'        End If
        If Right$(codutil, 1) <> "*" Then
            codutil = codutil + "*"
        End If
        sql = "SELECT U_Num, U_Nom, U_Prenom, U_Labo" _
            & " FROM Utilisateur, UtilAppli" _
            & " WHERE U_kb_actif=True AND U_Num>1" _
            & " AND UAPP_UNum=U_Num" _
            & " AND UAPP_APPNum=" & p_appli_kalidoc _
            & " AND UAPP_Code LIKE " & Odbc_String(codutil) _
            & " AND U_Actif=True" _
            & g_sql_exclu _
            & " ORDER BY U_Nom"
    End If
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        verif_util = P_ERREUR
        Exit Function
    End If
    If rs.EOF Then
        rs.Close
        verif_util = P_NON
        Exit Function
    End If

lab_affiche:
    Call CL_Init
    Call CL_InitTitreHelp("Personnes ayant le critère recherché", "")
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    n = 0
    nbtot = 0
    While Not rs.EOF
        nbtot = nbtot + 1
        ' Utilisateur fait partie des labos indiqués ?
        If g_ssite <> "" Then
            est_ok = False
            nbchp_u = STR_GetNbchamp(rs("U_Labo").Value, ";")
            nbchp_p = STR_GetNbchamp(g_ssite, ";")
            For i = 0 To nbchp_u - 1
                numlabo = Mid$(STR_GetChamp(rs("U_Labo").Value, ";", i), 2)
                For j = 0 To nbchp_p - 1
                    If numlabo = Mid$(STR_GetChamp(g_ssite, ";", j), 2) Then
                        est_ok = True
                        Exit For
                    End If
                Next j
            Next i
            If Not est_ok Then
                GoTo lab_suivant
            End If
        End If
        If recup_fsp(rs("U_Num").Value, nomutil, libfct, libspm) = P_ERREUR Then
            verif_util = P_ERREUR
            Exit Function
        End If
        Call CL_AddLigne(rs("U_Nom").Value & " " & rs("U_Prenom").Value & vbTab & libfct & vbTab & libspm, _
                         rs("U_Num").Value, _
                         "", _
                         False)
        n = n + 1
lab_suivant:
        rs.MoveNext
    Wend
    rs.Close

    If n = 0 Then
        Call MsgBox("Aucune personne faisant partie des sites indiqués n'a été trouvé avec les critères désirés.", _
                     vbInformation + vbOKOnly, "")
        verif_util = P_NON
        Exit Function
    End If

    ' Un seul utilisateur trouvé -> on le sélectionne d'office
'    If n = 1 And nbtot = 1 Then
'        p_siz_tblu_sel = 0
'        ReDim p_tblu_sel(0) As Long
'        p_tblu_sel(0) = CL_liste.lignes(0).num
'        GoTo lab_fin
'    End If

    Call CL_InitTaille(0, -20)
'    If g_plusieurs_util_autor Then
'        Call CL_InitMultiSelect(True)
'        ChoixListe.Show 1
'        If CL_liste.retour = 1 Then
'            verif_util = P_NON
'            Exit Function
'        End If
'        p_siz_tblu_sel = -1
'        For i = 0 To n - 1
'            If CL_liste.lignes(i).selected Then
'                p_siz_tblu_sel = p_siz_tblu_sel + 1
'                ReDim Preserve p_tblu_sel(p_siz_tblu_sel) As Long
'                p_tblu_sel(p_siz_tblu_sel) = CL_liste.lignes(i).num
'            End If
'        Next i
'    Else
        ' Ne pas supprimer : sinon txt_LostFocus reprend la main et ChoixListe
        ' est lancée une 2e fois ...
        g_mode_saisie = False
        ChoixListe.Show 1
        If CL_liste.retour = 1 Then
            verif_util = P_NON
            Exit Function
        End If
        p_siz_tblu_sel = 0
        ReDim p_tblu_sel(0) As Long
        p_tblu_sel(0) = CL_liste.lignes(CL_liste.pointeur).num
'    End If

lab_fin:
    g_scr = "U"
    Unload Me

    verif_util = P_OK

End Function

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case CMD_CHOIX_FONCTION
            Call choisir_fonction
        Case CMD_CHOIX_SERVICE
            Call choisir_service
        Case CMD_FERMER
            Call quitter
        Case CMD_OK
            If verif_util() <> P_OUI Then
                txt(TXT_NOM).Text = ""
                txt(TXT_CODE).Text = ""
                txt(TXT_NOM).SetFocus
            End If
    End Select

End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Index = CMD_FERMER Then g_mode_saisie = False

End Sub

Private Sub Form_Activate()

    If g_form_active Then Exit Sub

    g_form_active = True
    Call initialiser

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        Call quitter
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
   End If

End Sub

Private Sub Form_Load()

    g_mode_saisie = False
    g_form_active = False

End Sub

Private Sub txt_GotFocus(Index As Integer)

    g_txt_avant = txt(TXT_NOM).Text

End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        KeyCode = 0
        If Index = TXT_NOM Then
            Call choisir_utilisateur_fsu("", "")
        End If
    End If

End Sub

Private Sub txt_LostFocus(Index As Integer)

    If g_mode_saisie Then
        If g_txt_avant <> txt(Index).Text Then
            If Index = TXT_NOM Or Index = TXT_CODE Then
                If verif_util() <> P_OUI Then
                    txt(TXT_NOM).Text = ""
                    txt(TXT_CODE).Text = ""
                    txt(TXT_NOM).SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If

End Sub
