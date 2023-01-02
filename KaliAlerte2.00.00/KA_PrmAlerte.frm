VERSION 5.00
Begin VB.Form KA_PrmAlerte 
   Caption         =   "KaliAlerte - V.2.00.00"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   5250
   Icon            =   "KA_PrmAlerte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H8000000B&
      Caption         =   "Identification"
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
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.CheckBox chk 
         BackColor       =   &H8000000B&
         Caption         =   "Lancer KaliAlerte au d�marrage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   9
         Top             =   2280
         Width           =   3375
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H8000000B&
         Caption         =   "M�moriser code et mot de passe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   8
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1140
         Width           =   1935
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   2040
         TabIndex        =   3
         Top             =   555
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         Index           =   0
         X1              =   240
         X2              =   4920
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lbl 
         BackColor       =   &H8000000A&
         Caption         =   "Mot de passe"
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
         Left            =   600
         TabIndex        =   2
         Top             =   1150
         Width           =   1455
      End
      Begin VB.Label lbl 
         BackColor       =   &H8000000A&
         Caption         =   "Code d'acc�s"
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
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame frmBtn 
      BackColor       =   &H00808080&
      Height          =   780
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Param�trer"
      Top             =   2640
      Width           =   5295
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "KA_PrmAlerte.frx":144A
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   2
         Left            =   2520
         Picture         =   "KA_PrmAlerte.frx":19A6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Choisir le fichier d'initialisation"
         Top             =   200
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         DisabledPicture =   "KA_PrmAlerte.frx":1F6D
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   480
         Picture         =   "KA_PrmAlerte.frx":24C9
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Enregistrer"
         Top             =   200
         Width           =   550
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   0
         Left            =   4320
         Picture         =   "KA_PrmAlerte.frx":2A32
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fermer KaliAlerte"
         Top             =   200
         Width           =   550
      End
   End
   Begin VB.Menu mnuFCT 
      Caption         =   "Liste des alertes"
      Begin VB.Menu mnuAction 
         Caption         =   "&Actions"
      End
      Begin VB.Menu mnuDem 
         Caption         =   "&Demandes"
      End
      Begin VB.Menu mnuAR 
         Caption         =   "&Accus� r�ception"
      End
      Begin VB.Menu mnuInfos 
         Caption         =   "&Informations ANCIEN"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKM 
         Caption         =   "&KaliMails"
      End
      Begin VB.Menu mnuClasseurs 
         Caption         =   "&Classeurs"
      End
      Begin VB.Menu mnuForm 
         Caption         =   "&Formulaires re�us"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFEI 
         Caption         =   "&Fiche d'�v�nement ind�sirables"
      End
   End
   Begin VB.Menu mPopUpSys 
      Caption         =   "&Systray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "&Propri�t�s"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAffichAlerte 
         Caption         =   "&Ouvrir ..."
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKD 
         Caption         =   "Acc�s KaliDoc"
      End
      Begin VB.Menu mnuKR 
         Caption         =   "Acc�s KaliGdR"
      End
      Begin VB.Menu mnuKW 
         Caption         =   "Acc�s KaliWeb"
      End
      Begin VB.Menu mnuKRP 
         Caption         =   "Acc�s KaliRP"
      End
      Begin VB.Menu mnuScanX 
         Caption         =   "Acc�s KaliScanX"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Quitter KaliAlerte"
      End
   End
End
Attribute VB_Name = "KA_PrmAlerte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SC_CLOSE = &HF060&
Private Const MF_BYCOMMAND = &H0&

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, _
    ByVal wFlags As Long) As Long

Private Const CMD_OK = 1
Private Const CMD_QUITTER = 0
Private Const CMD_PARAM = 2

Private Const CHK_SAVE = 0
Private Const CHK_START = 1

Private Const ACTION = 0
Private Const DEMANDE = 1
Private Const AR = 2
Private Const INFO = 3
Private Const KM = 4
Private Const CLASSEUR = 5
Private Const Form = 6
Private Const GDR = 7

Private Const TXT_CODE = 0
Private Const TXT_MDP = 1

Private g_nbconnex As Integer

Private Function afficher_infos() As Integer

    Dim scode As String
    Dim sMdP As String
    Dim sMdPClair As String
    Dim sTypeCrypt As String
    Dim i As Integer
    Dim lnumappli As Long
    Dim oMD5 As CMD5
    
    ' Pour le cryptage MD5
    Set oMD5 = New CMD5

    'Verification de l'identification de la personne
    
    If GetSetting(p_nomappli, "Section", "NumUtil") <> "" Then
        p_numUtil = GetSetting(p_nomappli, "Section", "NumUtil")
        If p_numUtil > 0 Then
            Me.WindowState = vbMinimized
            chk(CHK_SAVE).Value = 1
        
            'R�cup�ration des l'application
            If Odbc_RecupVal("SELECT APP_Num FROM Application WHERE APP_Code='KALIDOC'", lnumappli) = P_ERREUR Then
                afficher_infos = P_ERREUR
                Exit Function
            End If
        
            'Identifiant
            If Odbc_RecupVal("SELECT UAPP_Code FROM UtilAppli WHERE UAPP_APPNum=" & lnumappli & " AND UAPP_UNum=" & p_numUtil, scode) = P_ERREUR Then
                afficher_infos = P_ERREUR
                Exit Function
            End If
        
            txt(TXT_CODE).Text = scode
            ' decrypter le mot de passe
            If Odbc_RecupVal("SELECT UAPP_MotPasse,UAPP_TypeCrypt FROM UtilAppli WHERE UAPP_APPNum=" & lnumappli & " AND UAPP_UNum=" & p_numUtil, sMdP, sTypeCrypt) = P_ERREUR Then
                afficher_infos = P_ERREUR
                Exit Function
            End If
            If sTypeCrypt = "kalidoc" Or sTypeCrypt = "" Then
                sMdPClair = STR_Decrypter(sMdP)
            ElseIf sTypeCrypt = "kalidocnew" Then
                sMdPClair = STR_Decrypter_New(sMdP)
            ElseIf sTypeCrypt = "md5" Then
                sMdPClair = oMD5.MD5(txt(TXT_MDP).Text)
            End If
            txt(TXT_MDP).Text = sMdPClair
            'txt(TXT_MDP).Text = scode
        End If
    End If
    
    'Gestion des alertes
    p_salerte = GetSetting(p_nomappli, "Section", "Alerte")
    If p_salerte <> "" Then
                
        'ACTIONS
        If STR_GetChamp(p_salerte, ";", ACTION) = 1 Then
            mnuAction.Checked = True
        End If
        
        'DEMANDES
        If STR_GetChamp(p_salerte, ";", DEMANDE) = 1 Then
            mnuDem.Checked = True
        End If
        
        'AR
        If STR_GetChamp(p_salerte, ";", AR) = 1 Then
            mnuAR.Checked = True
        End If
        
        'INFOS
        If STR_GetChamp(p_salerte, ";", INFO) = 1 Then
            mnuInfos.Checked = True
        End If
        
        'KM
        If STR_GetChamp(p_salerte, ";", KM) = 1 Then
            mnuKM.Checked = True
        End If
        
        'CLASSEURS
        If STR_GetChamp(p_salerte, ";", CLASSEUR) = 1 Then
            mnuClasseurs.Checked = True
        End If
        
        'FORM
        If STR_GetChamp(p_salerte, ";", Form) = 1 Then
            mnuForm.Checked = True
        End If
        
        'KGDR
        If STR_GetChamp(p_salerte, ";", GDR) = 1 Then
            mnuFEI.Checked = True
        End If

    End If
    
    'Chargement au d�marrage
    If IsRunningOnStartup(p_nomappli) Then
        chk(CHK_START).Value = 1
    Else
        'MsgBox "is not running"
    End If
    
End Function

Private Function autor_fct() As Integer

    Dim sql As String
    Dim lnb As Long

    ' Verification de l'autorisation d'acc�s a KaliGdR
    sql = "SELECT Count(*) " _
        & "FROM Fctok_Util " _
        & "WHERE FU_Unum=" & p_numUtil _
        & " AND FU_FCTNum=" _
        & " (SELECT FCT_Num " _
        & " FROM Fonction " _
        & "WHERE FCT_Code='GDR_QUALIF')"
    
    If Odbc_Count(sql, lnb) = P_ERREUR Then
        autor_fct = P_ERREUR
        Exit Function
    End If

    If lnb > 0 Then
        mnuFEI.Visible = True
        mnuSep3.Visible = True
    End If

    autor_fct = P_OK
    
End Function

Private Sub initialiser()

    Dim frm As Form

    'V�rifier si application disponible
    If FICH_FichierExiste(p_chemin_appli & "\kalidoc.exe") Then
        mnuKD.Visible = True
    Else
        mnuKD.Visible = False
    End If
        
    If FICH_FichierExiste(p_chemin_appli & "\kaligdr.exe") Then
        mnuKR.Visible = True
    Else
        mnuKR.Visible = False
    End If

    If FICH_FichierExiste(p_chemin_appli & "\kaliRP.exe") Then
        mnuKRP.Visible = True
    Else
        mnuKRP.Visible = False
    End If
    
    If FICH_FichierExiste(p_chemin_appli & "\kaliScanX.exe") Then
        mnuScanX.Visible = True
    Else
        mnuScanX.Visible = False
    End If
    
    If afficher_infos = P_ERREUR Then
        Exit Sub
    End If

    If p_numUtil > 0 And p_nomini <> "" And p_salerte <> "" Then
        If autor_fct = P_ERREUR Then
            Exit Sub
        End If
    
        Set frm = KA_Alerte
        Call KA_Alerte.AppelFrm
        Set frm = Nothing
    End If

End Sub

Private Function maj_alerte() As Integer

    p_salerte = IIf(mnuAction.Checked, 1, 0) & ";" _
                & IIf(mnuDem.Checked, 1, 0) & ";" _
                & IIf(mnuAR.Checked, 1, 0) & ";" _
                & IIf(mnuInfos.Checked, 1, 0) & ";" _
                & IIf(mnuKM.Checked, 1, 0) & ";" _
                & IIf(mnuClasseurs.Checked, 1, 0) & ";" _
                & IIf(mnuForm.Checked, 1, 0) & ";" _
                & IIf(mnuFEI.Checked, 1, 0) & ";"
    
    SaveSetting p_nomappli, "Section", "Alerte", p_salerte

    Unload KA_Alerte
    
    If p_numUtil > 0 Then
        Call KA_Alerte.AppelFrm
    End If

    maj_alerte = P_OK

End Function

Private Sub selectionTextBox(ByVal v_otxt As Object)
    
    Dim LongueurTexte As Long

    On Error GoTo Errorhandler
    LongueurTexte = Len(v_otxt.Text)
    v_otxt.SelStart = 0
    v_otxt.SelLength = LongueurTexte
    
Errorhandler:
    Exit Sub

End Sub

Private Function verif_ident() As Integer

    Dim sql As String
    Dim rs As rdoResultset
    Dim appli_kalidoc As Integer, i As Integer
    Dim bad_util As Boolean
    Dim oMD5 As CMD5
    
    ' Pour le cryptage MD5
    Set oMD5 = New CMD5
    
    ' Initialisation de l'application
    sql = "select APP_Num from Application" _
        & " where APP_Code='KALIDOC'"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        verif_ident = P_ERREUR
        Exit Function
    End If
    If Not rs.EOF Then
        appli_kalidoc = rs("APP_Num").Value
    End If
    rs.Close
    
    'Recherche de cet utilisateur
    sql = "select U_Num, UAPP_MotPasse, UAPP_TypeCrypt from Utilisateur, UtilAppli" _
        & " where UAPP_Code='" & UCase(txt(TXT_CODE).Text) & "'" _
        & " and UAPP_APPNum=" & appli_kalidoc _
        & " and U_Actif=True" _
        & " and U_Num=UAPP_UNum"
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        verif_ident = P_ERREUR
        Exit Function
    End If
    If rs.EOF Then
        rs.Close
        bad_util = True
    Else
        If rs("UAPP_MotPasse").Value <> "" Then
            If rs("UAPP_TypeCrypt").Value = "kalidoc" Or rs("UAPP_TypeCrypt").Value = "" Then
                If STR_Decrypter(rs("UAPP_MotPasse").Value) <> UCase(txt(TXT_MDP).Text) Then
                    bad_util = True
                Else
                    p_numUtil = rs("U_Num").Value
                End If
            ElseIf rs("UAPP_TypeCrypt").Value = "kalidocnew" Then
                If STR_Decrypter_New(rs("UAPP_MotPasse").Value) <> UCase(txt(TXT_MDP).Text) Then
                    bad_util = True
                Else
                    p_numUtil = rs("U_Num").Value
                End If
            ElseIf rs("UAPP_TypeCrypt").Value = "md5" Then
                If rs("UAPP_MotPasse").Value <> oMD5.MD5(txt(TXT_MDP).Text) Then
                    bad_util = True
                Else
                    p_numUtil = rs("U_Num").Value
                End If
            Else
                p_numUtil = rs("U_Num").Value
                rs.Close
                Me.WindowState = vbMinimized
                GoTo lab_ok
            End If
        Else
            bad_util = False
        End If
        rs.Close
    End If
    If bad_util Then
        MsgBox "Identification inconnue.", vbOKOnly + vbExclamation, ""
        g_nbconnex = g_nbconnex + 1
        If g_nbconnex > 3 Then
            verif_ident = P_ERREUR
            End
            Exit Function
        End If
        txt(TXT_MDP).Text = ""
        GoTo Fin_Bad_Util
    End If
    
lab_ok:
    ' IDENTIFICATION AUTO
    If chk(CHK_SAVE).Value = 1 Then
        ' AJOUT DE P_NUMUTIL DANS LA BASE DE REGISTRE
        SaveSetting p_nomappli, "Section", "NumUtil", p_numUtil
    Else
        ' SUPPRESSION DANS LA BASE DE REGISTRE
        If GetSetting(p_nomappli, "Section", "NumUtil") <> "" Then
            DeleteSetting p_nomappli, "Section", "NumUtil"
        End If
    End If

    ' LANCEMENT AU CHARGEMENT
    If chk(CHK_START).Value = 1 Then
        ' Verification
        If Not IsRunningOnStartup(p_nomappli) Then
            ' Ajout
            Call RunAtStartUp(p_nomappli, App.Path & "\" & p_nomappli & ".exe " & p_nomini)
        End If
    Else
        If IsRunningOnStartup(p_nomappli) Then
            ' Suppression
            Call StopRunningStartUp(p_nomappli)
        End If
    End If
    
fin:
    
    If afficher_infos = P_ERREUR Then
        verif_ident = P_ERREUR
        Exit Function
    End If
    
    If autor_fct = P_ERREUR Then
        verif_ident = P_OK
        Exit Function
    End If
    
    verif_ident = P_OK
    Exit Function

Fin_Bad_Util:
    verif_ident = P_BAD_UTIL
    Exit Function

End Function

Private Sub cmd_Click(Index As Integer)
    Dim bVerif As Integer
    
    If Index = CMD_OK Then
        bVerif = verif_ident
        If bVerif = P_BAD_UTIL Then
            Me.txt(TXT_CODE).SetFocus
            Exit Sub
        ElseIf bVerif <> P_OK Then
            Unload Me
        End If
        Me.WindowState = vbMinimized
    ElseIf Index = CMD_PARAM Then
        If saisi_nomini = P_OUI Then
            txt(TXT_CODE).Text = ""
            txt(TXT_MDP).Text = ""
        End If
    Else
        Me.WindowState = vbMinimized
    End If

End Sub

Private Sub Form_Load()
    
    Dim hSysMenu As Long
    
    hSysMenu = GetSystemMenu(Me.hwnd, False)
    RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND
   
   
    'la forme doit �tre enti�rement visible avant d'appeler Shell_NotifyIcon
    Me.Show
    Me.Refresh
   
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "KaliAlerte" & vbNullChar
    End With
    
    Call initialiser
   
    ' Ajout dans la barre des taches
    Shell_NotifyIcon NIM_ADD, nid

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   
    'cette proc�dure re�oit les rappels de l'ic�ne de barre d'�tat syst�me.
    Dim msg As Long
    p_numUtil = 10
    If p_numUtil > 0 Then
        'the value of X will vary depending upon the scalemode setting
        If Me.ScaleMode = vbPixels Then
            msg = x
        Else
            msg = x / Screen.TwipsPerPixelX
        End If
        
        Select Case msg
            Case WM_LBUTTONUP                  '514 restore form window
                p_Result = SetForegroundWindow(Me.hwnd)
                Me.PopupMenu Me.mPopUpSys
            Case WM_LBUTTONDBLCLK     '515 restore form window
                p_Result = SetForegroundWindow(Me.hwnd)
                Me.PopupMenu Me.mPopUpSys
            Case WM_RBUTTONUP                  '517 display popup menu
                p_Result = SetForegroundWindow(Me.hwnd)
                Me.PopupMenu Me.mPopUpSys
        End Select
    End If
End Sub

Private Sub Form_Resize()

    'N�cessaire pour assurer que la fen�tre r�duite soit masqu�e
    If Me.WindowState = vbMinimized Then
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'cela supprime l'ic�ne de la barre d'�tat syst�me
    Unload KA_Alerte
    Shell_NotifyIcon NIM_DELETE, nid
    
End Sub

Private Sub mnuAction_Click()

    If mnuAction.Checked Then
        mnuAction.Checked = False
    Else
        mnuAction.Checked = True
    End If
    
    Call maj_alerte
    
End Sub

Private Sub mnuAffichAlerte_Click()

    'Affiche les alertes
    If p_salerte <> "" Then
        KA_Alerte.P_affiche_new
    Else
        Call MsgBox("Pensez � d�finir vos indicateurs", vbOKOnly + vbExclamation, "")
    End If

End Sub

Private Sub mnuAR_Click()

    If mnuAR.Checked Then
        mnuAR.Checked = False
    Else
        mnuAR.Checked = True
    End If
    
    Call maj_alerte
    
End Sub

Private Sub mnuClasseurs_Click()

    If mnuClasseurs.Checked Then
        mnuClasseurs.Checked = False
    Else
        mnuClasseurs.Checked = True
    End If
    
    Call maj_alerte
    
End Sub

Private Sub mnuDem_Click()

    If mnuDem.Checked Then
        mnuDem.Checked = False
    Else
        mnuDem.Checked = True
    End If
    
    Call maj_alerte
    
End Sub

Private Sub mnuFEI_Click()

    If mnuFEI.Checked Then
        mnuFEI.Checked = False
    Else
        mnuFEI.Checked = True
    End If
    
    Call maj_alerte
    
End Sub

Private Sub mnuForm_Click()

    If mnuForm.Checked Then
        mnuForm.Checked = False
    Else
        mnuForm.Checked = True
    End If
    
    Call maj_alerte
    
End Sub

Private Sub mnuInfos_Click()

    If mnuInfos.Checked Then
        mnuInfos.Checked = False
    Else
        mnuInfos.Checked = True
    End If
    
    Call maj_alerte
    
End Sub

Private Sub mnuKD_Click()

    Dim nomIni As String
    Dim cheminIni As String
    Dim cmd As String
    
    nomIni = SYS_GetIni("LOGICIEL", "KD", p_nomini)
    If nomIni = "" Then nomIni = "kalidoc.ini"
    cheminIni = p_chemin_appli & "\" & nomIni
    If Not FICH_FichierExiste(cheminIni) Then
        MsgBox cheminIni & " introuvable"
    Else
        cmd = p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";kalidoc;" & nomIni & ";CONNEXION=" & p_numUtil
        'MsgBox cmd
        Shell cmd, vbNormalFocus
    End If
End Sub

Private Sub mnuKM_Click()

    If mnuKM.Checked Then
        mnuKM.Checked = False
    Else
        mnuKM.Checked = True
    End If
    
    Call maj_alerte
    
End Sub

Private Sub mnuKR_Click()

    Dim nomIni As String
    Dim cheminIni As String
    Dim cmd As String
    
    nomIni = SYS_GetIni("LOGICIEL", "KGDR", p_nomini)
    If nomIni = "" Then nomIni = "kaligdr.ini"
    cheminIni = p_chemin_appli & "\" & nomIni
    If Not FICH_FichierExiste(cheminIni) Then
        MsgBox cheminIni & " introuvable"
    Else
        'cmd = p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";kalidoc;" & cheminIni & ";CONNEXION=" & p_numUtil
        cmd = p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";kaligdr;" & nomIni & ";" & p_numUtil
        'Shell p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";kaligdr;kaligdr.ini;" & p_numUtil, vbNormalFocus
        'MsgBox cmd
        Shell cmd, vbNormalFocus
    End If

End Sub

Private Sub mnuKRP_Click()
    Dim nomIni As String
    Dim cheminIni As String
    Dim cmd As String
    
    nomIni = SYS_GetIni("LOGICIEL", "KRP", p_nomini)
    If nomIni = "" Then nomIni = "kalidoc.ini"
    cheminIni = p_chemin_appli & "\" & nomIni
    If Not FICH_FichierExiste(cheminIni) Then
        MsgBox cheminIni & " introuvable"
    Else
        's = p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";kaliRP;" & tb(UBound(tb)) & ";CONNEXION=" & p_numUtil
        cmd = p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";kaliRP;" & nomIni & ";CONNEXION=" & p_numUtil
        'Shell p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";kaligdr;kaligdr.ini;" & p_numUtil, vbNormalFocus
        'MsgBox cmd
        Shell cmd, vbNormalFocus
    End If

End Sub

Private Sub mnuKW_Click()

    Dim snumutil As String
    
    snumutil = STR_CrypterNombre(Format(p_numUtil, "#0000000"))
    Shell "C:\Program Files\Internet Explorer\iexplore.exe " & p_cheminphp & "/pident.php?in=kaliweb&V_util=" & snumutil, vbMaximizedFocus

End Sub

Private Sub mnuScanX_Click()
    Dim nomIni As String
    Dim cheminIni As String
    Dim cmd As String
    
    nomIni = SYS_GetIni("LOGICIEL", "KScanX", p_nomini)
    If nomIni = "" Then nomIni = "kalidoc.ini"
    cheminIni = p_chemin_appli & "\" & nomIni
    If Not FICH_FichierExiste(cheminIni) Then
        MsgBox cheminIni & " introuvable"
    Else
        'Shell p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";kalidoc;" & tb(UBound(tb)) & ";CONNEXION=" & p_numUtil, vbNormalFocus
        cmd = p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";kaliScanX;" & nomIni & ";CONNEXION=" & p_numUtil
        'Shell p_chemin_appli & "\Lance.exe " & p_chemin_appli & ";kaligdr;kaligdr.ini;" & p_numUtil, vbNormalFocus
        'MsgBox cmd
        Shell cmd, vbNormalFocus
    End If
End Sub

Private Sub mPopExit_Click()
'appel�e quand l'utilisateur clique sur le menu contextuel Quitter
    
    Dim reponse As Integer
    
    reponse = MsgBox("Souhaitez-vous r�ellement quitter KaliAlerte ?" & vbLf & vbLf _
                        & "Vous n'aurez plus d'alertes !", vbQuestion + vbYesNo, "Quitter KaliAlerte ?")
    If reponse = vbYes Then
        Unload KA_Alerte
        Unload Me
    End If
    
End Sub

Private Sub mPopRestore_Click()

    'appel�e quand l'utilisateur clique sur le menu contextuel Agrandir
    Me.WindowState = vbNormal
    p_Result = SetForegroundWindow(Me.hwnd)
    Me.Show
    
 End Sub

Private Sub txt_GotFocus(Index As Integer)

    selectionTextBox txt(Index)

End Sub