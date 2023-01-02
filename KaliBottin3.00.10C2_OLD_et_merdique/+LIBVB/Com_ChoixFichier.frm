VERSION 5.00
Begin VB.Form Com_ChoixFichier 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Choix d'un fichier"
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
      Height          =   7485
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11925
      Begin VB.ListBox flb 
         ForeColor       =   &H00800000&
         Height          =   4935
         Left            =   5520
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   2160
         Width           =   5955
      End
      Begin VB.PictureBox pct 
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   1
         Left            =   4920
         Picture         =   "Com_ChoixFichier.frx":0000
         ScaleHeight     =   345
         ScaleWidth      =   375
         TabIndex        =   10
         Top             =   3600
         Width           =   375
      End
      Begin VB.PictureBox pct 
         BorderStyle     =   0  'None
         Height          =   345
         Index           =   0
         Left            =   1740
         Picture         =   "Com_ChoixFichier.frx":0876
         ScaleHeight     =   345
         ScaleWidth      =   375
         TabIndex        =   7
         Top             =   1440
         Width           =   375
      End
      Begin VB.DirListBox dlb 
         ForeColor       =   &H00800000&
         Height          =   5040
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   4455
      End
      Begin VB.DriveListBox drv 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1050
         TabIndex        =   4
         Top             =   1080
         Width           =   6135
      End
      Begin VB.Image img 
         Height          =   975
         Left            =   9000
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblComm 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   8325
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fichiers du répertoire"
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
         Left            =   5520
         TabIndex        =   9
         Top             =   1920
         Width           =   5955
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Répertoires du lecteur"
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
         TabIndex        =   8
         Top             =   1920
         Width           =   4455
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lecteur"
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
         Left            =   210
         TabIndex        =   6
         Top             =   1140
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Height          =   855
      Left            =   -30
      TabIndex        =   0
      Top             =   7350
      Width           =   11955
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
         Left            =   8970
         Picture         =   "Com_ChoixFichier.frx":10EC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Quitter"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
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
         Left            =   2040
         Picture         =   "Com_ChoixFichier.frx":16A5
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Sélectionner"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
End
Attribute VB_Name = "Com_ChoixFichier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Entrée : Drive|Path|Pattern
' Sortie : rien ou fichier sélectionné avec chemin complet

Const CMD_OK = 0
Const CMD_QUITTER = 1

Private g_selrep As Boolean
Private g_nomfich As String
Private g_pattern As String

Public Function AppelFrm(ByVal v_titre As String, _
                         ByVal v_drive As String, _
                         ByVal v_path As String, _
                         ByVal v_pattern As String, _
                         ByVal v_selrep As Boolean) As String
                         
    lblComm.Caption = v_titre
    
    If v_pattern <> "" Then
        g_pattern = v_pattern
    End If
    g_selrep = v_selrep
    If v_drive <> "" Then
        drv.Drive = v_drive
    End If
    If v_path <> "" Then
        dlb.Path = v_path
    End If
    
    Com_ChoixFichier.Show 1
    
    AppelFrm = g_nomfich
    
End Function

Private Sub initialiser()

    dlb.SetFocus
    
    If g_selrep Then
        flb.Enabled = False
    End If
    
End Sub

Private Sub quitter()

    g_nomfich = ""
    Unload Me

End Sub

Private Sub valider()

    Dim s As String
    
    s = dlb.List(dlb.ListIndex)
    If Right$(s, 1) = "\" Then
        s = left$(s, Len(s) - 1)
    End If
    If flb.List(flb.ListIndex) <> "" Then
        s = s + "\" + flb.List(flb.ListIndex)
    ElseIf Not g_selrep Then
        Call MsgBox("Vous devez sélectionner un fichier et non un répertoire.", vbOKOnly + vbInformation, "")
        dlb.SetFocus
        Exit Sub
    End If
    g_nomfich = s
    Unload Me
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
    Case CMD_OK
        Call valider
    Case CMD_QUITTER
        Call quitter
    End Select
    
End Sub

Private Sub dlb_change()

    Dim strCurrentPath As String, strFilename As String, sp As String
    Dim i As Integer, n As Integer
    
    If Right$(dlb.Path, 1) = "\" Then
        strCurrentPath = dlb.Path
    Else
        strCurrentPath = dlb.Path & "\"
    End If

    ' Clear the Listbox.
    flb.Clear

    ' Populate the Listbox with the file names.
    If g_pattern = "" Then
        strFilename = Dir(strCurrentPath)
        Do While strFilename <> ""
            flb.AddItem strFilename
            strFilename = Dir
        Loop
    Else
        n = STR_GetNbchamp(g_pattern, ";")
        For i = 0 To n - 1
            sp = STR_GetChamp(g_pattern, ";", i)
            strFilename = Dir(strCurrentPath & sp)
            Do While strFilename <> ""
                flb.AddItem strFilename
                strFilename = Dir
            Loop
        Next i
    End If
    
End Sub

Private Sub dlb_Click()

    dlb.Path = dlb.List(dlb.ListIndex)
    
End Sub

Private Sub drv_Change()

    On Error GoTo lab_no_drv
    dlb.Path = drv.List(drv.ListIndex)
    On Error GoTo 0
    Exit Sub
    
lab_no_drv:
    On Error GoTo 0
    MsgBox "Ce lecteur n'est pas disponible.", vbInformation + vbOKCancel, ""
    
End Sub

Private Sub flb_Click()

    ' Afficher un apperçu de l'image (si on a choisi une :))
    On Error Resume Next
    img.Stretch = True
    Set img.Picture = LoadPicture(dlb.List(dlb.ListIndex) & "\" & flb.List(flb.ListIndex))

End Sub

Private Sub flb_DblClick()

    Call valider
    
End Sub

Private Sub Form_Activate()

    Call initialiser
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        Call quitter
    ElseIf (KeyCode = vbKeyO And Shift = vbAltMask) Or KeyCode = vbKeyF1 Then
        KeyCode = 0
        Call valider
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Call quitter
    End If
    
End Sub
