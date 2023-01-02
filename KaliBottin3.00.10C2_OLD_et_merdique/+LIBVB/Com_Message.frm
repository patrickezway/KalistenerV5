VERSION 5.00
Begin VB.Form Com_Message 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBidon 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   240
      Width           =   255
   End
   Begin VB.Frame frm 
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
      ForeColor       =   &H00800080&
      Height          =   2565
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7515
      Begin VB.Image img 
         Height          =   600
         Left            =   3510
         Picture         =   "Com_Message.frx":0000
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00800000&
         Height          =   1515
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   990
         Width           =   7035
      End
   End
   Begin VB.Frame frmfct 
      BackColor       =   &H00808080&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   2430
      Width           =   7515
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
         Height          =   615
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1635
      End
   End
End
Attribute VB_Name = "Com_Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private g_nomhelp As String

Private g_ind_esc As Integer
Private g_ret As Integer

Public Function AppelFrm(ByVal v_mess As Variant, _
                         ByVal v_nomhelp As String, _
                         ByRef v_tblcmd_libelle() As String, _
                         ByRef v_tblcmd_tooltip() As String) As Integer

    Dim n As Integer, I As Integer
    
    g_nomhelp = v_nomhelp
    
    'txt1.Text = v_mess
    lbl(1).Caption = v_mess
    
    g_ind_esc = -1
    
    n = -1
    On Error Resume Next
    n = UBound(v_tblcmd_libelle())
    On Error GoTo 0
    
    For I = 0 To n
        If I = 0 Then
            cmd(0).Visible = True
        Else
            Load cmd(I)
            cmd(I).Visible = True
            cmd(I).Top = cmd(0).Top
        End If
        Call init_bouton(I, v_tblcmd_libelle(I), v_tblcmd_tooltip(I))
    Next I
    Call aligner_boutons
    
    Com_Message.Show 1
    
    AppelFrm = g_ret
    
End Function

Private Sub aligner_boutons()

    Dim I As Integer, n As Integer, inter As Integer
    
    n = cmd.Count
    inter = (frmfct.width - cmd(0).left) / n
    For I = 1 To n - 1
        cmd(I).left = cmd(0).left + (I * inter)
    Next I
    
End Sub

Private Sub init_bouton(ByVal v_indcmd As Integer, _
                        ByVal v_libelle As String, _
                        ByVal v_tooltip As String)
                        
    If v_libelle = "Quitter" Then
        g_ind_esc = v_indcmd
        cmd(v_indcmd).Picture = CM_LoadPicture(p_chemin_appli + "\btnporte.gif")
        cmd(v_indcmd).Caption = ""
    Else
        cmd(v_indcmd).Picture = CM_LoadPicture("")
        cmd(v_indcmd).Caption = v_libelle
    End If
    cmd(v_indcmd).ToolTipText = v_tooltip
    
End Sub

Private Sub initialiser()

    txtBidon.SetFocus
    
End Sub

Private Sub quitter(ByVal v_index As Integer)

    g_ret = v_index
    
    Unload Me

End Sub

Private Sub cmd_Click(Index As Integer)

    Call quitter(Index)
    
End Sub

Private Sub Form_Activate()

    Call initialiser
    txtBidon.ZOrder 1
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyH And Shift = vbAltMask Then
        KeyCode = 0
        Call HtmlHelp(0, STR_GetChamp(g_nomhelp, ";", 0), HH_DISPLAY_TOPIC, STR_GetChamp(g_nomhelp, ";", 1))
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        If g_ind_esc >= 0 Then
            g_ret = g_ind_esc
            Unload Me
        End If
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Call quitter(cmd.Count - 1)
    End If
    
End Sub

