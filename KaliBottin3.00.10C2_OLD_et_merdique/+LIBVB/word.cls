VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Public WithEvents doc As Document
Attribute doc.VB_VarHelpID = -1

Private Const S_Fichier = "&Fichier"
Private Const S_Imprimer = "&Imprimer"
Private Const S_Apercu = "&Aperçu avant impression"
Private Const S_ImprimerMID = "&Imprimer..."
Private Const S_Quitter = "&Quitter"
Private Const S_Retour = "&Retour à KaliDoc"

Public lafin As Boolean

Private g_bdisplay_recent As Boolean
Private g_tblcmd() As Integer

Private g_tblmnbar() As Integer
Private g_tblmnbar1() As Integer
Private g_tblstd() As Integer
Private g_id_quitter As Integer
Private g_cmd_standard As Boolean
Private g_cmd_formatting As Boolean

Public Function InitConfig(ByVal v_type As Integer, _
                           ByVal v_nomdot As String, _
                           ByVal v_bfin As Boolean) As Integer

    Dim nomb As String
    Dim i As Integer, nbcmd As Integer, n As Integer
    Dim v As Variant, vmnufich As Variant, v2 As Variant
    Dim cmdb As CommandBar
    
    lafin = False
    
    On Error GoTo lab_fin_err
    g_bdisplay_recent = App.DisplayRecentFiles
    App.DisplayRecentFiles = False
    
    nbcmd = 0
    Erase g_tblcmd()
    
    doc.AttachedTemplate = v_nomdot
    
    If v_type = 1 Then GoTo lab_fin
    
    i = 1
    For Each cmdb In App.CommandBars
        If Not cmdb.Visible Then GoTo lab_suiv_cmd
        nomb = cmdb.Name
        ' La barre de menu reste visible dans tous les cas
        If nomb = "Menu Bar" Then GoTo lab_suiv_cmd
        ' La standard barre reste visible dans le cas 2 seulement
        If nomb = "Standard" And v_type = 2 Then GoTo lab_suiv_cmd
        ReDim Preserve g_tblcmd(nbcmd) As Integer
        g_tblcmd(nbcmd) = i
        nbcmd = nbcmd + 1
        On Error Resume Next
        cmdb.Visible = False
        On Error GoTo 0
lab_suiv_cmd:
        i = i + 1
    Next cmdb
    
lab_fin:
    ' Ne demandera pas de réécrire le modèle global (Normal.dot)
    doc.Saved = True
    App.NormalTemplate.Saved = True
    doc.AttachedTemplate.Saved = True

    On Error GoTo 0
    InitConfig = P_OK
    Exit Function
    
lab_fin_err:
    MsgBox "Erreur InitConfig " & Err.Description & vbCr & vbLf & Err.Number, vbOKOnly + vbCritical
    On Error GoTo 0
    InitConfig = P_ERREUR
    
End Function

'Public Sub App_Quit()

'    On Error GoTo lab_fin
'    Set App = Nothing
'    If g_bfin Then End
'    On Error GoTo 0
'    Exit Sub
    
'lab_fin:
'    Exit Sub
    
'End Sub

Private Sub SaveConfig(ByVal v_type As Integer)

    Dim i As Integer, nbcmd As Integer, nbmnbar As Integer, nbmnbar1 As Integer, nbstd As Integer
    Dim v As Variant, vmnufich As Variant
    
    nbcmd = 0
    Erase g_tblcmd()
    nbmnbar = 0
    Erase g_tblmnbar()
    nbmnbar1 = 0
    Erase g_tblmnbar1()
    nbstd = 0
    Erase g_tblstd()
    
    ' Menu Fichier - Quitter devient Retour à KaliDoc
    g_id_quitter = 0
    Set v = App.CommandBars("Menu Bar").Controls
    If v.Count >= 1 Then
        Set v = v(1).Controls
        For i = 1 To v.Count
            If v(i).Caption = S_Quitter Then
                g_id_quitter = v(i).Id
                v(i).Caption = S_Retour
                Exit For
            End If
        Next i
    End If
        
    If v_type = 1 Then Exit Sub
    
    Set v = App.CommandBars
    For i = 1 To App.CommandBars.Count
        If Not v(i).Visible Then GoTo lab_suiv_cmd
        ' La barre de menu reste visible dans tous les cas
        If v(i).Name = "Menu Bar" Then GoTo lab_suiv_cmd
        ' La standard barre reste visible dans le cas 2 seulement
        If v(i).Name = "Standard" And v_type = 2 Then GoTo lab_suiv_cmd
        On Error Resume Next
        ReDim Preserve g_tblcmd(nbcmd) As Integer
        g_tblcmd(nbcmd) = i
        nbcmd = nbcmd + 1
        v(i).Visible = False
        On Error GoTo 0
lab_suiv_cmd:
    Next i
    
    ' Dans le menu, il n'y a que Fichier qui reste visible
    Set v = App.CommandBars("Menu Bar").Controls
    For i = 1 To v.Count
        If Not v(i).Visible Then GoTo lab_suiv_mnbar
        If v(i).Caption = S_Fichier Then GoTo lab_suiv_mnbar
        ReDim Preserve g_tblmnbar(nbmnbar) As Integer
        g_tblmnbar(nbmnbar) = i
        nbmnbar = nbmnbar + 1
        v(i).Visible = False
lab_suiv_mnbar:
    Next i
    
    If v_type = 2 Then
        ' Standard : bouton Impression et Aperçu seulement sont visibles
        Set v = App.CommandBars("Standard").Controls
        For i = 1 To v.Count
            If Not v(i).Visible Then GoTo lab_suiv_std
            If v(i).Caption = S_Imprimer Then GoTo lab_suiv_std
            If v(i).Caption = S_Apercu Then GoTo lab_suiv_std
            On Error Resume Next
            ReDim Preserve g_tblstd(nbstd) As Integer
            g_tblstd(nbstd) = i
            nbstd = nbstd + 1
            v(i).Visible = False
            On Error GoTo 0
lab_suiv_std:
        Next i
        ' Menu Fichier : Quitter, Impression et Aperçu seulement
        Set v = App.CommandBars("Menu Bar").Controls
        Set vmnufich = v(1).CommandBar.Controls
        For i = 1 To vmnufich.Count
            Set v = vmnufich(i)
            If Not v.Visible Then GoTo lab_suiv_mnbar1
            If v.Caption = S_Retour Then GoTo lab_suiv_mnbar1
            If v.Caption = S_Apercu Then GoTo lab_suiv_mnbar1
            If v.Caption = S_ImprimerMID Then GoTo lab_suiv_mnbar1
            On Error Resume Next
            ReDim Preserve g_tblmnbar1(nbmnbar1) As Integer
            g_tblmnbar1(nbmnbar1) = i
            nbmnbar1 = nbmnbar1 + 1
            v.Visible = False
            On Error GoTo 0
lab_suiv_mnbar1:
        Next i
    ElseIf v_type = 3 Then
        ' Menu Fichier : Quitter seulement
        Set v = App.CommandBars("Menu Bar").Controls
        Set vmnufich = v(1).CommandBar.Controls
        For i = 1 To vmnufich.Count
            Set v = vmnufich(i)
            If Not v.Visible Then GoTo lab_suiv_mnbar1b
            If v.Caption = S_Retour Then GoTo lab_suiv_mnbar1b
            On Error Resume Next
            ReDim Preserve g_tblmnbar1(nbmnbar1) As Integer
            g_tblmnbar1(nbmnbar1) = i
            nbmnbar1 = nbmnbar1 + 1
            v.Visible = False
            On Error GoTo 0
lab_suiv_mnbar1b:
        Next i
    End If
        
End Sub

Private Sub docClose1()

    Dim i As Integer, size As Integer
    Dim v As Variant, v2 As Variant
    
    If g_id_quitter <> 0 Then
        Set v = App.CommandBars("Menu Bar").Controls(1).Controls
        For i = 1 To v.Count
            Set v2 = v(i)
            If v2.Id = g_id_quitter Then
                v2.Caption = S_Quitter
                Exit For
            End If
        Next i
    End If
    
    On Error Resume Next
    
    size = -1
    size = UBound(g_tblcmd)
    For i = 0 To size
        App.CommandBars(g_tblcmd(i)).Visible = True
    Next i
    
    size = -1
    size = UBound(g_tblstd)
    Set v = App.CommandBars("Standard").Controls
    For i = 0 To size
        v(g_tblstd(i)).Visible = True
    Next i
    
    size = -1
    size = UBound(g_tblmnbar)
    Set v = App.CommandBars("Menu Bar").Controls
    For i = 0 To size
        v(g_tblmnbar(i)).Visible = True
    Next i
    
    size = -1
    size = UBound(g_tblmnbar1)
    Set v = App.CommandBars("Menu Bar").Controls
    Set v2 = v(1).CommandBar.Controls
    For i = 0 To size
        v2(g_tblmnbar1(i)).Visible = True
    Next i
    
    On Error GoTo 0
    
End Sub

Private Sub App_DocumentChange()

    lafin = True
    
End Sub

Private Sub App_Quit()

    lafin = True
    
End Sub

Private Sub Doc_Close()

    Dim i As Integer, size As Integer
    Dim v As Variant, v2 As Variant
    
    size = -1
    On Error Resume Next
    size = UBound(g_tblcmd)
    For i = 0 To size
        App.CommandBars(g_tblcmd(i)).Visible = True
    Next i
    
    App.DisplayRecentFiles = g_bdisplay_recent
    
    On Error GoTo 0
    
    lafin = True
    
End Sub

