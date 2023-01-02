Attribute VB_Name = "MPowerPoint"
Option Explicit

Public Ppt_obj As PowerPoint.Application
Public Ppt_doc As Presentation

Public Const PPT_544_376 = 1
Public Const PPT_640_480 = 2
Public Const PPT_720_512 = 3
Public Const PPT_800_600 = 4
Public Const PPT_1024_768 = 5
Public Const PPT_1152_882 = 6
Public Const PPT_1152_900 = 7
Public Const PPT_1280_1024 = 8
Public Const PPT_1600_1200 = 9
Public Const PPT_1800_1440 = 10
Public Const PPT_1920_1200 = 11

Public Function PowPoint_AfficherDoc(ByVal v_nomdoc As String, _
                                  ByVal v_passwd As String, _
                                  ByVal v_fimprime As Boolean, _
                                  ByVal v_fmodif As Boolean) As Integer

    Dim encore As Boolean
    
    If PowPoint_Init() = P_ERREUR Then
        PowPoint_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Not FICH_FichierExiste(v_nomdoc) Then
        MsgBox "Impossible d'ouvrir " & v_nomdoc & vbCr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
        PowPoint_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    Ppt_obj.Visible = True
    If PowPoint_OuvrirDoc(v_nomdoc, _
                          v_passwd, _
                          Ppt_doc) = P_ERREUR Then
        PowPoint_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Ppt_obj.WindowState <> ppWindowMaximized Then
        Ppt_obj.WindowState = ppWindowMaximized
    End If
    
    encore = True
    Do
        Call SYS_Sleep(10)
        On Error Resume Next
        If Not Ppt_obj.Visible Then
            encore = False
        End If
    Loop Until Not encore
    On Error GoTo 0
    
    Set Ppt_obj = Nothing
    
    PowPoint_AfficherDoc = P_OK

End Function

Public Sub PowPoint_Imprimer(ByVal v_nomdoc As String, _
                             ByVal v_passwd As String, _
                             ByVal v_nbex As Integer)

    If PowPoint_Init() = P_ERREUR Then
        Exit Sub
    End If
    
    If PowPoint_OuvrirDoc(v_nomdoc, _
                          v_passwd, _
                          Ppt_doc) = P_ERREUR Then
        Exit Sub
    End If
    
    Ppt_doc.PrintOut (v_nbex)
    
    Ppt_obj.Quit
    Set Ppt_obj = Nothing
    
End Sub

Public Function PowPoint_Init()

    On Error GoTo err_create_obj
    Set Ppt_obj = CreateObject("powerpoint.application")
    On Error GoTo 0
'    Ppt_obj.Visible = msoTrue
    
    PowPoint_Init = P_OK
    Exit Function

err_create_obj:
    MsgBox "Impossible de créer l'objet POWER POINT." & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    PowPoint_Init = P_ERREUR
    Exit Function

End Function

Public Function PowPoint_OuvrirDoc(ByVal v_nomdoc As String, _
                               ByVal v_passwd As String, _
                               ByRef r_doc As Presentation) As Integer

    On Error GoTo err_open_ficr
    If Ppt_obj.Visible Then
        Set r_doc = Ppt_obj.Presentations.Open(FileName:=v_nomdoc)
    Else
        Set r_doc = Ppt_obj.Presentations.Open(FileName:=v_nomdoc, withwindow:=False)
    End If
    On Error GoTo 0
'    r_doc.Windows(1).Activate
   
    PowPoint_OuvrirDoc = P_OK
    Exit Function
    
err_open_ficr:
    MsgBox "Impossible d'ouvrir le fichier " & v_nomdoc & vbCr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    PowPoint_OuvrirDoc = P_ERREUR
    Exit Function
    
End Function




