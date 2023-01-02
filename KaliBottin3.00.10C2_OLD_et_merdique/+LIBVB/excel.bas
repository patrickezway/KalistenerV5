Attribute VB_Name = "Mexcel"
Option Explicit

Public Exc_obj As Excel.Application
Public Exc_doc As Workbook

Public Function Excel_AfficherDoc(ByVal v_nomdoc As String, _
                                  ByVal v_passwd As String, _
                                  ByVal v_fimprime As Boolean, _
                                  ByVal v_fmodif As Boolean) As Integer

    Dim encore As Boolean
    
    If Excel_Init() = P_ERREUR Then
        Excel_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Not FICH_FichierExiste(v_nomdoc) Then
        MsgBox "Impossible d'ouvrir " & v_nomdoc & vbCr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
        Excel_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    If Excel_OuvrirDoc(v_nomdoc, _
                       v_passwd, _
                       Exc_doc) = P_ERREUR Then
        Excel_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    Exc_obj.Visible = True
    If Exc_obj.WindowState <> xlMaximized Then
        Exc_obj.WindowState = xlMaximized
    End If
    If Exc_obj.ActiveWindow.WindowState <> xlMaximized Then
        Exc_obj.ActiveWindow.WindowState = xlMaximized
    End If
    
    encore = True
    Do
        Call SYS_Sleep(10)
        On Error Resume Next
        If Not Exc_obj.Visible Then
            encore = False
        End If
    Loop Until Not encore
    On Error GoTo 0
    
    Set Exc_obj = Nothing
    
    Excel_AfficherDoc = P_OK

End Function

Public Sub Excel_Imprimer(ByVal v_nomdoc As String, _
                          ByVal v_passwd As String, _
                          ByVal v_nbex As Integer)

    If Excel_Init() = P_ERREUR Then
        Exit Sub
    End If
    
    If Excel_OuvrirDoc(v_nomdoc, _
                       v_passwd, _
                       Exc_doc) = P_ERREUR Then
        Exit Sub
    End If
    
    Call Exc_doc.PrintOut(, , v_nbex)
    
    Exc_obj.Quit
    Set Exc_obj = Nothing
    
End Sub

Public Function Excel_Init()

    On Error GoTo err_create_obj
    Set Exc_obj = CreateObject("excel.application")
    On Error GoTo 0
    
    Excel_Init = P_OK
    Exit Function

err_create_obj:
    MsgBox "Impossible de créer l'objet EXCEL." & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    Excel_Init = P_ERREUR
    Exit Function

End Function

Public Function Excel_OuvrirDoc(ByVal v_nomdoc As String, _
                               ByVal v_passwd As String, _
                               ByRef r_doc As Workbook) As Integer

    On Error GoTo err_open_ficr
    Set r_doc = Exc_obj.Workbooks.Open(FileName:=v_nomdoc, _
                                        ReadOnly:=False, _
                                        Password:=v_passwd)
    On Error GoTo 0
    
    Excel_OuvrirDoc = P_OK
    Exit Function
    
err_open_ficr:
    MsgBox "Impossible d'ouvrir le fichier " & v_nomdoc & vbCr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    Excel_OuvrirDoc = P_ERREUR
    Exit Function
    
End Function


