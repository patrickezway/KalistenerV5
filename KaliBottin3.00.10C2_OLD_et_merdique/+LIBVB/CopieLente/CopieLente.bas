Attribute VB_Name = "MCopieLente"
Option Explicit

Public Sub Main()

    Dim scmd As String, nomfich_src As String, nomfich_dest As String
    Dim nomfich_dest_tmp As String
    Dim buf As String * 512
    Dim n As Integer, fdr As Integer, fdw As Integer, pos As Integer
    Dim i As Integer
    
    scmd = Command()
    n = STR_GetNbchamp(scmd, ";")
    If n < 2 Then
        Call MsgBox("Usage : CopieLente <Nomfichier source>;<Nomfichier destinatation>" & vbCr & vbLf _
                    & "scmd:" & scmd, vbInformation + vbOKOnly, "")
        End
    End If
    ' Fichier source
    nomfich_src = STR_GetChamp(scmd, ";", 0)
    If nomfich_src = "" Then
        Call MsgBox("'Nomfichier source' est vide." & vbCr & vbLf _
                    & scmd & vbCr & vbLf _
                    & "Usage : CopieLente <Nomfichier source>;<Nomfichier destination>", vbInformation + vbOKOnly, "")
        End
    End If
    ' Fichier destination
    nomfich_dest = STR_GetChamp(scmd, ";", 1)
    If nomfich_src = "" Then
        Call MsgBox("'Nomfichier destination' est vide." & vbCr & vbLf _
                    & scmd & vbCr & vbLf _
                    & "Usage : CopieLente <Nomfichier source>;<Nomfichier destination>", vbInformation + vbOKOnly, "")
        End
    End If
    
    ' Fichier source inaccessible
    If Not FICH_FichierExiste(nomfich_src) Then
        Call MsgBox("Impossible d'accéder à " & nomfich_src & ".", vbCritical + vbOKOnly, "")
        End
    End If
    
    fdr = FreeFile
    Open nomfich_src For Binary Access Read Shared As #fdr
    
    ' On change l'extension du fichier dest le temps du transfert
    i = 1
    pos = STR_InstrInverse(nomfich_dest, ".")
    Do
        If pos = 0 Then
            nomfich_dest_tmp = nomfich_dest & ".tmp" & i
        Else
            If InStr(Mid$(nomfich_dest, pos), "\") > 0 Then
                nomfich_dest_tmp = nomfich_dest & ".tmp" & i
            Else
                nomfich_dest_tmp = Left$(nomfich_dest, pos - 1) & ".tmp" & i
            End If
        End If
        i = i + 1
    Loop Until Not FICH_FichierExiste(nomfich_dest_tmp)
    
    fdw = FreeFile
    ' Le fichier est créé s'il n'existe pas
    Open nomfich_dest_tmp For Binary Access Write As #fdw
    
    Do While Not EOF(fdr)
        ' On efface le fichier dest s'il existe
        If FICH_FichierExiste(nomfich_dest) Then
            Call FICH_EffacerFichier(nomfich_dest)
        End If
        buf = Input(512, fdr)
        Put #fdw, , buf
        SYS_Sleep (1000)
    Loop
    Close #fdr
    Close #fdw
    
    Call FICH_EffacerFichier(nomfich_src)
    Call FICH_RenommerFichier(nomfich_dest_tmp, nomfich_dest)
    
    End
    
End Sub
