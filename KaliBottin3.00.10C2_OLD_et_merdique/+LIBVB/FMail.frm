VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form FMail 
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
   Begin MSMAPI.MAPIMessages MAPIMessage 
      Left            =   2400
      Top             =   1410
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession 
      Left            =   540
      Top             =   1470
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "FMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const conMailLongDate = 0
Private Const conMailListView = 1

' Constante pour le type de la boîte de dialogue d'options - Options générales.
Private Const conOptionGeneral = 1
' Constante pour le type de la boîte de dialogue d'options - Options du message.
Private Const conOptionMessage = 2
' Constante pour chaîne réprésentent les messages non lus
Private Const conUnreadMessage = "*"

Private Const vbRecipTypeTo = 1
Private Const vbRecipTypeCc = 2

Private Const vbMessageFetch = 1
Private Const vbMessageSendDlg = 2
Private Const vbMessageSend = 3
Private Const vbMessageSaveMsg = 4
Private Const vbMessageCopy = 5
Private Const vbMessageCompose = 6
Private Const vbMessageReply = 7
Private Const vbMessageReplyAll = 8
Private Const vbMessageForward = 9
Private Const vbMessageDelete = 10
Private Const vbMessageShowAdBook = 11
Private Const vbMessageShowDetails = 12
Private Const vbMessageResolveName = 13
Private Const vbRecipientDelete = 14
Private Const vbAttachmentDelete = 15

Private Const vbAttachTypeData = 0
Private Const vbAttachTypeEOLE = 1
Private Const vbAttachTypeSOLE = 2

Private Type ListDisplay
    Name As String * 20
    Subject As String * 40
    Date As String * 20
End Type

Private currentRCIndex As Integer
Private UnRead As Integer
Private SendWithMapi As Integer
Private ReturnRequest As Integer
Private OptionType As Integer

Private Function ConnectMAPI() As Integer
    
    Dim nb_mess As Integer, nb_new_mess As Integer, i As Integer
    
    On Error GoTo lab_erreur
    MAPISession.action = 1
    MAPISession.DownLoadMail = False
    MAPIMessage.SessionID = MAPISession.SessionID
    MAPIMessage.FetchUnreadOnly = False
'    MAPIMessage.Fetch
'    nb_mess = MAPIMessage.MsgCount - 1
'    For i = 0 To nb_mess
'        MAPIMessage.MsgIndex = i
'        If Not MAPIMessage.MsgRead Then
'            If InStr(MAPIMessage.MsgSubject, "KaliDoc") > 0 Then
'                nb_new_mess = nb_new_mess + 1
'            End If
'        End If
'    Next i
'    If nb_new_mess > 0 Then
'        MsgBox "Vous avez " & Format$(MAPIMessage.MsgCount) + " Nouveaux Messages"
'    End If
    
    ConnectMAPI = 1
    On Error GoTo 0
    Exit Function
    
lab_erreur:
    If Err = 32050 Then
        Resume Next
    ElseIf Err = 32003 Then
        ' Échec de la connexion
        Call MsgBox("Echec lors de la connexion Mail", vbOKOnly + vbCritical, "")
        ConnectMAPI = 0
        On Error GoTo 0
        Exit Function
    Else
        Call MsgBox("Erreur lors de la connexion Mail : " & Err.Description & " (" & Err.Number & ")", vbOKOnly + vbCritical, "")
        ConnectMAPI = 0
        On Error GoTo 0
        Exit Function
    End If
    
End Function

Public Function EnvoiMessage(ByVal v_nomdest As String, _
                             ByVal v_adrdest As String, _
                             ByVal v_subject As String, _
                             ByVal v_note As String, _
                             ByVal v_filename As String, _
                             ByVal v_LibFilename As String, _
                             ByVal v_BackGround As Boolean) As Integer
    Dim action As Integer
    
    On Error GoTo lab_erreur
lab_Compose:
    MAPIMessage.action = vbMessageCompose
    MAPIMessage.MsgNoteText = v_note
    ' Pas de message "Lu" renvoyé à l'émetteur qd le destinataire a ouvert le message
    MAPIMessage.MsgReceiptRequested = False
    MAPIMessage.MsgSubject = v_subject
    MAPIMessage.RecipAddress = v_adrdest
'MAPIMessage.RecipAddress = "ali@kalitech.fr"
    MAPIMessage.ResolveName
    
    ' Fichier joint
    If Len(v_filename) > 0 Then
        If Len(Dir$(v_filename)) > 0 Then
            MAPIMessage.AttachmentIndex = 0
            MAPIMessage.AttachmentName = v_LibFilename
            MAPIMessage.AttachmentPathName = v_filename
            MAPIMessage.AttachmentType = 2
        End If
    End If
    
    If v_BackGround Then
        action = vbMessageSend
    Else
        action = vbMessageSendDlg
    End If
    MAPIMessage.action = action
    ' ou mapimessage.send not v_background
    
lab_fin:
    On Error GoTo 0
    EnvoiMessage = 0
    Exit Function

lab_erreur:
    If Err = 32053 Then
        If ConnectMAPI() = 0 Then
            ' pas la peine de continuer
            On Error GoTo 0
            EnvoiMessage = -1
            Exit Function
        End If
        Resume Next
    ElseIf Err = 32051 Then
        ' La propriété est en lecture seule lorsque le tampon de
        ' composition n'est pas utilisé. Définissez MsgIndex = -1
        Resume lab_Compose
    ElseIf Err = 32014 Then
        ' Destinataire inconnu
        Resume Next
    Else
        Call MsgBox(Err & " " & Error$, vbCritical + vbOKOnly, "")
        Resume Next
    End If
    
End Function

Private Sub Attachments(Msg As Form)
    
    Dim s As String
    Dim i As Integer
    
    ' Efface la liste des pièces jointes en cours.
    Msg.aList.Clear

    ' S'il y a des pièces jointes, les charge dans la zone de liste.
    If MAPIMessage.AttachmentCount Then
        Msg.NumAtt = MAPIMessage.AttachmentCount & " fichiers"
        For i = 0 To MAPIMessage.AttachmentCount - 1
            MAPIMessage.AttachmentIndex = i
            s = MAPIMessage.AttachmentName
            Select Case MAPIMessage.AttachmentType
            Case vbAttachTypeData
                s = s + " (Fichier de données)"
            Case vbAttachTypeEOLE
                s = s + " (Objet OLE incorporé)"
            Case vbAttachTypeSOLE
                s = s + " (Objet OLE statique)"
            Case Else
                s = s + " (Type de pièce jointe inconnu)"
            End Select
            Msg.aList.AddItem s
        Next i
        
        If Not Msg.AttachWin.Visible Then
            Msg.AttachWin.Visible = True
            Call SizeMessageWindow(Msg)
            ' If Msg.WindowState = 0 Then
            '    Msg.Height = Msg.Height + Msg.AttachWin.Height
            ' End If
        End If
    
    Else
        If Msg.AttachWin.Visible Then
            Msg.AttachWin.Visible = False
            Call SizeMessageWindow(Msg)
            ' If Msg.WindowState = 0 Then
            '    Msg.Height = Msg.Height - Msg.AttachWin.Height
            ' End If
        End If
    End If
    Msg.Refresh
End Sub

Private Sub CopyNamestoMsgBuffer(Msg As Form, fResolveNames As Integer)

    Call KillRecips(MAPIMessage)
    Call SetRCList(Msg.txtTo, MAPIMessage, vbRecipTypeTo, fResolveNames)
    Call SetRCList(Msg.txtcc, MAPIMessage, vbRecipTypeCc, fResolveNames)

End Sub

' Cette procédure formate une date MAPI dans un
' des deux formats pour visualiser le message.
Private Function DateFromMapiDate(ByVal v_str As String, _
                                  ByVal v_format As Integer) As String
    
    Dim sa As String, sm As String, sj As String, sh As String
    Dim sf As String
    Dim dates As Date
    
    sa = left$(v_str, 4)
    sm = Mid$(v_str, 6, 2)
    sj = Mid$(v_str, 9, 2)
    sh = Mid$(v_str, 12)
    dates = DateValue(sm + "/" + sj + "/" + sa) + TimeValue(sh)
    Select Case v_format
    Case conMailLongDate
        sf = "dddd d mmmm yyyy, hh:mm"
    Case conMailListView
        sf = "dd/mm/yy hh:mm"
    End Select
    
    DateFromMapiDate = Format$(dates, sf)

End Function

Private Sub DisplayAttachedFile(ByVal v_filename As String)
    
    On Error Resume Next
    
    ' Détermine l'extension du nom de fichier
'        ext$ = FileName
        ' Obtient la feuille de l'application depuis le fichier WIN.INI.
'        buffer$ = String$(256, " ")
        'errCode% = GetProfileString("Extensions", ext$, "NOTFOUND", buffer$, Len(Left(buffer$, Chr(0)) - 1))
'        If errCode% Then
'            buffer$ = Mid$(buffer$, 1, InStr(buffer$, Chr(0)) - 1)
'            If buffer$ <> "NOTFOUND" Then
'                ' Enlève l'information .EXT de la chaîne.
'                EXEName$ = Token$(buffer$, " ")
'                errCode% = Shell(EXEName$ + " " + FileName, 1)
'                If Err Then
'                    MsgBox "Une erreur s'est produite durant l'instruction Shell: " + Error$
'                End If
'            Else
'                MsgBox "L'application qui utilise l'extension <" + ext$ + "> n'a pas été trouvée dans WIN.INI"
'            End If
'        End If
End Sub

Private Function GetHeader(Msg As Control) As String

    Dim CR As String
    
'    CR = Chr$(13) + Chr$(10)
'
'      Header$ = String$(25, "-") + CR
'      Header$ = Header$ + "De: " + Msg.MsgOrigDisplayName + CR
'      Header$ = Header$ + "A: " + GetRCList(Msg, vbRecipTypeTo) + CR
'      Header$ = Header$ + "Cc: " + GetRCList(Msg, vbRecipTypeCc) + CR
'      Header$ = Header$ + "Sujet: " + Msg.MsgSubject + CR
'      Header$ = Header$ + "Date: " + DateFromMapiDate$(Msg.MsgDateReceived, conMailLongDate) + CR + CR
'      GetHeader = Header$
End Function

' Lit tous les messages de la Messagerie et affiche le compteur.
Private Sub GetMessageCount()
    
    Screen.MousePointer = 11
    MAPIMessage.FetchUnreadOnly = 0
    MAPIMessage.action = vbMessageFetch
    MsgBox Format$(MAPIMessage.MsgCount) + " Messages"
    Screen.MousePointer = 0

End Sub

' En donnant une liste de destinataires, cette fonction retourne
' une liste de destinataires avec le type spécifié dans le format
' suivant: Personne 1; Personne 2; Personne 3
Private Function GetRCList(Msg As Control, RCType As Integer) As String

    Dim s As String
    Dim i As Integer
    
    s = ""
    For i = 0 To Msg.RecipCount - 1
        Msg.RecipIndex = i
        If RCType = Msg.RecipType Then
            s = s + ";" + Msg.RecipDisplayName
        End If
    Next i
    ' Enlève le ";" final
    If s <> "" Then
       s = left$(s, Len(s) - 1)
    End If
    
    GetRCList = s
    
End Function

Private Sub KillRecips(MsgControl As Control)
    
    ' Supprime chaque destinataire. Itération en boucle jusqu'à ne plus avoir de destinataires.
    While MsgControl.RecipCount
        MsgControl.action = vbRecipientDelete
    Wend
    
End Sub

Private Function Logon_Mail(MAPIMess As Control, MAPISess As Control)
    
    ' Se connecte à la messagerie.
    On Error GoTo Err
    MAPISess.action = 1
    If Err <> 0 Then
        MsgBox "Échec de connexion: " + Error$
    Else
        Screen.MousePointer = 11
        MAPIMess.SessionID = MAPISess.SessionID
        'GoTo suite
        ' Obtient le nombre de messages.
        GetMessageCount
        ' Charge la liste des messages avec l'information enveloppe.
        Screen.MousePointer = 11
        Call LoadList(MAPIMessage)
Suite:          ' pour l'instant hfhf
        Screen.MousePointer = 0
      End If
Fin:
    On Error GoTo 0
    Exit Function
Err:
    If Err = 32050 Then
        Resume Fin
    Else
        MsgBox Err & " " & Error$
        Resume Next
    End If
    
End Function

Private Function Send_Mail(MAPIMess As Control, MAPISess As Control)
    'Private Sub SendCtl_Click(Index As Integer)
    'Dim NewMessage As New NewMsg
    Dim Adr, TS, TL As String
    On Error Resume Next

    ' Index = 6: Composer un nouveau message.
    '       = 7: Répondre.
    '       = 8: Répondre à tous.
    '       = 9: Renvoyer.

    ' Enregistre l'information de l'en-tête et le texte en cours du message.
    ''If Index > 6 Then
    ''    ' SVNote = GetHeader(mapimessage) + mapimessage.MsgNoteText
    ''    SVNote = mapimessage.MsgNoteText
    ''    SVNote = GetHeader(mapimessage) + SVNote
    ''End If

    MAPIMess.action = 6

    ' Définit le nouveau texte du message.
    ''If Index > 6 Then
    ''    mapimessage.MsgNoteText = SVNote
    ''End If

    If SendWithMapi Then
        '
        'strmail = "Adr=csth@ccml.com|TS=Coucou|TL=quiquitruc"
        'Adr = recup_valeurs_spécifiques("Adr", strmail)
        'TS = recup_valeurs_spécifiques("TS", strmail)
        'TL = recup_valeurs_spécifiques("TL", strmail)
        '
        MAPIMess.MsgSubject = TS
        MAPIMess.MsgNoteText = TL
        'mapimessage.RecipType = 1
        If Len(Trim(Adr)) > 0 Then
            '
            MAPIMess.AddressResolveUI = True
            MAPIMess.RecipAddress = Adr
            MAPIMess.ResolveName
            
            MsgBox MAPIMess.MsgSubject
            MsgBox MAPIMess.MsgNoteText
            MsgBox MAPIMess.AddressResolveUI
            MsgBox MAPIMess.RecipAddress
            MsgBox MAPIMess.ResolveName
        
        End If

        '
        MAPIMess.action = vbMessageSendDlg
    Else
        ''Call LoadMessage(-1, NewMessage)            ' Charge le message dans la fenêtre VBMail.NewMSG.
    End If
    
End Function

' Cette procédure charge les en-têtes des messages de la Messagerie
' dans la liste MailLst.MList. Les messages non lus ont un caractère
' conUnreadMessage placé au début de la chaîne.
Private Sub LoadList(mailctl As Control)
    
'    FMail.MailLst.Clear
'    UnRead = 0
'    StartIndex = 0
'    For i = 0 To mailctl.MsgCount - 1
'        mailctl.MsgIndex = i
'        If Not mailctl.MsgRead Then
'            a$ = conUnreadMessage + " "
'            If UnRead = 0 Then
'                StartIndex = i  ' Position de départ de la liste des messages.
'            End If
'            UnRead = UnRead + 1
'        Else
'            a$ = "  "
'        End If
'        a$ = a$ + Mid$(Format$(mailctl.MsgOrigDisplayName, "!" + String$(10, "@")), 1, 10)
'        If mailctl.MsgSubject <> "" Then
'            b$ = Mid$(Format$(mailctl.MsgSubject, "!" + String$(35, "@")), 1, 35)
'        Else
'            b$ = String$(30, " ")
'        End If
'        c$ = Mid$(Format$(DateFromMapiDate(mailctl.MsgDateReceived, conMailListView), "!" + String$(15, "@")), 1, 15)
'        FMail.MailLst.AddItem a$ + Chr$(9) + b$ + Chr$(9) + c$
        'MsgBox a$ + Chr$(9) + b$ + Chr$(9) + c$
        'MailLst.MList.Refresh
'    Next i

    'FMail.MailLst.ListIndex = StartIndex
'    Exit Sub
    ' Active les boutons correspondants.
'    VBMail.Next.Enabled = True
'    VBMail.Previous.Enabled = True
'    VBMail![Delete].Enabled = True'

    ' Ajuste la valeur des étiquettes affichant le compteur de messages.
'    If UnRead Then
'        VBMail.UnreadLbl = " - " + Format$(UnRead) + " Non lu"
'        MailLst.Icon = MailLst.NewMail.Picture
'    Else
'        VBMail.UnreadLbl = ""
'        MailLst.Icon = MailLst.nonew.Picture
'    End If

End Sub
    
Private Sub LogOffUser()
    
'    On Error Resume Next
'    MAPISession.action = 2
'    If Err <> 0 Then
'        MsgBox "Echec de la connexion: " + Error
'    Else
'        MAPIMessage.SessionID = 0
'        ' Ajuste les éléments du menu.
'        VBMail.LogOff.Enabled = 0
'        VBMail.Logon.Enabled = -1
'        ' Décharge toutes les feuilles, sauf la feuille MDI principale.
'        Do Until Forms.Count = 1
'            i = Forms.Count - 1
'            If TypeOf Forms(i) Is MDIForm Then
'                ' Ne rien faire.
'            Else
'                Unload Forms(i)
'            End If
'        Loop
'        ' Désactive les boutons de la barre d'outils.
'        VBMail.Next.Enabled = False
'        VBMail.Previous.Enabled = False
'        VBMail![Delete].Enabled = False
'        VBMail.SendCtl(vbMessageCompose).Enabled = False
'        VBMail.SendCtl(vbMessageReplyAll).Enabled = False
'        VBMail.SendCtl(vbMessageReply).Enabled = False
'        VBMail.SendCtl(vbMessageForward).Enabled = False
'        VBMail.rMsgList.Enabled = False
'        VBMail.PrintMessage.Enabled = False
'        VBMail.DispTools.Enabled = False
'        VBMail.EditDelete.Enabled = False
'
'        ' Réinitialise les étiquettes de la barre d'état.
'        VBMail.MsgCountLbl = "Hors ligne"
'        VBMail.UnreadLbl = ""
'    End If

End Sub

' Pour une liste de destinataires donnée, sous la forme
'       Personne 1; Personne 2; Personne 3
' place les noms dans les structures Msg.Recip.
Private Sub SetRCList(ByVal NameList As String, Msg As Control, RCType As Integer, fResolveNames As Integer)
    
    Dim i As Integer
    
    If NameList = "" Then
        Exit Sub
    End If

    i = Msg.RecipCount
    Do
        Msg.RecipIndex = i
        Msg.RecipDisplayName = Trim$(Token(NameList, ";"))
        If fResolveNames Then
            Msg.action = vbMessageResolveName
        End If
        Msg.RecipType = RCType
        i = i + 1
    Loop Until (NameList = "")

End Sub

Private Sub SizeMessageWindow(MsgWindow As Form)
    
    Dim MinSize As Long, X As Long
    
    If MsgWindow.WindowState <> 1 Then
        ' Détermine la taille minimum de la fenêtre en se basant
        ' sur la visibilité d'AttachWin (fenêtre de pièce jointe).
        If MsgWindow.AttachWin.Visible Then    ' Fenêtre de pièce jointe.
            MinSize = 3700
        Else
            MinSize = 3700 - MsgWindow.AttachWin.Height
        End If

        ' Maintient la taille minimum de la feuille.
        If MsgWindow.Height < MinSize And (MsgWindow.WindowState = 0) Then
            MsgWindow.Height = MinSize
            Exit Sub

        End If
        ' Ajuste la taille de la zone de texte.
        If MsgWindow.ScaleHeight > MsgWindow.txtNoteText.Top Then
            If MsgWindow.AttachWin.Visible Then
                X = MsgWindow.AttachWin.Height
            Else
                X = 0
            End If
            MsgWindow.txtNoteText.Height = MsgWindow.ScaleHeight - MsgWindow.txtNoteText.Top - X
            MsgWindow.txtNoteText.width = MsgWindow.ScaleWidth
        End If
    End If

End Sub

Private Function Token$(tmp$, search$)

'    X = InStr(1, tmp$, search$)
'    If X Then
'       Token$ = Mid$(tmp$, 1, X - 1)
'       tmp$ = Mid$(tmp$, X + 1)
'    Else
'       Token$ = tmp$
'       tmp$ = ""
'    End If

End Function

' Cette procédure met à jour les champs d'édition corrects et
' l'information de destinataire.
Private Sub UpdateRecips(Msg As Form)
    
    Msg.txtTo.Text = GetRCList(MAPIMessage, vbRecipTypeTo)
    Msg.txtcc.Text = GetRCList(MAPIMessage, vbRecipTypeCc)

End Sub

