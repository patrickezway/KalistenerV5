Attribute VB_Name = "Mail"
Option Explicit
Public Const gMessageClass = ""
'-----------------------------------------------------------------------
'               Copyright (C) 1993 Microsoft Corporation
'
' You have a royalty-free right to use, modify, reproduce and distribute
' the Sample Application Files (and/or any modified version) in any way
' you find useful, provided that you agree that Microsoft has no warranty,
' obligations or liability for any Sample Application Files.
'
' -----------------------------------------------------------------------

'***************************************************
'Application specific globals
'***************************************************
Global gMAPISession As Long

'***************************************************
'   MAPI Message holds information about a message
'***************************************************

Type MAPIMessage
    Reserved As Long
    Subject As String
    NoteText As String
    MessageType As String
    DateReceived As String
    ConversationID As String
    flags As Long
'    Originator As Long
    RecipCount As Long
    FileCount As Long
End Type


'************************************************
'   MAPIRecip holds information about a message
'   originator or recipient
'************************************************

Type MAPIRecip
    Reserved As Long
    RecipClass As Long
    Name As String
    Address As String
    EIDSize As Long
    EntryID As String
End Type


'******************************************************
'   MapiFile holds information about file attachments
'******************************************************

Type MAPIFile
    Reserved As Long
    flags As Long
    Position As Long
    PathName As String
    FileName As String
    FileType As String
End Type


'***************************
'   FUNCTION Declarations
'***************************

Declare Function MAPISendMail Lib "MAPI32.DLL" Alias "BMAPISendMail" (ByVal Session&, ByVal UIParam&, message As MAPIMessage, Recipient() As MAPIRecip, File() As MAPIFile, ByVal flags&, ByVal Reserved&) As Long
Declare Function MAPILogoff Lib "MAPI32.DLL" (ByVal Session&, ByVal UIParam&, ByVal flags&, ByVal Reserved&) As Long
Declare Function MAPILogon Lib "MAPI32.DLL" (ByVal UIParam&, ByVal User$, ByVal Password$, ByVal flags&, ByVal Reserved&, Session&) As Long
Declare Function BMAPIAddress Lib "MAPI32.DLL" (lInfo&, ByVal Session&, ByVal UIParam&, Caption$, ByVal nEditFields&, label$, nRecipients&, Recip() As MAPIRecip, ByVal flags&, ByVal Reserved&) As Long
Declare Function BMAPIGetAddress Lib "MAPI32.DLL" (ByVal lInfo&, ByVal nRecipients&, Recipients() As MAPIRecip) As Long


'**************************
'   CONSTANT Declarations
'**************************
'

Global Const SUCCESS_SUCCESS = 0
Global Const MAPI_USER_ABORT = 1
Global Const MAPI_E_FAILURE = 2
Global Const MAPI_E_LOGIN_FAILURE = 3
Global Const MAPI_E_DISK_FULL = 4
Global Const MAPI_E_INSUFFICIENT_MEMORY = 5
Global Const MAPI_E_BLK_TOO_SMALL = 6
Global Const MAPI_E_TOO_MANY_SESSIONS = 8
Global Const MAPI_E_TOO_MANY_FILES = 9
Global Const MAPI_E_TOO_MANY_RECIPIENTS = 10
Global Const MAPI_E_ATTACHMENT_NOT_FOUND = 11
Global Const MAPI_E_ATTACHMENT_OPEN_FAILURE = 12
Global Const MAPI_E_ATTACHMENT_WRITE_FAILURE = 13
Global Const MAPI_E_UNKNOWN_RECIPIENT = 14
Global Const MAPI_E_BAD_RECIPTYPE = 15
Global Const MAPI_E_NO_MESSAGES = 16
Global Const MAPI_E_INVALID_MESSAGE = 17
Global Const MAPI_E_TEXT_TOO_LARGE = 18
Global Const MAPI_E_INVALID_SESSION = 19
Global Const MAPI_E_TYPE_NOT_SUPPORTED = 20
Global Const MAPI_E_AMBIGUOUS_RECIPIENT = 21
Global Const MAPI_E_MESSAGE_IN_USE = 22
Global Const MAPI_E_NETWORK_FAILURE = 23
Global Const MAPI_E_INVALID_EDITFIELDS = 24
Global Const MAPI_E_INVALID_RECIPS = 25
Global Const MAPI_E_NOT_SUPPORTED = 26

Global Const MAPI_ORIG = 0
Global Const MAPI_TO = 1
Global Const MAPI_CC = 2
Global Const MAPI_BCC = 3

Global Const MAPI_UNREAD = &H1
Global Const MAPI_RECEIPT_REQUESTED = &H2
Global Const MAPI_SENT = &H4


'***********************
'   FLAG Declarations
'***********************

Global Const MAPI_LOGON_UI = &H1
Global Const MAPI_NEW_SESSION = &H2
Global Const MAPI_DIALOG = &H8
Global Const MAPI_UNREAD_ONLY = &H20
Global Const MAPI_ENVELOPE_ONLY = &H40
Global Const MAPI_PEEK = &H80
Global Const MAPI_GUARANTEE_FIFO = &H100
Global Const MAPI_BODY_AS_FILE = &H200
Global Const MAPI_AB_NOMODIFY = &H400
Global Const MAPI_SUPPRESS_ATTACH = &H800
Global Const MAPI_FORCE_DOWNLOAD = &H1000

Global Const MAPI_OLE = &H1
Global Const MAPI_OLE_STATIC = &H2

Function CopyFiles(MfIn As MAPIFile, MfOut As MAPIFile) As Long

    MfOut.FileName = MfIn.FileName
    MfOut.PathName = MfIn.PathName
    MfOut.Reserved = MfIn.Reserved
    MfOut.flags = MfIn.flags
    MfOut.Position = MfIn.Position
    MfOut.FileType = MfIn.FileType
    CopyFiles = 1&
    
End Function

Function CopyRecipient(MrIn As MAPIRecip, MrOut As MAPIRecip) As Long

    MrOut.Name = MrIn.Name
    MrOut.Address = MrIn.Address
    MrOut.EIDSize = MrIn.EIDSize
    MrOut.EntryID = MrIn.EntryID
    MrOut.Reserved = MrIn.Reserved
    MrOut.RecipClass = MrIn.RecipClass

    CopyRecipient = 1&
    
End Function

Function GetMAPIErrorText(ByVal errorCode As Integer) As String
    
    Dim s As String

    Select Case errorCode

        Case MAPI_USER_ABORT
            s = "The operation was cancelled by the user"
        Case MAPI_E_FAILURE
            s = "MAPI failure"
        Case MAPI_E_LOGIN_FAILURE
            s = "Login failure"
        Case MAPI_E_DISK_FULL
            s = "Disk Full"
        Case MAPI_E_INSUFFICIENT_MEMORY
            s = "Insufficient memory to complete operation"
        Case MAPI_E_BLK_TOO_SMALL
            s = "Access denied"
        Case MAPI_E_TOO_MANY_SESSIONS
            s = "Too many sessions"
        Case MAPI_E_TOO_MANY_FILES
            s = "Too many files"
        Case MAPI_E_TOO_MANY_RECIPIENTS
            s = "Too many recipients"
        Case MAPI_E_ATTACHMENT_NOT_FOUND
            s = "Attachment not found"
        Case MAPI_E_ATTACHMENT_OPEN_FAILURE
            s = "Attachment open failure"
        Case MAPI_E_ATTACHMENT_WRITE_FAILURE
            s = "Attachment write failure"
        Case MAPI_E_UNKNOWN_RECIPIENT
            s = "Some names could not be matched to names in the address list"
        Case MAPI_E_BAD_RECIPTYPE
            s = "Bad recipient type"
        Case MAPI_E_NO_MESSAGES
            s = "No messages"
        Case MAPI_E_INVALID_MESSAGE
            s = "Invalid message"
        Case MAPI_E_TEXT_TOO_LARGE
            s = "Text too large"
        Case MAPI_E_INVALID_SESSION
            s = "Invalid session"
        Case MAPI_E_TYPE_NOT_SUPPORTED
            s = "Type not supported"
        Case MAPI_E_AMBIGUOUS_RECIPIENT
            s = "Some names could not be matched to a name in the address list. Please refine the currently selected name in the list"
        Case MAPI_E_MESSAGE_IN_USE
            s = "Message is in use by another process"
        Case MAPI_E_NETWORK_FAILURE
            s = "Network Failure"
        Case MAPI_E_INVALID_EDITFIELDS
            s = "Invalid Edit Fields"
        Case MAPI_E_INVALID_RECIPS
            s = "Invalid Recips"
        Case MAPI_E_NOT_SUPPORTED
            s = "MAPI: Not Supported"
        Case Else
            ' default to MAPI FAILURE for unknown mapi errors.
            s = "MAPI failure"
    End Select

    GetMAPIErrorText = s

End Function

Function MAPIAddress(Session As Long, UIParam As Long, Caption As String, nEditFields As Long, label As String, nRecipients As Long, Recips() As MAPIRecip, flags As Long, Reserved As Long) As Long


    Dim Info&
    Dim rc&, i As Integer, ignore&
    Dim nRecips As Long

    ReDim rec(0 To nRecipients) As MAPIRecip
    ' Use local since BMAPIAddress changes passed value
    nRecips = nRecipients
    '*****************************************************
    ' Copy input recipient structure into local
    ' recipient structure used as input to BMAPIAddress
    '*****************************************************

    For i = 0 To nRecipients - 1
        ignore& = CopyRecipient(Recips(i), rec(i))
    Next i

    rc& = BMAPIAddress(Info&, Session&, UIParam&, Caption$, nEditFields&, label$, nRecips&, rec(), 0&, 0&)
    
    If (rc& = SUCCESS_SUCCESS) Then

        '**************************************************
        ' New recipients are now in the memory referenced
        ' by Info (HANDLE). nRecipients is the number of
        ' new recipients.
        '**************************************************
        nRecipients = nRecips   ' copy back to parameter
        If (nRecipients > 0) Then
            ReDim rec(0 To nRecipients - 1) As MAPIRecip
            rc& = BMAPIGetAddress(Info&, nRecipients&, rec())
                                                  
            '*********************************************
            ' Copy local recipient structure to
            ' recipient structure passed as procedure
            ' parameter.  We can't pass the procedure
            ' parameter directory to the BMAPI.DLL Address routine
            '*********************************************

            ReDim Recips(0 To nRecipients - 1) As MAPIRecip

            For i = 0 To nRecipients - 1
                ignore& = CopyRecipient(rec(i), Recips(i))
            Next i

        End If

    End If

    MAPIAddress = rc&
    
End Function
Sub Domapilogoff()
    Dim iStatus As Long

    On Error Resume Next

    iStatus = MAPILogoff(gMAPISession, 0, 0, 0)
    On Error GoTo 0
    'on remet l'id a zero ainsi on sait qu'aucune session n'est ouverte
    gMAPISession = 0
    
End Sub
Private Function DoMAPIlogon() As Long
    
    On Error GoTo DoMAPILogonErr

    If gMAPISession = 0 Then
        'aucune session n'est ouverte.on etablit alors une session
        DoMAPIlogon = MAPILogon(0&, "", "", MAPI_LOGON_UI, 0&, gMAPISession)
    Else
        'une session valide est deja ouverte
        DoMAPIlogon = SUCCESS_SUCCESS
    End If

    Exit Function

DoMAPILogonErr:
    Call MsgBox("Erreur DoMAPILogon survenue", vbOKOnly + vbCritical, "")
    DoMAPIlogon = -1
    Exit Function

End Function

' ex : call envoi_message( "hf", "", "bonjour", null)
Function Envoi_Message(sTO As String, sTOadd As String, sSubject As String, sNote As String, sFileName As Variant) As Integer

    Dim iStatus As Long
    
   'si aucun destinataire n'est spécifié
    If IsNull(sTO) Then
        MsgBox "Aucun destinataire spécifié."
        Envoi_Message = -1
        Exit Function
    End If

    'declaration des variables
    Dim iStatus As Long
    Dim message As MAPIMessage
    Dim nbRecips As Integer
    ReDim tFiles(0) As MAPIFile
    Dim nbFiles As Integer
    ReDim tRecip(0) As MAPIRecip
    
    'ouverture d'une session.
    iStatus = DoMAPIlogon()
    If iStatus <> SUCCESS_SUCCESS Then
        MsgBox GetMAPIErrorText(iStatus)
        Envoi_Message = -1
        Exit Function
    End If

    'Initialisation du destinataire
    nbRecips = 1
'    ReDim tRecip(nbRecips - 1) As MAPIRecip
    tRecip(0).Reserved = 0
    tRecip(0).RecipClass = MAPI_TO
    tRecip(0).Name = sTO
    tRecip(0).Address = sTOadd
    tRecip(0).EIDSize = 0
    tRecip(0).EntryID = ""

    'Initialisation du fichier a rattacher
    If IsNull(sFileName) Then
        nbFiles = 0
    Else
        nbFiles = 1
'        ReDim tFiles(0) As MAPIFile
        tFiles(0).PathName = CStr(sFileName)
        tFiles(0).Reserved = 0&
        tFiles(0).Position = -1
        tFiles(0).FileName = ""
        tFiles(0).FileType = ""
    End If

    'Initialisation du message
    message.RecipCount = nbRecips
    message.FileCount = nbFiles
    If Not IsNull(sSubject) Then
        message.Subject = sSubject
    End If
    If Not IsNull(sNote) Then
        message.NoteText = sNote
    End If
    message.flags = MAPI_RECEIPT_REQUESTED
    message.MessageType = gMessageClass

    'Envoi du message
    If nbFiles = 0 Then
        Envoi_Message = MAPISendMail(gMAPISession, 0&, message, tRecip(), tFiles(), MAPI_LOGON_UI, 0&)
    Else
        'Envoi_Message = MAPISendMail(gMAPISession, 0&, tmessage, tRecip(1), tFiles(1), MAPI_LOGON_UI, 0&)
    End If
    If iStatus <> SUCCESS_SUCCESS Then
        MsgBox GetMAPIErrorText(iStatus)
        Envoi_Message = -1
        Exit Function
    End If

    'fermeture de la session
    Domapilogoff
    
    Envoi_Message = 0
    
End Function

Function envoimessage2(sTO As String, sSubject As String, sNote As String, sFileName As Variant) As Long

   'si aucun destinataire n'est spécifié
    If IsNull(sTO) Then
        MsgBox "Aucun destinataire spécifié."
        Exit Function
    End If

    'declaration des variables
    Dim iStatus As Long
    Dim tmessage As MAPIMessage
    Dim iNumRecips As Integer
    ReDim tFiles(0) As MAPIFile
    Dim inumFiles As Integer
    
    'ouverture d'une session.
  '  iStatus = DoMAPILogon()
  '  If iStatus <> SUCCESS_SUCCESS Then
  '      MsgBox GetMAPIErrorText(iStatus)
  '      Exit Function
  '  End If

    '---------------------------------------------------
    'initialisation du destinataire
    '---------------------------------------------------
    
    iNumRecips = 2
    

    'redimensionner le tableau des destinataires
    ReDim tRecip(1 To iNumRecips) As MAPIRecip

    'caracteristiques du destinataire
    tRecip(1).Name = sTO
    tRecip(1).RecipClass = MAPI_TO
    tRecip(2).Name = "Tahar Lalmi"
    tRecip(2).RecipClass = MAPI_CC
    

    
    'initialisation du nombre de destinataires du message
    tmessage.RecipCount = iNumRecips

    '---------------------------------------------------
    'initialisation du fichier a rattacher
    '---------------------------------------------------
    If IsNull(sFileName) Then
        tmessage.FileCount = 0
    Else
        tmessage.FileCount = 1
    End If

    If tmessage.FileCount = 1 Then
        ReDim tFiles(1) As MAPIFile
        tFiles(1).PathName = CStr(sFileName)
        tFiles(1).Reserved = 0&
        tFiles(1).Position = -1
        tFiles(1).FileName = ""
        tFiles(1).FileType = ""
    End If

    '---------------------------------------------------
    'initialisation du message
    '---------------------------------------------------
    If Not IsNull(sSubject) Then
        tmessage.Subject = sSubject
    End If

    If Not IsNull(sNote) Then
        tmessage.NoteText = sNote
    End If

        
        tmessage.flags = MAPI_RECEIPT_REQUESTED
    

    tmessage.MessageType = "IPM.{3FAE7021-58C9-11D0-A8E2-00609711C6A7}" 'gMessageClass

    '---------------------------------------------------
    'envoi du message
    '---------------------------------------------------
    If tmessage.FileCount = 0 Then
        envoimessage2 = MAPISendMail(gMAPISession, 0&, tmessage, tRecip(), tFiles(), MAPI_LOGON_UI, 0&)
    Else
        envoimessage2 = MAPISendMail(gMAPISession, 0&, tmessage, tRecip(), tFiles(), MAPI_LOGON_UI, 0&)
    End If
    If iStatus <> SUCCESS_SUCCESS Then
        MsgBox GetMAPIErrorText(iStatus)
        Exit Function
    End If

    'fermeture de la session
   ' DoMAPILogoff
    
End Function

Function envoimessage3(sTO As String, sSubject As String, sNote As String, sFileName As Variant) As Long

   'si aucun destinataire n'est spécifié
    If IsNull(sTO) Then
        MsgBox "Aucun destinataire spécifié."
        Exit Function
    End If

    'declaration des variables
    Dim iStatus As Long
    Dim tmessage As MAPIMessage
    Dim iNumRecips As Integer
    ReDim tFiles(0) As MAPIFile
    Dim inumFiles As Integer
    
    '---------------------------------------------------
    'initialisation du destinataire
    '---------------------------------------------------
    
    iNumRecips = 1
    

    'redimensionner le tableau des destinataires
    ReDim tRecip(1 To iNumRecips) As MAPIRecip

    'caracteristiques du destinataire
    tRecip(1).Name = sTO
    tRecip(1).RecipClass = MAPI_TO
    

    
    'initialisation du nombre de destinataires du message
    tmessage.RecipCount = iNumRecips

    '---------------------------------------------------
    'initialisation du fichier a rattacher
    '---------------------------------------------------
    If IsNull(sFileName) Then
        tmessage.FileCount = 0
    Else
        tmessage.FileCount = 1
    End If

    If tmessage.FileCount = 1 Then
        ReDim tFiles(1) As MAPIFile
        tFiles(1).PathName = CStr(sFileName)
        tFiles(1).Reserved = 0&
        tFiles(1).Position = -1
        tFiles(1).FileName = ""
        tFiles(1).FileType = ""
    End If

    '---------------------------------------------------
    'initialisation du message
    '---------------------------------------------------
    If Not IsNull(sSubject) Then
        tmessage.Subject = sSubject
    End If

    If Not IsNull(sNote) Then
        tmessage.NoteText = sNote
    End If

        
        tmessage.flags = MAPI_RECEIPT_REQUESTED
    

    tmessage.MessageType = gMessageClass

    '---------------------------------------------------
    'envoi du message
    '---------------------------------------------------
    If tmessage.FileCount = 0 Then
        envoimessage3 = MAPISendMail(gMAPISession, 0&, tmessage, tRecip(), tFiles(), MAPI_LOGON_UI, 0&)
    Else
        envoimessage3 = MAPISendMail(gMAPISession, 0&, tmessage, tRecip(), tFiles(), MAPI_LOGON_UI, 0&)
    End If
    If iStatus <> SUCCESS_SUCCESS Then
        MsgBox GetMAPIErrorText(iStatus)
        Exit Function
    End If

End Function
Function SendMailBasic()
    'This function simply calls the MAPI functions
    'to prompt the user for login, addressee information, etc.
    '
    'This is the most basic of mail-enabling routines because
    'we're allowing MAPI to handle all aspects of the creation
    'of the message.

    On Error GoTo SendMailBasicErr

    Dim iStatus As Long
    Dim tmessage As MAPIMessage
    Dim tRecips() As MAPIRecip
    Dim tFiles() As MAPIFile
    
    'make sure we're logged on.
    iStatus = DoMAPIlogon()
    If iStatus <> SUCCESS_SUCCESS Then
        MsgBox GetMAPIErrorText(iStatus)
        Exit Function
    End If

    'next, call up the dialog box.  Since we're not specifying
    'basic information and we are requesting the standard dialog box
    'by specifying MAPI_DIALOG, the result is the presentation of
    'the standard send note dialog box.
    iStatus = MAPISendMail(gMAPISession, 0&, tmessage, tRecips(), tFiles(), MAPI_DIALOG, 0&)
    If iStatus <> SUCCESS_SUCCESS Then
        MsgBox GetMAPIErrorText(iStatus)
        Exit Function
    End If

    'logoff of mail, release our session
    Domapilogoff

    Exit Function
    
SendMailBasicErr:
    MsgBox Error$
    Exit Function

End Function
