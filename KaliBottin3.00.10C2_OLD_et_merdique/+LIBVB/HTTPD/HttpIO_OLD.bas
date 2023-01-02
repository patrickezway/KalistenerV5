Attribute VB_Name = "HttpIO_Old"
Option Explicit

Public p_HTTP_AdrServeur As String
Public p_HTTP_CheminDépot As String
Public p_HTTP_Drive_Modeles_Serveur As String
Public p_HTTP_Chemin_Modeles_Serveur As String
Public p_HTTP_strHTTP As String
Public p_HTTP_TailleFichier As Long
Public p_HTTP_Form_Frame As Form
Public p_HTTP_Form_Hwnd As Window

' Pour mémoriser les chargements HTTP
Public Type HTTP_Fichiers_Chargés
    HTTP_Fullname_Serveur As String
    HTTP_Chemin_Serveur As String
    HTTP_NomFichier_Serveur As String
    HTTP_Extension_Serveur As String
    HTTP_Fullname_Local As String
    HTTP_Name As String
    HTTP_Chargé As Boolean
    HTTP_Locké As Boolean
    HTTP_SESSION As String
End Type
Public p_tbl_HTTP_Fichiers_Chargés() As HTTP_Fichiers_Chargés

Public p_bool_HTTP_Fichiers_Chargés As Boolean

Public Const HTTP_TAILLE_OK = 0
Public Const HTTP_TAILLE_ERREUR = 1
Public HTTP_TAILLE_LIB As String

Public Const HTTP_GET_OK = 0
Public Const HTTP_GET_ERREUR = 1
Public Const HTTP_GET_LOCKE = 2
Public Const HTTP_GET_OK_VIDE = 3
Public Const HTTP_GET_FIC_INTROUVABLE = 4
Public Const HTTP_GET_DEJA_EN_LOCAL = 5
Public Const HTTP_GET_PAS_COMPLET = 6
Public HTTP_GET_LIB As String

Public Const HTTP_PUT_OK = 0
Public Const HTTP_PUT_ERREUR = 1
Public Const HTTP_PUT_PAS_COMPLET = 2
Public HTTP_PUT_LIB As String

Public Const HTTP_DEL_OK = 0
Public Const HTTP_DEL_ERREUR = 1
Public HTTP_DEL_LIB As String

Public Const HTTP_LOCK_OK = 0
Public Const HTTP_LOCK_ERREUR = 1
Public Const HTTP_LOCK_AUTRE_USER = 5
Public Const HTTP_LOCK_PASFAIT = 6
Public HTTP_LOCK_LIB As String

Public Enum INTERNET_DEF
    INTERNET_DEFAULT_HTTP_PORT2 = 80
    INTERNET_DEFAULT_HTTPS_PORT = 443
End Enum

Public p_lInternetSession   As Long
Private m_lInternetConnect  As Long
Private m_lHttpRequest      As Long

Public Function HTTP_InitialHttpConnect(ByVal strURL As String) As Boolean
    
Dim iPort As Integer
Dim strObject As String
Dim intPos As Integer
    
    If left$(LCase(strURL), 7) = "http://" Then
        strURL = Right$(strURL, Len(strURL) - 7)
    Else
        If left$(LCase(strURL), 8) = "https://" Then
            strURL = Right$(strURL, Len(strURL) - 8)
        End If
    End If
    
    intPos = InStr(1, strURL, "/")
    If intPos > 0 Then
        strObject = Right$(strURL, Len(strURL) - intPos + 1)
        strURL = left$(strURL, intPos - 1)
    End If
    
    intPos = InStr(1, strURL, ":")
    If intPos > 0 Then
        iPort = val(Right$(strURL, Len(strURL) - intPos))
        strURL = left$(strURL, intPos - 1)
    Else
        iPort = INTERNET_DEFAULT_HTTP_PORT
    End If
    
    p_lInternetSession = InternetOpen("KaliDoc", 0, vbNullString, vbNullString, 0)

    m_lInternetConnect = InternetConnect(p_lInternetSession, strURL, iPort, _
                         vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    
    m_lHttpRequest = HttpOpenRequest(m_lInternetConnect, "POST", strObject, _
                     "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_NO_CACHE_WRITE, 0)
                                    
    HTTP_InitialHttpConnect = CBool(m_lHttpRequest)
End Function

Public Function HTTP_deletefile(v_sURL As String, v_CheminHTTP As String, v_CheminFichier As String, v_Session As String) As Integer
    
    'Dim stStatusCode As String, stStatusText As String
    Dim bresult As Boolean
    Dim lgTotal As Long
    Dim stLoad As String
    Dim stPost As String
    Dim sret As String
    Dim ret As Integer
    Dim maxn As Long
    Dim hFileLocal As Long
    Dim buf As String
    Dim n As Long, hindex As Long
    Dim nb_ecrits As Long
    Dim fpIn As Integer, ligne As String
    Dim RetClose As Long
            
    ' v_Chemin = Replace(v_Chemin, "\\", "\")
    v_sURL = v_sURL & "?v_CheminHTTP=" & v_CheminHTTP & "&v_CheminFichier=" & v_CheminFichier & "&v_Session=" & v_Session & "&v_NumUtil=" & p_NumUtil
    
    ret = HTTP_InitialHttpConnect(v_sURL)
    If CBool(ret) = False Then GoTo ErrorHandle

    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = ""
    
    ret = HttpSendRequest(m_lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        HTTP_deletefile = HTTP_DEL_ERREUR
        HTTP_DEL_LIB = "DeleteFile : HttpSendRequest=0 : Apache arrêté ?"
        GoTo ErrorHandle
    End If
    
    
    
    
    lgTotal = 0
    maxn = 1024
    hFileLocal = 1
    HTTP_deletefile = HTTP_DEL_OK
    Do While (hFileLocal > 0)
        buf = String(maxn, Chr(0))
        ret = InternetReadFile(m_lHttpRequest, buf, maxn, n)
        lgTotal = lgTotal + n
                
        If left(buf, 6) = "ERREUR" Or InStr(LCase(buf), "warning") > 0 Or InStr(buf, "404") Then
            HTTP_deletefile = HTTP_DEL_ERREUR
            HTTP_DEL_LIB = Mid(STR_GetChamp(buf, "|", 2), InStr(STR_GetChamp(buf, "|", 2), "mod_"))
            HTTP_DEL_LIB = Replace(HTTP_DEL_LIB, "mod_", "")
            GoTo ErrorHandle
        End If
        If (n = 0) Then
            hFileLocal = 0
        End If
    Loop
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(m_lHttpRequest, HTTP_QUERY_CONTENT_TYPE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_deletefile = HTTP_DEL_ERREUR
        HTTP_DEL_LIB = "DeleteFile : Erreur HttpQueryInfo_1 "
        GoTo ErrorHandle
    End If
    sret = "HTTP_QUERY_CONTENT_TYPE=" & left(buf, n) & vbCrLf
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(m_lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_deletefile = HTTP_DEL_ERREUR
        HTTP_DEL_LIB = "DeleteFile : Erreur HttpQueryInfo_2 "
        GoTo ErrorHandle
    End If
    sret = sret & "HTTP_QUERY_STATUS_CODE=" & left(buf, n) & vbCrLf
    If val(left(Trim(buf), 1)) > 3 Then
        HTTP_deletefile = HTTP_DEL_ERREUR
        HTTP_DEL_LIB = "DeleteFile : Erreur HttpQueryInfo_2 sret=" & left(Trim(buf), 3)
        GoTo ErrorHandle
    End If
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(m_lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_deletefile = HTTP_DEL_ERREUR
        HTTP_DEL_LIB = "DeleteFile : Erreur HttpQueryInfo"
        GoTo ErrorHandle
    End If
    sret = sret & "HTTP_QUERY_STATUS_CODE=" & left(buf, n) & vbCrLf
    
    HTTP_CloseConnect
    
    Exit Function

ErrorHandle:
    'Debug.Print Err.Description
    'Debug.Print Err.HelpContext
    'Debug.Print Err.LastDllError
    Err.Clear
    HTTP_CloseConnect

End Function

Public Function HTTP_deletefile_simple(v_sURL As String, v_CheminHTTP As String, v_CheminFichier As String) As Integer
    
    'Dim stStatusCode As String, stStatusText As String
    Dim bresult As Boolean
    Dim lgTotal As Long
    Dim stLoad As String
    Dim stPost As String
    Dim sret As String
    Dim ret As Integer
    Dim maxn As Long
    Dim hFileLocal As Long
    Dim buf As String
    Dim n As Long, hindex As Long
    Dim nb_ecrits As Long
    Dim fpIn As Integer, ligne As String
    Dim RetClose As Long
            
    HTTP_deletefile_simple = HTTP_DEL_OK
    
    v_sURL = v_sURL & "?v_CheminHTTP=" & v_CheminHTTP & "&v_CheminFichier=" & v_CheminFichier & "&v_NumUtil=" & p_NumUtil
    
    ret = HTTP_InitialHttpConnect(v_sURL)
    If CBool(ret) = False Then GoTo ErrorHandle

    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = ""
    
    ret = HttpSendRequest(m_lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        HTTP_deletefile_simple = HTTP_DEL_ERREUR
        HTTP_DEL_LIB = "DeleteFileSimple : HttpSendRequest=0 : Apache arrêté ?"
        GoTo ErrorHandle
    End If
    
    lgTotal = 0
    maxn = 1024
    hFileLocal = 1
    HTTP_deletefile_simple = HTTP_DEL_OK
    Do While (hFileLocal > 0)
        buf = String(maxn, Chr(0))
        ret = InternetReadFile(m_lHttpRequest, buf, maxn, n)
        lgTotal = lgTotal + n
                
        If left(buf, 6) = "ERREUR" Then
            HTTP_deletefile_simple = HTTP_DEL_ERREUR
            If InStr(buf, "mod_") > 0 Then
                HTTP_DEL_LIB = Mid(STR_GetChamp(buf, "|", 2), InStr(STR_GetChamp(buf, "|", 2), "mod_"))
                HTTP_DEL_LIB = Replace(HTTP_DEL_LIB, "mod_", "")
            Else
                HTTP_DEL_LIB = STR_GetChamp(buf, "|", 2)
            End If
            GoTo ErrorHandle
        ElseIf InStr(LCase(buf), "warning") > 0 Or InStr(buf, "404") Then
            HTTP_deletefile_simple = HTTP_DEL_ERREUR
            If InStr(buf, "mod_") > 0 Then
                HTTP_DEL_LIB = Mid(STR_GetChamp(buf, "|", 2), InStr(STR_GetChamp(buf, "|", 2), "mod_"))
                HTTP_DEL_LIB = Replace(HTTP_DEL_LIB, "mod_", "")
            Else
                HTTP_DEL_LIB = STR_GetChamp(buf, "|", 2)
            End If
            GoTo ErrorHandle
        End If
        If (n = 0) Then
            hFileLocal = 0
        End If
    Loop
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(m_lHttpRequest, HTTP_QUERY_CONTENT_TYPE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_deletefile_simple = HTTP_DEL_ERREUR
        HTTP_DEL_LIB = "DeleteFileSimple : Erreur HttpQueryInfo_1 "
        GoTo ErrorHandle
    End If
    sret = "HTTP_QUERY_CONTENT_TYPE=" & left(buf, n) & vbCrLf
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(m_lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_deletefile_simple = HTTP_DEL_ERREUR
        HTTP_DEL_LIB = "DeleteFileSimple : Erreur HttpQueryInfo_2 "
        GoTo ErrorHandle
    End If
    sret = sret & "HTTP_QUERY_STATUS_CODE=" & left(buf, n) & vbCrLf
    If val(left(Trim(buf), 1)) > 3 Then
        HTTP_deletefile_simple = HTTP_DEL_ERREUR
        HTTP_DEL_LIB = "DeleteFileSimple : Erreur HttpQueryInfo_2 sret=" & left(Trim(buf), 3)
        GoTo ErrorHandle
    End If
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(m_lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_deletefile_simple = HTTP_DEL_ERREUR
        HTTP_DEL_LIB = "DeleteFileSimple : Erreur HttpQueryInfo"
        GoTo ErrorHandle
    End If
    sret = sret & "HTTP_QUERY_STATUS_CODE=" & left(buf, n) & vbCrLf
    
    HTTP_CloseConnect
    
    Exit Function

ErrorHandle:
    'Debug.Print Err.Description
    'Debug.Print Err.HelpContext
    'Debug.Print Err.LastDllError
    Err.Clear
    HTTP_CloseConnect

End Function

Public Function HTTP_Appel_gerer_lockerfile(ByRef r_HTTP_LOCK_LIB As String, ByVal v_Trait As String, ByVal v_FichServeur As String, ByVal v_bMessage As Boolean) As Integer
    Dim FichServeur_Chemin As String, FichServeur_Fichier As String, FichServeur_Extension As String
    Dim FichLocal As String, strChemin As String, Session As String
    Dim iRet As Integer
    
    p_HTTP_AdrServeur = "\\192.168.101.20"
    p_HTTP_strHTTP = "http:" & p_HTTP_AdrServeur & "/TRSF_HTTP/locker_file.php"
    p_HTTP_strHTTP = Replace(p_HTTP_strHTTP, "\", "/")
    Session = HTTP_RandomAlphaNumString(5)
    p_HTTP_CheminDépot = "/usr/kalitech/kalidoc/TRSF_HTTP/HTTP_IO/"
    
    Menu.mnuHTTPDConfig1.visible = True
    Menu.mnuHTTPDConfig1.Caption = "Chemin de transfert = " & p_HTTP_CheminDépot
    Menu.mnuHTTPDConfig2.visible = True
    Menu.mnuHTTPDConfig2.Caption = "Adresse du serveur = " & p_HTTP_AdrServeur
    
    ' décomposer FichServeur
    strChemin = STR_GetChamp(v_FichServeur, "/", STR_GetNbchamp(v_FichServeur, "/") - 1)
    FichServeur_Extension = STR_GetChamp(strChemin, ".", 1)
    FichServeur_Extension = Replace(FichServeur_Extension, "mod_", "")
    FichServeur_Extension = Replace(FichServeur_Extension, "_" & p_NumUtil, "")
    FichServeur_Fichier = STR_GetChamp(strChemin, ".", 0)
    FichServeur_Chemin = Mid(v_FichServeur, 1, Len(v_FichServeur) - Len(strChemin) - 1)
    'iRet = HTTP_getfile(r_HTTP_GET_LIB, Session, p_http_strHTTP, p_http_CheminDépot, FichServeur_Chemin, FichServeur_Fichier, FichServeur_Extension, V_FichLocal, v_bMessage, v_bLocker)
    iRet = HTTP_lockerfile(r_HTTP_LOCK_LIB, v_Trait, p_HTTP_strHTTP, p_HTTP_CheminDépot, FichServeur_Chemin, FichServeur_Fichier, FichServeur_Extension, True, Session)
        
    
    HTTP_Appel_gerer_lockerfile = iRet
End Function


Public Function HTTP_Appel_getfile(ByRef r_HTTP_GET_LIB As String, ByVal v_FichServeur As String, ByVal V_FichLocal As String, ByVal v_bMessage As Boolean, ByVal v_bLocker As Boolean) As Integer
    Dim FichServeur_Chemin As String, FichServeur_Fichier As String, FichServeur_Extension As String
    Dim FichLocal As String, strChemin As String, Session As String
    Dim iRet As Integer
    
    p_HTTP_AdrServeur = "\\192.168.101.20"
    p_HTTP_strHTTP = "http:" & p_HTTP_AdrServeur & "/TRSF_HTTP/get_file.php"
    p_HTTP_strHTTP = Replace(p_HTTP_strHTTP, "\", "/")
    Session = HTTP_RandomAlphaNumString(5)
    p_HTTP_CheminDépot = "/usr/kalitech/kalidoc/TRSF_HTTP/HTTP_IO/"
    
    Menu.mnuHTTPDConfig1.visible = True
    Menu.mnuHTTPDConfig1.Caption = "Chemin de transfert = " & p_HTTP_CheminDépot
    Menu.mnuHTTPDConfig2.visible = True
    Menu.mnuHTTPDConfig2.Caption = "Adresse du serveur = " & p_HTTP_AdrServeur
    
    ' décomposer FichServeur
    strChemin = STR_GetChamp(v_FichServeur, "/", STR_GetNbchamp(v_FichServeur, "/") - 1)
    FichServeur_Extension = STR_GetChamp(strChemin, ".", 1)
    FichServeur_Fichier = STR_GetChamp(strChemin, ".", 0)
    FichServeur_Chemin = Mid(v_FichServeur, 1, Len(v_FichServeur) - Len(strChemin) - 1)
    iRet = HTTP_getfile(r_HTTP_GET_LIB, Session, p_HTTP_strHTTP, p_HTTP_CheminDépot, FichServeur_Chemin, FichServeur_Fichier, FichServeur_Extension, V_FichLocal, v_bMessage, v_bLocker)
        
    If iRet = HTTP_GET_LOCKE Then
    ElseIf iRet = HTTP_GET_ERREUR Then
    ElseIf iRet = HTTP_GET_PAS_COMPLET Then
    ElseIf iRet = HTTP_GET_FIC_INTROUVABLE Then
    ElseIf iRet = HTTP_GET_DEJA_EN_LOCAL Then
    ElseIf iRet = HTTP_GET_OK Then
    End If
    
    HTTP_Appel_getfile = iRet
End Function

Public Function HTTP_getfile(ByRef r_HTTP_GET_LIB As String, ByRef r_Session As String, v_sURL As String, v_chemin As String, v_CheminFichier_Serveur As String, v_NomFichier_Serveur As String, v_ExtensionFichier_Serveur As String, v_nomfich_Copie As String, v_bool_message As Boolean, ByVal v_locker As Boolean) As Integer

' Optional stUser As String = vbNullString,
' Optional stPass As String = vbNullString

    Dim stStatusCode As String, stStatusText As String

    Dim lgTotal As Long
    Dim stLoad As String
    Dim stPost As String
    Dim sret As String
    Dim ret As Integer
    Dim maxn As Long
    Dim hFileLocal As Long
    Dim buf As String
    Dim n As Long, hindex As Long
    Dim nb_ecrits As Long
    Dim fpIn As Integer, ligne As String
    Dim RetClose As Long
    Dim nomfich_Serveur As String
    Dim nomFicRenomme As String
    Dim bresult As Boolean
    Dim Locker As String
    Dim CheminTmp As String
    
    Dim TimeDébut As Date
    Dim TimePrem As Date
    Dim ResteSeconde As Long
    Dim TailleChargement As Long
    Dim iRetTaille As Integer
    Dim bPrem As Boolean
    Dim NbSeconde As Long
    
    HTTP_GET_LIB = ""
    
    ' Récupérer la taille du fichier
    p_HTTP_TailleFichier = 0
    iRetTaille = HTTP_gettaille(v_sURL, v_CheminFichier_Serveur, v_NomFichier_Serveur, v_ExtensionFichier_Serveur)
    If iRetTaille = HTTP_TAILLE_OK Then
        p_HTTP_TailleFichier = HTTP_TAILLE_LIB
        maxn = 1024
        TailleChargement = HTTP_TAILLE_LIB + maxn
        'MsgBox "TailleChargement=" & TailleChargement
    End If
        
    If iRetTaille = HTTP_TAILLE_OK And p_HTTP_Form_Frame.Name <> "" Then
        p_HTTP_Form_Frame.FrmHTTPD.visible = True
        p_HTTP_Form_Frame.FrmHTTPD.ZOrder 0
        p_HTTP_Form_Frame.lblHTTPD.Caption = "Chargement de " & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur & " (" & (TailleChargement / 1024) & " K Octets)"
        p_HTTP_Form_Frame.PgbarHTTPDTaille.max = TailleChargement
        p_HTTP_Form_Frame.PgbarHTTPDTaille.Value = 0
        p_HTTP_Form_Frame.lblHTTPDTemps.Caption = "Temps restant"
        p_HTTP_Form_Frame.lblHTTPDTaille.Caption = "Volume chargé"
        p_HTTP_Form_Frame.Refresh
    End If
    
    TimeDébut = DateTime.Now()
    
    If (v_locker) Then
        Locker = "O"
    Else
        Locker = "N"
    End If
    
    v_sURL = v_sURL & "?v_Locker=" & Locker
    v_sURL = v_sURL & "&v_Session=" & r_Session

    ret = HTTP_InitialHttpConnect(v_sURL)
    If CBool(ret) = False Then GoTo ErrorHandle

    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = "v_CheminHTTP=" & v_chemin & "&"
    stPost = stPost & "v_CheminFichier=" & v_CheminFichier_Serveur & "&"
    stPost = stPost & "v_NomFichier=" & v_NomFichier_Serveur & "&"
    stPost = stPost & "v_ExtensionFichier=" & v_ExtensionFichier_Serveur & "&"
    stPost = stPost & "v_NumUtil=" & p_NumUtil
    
    nomfich_Serveur = v_CheminFichier_Serveur & "/" & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur
    
    ret = HttpSendRequest(m_lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        HTTP_getfile = HTTP_GET_ERREUR
        HTTP_GET_LIB = "GetFile : HttpSendRequest=0 : Apache arrêté ?"
        GoTo ErrorHandle
    End If
    
    v_nomfich_Copie = v_nomfich_Copie & "_Session_" & r_Session
    
    hFileLocal = CreateFile(v_nomfich_Copie, GENERIC_WRITE Or GENERIC_READ, _
                        0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If (hFileLocal < 0) Then
        HTTP_getfile = HTTP_GET_ERREUR
        HTTP_GET_LIB = "GetFile : CreateFile " & v_nomfich_Copie
        GoTo ErrorHandle
    End If
        
    lgTotal = 0
    maxn = 1024
    Do While (hFileLocal > 0)
        buf = String(maxn, Chr(0))
        ret = InternetReadFile(m_lHttpRequest, buf, maxn, n)
        lgTotal = lgTotal + n
        
        If iRetTaille = HTTP_TAILLE_OK Then
            'p_HTTP_Form_Frame.lblHTTPDTaille.Caption = "Chargés " & lgTotal / 1024 & " K Octets"
            If p_HTTP_Form_Frame.PgbarHTTPDTaille.max < lgTotal Then
                p_HTTP_Form_Frame.PgbarHTTPDTaille.max = lgTotal
            End If
            p_HTTP_Form_Frame.PgbarHTTPDTaille.Value = lgTotal
        
            TimePrem = DateTime.Now()
            If n > 0 Then
                NbSeconde = DateDiff("s", TimeDébut, TimePrem)
                If NbSeconde = 0 Then NbSeconde = 1
                'Debug.Print p_HTTP_Form_Frame.PgbarHTTPDTemps.max & " et " & NbSeconde / n * p_HTTP_TailleFichier
                p_HTTP_Form_Frame.PgbarHTTPDTemps.max = NbSeconde / p_HTTP_TailleFichier
            
                ResteSeconde = (p_HTTP_TailleFichier - lgTotal) / NbSeconde
                If ResteSeconde < 0 Then
                    ResteSeconde = 0
                    p_HTTP_Form_Frame.PgbarHTTPDTemps.Value = ResteSeconde
                    p_HTTP_Form_Frame.lblHTTPDTemps.Caption = "Terminé"
                Else
                    p_HTTP_Form_Frame.PgbarHTTPDTemps.Value = ResteSeconde
                End If
            End If
            p_HTTP_Form_Frame.Refresh
        End If
                        
        If (n > 0) Then
            ret = WriteFile(hFileLocal, ByVal buf, n, nb_ecrits, ByVal 0&)
            If (nb_ecrits < 1) Then
                ret = GetLastError
            End If
        Else
            RetClose = CloseHandle(hFileLocal)
            If RetClose <> 1 Then
                HTTP_getfile = HTTP_GET_ERREUR
                HTTP_GET_LIB = "GetFile : Impossible de fermer " & v_nomfich_Copie
                Kill v_nomfich_Copie
                GoTo ErrorHandle
            End If
            hFileLocal = 0
        End If
    Loop
    If lgTotal = 0 Then
        HTTP_getfile = HTTP_GET_OK_VIDE
        HTTP_GET_LIB = "GetFile : Fichier Vide " & v_nomfich_Copie
        'Kill v_nomfich_Copie
        'GoTo ErrorHandle
    End If
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(m_lHttpRequest, HTTP_QUERY_CONTENT_TYPE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_getfile = HTTP_GET_ERREUR
        HTTP_GET_LIB = "GetFile : Erreur HttpQueryInfo_1 " & v_nomfich_Copie
        GoTo ErrorHandle
    End If
    sret = "HTTP_QUERY_CONTENT_TYPE=" & left(buf, n) & vbCrLf
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(m_lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_getfile = HTTP_GET_ERREUR
        HTTP_GET_LIB = "GetFile : Erreur HttpQueryInfo_2 " & v_nomfich_Copie
        GoTo ErrorHandle
    End If
    sret = sret & "HTTP_QUERY_STATUS_CODE=" & left(buf, n) & vbCrLf
    If val(left(Trim(buf), 1)) > 3 Then
        HTTP_getfile = HTTP_GET_ERREUR
        HTTP_GET_LIB = "GetFile : Erreur HttpQueryInfo_2 sret=" & left(Trim(buf), 3) & " " & v_nomfich_Copie
        GoTo ErrorHandle
    End If
    
    HTTP_CloseConnect
    
    If HTTP_getfile = HTTP_GET_OK_VIDE Then
        GoTo SuiteGet
    End If
    
    HTTP_getfile = HTTP_GET_OK
    ' Voir si erreur
    If FICH_FichierExiste(v_nomfich_Copie) Then
    
        fpIn = FreeFile
        FICH_OuvrirFichier v_nomfich_Copie, FICH_LECTURE, fpIn
            
        If Not EOF(fpIn) Then
            Line Input #fpIn, ligne
            If left(ligne, 6) = "ERREUR" Then
                If STR_GetChamp(ligne, "|", 1) = 5 Then
                    HTTP_getfile = HTTP_GET_LOCKE
                    HTTP_GET_LIB = Mid(STR_GetChamp(ligne, "|", 2), InStr(STR_GetChamp(ligne, "|", 2), "mod_"))
                    HTTP_GET_LIB = Replace(HTTP_GET_LIB, "mod_", "")
                ElseIf InStr(LCase(ligne), nomfich_Serveur & " introuvable") > 0 Then
                    HTTP_getfile = HTTP_GET_FIC_INTROUVABLE
                    HTTP_GET_LIB = "GetFile : " & nomfich_Serveur & " introuvable"
                    If v_bool_message Then
                        MsgBox "Erreur " & HTTP_GET_LIB
                    End If
                Else
                    HTTP_getfile = HTTP_GET_ERREUR
                    HTTP_GET_LIB = "GetFile : " & STR_GetChamp(ligne, "|", 1) & " " & STR_GetChamp(ligne, ":", 2) & Chr(13) & Chr(10) & "Url : " & v_sURL & " lors de la copie de " & nomfich_Serveur & " vers " & v_nomfich_Copie
                    If v_bool_message Then
                        MsgBox "Erreur " & STR_GetChamp(ligne, "|", 1) & " " & STR_GetChamp(ligne, ":", 2) & Chr(13) & Chr(10) & "Url : " & v_sURL & " lors de la copie de " & nomfich_Serveur & " vers " & v_nomfich_Copie
                    End If
                End If
            ElseIf InStr(LCase(ligne), "warning") > 0 Or InStr(LCase(ligne), "parse") > 0 Or InStr(ligne, "404") Then
                HTTP_getfile = HTTP_GET_ERREUR
                HTTP_GET_LIB = "GetFile : " & ligne
                If v_bool_message Then
                    MsgBox "Erreur " & HTTP_GET_LIB
                End If
            Else
                HTTP_getfile = HTTP_GET_OK
            End If
        Else
            HTTP_getfile = HTTP_GET_OK
        End If
        
        If HTTP_getfile = HTTP_GET_OK Then
            If iRetTaille = HTTP_TAILLE_OK Then
                'MsgBox FileLen(v_nomfich_Copie) & " " & p_HTTP_TailleFichier
                ' comparer la taille des deux fichiers
                If p_HTTP_TailleFichier <> FileLen(v_nomfich_Copie) Then
                    HTTP_getfile = HTTP_GET_PAS_COMPLET
                    HTTP_GET_LIB = "Taille du fichier sur le serveur " & p_HTTP_TailleFichier & Chr(13) & Chr(10) & "Taille du fichier chargé " & FileLen(v_nomfich_Copie)
                    p_HTTP_Form_Frame.lblHTTPD.Caption = HTTP_GET_LIB
                Else
                    p_HTTP_Form_Frame.lblHTTPD.Caption = "Chargement terminé avec succès"
                End If
            End If
        End If

        Close (fpIn)
    Else
        HTTP_getfile = HTTP_GET_ERREUR
        HTTP_GET_LIB = "GetFile : Fichier non récupéré : " & v_nomfich_Copie
    End If
    If HTTP_getfile = HTTP_GET_OK Then
SuiteGet:
        ' renommer pour enlever la session
        ' si existe déjà => renommer avec date et heure
        If FICH_FichierExiste(Replace(v_nomfich_Copie, "_Session_" & r_Session, "")) Then
            nomFicRenomme = Replace(v_nomfich_Copie, "_Session_" & r_Session, "")
            nomFicRenomme = Replace(nomFicRenomme, "." & v_ExtensionFichier_Serveur, "_Date_" & Format(Date, "yyyy_mm_dd") & "_Heure_" & Format(Time, "hh_nn_ss") & "." & v_ExtensionFichier_Serveur)
            HTTP_GET_LIB = "Le fichier " & Replace(v_nomfich_Copie, "_Session_" & r_Session, "") & " a été renommé en " & nomFicRenomme
            Call FICH_RenommerFichier(Replace(v_nomfich_Copie, "_Session_" & r_Session, ""), nomFicRenomme)
            HTTP_getfile = HTTP_GET_DEJA_EN_LOCAL
        End If
        Call FICH_RenommerFichier(v_nomfich_Copie, Replace(v_nomfich_Copie, "_Session_" & r_Session, ""))
    ElseIf HTTP_getfile = HTTP_GET_LOCKE Or HTTP_getfile = HTTP_GET_ERREUR Or HTTP_getfile = HTTP_GET_FIC_INTROUVABLE Then
        ' supprimer le fichier généré
        If FICH_FichierExiste(v_nomfich_Copie) Then
            Kill (v_nomfich_Copie)
        End If
    End If
    ' supprimer les fichiers temporaires sur le serveur
    If HTTP_getfile = HTTP_GET_OK Then
        CheminTmp = v_chemin & v_NomFichier_Serveur & ".mod_" & v_ExtensionFichier_Serveur & "_" & p_NumUtil & "_Session_" & r_Session
        If HTTP_Supprimer_Fichier_Temporaire(v_chemin, CheminTmp) = HTTP_DEL_ERREUR Then
            CheminTmp = Replace(CheminTmp, ".mod_", ".")
            CheminTmp = Replace(CheminTmp, "_" & p_NumUtil, "")
            Call HTTP_Supprimer_Fichier_Temporaire(v_chemin, Replace(CheminTmp, ".mod_", "."))
        End If
    End If

ErrorHandle:
    
    CheminTmp = v_chemin & v_NomFichier_Serveur & ".mod_" & v_ExtensionFichier_Serveur & "_" & p_NumUtil & "_Session_" & r_Session
    
    Call HTTP_Supprimer_Fichier_Temporaire(v_chemin, CheminTmp)
    
    Err.Clear
    
    HTTP_CloseConnect
    
End Function

Public Function HTTP_Supprimer_Fichier_Temporaire(v_cheminDépot As String, v_NomFichier As String)
    Dim p_HTTP_strHTTP As String
    Dim iRet As Integer
    
    p_HTTP_strHTTP = "http:" & p_HTTP_AdrServeur & "/TRSF_HTTP/delete_file_simple.php"
    p_HTTP_strHTTP = Replace(p_HTTP_strHTTP, "\", "/")
            'nomfichier = p_http_CheminDépot & nomIn_Fichier & ".mod_" & nomIn_Extension & "_" & p_NumUtil
    iRet = HTTP_deletefile_simple(p_HTTP_strHTTP, v_cheminDépot, v_NomFichier)
    
    HTTP_Supprimer_Fichier_Temporaire = iRet
End Function

Public Function HTTP_lockerfile(ByRef r_HTTP_GET_LIB As String, v_Trait As String, v_sURL As String, v_chemin As String, v_CheminFichier_Serveur As String, v_NomFichier_Serveur As String, v_ExtensionFichier_Serveur As String, v_bool_message As Boolean, v_Session As String) As Integer

    Dim stStatusCode As String, stStatusText As String

    Dim lgTotal As Long
    Dim stLoad As String
    Dim stPost As String
    Dim sret As String
    Dim ret As Integer
    Dim maxn As Long
    Dim hFileLocal As Long
    Dim buf As String
    Dim n As Long, hindex As Long
    Dim nb_ecrits As Long
    Dim fpIn As Integer, ligne As String
    Dim RetClose As Long
    Dim nomfich_Serveur As String
    Dim bresult As Boolean
    Dim Locker As String
    
    HTTP_LOCK_LIB = ""
    
    v_sURL = v_sURL & "?v_Locker=" & v_Trait
    v_sURL = v_sURL & "&v_Session=" & v_Session

    ret = HTTP_InitialHttpConnect(v_sURL)
    If CBool(ret) = False Then GoTo ErrorHandle

    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = "v_CheminHTTP=" & v_chemin & "&"
    stPost = stPost & "v_CheminFichier=" & v_CheminFichier_Serveur & "&"
    stPost = stPost & "v_NomFichier=" & v_NomFichier_Serveur & "&"
    stPost = stPost & "v_ExtensionFichier=" & v_ExtensionFichier_Serveur & "&"
    stPost = stPost & "v_NumUtil=" & p_NumUtil
    
    nomfich_Serveur = v_CheminFichier_Serveur & "/" & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur
    
    ret = HttpSendRequest(m_lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        HTTP_lockerfile = HTTP_LOCK_ERREUR
        HTTP_LOCK_LIB = "LockerFile : HttpSendRequest=0 : Apache arrêté ?"
        GoTo ErrorHandle
    End If
    
    lgTotal = 0
    maxn = 1024
    hFileLocal = 1
    HTTP_lockerfile = HTTP_LOCK_OK
    Do While (hFileLocal > 0)
        buf = String(maxn, Chr(0))
        ret = InternetReadFile(m_lHttpRequest, buf, maxn, n)
        lgTotal = lgTotal + n
                
        If left(buf, 6) = "ERREUR" Then
            HTTP_lockerfile = STR_GetChamp(buf, "|", 1)
            HTTP_LOCK_LIB = STR_GetChamp(buf, "|", 2)
            HTTP_LOCK_LIB = Replace(HTTP_LOCK_LIB, "mod_", "")
            GoTo ErrorHandle
        ElseIf InStr(LCase(buf), "warning") > 0 Or InStr(buf, "404") Then
            HTTP_lockerfile = HTTP_LOCK_ERREUR
            HTTP_LOCK_LIB = buf
            GoTo ErrorHandle
        End If
        If (n = 0) Then
            hFileLocal = 0
        End If
    Loop
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(m_lHttpRequest, HTTP_QUERY_CONTENT_TYPE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_lockerfile = HTTP_LOCK_ERREUR
        HTTP_LOCK_LIB = "LockerFile : Erreur HttpQueryInfo_1 "
        GoTo ErrorHandle
    End If
    sret = "HTTP_QUERY_CONTENT_TYPE=" & left(buf, n) & vbCrLf
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(m_lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_lockerfile = HTTP_LOCK_ERREUR
        HTTP_LOCK_LIB = "LockerFile : Erreur HttpQueryInfo_2 "
        GoTo ErrorHandle
    End If
    sret = sret & "HTTP_QUERY_STATUS_CODE=" & left(buf, n) & vbCrLf
    If val(left(Trim(buf), 1)) > 3 Then
        HTTP_lockerfile = HTTP_LOCK_ERREUR
        HTTP_LOCK_LIB = "LockerFile : Erreur HttpQueryInfo_2 sret=" & left(Trim(buf), 3)
        GoTo ErrorHandle
    End If
    
    HTTP_CloseConnect
    
    HTTP_lockerfile = HTTP_LOCK_OK
    
    Exit Function

ErrorHandle:
    Debug.Print Err.Description
    Debug.Print Err.HelpContext
    Debug.Print Err.LastDllError
    Err.Clear
    HTTP_CloseConnect
End Function

Public Function HTTP_gettaille(ByVal v_sURL As String, ByVal v_CheminFichier_Serveur As String, ByVal v_NomFichier_Serveur As String, ByVal v_ExtensionFichier_Serveur As String) As Long

    Dim stStatusCode As String, stStatusText As String

    Dim lgTotal As Long
    Dim stLoad As String
    Dim stPost As String
    Dim sret As String
    Dim ret As Integer
    Dim maxn As Long
    Dim hFileLocal As Long
    Dim buf As String
    Dim n As Long, hindex As Long
    Dim nb_ecrits As Long
    Dim fpIn As Integer, ligne As String
    Dim RetClose As Long
    Dim nomfich_Serveur As String
    Dim bresult As Boolean
    Dim Locker As String
    
    HTTP_TAILLE_LIB = ""
    
    v_sURL = Replace(v_sURL, "get_file", "get_taille")
    v_sURL = Replace(v_sURL, "put_file", "get_taille")

    ret = HTTP_InitialHttpConnect(v_sURL)
    If CBool(ret) = False Then GoTo ErrorHandle

    stLoad = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    stPost = "v_CheminHTTP=" & p_HTTP_CheminDépot & "&"
    stPost = stPost & "v_Chemin=" & v_CheminFichier_Serveur & "&"
    stPost = stPost & "v_Fichier=" & v_NomFichier_Serveur & "&"
    stPost = stPost & "v_Extension=" & v_ExtensionFichier_Serveur & "&"
    stPost = stPost & "v_NumUtil=" & p_NumUtil
    
    'nomfich_Serveur = v_CheminFichier_Serveur & "/" & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur
    
    ret = HttpSendRequest(m_lHttpRequest, stLoad, Len(stLoad), stPost, Len(stPost))
    If ret = 0 Then
        HTTP_gettaille = HTTP_TAILLE_ERREUR
        HTTP_TAILLE_LIB = "GetTaille : HttpSendRequest=0 : Apache arrêté ?"
        GoTo ErrorHandle
    End If
    
    lgTotal = 0
    maxn = 1024
    hFileLocal = 1
    HTTP_gettaille = HTTP_TAILLE_OK
    Do While (hFileLocal > 0)
        buf = String(maxn, Chr(0))
        ret = InternetReadFile(m_lHttpRequest, buf, maxn, n)
        lgTotal = lgTotal + n
                
        If left(buf, 2) = "OK" Then
            HTTP_gettaille = HTTP_TAILLE_OK
            HTTP_TAILLE_LIB = STR_GetChamp(buf, "|", 1)
            HTTP_TAILLE_LIB = STR_GetChamp(HTTP_TAILLE_LIB, " ", 0)
            n = 0
        ElseIf left(buf, 6) = "ERREUR" Then
            HTTP_gettaille = HTTP_TAILLE_ERREUR
            HTTP_TAILLE_LIB = STR_GetChamp(buf, "|", 1)
            GoTo ErrorHandle
        ElseIf InStr(LCase(buf), "warning") > 0 Or InStr(buf, "404") Then
            HTTP_gettaille = HTTP_TAILLE_ERREUR
            HTTP_TAILLE_LIB = buf
            GoTo ErrorHandle
        End If
        If (n = 0) Then
            hFileLocal = 0
        End If
    Loop
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(m_lHttpRequest, HTTP_QUERY_CONTENT_TYPE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_gettaille = HTTP_TAILLE_ERREUR
        HTTP_TAILLE_LIB = "GetTaille : Erreur HttpQueryInfo_1 "
        GoTo ErrorHandle
    End If
    sret = "HTTP_QUERY_CONTENT_TYPE=" & left(buf, n) & vbCrLf
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(m_lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_gettaille = HTTP_TAILLE_ERREUR
        HTTP_TAILLE_LIB = "GetTaille : Erreur HttpQueryInfo_2 "
        GoTo ErrorHandle
    End If
    sret = sret & "HTTP_QUERY_STATUS_CODE=" & left(buf, n) & vbCrLf
    If val(left(Trim(buf), 1)) > 3 Then
        HTTP_gettaille = HTTP_TAILLE_ERREUR
        HTTP_TAILLE_LIB = "GetTaille : Erreur HttpQueryInfo_2 sret=" & left(Trim(buf), 3)
        GoTo ErrorHandle
    End If
    
    HTTP_CloseConnect
    
    Exit Function

ErrorHandle:
    Err.Clear
    HTTP_CloseConnect
End Function

' generates a random alphanumeirc string of a given length
Public Function HTTP_RandomAlphaNumString(ByVal intLen As Integer)
    Dim StrReturn As String
    
    Dim X As Integer
    Dim c As Byte
    
    Randomize
    
    For X = 1 To intLen
        c = Int(Rnd() * 127)
    
        If (c >= Asc("0") And c <= Asc("9")) Or _
           (c >= Asc("A") And c <= Asc("Z")) Or _
           (c >= Asc("a") And c <= Asc("z")) Then
           
            StrReturn = StrReturn & Chr(c)
        Else
            X = X - 1
        End If
    Next X
    
    HTTP_RandomAlphaNumString = StrReturn
End Function

Public Function HTTP_Appel_putfile(ByRef r_HTTP_PUT_LIB As String, ByVal v_FichServeur As String, ByVal V_FichLocal As String, ByVal v_bMessage As Boolean, ByVal v_bDeLocker As Boolean) As Integer
    Dim FichServeur_Chemin As String, FichServeur_Fichier As String, FichServeur_Extension As String
    Dim FichLocal As String, strChemin As String, Session As String
    Dim FichTmp As String
    Dim iRet As Integer
    
    p_HTTP_AdrServeur = "\\192.168.101.20"
    p_HTTP_strHTTP = "http:" & p_HTTP_AdrServeur & "/TRSF_HTTP/put_file.php"
    p_HTTP_strHTTP = Replace(p_HTTP_strHTTP, "\", "/")
    Session = HTTP_RandomAlphaNumString(5)
    p_HTTP_CheminDépot = "/usr/kalitech/kalidoc/TRSF_HTTP/HTTP_IO/"

    Menu.mnuHTTPDConfig1.visible = True
    Menu.mnuHTTPDConfig1.Caption = "Chemin de transfert = " & p_HTTP_CheminDépot
    Menu.mnuHTTPDConfig2.visible = True
    Menu.mnuHTTPDConfig2.Caption = "Adresse du serveur = " & p_HTTP_AdrServeur
    
    ' décomposer FichServeur
    strChemin = STR_GetChamp(v_FichServeur, "/", STR_GetNbchamp(v_FichServeur, "/") - 1)
    FichServeur_Extension = STR_GetChamp(strChemin, ".", 1)
    FichServeur_Fichier = STR_GetChamp(strChemin, ".", 0)
    FichServeur_Chemin = Mid(v_FichServeur, 1, Len(v_FichServeur) - Len(strChemin) - 1)
    
    iRet = HTTP_putfile(r_HTTP_PUT_LIB, p_HTTP_strHTTP, FichServeur_Chemin, FichServeur_Fichier, FichServeur_Extension, V_FichLocal, v_FichServeur, v_bDeLocker)
    
    HTTP_Appel_putfile = iRet
End Function

Public Function HTTP_putfile(ByRef r_HTTP_PUT_LIB As String, ByVal v_sURL As String, ByVal v_CheminFichier_Serveur As String, ByVal v_NomFichier_Serveur As String, ByVal v_ExtensionFichier_Serveur As String, ByVal nomFichTmp, nomFichDest As String, ByRef v_DeLocker As Boolean) As Integer

    Dim stLoad As String
    Dim stPost1 As String, stPost2 As String
    Dim strBoundary As String
    
    Dim MimeType As String
    Dim lgTot As Long
    Dim sret As String
    
    Dim ret As Integer
    Dim bresult As Boolean
    Dim lBufferLength   As Long
    Dim BufferIn As INTERNET_BUFFERS
    Dim maxn As Long
        
    Dim RetClose As Long
    Dim hFileLocal As Long
    Dim buf As String
    Dim n As Long, nb_transmis As Long, hindex As Long
    Dim nb_total As Long
    Dim fpIn As Integer, ligne As String
    
    Dim TimeDébut As Date
    Dim lgTotal As Long
    Dim TimePrem As Date
    Dim ResteSeconde As Long
    Dim TailleChargement As Long
    Dim iRetTaille As Integer
    Dim bPrem As Boolean
    Dim NbSeconde As Long
    Dim sUrl As String
    
    HTTP_TAILLE_LIB = ""
    
    ' Récupérer la taille du fichier
    p_HTTP_TailleFichier = 0
    iRetTaille = HTTP_gettaille(v_sURL, v_CheminFichier_Serveur, v_NomFichier_Serveur, v_ExtensionFichier_Serveur)
    If iRetTaille = HTTP_TAILLE_OK Then
        p_HTTP_TailleFichier = HTTP_TAILLE_LIB
        maxn = 1024
        TailleChargement = HTTP_TAILLE_LIB + maxn
        'MsgBox "TailleChargement=" & TailleChargement
    End If
    
    If iRetTaille = HTTP_TAILLE_OK And p_HTTP_Form_Frame.Name <> "" Then
        p_HTTP_Form_Frame.FrmHTTPD.visible = True
        p_HTTP_Form_Frame.FrmHTTPD.ZOrder 0
        p_HTTP_Form_Frame.lblHTTPD.Caption = "Chargement de " & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur & " (" & (TailleChargement / 1024) & " K Octets)"
        p_HTTP_Form_Frame.PgbarHTTPDTaille.max = TailleChargement
        p_HTTP_Form_Frame.PgbarHTTPDTaille.Value = 0
        p_HTTP_Form_Frame.lblHTTPDTemps.Caption = "Temps restant"
        p_HTTP_Form_Frame.lblHTTPDTaille.Caption = "Volume chargé"
        p_HTTP_Form_Frame.Refresh
    End If
    
    TimeDébut = DateTime.Now()
    
    MimeType = "application/octet-stream"
    strBoundary = HTTP_RandomAlphaNumString(32)
    
    nomFichDest = v_CheminFichier_Serveur & "/" & v_NomFichier_Serveur & "." & v_ExtensionFichier_Serveur

    stPost1 = "--" & strBoundary & vbCrLf & _
             "Content-Disposition: form-data; " & _
             "name=""" & nomFichTmp & """; " & _
             "filename=""" & nomFichDest & """" & vbCrLf & _
             "Content-Type: " & MimeType & vbCrLf & vbCrLf
    
    stPost2 = vbCrLf & "--" & strBoundary & "--"
    ' find the length of the request body - this is required for the Content-Length header
    lgTot = Len(stPost1) + FileLen(nomFichTmp) + Len(stPost2)
    
    ' headers
    stLoad = "Content-Type: multipart/form-data, boundary=" & strBoundary & vbCrLf & _
             "Content-Length: " & lgTot & vbCrLf & vbCrLf
    
    On Error GoTo ErrorHandle
    sUrl = v_sURL
    sUrl = sUrl & "?filename=" & nomFichDest & "&"
    sUrl = sUrl & "v_CheminFichier=" & v_CheminFichier_Serveur & "&"
    sUrl = sUrl & "v_NomFichier=" & v_NomFichier_Serveur & "&"
    sUrl = sUrl & "v_ExtensionFichier=" & v_ExtensionFichier_Serveur & "&"
    sUrl = sUrl & "v_CheminHTTP=" & p_HTTP_CheminDépot & "&"
    sUrl = sUrl & "v_NumUtil=" & p_NumUtil & "&"

    If v_DeLocker Then
        sUrl = sUrl & "v_DeLocker=O"
    Else
        sUrl = sUrl & "v_DeLocker=N"
    End If
    
    bresult = HTTP_InitialHttpConnect(sUrl)
    If bresult = False Then
        HTTP_putfile = HTTP_PUT_ERREUR
        HTTP_PUT_LIB = "PutFile : Erreur InitialHttpConnect sret=" & left(Trim(buf), 3) & " " & sUrl
        GoTo ErrorHandle
    End If
    
    ret = HttpAddRequestHeaders(m_lHttpRequest, stLoad, Len(stLoad), HTTP_ADDREQ_FLAG_ADD)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_putfile = HTTP_PUT_ERREUR
        HTTP_PUT_LIB = "PutFile : Erreur HttpAddRequestHeaders"
        GoTo ErrorHandle
    End If
    
    BufferIn.dwStructSize = 40
    BufferIn.Next = 0
    BufferIn.lpcszHeader = 0
    BufferIn.dwHeadersLength = 0
    BufferIn.dwHeadersTotal = 0
    BufferIn.lpvBuffer = 0
    BufferIn.dwBufferLength = 0
    BufferIn.dwBufferTotal = lgTot
    BufferIn.dwOffsetLow = 0
    BufferIn.dwOffsetHigh = 0
    
    ret = HttpSendRequestEx(m_lHttpRequest, BufferIn, 0, 0, 0)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_putfile = HTTP_PUT_ERREUR
        HTTP_PUT_LIB = "PutFile : Erreur HttpSendRequestEx"
        GoTo ErrorHandle
    End If
    
    nb_transmis = 0
    nb_total = 0
    ret = InternetWriteFile(m_lHttpRequest, stPost1, Len(stPost1), nb_transmis)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_putfile = HTTP_PUT_ERREUR
        HTTP_PUT_LIB = "PutFile : Erreur InternetWriteFile"
        GoTo ErrorHandle
    End If
    
    maxn = 1024
    
    hFileLocal = CreateFile(nomFichTmp, GENERIC_READ, _
                        FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                        0&, OPEN_EXISTING, 0, 0)
    
    lgTotal = 0
    Do While hFileLocal > 0
        buf = String(maxn, Chr(0))
        ret = ReadFile(hFileLocal, ByVal buf, maxn, n, ByVal 0&)
        If n > 0 Then
            ret = InternetWriteFile(m_lHttpRequest, buf, n, nb_transmis)
        
            If iRetTaille = HTTP_TAILLE_OK Then
                lgTotal = lgTotal + n
                If p_HTTP_Form_Frame.PgbarHTTPDTaille.max < lgTotal Then
                    p_HTTP_Form_Frame.PgbarHTTPDTaille.max = lgTotal
                End If
                p_HTTP_Form_Frame.PgbarHTTPDTaille.Value = lgTotal
            
                TimePrem = DateTime.Now()
                If n > 0 Then
                    NbSeconde = DateDiff("s", TimeDébut, TimePrem)
                    If NbSeconde = 0 Then NbSeconde = 1
                    p_HTTP_Form_Frame.PgbarHTTPDTemps.max = NbSeconde / n * p_HTTP_TailleFichier
                
                    ResteSeconde = NbSeconde / n * (p_HTTP_TailleFichier - lgTotal)
                    If ResteSeconde < 0 Then
                        ResteSeconde = 0
                        p_HTTP_Form_Frame.PgbarHTTPDTemps.Value = ResteSeconde
                        p_HTTP_Form_Frame.lblHTTPDTemps.Caption = "Terminé"
                    Else
                        p_HTTP_Form_Frame.PgbarHTTPDTemps.Value = ResteSeconde
                    End If
                End If
                p_HTTP_Form_Frame.Refresh
            End If
        
            bresult = CBool(ret)
            If bresult = False Then
                HTTP_putfile = HTTP_PUT_ERREUR
                HTTP_PUT_LIB = "PutFile : Erreur InternetWriteFile"
                GoTo ErrorHandle
            End If
            nb_total = nb_total + n
        Else
            RetClose = CloseHandle(hFileLocal)
            If RetClose <> 1 Then
                HTTP_putfile = HTTP_PUT_ERREUR
                HTTP_PUT_LIB = "PutFile : Impossible de fermer " & nomFichTmp
                Kill nomFichTmp
                GoTo ErrorHandle
            End If
            hFileLocal = -1
        End If
    Loop
    
    ' moment de l'écriture du fichier et du déplacement
    ret = InternetWriteFile(m_lHttpRequest, stPost2, Len(stPost2), nb_transmis)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_putfile = HTTP_PUT_ERREUR
        HTTP_PUT_LIB = "PutFile : Erreur InternetWriteFile"
        GoTo ErrorHandle
    End If
    
    ret = HttpEndRequest(m_lHttpRequest, 0, 0, 0)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_putfile = HTTP_PUT_ERREUR
        HTTP_PUT_LIB = "PutFile : Erreur HttpEndRequest"
        GoTo ErrorHandle
    End If
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(m_lHttpRequest, HTTP_QUERY_CONTENT_TYPE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_putfile = HTTP_PUT_ERREUR
        HTTP_PUT_LIB = "PutFile : Erreur HttpQueryInfo"
        GoTo ErrorHandle
    End If
    sret = "HTTP_QUERY_CONTENT_TYPE=" & left(buf, n) & vbCrLf
    
    buf = String(maxn, Chr(0))
    n = maxn
    hindex = 0
    ret = HttpQueryInfo(m_lHttpRequest, HTTP_QUERY_STATUS_CODE, ByVal buf, n, hindex)
    bresult = CBool(ret)
    If bresult = False Then
        HTTP_putfile = HTTP_PUT_ERREUR
        HTTP_PUT_LIB = "PutFile : Erreur HttpQueryInfo"
        GoTo ErrorHandle
    End If
    sret = sret & "HTTP_QUERY_STATUS_CODE=" & left(buf, n) & vbCrLf
    
    If val(left(Trim(buf), 1)) > 3 Then
        HTTP_putfile = HTTP_PUT_ERREUR
        HTTP_PUT_LIB = "PutFile : Erreur Trim(Buf) " & buf
        GoTo ErrorHandle
    End If
    
    
    If HTTP_putfile = HTTP_PUT_OK Then
        If iRetTaille = HTTP_TAILLE_OK Then
            'MsgBox FileLen(nomFichTmp) & " " & p_HTTP_TailleFichier
            ' comparer la taille des deux fichiers
            If p_HTTP_TailleFichier <> FileLen(nomFichTmp) Then
                HTTP_putfile = HTTP_PUT_PAS_COMPLET
                HTTP_PUT_LIB = "Taille du fichier sur le serveur " & p_HTTP_TailleFichier & Chr(13) & Chr(10) & "Taille du fichier chargé " & FileLen(nomFichTmp)
                p_HTTP_Form_Frame.lblHTTPD.Caption = HTTP_PUT_LIB
            Else
                p_HTTP_Form_Frame.lblHTTPD.Caption = "Chargement terminé avec succès"
            End If
        End If
    End If
    
    
    HTTP_CloseConnect
    
    Exit Function
        
ErrorHandle:
    'Debug.Print Err.Description
    'Debug.Print Err.HelpContext
    'Debug.Print Err.LastDllError
    Err.Clear
    HTTP_CloseConnect
End Function

Public Function HTTP_CloseConnect()
    
    Dim ret As Long
    Dim sT As String
    
    InternetCloseHandle (m_lHttpRequest)
    InternetCloseHandle (m_lInternetConnect)
    InternetCloseHandle (p_lInternetSession)

End Function

Public Function HTTP_getNomUtil(ByVal v_numutil As Integer) As String
    Dim sql As String, rs As rdoResultset
    
    sql = "select u_Nom, u_Prenom from utilisateur" _
        & " where u_Num = " & v_numutil
    If Odbc_SelectV(sql, rs) = P_ERREUR Then
        Exit Function
    End If
    If Not rs.EOF Then
        HTTP_getNomUtil = rs("u_Prenom") & "." & rs("u_Nom")
    End If
End Function
