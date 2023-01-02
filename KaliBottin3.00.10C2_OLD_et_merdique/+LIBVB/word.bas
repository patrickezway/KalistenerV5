Attribute VB_Name = "Mword"
Option Explicit

' Mode de travail de Word_Fusionner
Public Const WORD_IMPRESSION = 0
Public Const WORD_VISU = 1
Public Const WORD_MODIF = 2
Public Const WORD_CREATE = 3

' Ce qu'il y a à faire au début de l'appel à Word_Fusionner
Public Const WORD_DEB_CROBJ = 0
Public Const WORD_DEB_OUVDOC = 1
Public Const WORD_DEB_RIEN = 2

' Ce qu'il y a à faire à la fin de l'appel à Word_Fusionner
Public Const WORD_FIN_FERMDOC = 0
Public Const WORD_FIN_RAZOBJ = 1
Public Const WORD_FIN_RIEN = 2

Public Word_Doc As Document
Public Word_Obj As Word.Application
Public Word_EstActif As Boolean

' Pour Word_CreerModele
Public Type WORD_SSIGNET
    nom As String
    indice As Integer
End Type
Public Word_tblsignet() As WORD_SSIGNET
Public Word_nbsignet As Integer

' Pour Word_Fusionner
Private g_garder_bookmark As Boolean
Private Type W_STBLCHP
    nombk As String
    exist As Integer
End Type
' Valeurs de exist
Private Const W_AEVALUER = 0
Private Const W_NON = 1
Private Const W_OUI = 2


' Le fichier doit être en local dans tous les cas
Public Function Word_AfficherDoc(ByVal v_nomdoc As String, _
                               ByVal v_passwd As String, _
                               ByVal v_fimprime As Boolean, _
                               ByVal v_fmodif As Boolean, _
                               ByVal v_nomdata As String) As Integer

    Dim s As String, nombk As String, nomdoc As String, nomdot As String
    Dim fexist As Boolean, a_redim As Boolean
    Dim i As Integer, j As Integer, fd As Integer, n As Integer, pos As Integer
    Dim imode As Integer
    Dim s_bk As Variant
    Dim range As range
    Dim g_cmsword As New CWord
    
    If Not FICH_FichierExiste(v_nomdoc) Then
        MsgBox "Impossible d'ouvrir " & v_nomdoc & vbCr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
        Word_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
'    If Not v_fmodif Then
'        nomdoc = p_chemin_appli + "\tmp\" + p_CodeUtil + ".doc"
'        If FICH_CopierFichier(v_nomdoc, nomdoc) = P_ERREUR Then
'            Word_AfficherDoc = P_ERREUR
'            Exit Function
'        End If
'    Else
        nomdoc = v_nomdoc
'    End If
    
    If Word_Fusionner(nomdoc, _
                  "", _
                  v_nomdata, _
                  True, _
                  "", _
                  True, _
                  v_passwd, _
                  False, _
                  WORD_CREATE, _
                  0, _
                  WORD_DEB_CROBJ, _
                  WORD_FIN_RIEN) = P_ERREUR Then
        Word_AfficherDoc = P_ERREUR
        Exit Function
    End If
                        
    Set g_cmsword.App = Word_Obj

    If Not v_fmodif Then
        Word_Doc.Saved = False
        Word_Doc.Save
        Word_Doc.Close
        If Word_OuvrirDoc(nomdoc, _
                            Not v_fmodif, _
                            v_passwd, _
                            Word_Doc) = P_ERREUR Then
            Word_AfficherDoc = P_ERREUR
            Exit Function
        End If
    End If
    
    Set g_cmsword.doc = Word_Doc
    
    If Not v_fmodif Then
        If Not v_fimprime Then
            nomdot = p_chemin_appli + "\Modele\KaliDocNoFct.dot"
            imode = 3
        Else
            nomdot = p_chemin_appli + "\Modele\KaliDocImp.dot"
            imode = 2
        End If
        g_cmsword.doc.Protect wdAllowOnlyComments
    Else
        nomdot = p_chemin_appli + "\Modele\KaliDoc.dot"
        imode = 1
    End If
    
    If g_cmsword.InitConfig(imode, nomdot, False) = P_ERREUR Then
        Word_AfficherDoc = P_ERREUR
        Exit Function
    End If
    
    On Error GoTo lab_fin_err
    g_cmsword.App.Visible = True
    g_cmsword.App.ActiveWindow = True
    g_cmsword.doc.ActiveWindow.View.type = wdPageView
    If g_cmsword.App.WindowState <> wdWindowStateMaximize Then
        g_cmsword.App.WindowState = wdWindowStateMaximize
    End If
    g_cmsword.App.Activate
    
    On Error GoTo lab_fin
    While g_cmsword.App.Visible
        SYS_Sleep (500)
Debug.Print g_cmsword.lafin
'        DoEvents
    Wend
    
lab_fin:
Debug.Print "Fin Word_AfficherDoc " & Time()
    On Error GoTo 0
    Word_EstActif = False
    
    Word_AfficherDoc = P_OK
    Exit Function

lab_fin_err:
    MsgBox "Erreur WORD " & vbCr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
'    Set g_cmsword = Nothing
    Word_AfficherDoc = P_ERREUR
    Exit Function

End Function

Public Function Word_ChangerPasswd(ByVal v_nomdoc As String, _
                                   ByVal v_o_passwd As String, _
                                   ByVal v_n_passwd As String) As Integer
                             
    If Word_Init() = P_ERREUR Then
        Word_ChangerPasswd = P_OK
        Exit Function
    End If
    
    If Word_OuvrirDoc(v_nomdoc, False, v_o_passwd, Word_Doc) = P_ERREUR Then
        Word_ChangerPasswd = P_ERREUR
        Exit Function
    End If
    
    On Error GoTo err_word
    ' Ruse : sinon le document n'est pas enregistré ...
    Word_Doc.Saved = False
    Word_Doc.Password = v_n_passwd
    Word_Doc.Save
    Word_Doc.Close
    On Error GoTo 0
    
    Word_ChangerPasswd = P_OK
    Exit Function

err_word:
    MsgBox "Erreur WORD " & vbCr & vbLf & Err.Description, vbCritical + vbOKOnly, ""
    Word_ChangerPasswd = P_ERREUR
    Exit Function

End Function

Public Sub Word_ConvHTML(ByVal v_nomdoc As String, _
                         ByVal v_nomhtml As String, _
                         ByVal v_conv As Integer)
                         
    If Word_Init() = P_ERREUR Then
        Exit Sub
    End If
    
    If Word_OuvrirDoc(v_nomdoc, False, "", Word_Doc) = P_ERREUR Then
        Call Word_Quitter(WORD_FIN_RAZOBJ)
        Exit Sub
    End If
    
    Call Word_Doc.SaveAs(FileName:=v_nomhtml, FileFormat:=v_conv)
    
    Call Word_Quitter(WORD_FIN_RAZOBJ)
    
End Sub
                         
Public Sub Word_CopierModele(ByVal v_nomdoc As String, _
                             ByVal v_nommodele As String, _
                             ByVal v_bcopie_entete As Boolean, _
                             ByVal v_bcopie_corps As Boolean, _
                             ByVal v_passwd As String)

    Dim doc_modele As Document
    Dim arange As range

    If Word_Init() = P_ERREUR Then
        Exit Sub
    End If
    
    ' Ouvre le document d'origine
    If Word_OuvrirDoc(v_nomdoc, False, v_passwd, Word_Doc) = P_ERREUR Then
        Call Word_Quitter(WORD_FIN_RAZOBJ)
        Exit Sub
    End If
    
    ' Ouvre le modèle
    If Word_OuvrirDoc(v_nommodele, True, "", doc_modele) = P_ERREUR Then
        Call Word_Doc.Close(savechanges:=wdDoNotSaveChanges)
        Call Word_Quitter(WORD_FIN_RAZOBJ)
        Exit Sub
    End If
    
    ' Recopie de l'entete -> on part du modèle
    If v_bcopie_entete Then
        ' Recopie du corps
        If v_bcopie_corps Then
            Call w_suppr_bk_doublon(Word_Doc, doc_modele)
        Else
            ' Efface le corps du modèle
            Set arange = doc_modele.range
            arange.Select
            arange.Text = ""
        End If
        ' Recopie du corps du document dans le modèle
        Call w_copier_corps(Word_Doc, doc_modele)
        Call FICH_EffacerFichier(v_nomdoc, False)
        ' Le modèle devient le nouveau document
        Call doc_modele.SaveAs(FileName:=v_nomdoc, Password:=v_passwd)
        Call doc_modele.Close(savechanges:=wdDoNotSaveChanges)
    ' Pas d'entete -> on part du document
    Else
        ' Recopie du corps du modèle
        If v_bcopie_corps Then
            ' Recopie du corps du modèle dans le document
            Call w_copier_corps(doc_modele, Word_Doc)
        Else
            Call doc_modele.Close(savechanges:=wdDoNotSaveChanges)
        End If
        Call Word_Doc.Close(savechanges:=wdSaveChanges)
    End If
    
End Sub

Public Function Word_CreerModele(v_form As Form, _
                                 ByVal v_nomfich_chp As String, _
                                 ByVal v_nomdoc As String, _
                                 ByRef r_nomdoctmp As String)
                        
    Dim tbl_name() As String, ssys As String, nomdoc As String, nomdoc2 As String
    Dim nomecr As String, sdat_av As String, sdat_ap As String, nomlocal As String
    Dim ya_un_tab As Boolean, fmodif As Boolean, est_danstab As Boolean, inheadfoot As Boolean
    Dim tbl_inhf() As Boolean, trouve As Boolean
    Dim pos As Integer, i As Integer, notab As Integer, reponse As Integer, n As Integer
    Dim CR As Integer, ntab As Integer, lig_tab As Integer, j As Integer
    Dim siz_tab As Long, tbl_start() As Long, tbl_end() As Long, numecr As Long
    Dim arange As range, trange As range

    ' On vérifie l'existance de champ.txt
    If Not FICH_FichierExiste(v_nomfich_chp) Then
        Call MsgBox("Le fichier '" & v_nomfich_chp & "' étant inaccessible, vous ne pouvez pas accéder aux modèles.", vbInformation + vbOKOnly, "")
        Word_CreerModele = P_ERREUR
        Exit Function
    End If
    
    fmodif = True
lab_debut:
    ' Le .doc n'existe pas
    If Not FICH_FichierExiste(v_nomdoc) Then
        nomdoc = left$(v_nomdoc, Len(v_nomdoc) - 3) + "mod*"
        ' Le .mod existe (doc en cours de modif)
        If FICH_FichierExiste(nomdoc) Then
            nomdoc2 = Dir$(nomdoc)
            pos = InStr(nomdoc2, ".mod")
            numecr = Mid$(nomdoc2, pos + 4)
            If numecr = p_NumEcr Then
                nomdoc2 = left$(nomdoc, Len(nomdoc) - 4) & "mod" & numecr
                Call FICH_RenommerFichier(nomdoc2, v_nomdoc)
                GoTo lab_debut
            End If
            ' Récupère le N° d'écran effectuant la modif
            If Odbc_RecupVal("select E_Nom from Ecran where E_Num=" & numecr, _
                             nomecr) = P_ERREUR Then
                Word_CreerModele = P_ERREUR
                Exit Function
            End If
            ' Propose la lecture seule
            reponse = MsgBox("Le fichier '" & v_nomdoc & "' est en cours de modification sur '" & nomecr & "'." & vbCr & vbLf & vbCr & vbLf _
                             & "Voulez-vous y accéder en consultation seulement ?", vbQuestion + vbYesNo, "")
            If reponse = vbNo Then
                Word_CreerModele = P_NON
                Exit Function
            End If
            nomdoc = left$(v_nomdoc, Len(v_nomdoc) - 3) & "mod" & numecr
            fmodif = False
        Else
            Call MsgBox("Impossible d'accéder au fichier '" & v_nomdoc & "'.", vbInformation + vbOKOnly, "")
            Word_CreerModele = P_ERREUR
            Exit Function
        End If
    Else
        nomdoc = v_nomdoc
    End If
    
    nomlocal = p_chemin_appli & "\tmp\" & p_CodeUtil & Format(Time, "hhmmss") & ".doc"
    
    ' Lecture seulement
    If Not fmodif Then
        If FICH_CopierFichier(nomdoc, nomlocal) = P_ERREUR Then
            Word_CreerModele = P_ERREUR
            Exit Function
        End If
        Call Word_AfficherDoc(nomlocal, "", True, False, "")
        Call FICH_EffacerFichier(nomlocal, False)
        Word_CreerModele = P_NON
        Exit Function
    End If
    
    ' *** Ouverture en modif ***
    
    ' Renomme .doc en .mod sur le serveur
    nomdoc = left$(v_nomdoc, Len(v_nomdoc) - 3) & "mod" & p_NumEcr
    Call FICH_RenommerFichier(v_nomdoc, nomdoc)
    Call FICH_CopierFichier(nomdoc, nomlocal)
    sdat_av = FICH_FichierDateTime(nomlocal)
    
    ' L'utilisateur paramètre son document
    Call Word_AfficherDoc(nomlocal, "", True, True, "")
    
    sdat_ap = FICH_FichierDateTime(nomlocal)
    ' Pas de modification
    If sdat_ap = sdat_av Then
        ' Renomme .mod en .doc
        Call FICH_RenommerFichierNoMess(nomdoc, v_nomdoc)
        Call FICH_EffacerFichier(nomlocal, False)
        Word_CreerModele = P_NON
        Exit Function
    End If
    
    Call FRM_ResizeForm(v_form, v_form.width, v_form.Height)
    DoEvents
    v_form.Refresh
    
    If Word_Init() = P_ERREUR Then
        Word_CreerModele = P_ERREUR
        Exit Function
    End If
    
    If Word_OuvrirDoc(nomlocal, False, "", Word_Doc) = P_ERREUR Then
        Word_CreerModele = P_ERREUR
        Exit Function
    End If
    
    On Error GoTo err_word
'Word_Obj.Visible = True
    Word_Doc.ActiveWindow.View.ShowFieldCodes = False
    Word_Doc.ActiveWindow.View.type = wdPageView
    
    Call w_init_tblsignet
    
    ' En-tête
    For i = 1 To Word_Doc.Sections(1).Headers.Count
        Set trange = Word_Doc.Sections(1).Headers(i).range
        Call w_conv_champ_en_signet(trange)
    Next i
    ' Corps
    Word_Doc.ActiveWindow.View.type = wdPageView
    Word_Doc.ActiveWindow.View.SeekView = wdSeekMainDocument
    Set trange = Word_Doc.StoryRanges(wdMainTextStory)
    Call w_conv_champ_en_signet(trange)
' On laisse de coté pour l'instant
'    For i = 1 To Word_Doc.Shapes.Count
'        If Word_Doc.Shapes(i).TextFrame.HasText Then
'            Set trange = Word_Doc.Shapes(i).TextFrame.TextRange
'            Call w_conv_champ_en_signet(trange)
'        End If
'    Next i
    ' Pied
    For i = 1 To Word_Doc.Sections(1).Footers.Count
        Set trange = Word_Doc.Sections(1).Footers(i).range
        Call w_conv_champ_en_signet(trange)
    Next i

    ' On supprime les bookmark tableau glob
    i = 1
    While i <= Word_Doc.Bookmarks.Count
        If w_est_champ_tableau_global(Word_Doc.Bookmarks(i).Name) Then
            Word_Doc.Bookmarks(i).Delete
            i = i - 1
        End If
        i = i + 1
    Wend
    
    ' Insertion des bookmark tableau glob
    siz_tab = 0
    ya_un_tab = False
    For i = 1 To Word_Doc.Bookmarks.Count
        If w_est_champ_tableau(Word_Doc.Bookmarks(i).Name) = P_OUI Then
            est_danstab = w_ajouter_bkglob(Word_Doc.Bookmarks(i).Name, inheadfoot)
            If Not est_danstab Then GoTo lab_suivant
            notab = CInt(Mid$(Word_Doc.Bookmarks(i).Name, 2, 2))
            ya_un_tab = True
            If siz_tab < notab Then GoTo lab_cr_tab
lab_book_suiv:
            If tbl_start(notab) > Word_Doc.Bookmarks(i).start Then
                tbl_start(notab) = Word_Doc.Bookmarks(i).start
            End If
            If tbl_end(notab) < Word_Doc.Bookmarks(i).End Then
                tbl_end(notab) = Word_Doc.Bookmarks(i).End
            End If
            tbl_inhf(notab) = inheadfoot
        End If
lab_suivant:
    Next i
    If ya_un_tab Then
        For i = 1 To UBound(tbl_name)
            If tbl_name(i) <> "" Then
                If tbl_inhf(i) Then
                    trouve = False
                    For j = 1 To Word_Doc.Sections(1).Headers.Count
                        Word_Doc.Sections(1).Headers(j).range.Select
                        If w_ya_bktbl_dans_sel(tbl_name(i)) Then
                            Set arange = Word_Doc.Sections(1).Headers(j).range
                            arange.start = tbl_start(i)
                            arange.End = tbl_end(i) + 1
                            trouve = True
                            Exit For
                        End If
                    Next j
                    ' Pied
                    If Not trouve Then
                        For j = 1 To Word_Doc.Sections(1).Footers.Count
                            Word_Doc.Sections(1).Footers(j).range.Select
                            If w_ya_bktbl_dans_sel(tbl_name(i)) Then
                                Set arange = Word_Obj.Selection.range
                                arange.start = tbl_start(i)
                                arange.End = tbl_end(i) + 1
                                trouve = True
                                Exit For
                            End If
                        Next j
                    End If
                Else
                    If w_rangeb(tbl_start(i), tbl_end(i) + 1, arange) = P_OK Then
                        trouve = True
                    End If
                End If
                If trouve Then
                    If w_add_bookmark(tbl_name(i), arange) = P_ERREUR Then GoTo lab_fin_err
                End If
            End If
        Next i
    End If
    
    Call w_init_tblsignet
    
    GoTo lab_fin_ok
    
lab_cr_tab:
    siz_tab = notab
    ReDim Preserve tbl_name(notab) As String
    pos = InStr(Mid$(Word_Doc.Bookmarks(i).Name, 5), "_")
    tbl_name(notab) = left$(Word_Doc.Bookmarks(i).Name, pos + 3)
    ReDim Preserve tbl_start(notab) As Long
    tbl_start(notab) = 9999
    ReDim Preserve tbl_end(notab) As Long
    tbl_end(notab) = 0
    ReDim Preserve tbl_inhf(notab) As Boolean
    GoTo lab_book_suiv
    
err_word:
'Resume Next
    MsgBox "Erreur WORD ", Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    GoTo lab_fin_err

lab_fin_err:
    Call Word_Doc.Close(savechanges:=wdDoNotSaveChanges)
    Call FICH_EffacerFichier(nomlocal, False)
    Word_CreerModele = P_ERREUR
    Exit Function
    
lab_fin_ok:
    Call Word_Doc.Close(savechanges:=wdSaveChanges)
    ' Maj du document
    n = 1
    ' Boucle pour attendre que Word relache le fichier
    Do
        CR = FICH_CopierFichierNoMess(nomlocal, nomdoc)
        If CR = P_NON Then
            Call SYS_Sleep(10)
        End If
        n = n + 1
    Loop Until CR = P_OUI Or CR = P_ERREUR Or n > 100
    If CR <> P_OUI Then
        Call MsgBox("Impossible de copier " & nomlocal & " dans " & nomdoc & " (cr:" & CR & " n:" & n & ")", vbCritical + vbOKOnly, "")
        Call FICH_RenommerFichier(nomdoc, v_nomdoc)
        Word_CreerModele = P_ERREUR
        Exit Function
    End If
    ' Renomme .mod en .doc
    Call FICH_RenommerFichier(nomdoc, v_nomdoc)
    ' ON N'EFFACE pas "nomlocal"
    r_nomdoctmp = nomlocal
    Word_CreerModele = P_OUI
    Exit Function
    
End Function

Public Function Word_Fusionner(ByVal v_nommod As String, _
                           ByVal v_nominit As String, _
                           ByVal v_nomdata As String, _
                           ByVal v_garder_bookmark As Boolean, _
                           ByVal v_nomdest As String, _
                           ByVal v_ecraser As Boolean, _
                           ByVal v_passwd As String, _
                           ByVal v_word_visible As Boolean, _
                           ByVal v_word_mode As Integer, _
                           ByVal v_nbex As Integer, _
                           ByVal v_deb_mode As Integer, _
                           ByVal v_fin_mode As Integer) As Integer
    
    Dim s As String
    Dim str_entete As String, nomchp As String, str_data As String
    Dim chptab As String, nomtab As String, chp As String
    Dim sval As String, nombk As String
    Dim encore As Boolean, again As Boolean
    Dim ya_book_glob As Boolean, a_redim As Boolean, b_fairefusion As Boolean
    Dim frempl As Boolean, encore_bk As Boolean
    Dim fd As Integer, poschp As Integer, pos As Integer, pos2 As Integer
    Dim i As Integer, j As Integer, ntab As Integer, nlig As Integer, n As Integer
    Dim lig_tab As Integer, col As Integer, ind As Integer
    Dim idebtab As Long, ifintab As Long
    Dim tbl_chp() As W_STBLCHP
    Dim arange As range, arange2 As range
    Dim doc2 As Document
    Dim dochf As Variant
    
    If v_deb_mode = WORD_DEB_CROBJ Then
        If Word_Init() = P_ERREUR Then
            Word_Fusionner = P_ERREUR
            Exit Function
        End If
    End If
    
    If v_deb_mode <> WORD_DEB_RIEN Then
        If Word_OuvrirDoc(v_nommod, False, v_passwd, Word_Doc) = P_ERREUR Then
            Word_Fusionner = P_ERREUR
            Exit Function
        End If
    End If
    
    g_garder_bookmark = v_garder_bookmark
    
    b_fairefusion = True
    If v_nomdata = "" Then
        b_fairefusion = False
        GoTo lab_fin_fusion
    End If
    
'v_word_visible = True
    If v_word_visible Then
        Word_Obj.Visible = True
        Word_Obj.ActiveWindow = True
        a_redim = False
        On Error Resume Next
        If Word_Obj.WindowState <> wdWindowStateMaximize Then
            a_redim = True
        End If
        Word_Obj.Activate
        If a_redim Then Word_Obj.WindowState = wdWindowStateMaximize
        On Error GoTo 0
    Else
        Word_Obj.Visible = False
    End If
    
    If v_nomdest <> "" Then
        On Error GoTo err_sav_dest
        If v_ecraser Then
            If v_passwd <> "" Then Word_Doc.Password = v_passwd
            Word_Doc.SaveAs FileName:=v_nomdest
            On Error GoTo 0
        Else
            If v_deb_mode <> WORD_DEB_RIEN Then
                If Word_OuvrirDoc(v_nomdest, False, "", doc2) = P_ERREUR Then
                    GoTo lab_fin_err1
                End If
                While doc2.Bookmarks.Count > 0
                    doc2.Bookmarks(1).Delete
                Wend
                Set arange = Word_Doc.range
                arange.Copy
                Set arange2 = doc2.Content
                arange2.Collapse wdCollapseEnd
                On Error GoTo lab_err_paste
                arange2.Paste
                On Error GoTo 0
                Call Word_Doc.Close(savechanges:=wdDoNotSaveChanges)
                Set Word_Doc = Word_Obj.ActiveDocument
            Else
                While Word_Doc.Bookmarks.Count > 0
                    Word_Doc.Bookmarks(1).Delete
                Wend
                If Word_OuvrirDoc(v_nommod, False, "", doc2) = P_ERREUR Then
                    GoTo lab_fin_err1
                End If
                Set arange = doc2.range
                arange.Copy
                Set arange2 = Word_Doc.Content
                arange2.Collapse wdCollapseEnd
                On Error GoTo lab_err_paste
                arange2.Paste
                On Error GoTo 0
                Call doc2.Close(savechanges:=wdDoNotSaveChanges)
            End If
        End If
    End If
    
    On Error GoTo err_word1
    Word_Doc.MailMerge.MainDocumentType = wdNotAMergeDocument
    
    If Word_Doc.Bookmarks.Count < 1 Then GoTo lab_copie_fich

    On Error GoTo err_open_fus
    fd = FreeFile
    Open v_nomdata For Input As #fd
    On Error GoTo err_word2

    ' Ligne d'entete
    Line Input #fd, str_entete
    
    encore = True
    Do While encore
        If str_entete = "" Then GoTo lab_copie_fich
        poschp = InStr(str_entete, ";")
        If poschp = 0 Then
            ' fini
            nomchp = str_entete
            encore = False
        Else
            nomchp = left(str_entete, poschp - 1)
            str_entete = Right(str_entete, Len(str_entete) - poschp)
        End If
        pos = InStr(nomchp, "#")
        If pos > 0 Then
            ' c'est un tableau
            chptab = Mid$(nomchp, pos + 1)
            pos2 = InStr(pos + 1, chptab, "#")
            chptab = Mid$(chptab, pos2 + 1)
            nomtab = Mid$(nomchp, 2, pos2 - 1) & "_1"
            ya_book_glob = Word_Doc.Bookmarks.Exists(nomtab)
            ' Chargement du nom des champs
            again = True
            i = 0
            Do While again
                pos = InStr(chptab, "|")
                If pos > 0 Then
                    chp = left(chptab, pos - 1)
                    chptab = Right(chptab, Len(chptab) - pos)
                Else
                    chp = chptab
                    again = False
                End If
                i = i + 1
                ReDim Preserve tbl_chp(i) As W_STBLCHP
                tbl_chp(i).nombk = chp
                tbl_chp(i).exist = W_AEVALUER
            Loop
            ' Détermine s'il y a un tableau word
            If ya_book_glob Then
                lig_tab = -1
                Call w_bk_dans_tableau(nomtab, dochf, ntab, lig_tab)
                If lig_tab > 0 Then
                    ' Efface toutes les lignes du tableau sauf la 1e
                    For i = lig_tab + 1 To dochf.Tables(ntab).Rows.Count
                        dochf.Tables(ntab).Rows(lig_tab + 1).Delete
                    Next i
                    dochf.Tables(ntab).Rows(lig_tab).Select
                Else
                    ' Sauvegarde position bookmark tableau glob
                    idebtab = Word_Doc.Bookmarks(nomtab).start
                    Word_Doc.Bookmarks(nomtab).Select
                End If
                Word_Obj.Selection.Copy
            End If
            ' Traitement des lignes
            again = True
            frempl = True
            Do While again
                If w_lire_fich(fd, str_data) = P_NON Then Exit Do
                If Right$(str_data, 1) = ";" Then again = False
                If Not frempl Then GoTo lab_lig_suiv
                For col = 1 To UBound(tbl_chp())
                    pos = InStr(str_data, "|")
                    If pos > 0 Then
                        sval = left(str_data, pos - 1)
                        str_data = Right(str_data, Len(str_data) - pos)
                    Else
                        If Not again Then
                            sval = left$(str_data, Len(str_data) - 2)
                        Else
                            sval = left$(str_data, Len(str_data) - 1)
                        End If
                    End If
                    If sval = "" Then GoTo lab_chp_suiv
                    sval = STR_Remplacer(sval, "##", vbCr)
                    ' pour chaque donnée
                    If tbl_chp(col).exist = W_NON Then GoTo lab_chp_suiv
                    ind = 1
                    Do
                        encore_bk = False
                        nombk = left$(nomtab, Len(nomtab) - 2) & "_" & tbl_chp(col).nombk & "_" & ind
                        If tbl_chp(col).exist = W_AEVALUER Then
                            If Word_Doc.Bookmarks.Exists(nombk) = False Then
                                tbl_chp(col).exist = W_NON
                                GoTo lab_chp_suiv
                            Else
                                tbl_chp(col).exist = W_OUI
                            End If
                        End If
                        If tbl_chp(col).exist = W_OUI And Word_Doc.Bookmarks.Exists(nombk) Then
                            If w_put_txtbk(sval, nombk) = P_ERREUR Then GoTo lab_fin_err
                            If again And g_garder_bookmark And ya_book_glob Then Word_Doc.Bookmarks(nombk).Delete
                            encore_bk = True
                            ind = ind + 1
                        End If
                    Loop Until encore_bk = False
lab_chp_suiv:
                Next col
lab_lig_suiv:
                If Not ya_book_glob Then
                    frempl = False
                ElseIf again Then
                    If Word_Doc.Bookmarks.Exists(nomtab) Then Word_Doc.Bookmarks(nomtab).Delete
                    ' Recopie du bookmark tableau à la ligne précédente
                    If lig_tab <> -1 Then
                        dochf.Tables(ntab).Rows(lig_tab).Select
                        Word_Obj.Selection.Paste
                    Else
                        If w_range(idebtab, idebtab, arange) = P_ERREUR Then GoTo lab_fin_err
                        arange.InsertBefore vbCr
                        If w_range(idebtab, idebtab, arange) = P_ERREUR Then GoTo lab_fin_err
                        arange.Paste
                    End If
                End If
            Loop
        Else
            ' Récupère les données
            If w_lire_fich(fd, str_data) = P_NON Then GoTo lab_fin_err
            If str_data <> "" Then
                ' Remplace le bookmark
                ind = 1
                Do
                    encore_bk = False
                    nombk = nomchp & "_" & ind
                    If Word_Doc.Bookmarks.Exists(nombk) = True Then
                        If ind = 1 Then
                            str_data = left$(str_data, Len(str_data) - 1)
                            str_data = STR_Remplacer(str_data, "|", vbCr)
                        End If
                        If w_put_txtbk(str_data, nombk) = P_ERREUR Then GoTo lab_fin_err
                        encore_bk = True
                        ind = ind + 1
                    End If
                Loop Until encore_bk = False
            End If
        End If
    Loop

    ' Rapatriement du document à recopier
lab_copie_fich:
    If v_nominit <> "" Then
        If Word_OuvrirDoc(v_nominit, True, "", doc2) = P_ERREUR Then
            GoTo lab_fin_err2
        End If
        Call w_copier_corps(doc2, Word_Doc)
    End If
    
lab_fin_fusion:
'    If Not b_fairefusion Then
'        Word_Obj.ActivePrinter = Printer.DeviceName
'        Word_Obj.ActiveDocument.PrintOut Background:=False, Copies:=v_nbex
'        GoTo lab_fin
'    End If
    
    If v_word_mode = WORD_IMPRESSION Then
        ' Fax ?
        If w_lire_fich(fd, str_data) = P_OUI Then
            Close #fd
            On Error GoTo err_fax
            Word_Doc.SendFax left$(str_data, Len(str_data) - 1)
            On Error GoTo err_word2
        Else
            Close #fd
            Word_Obj.ActivePrinter = Printer.DeviceName
            If Word_Doc.Bookmarks.Exists("ImpPaysage") = True Then
                If w_put_txtbk("ImpPaysage", "") = P_ERREUR Then GoTo lab_fin_err
                Word_Doc.PageSetup.Orientation = wdOrientLandscape
            End If
            Word_Doc.PrintOut Background:=False, Copies:=v_nbex
        End If
    ElseIf v_word_mode = WORD_VISU Then
        Close #fd
        sval = Word_Doc.FullName
        Call Word_Doc.Close(savechanges:=wdSaveChanges)
        Word_Obj.Documents.Open FileName:=sval, ReadOnly:=True, passworddocument:=v_passwd
        GoTo lab_fin_visible
    ElseIf v_word_mode = WORD_MODIF Then
        Close #fd
        Word_Doc.Saved = True
        GoTo lab_fin_visible
    ElseIf v_word_mode = WORD_CREATE Then
        Close #fd
        GoTo lab_fin_create
    End If
    
lab_fin:
    If v_fin_mode <> WORD_FIN_RIEN Then
        Call Word_Doc.Close(savechanges:=wdDoNotSaveChanges)
    End If
    If v_fin_mode = WORD_FIN_RAZOBJ Then
'
    End If
    Call FICH_EffacerFichier(v_nomdata, False)
    Word_Fusionner = P_OK
    Exit Function

lab_err_paste:
    MsgBox "Erreur Paste" & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    Resume Next

lab_fin_create:
    Call FICH_EffacerFichier(v_nomdata, False)
    If v_fin_mode <> WORD_FIN_RIEN Then
        Call Word_Doc.Close(savechanges:=wdSaveChanges)
    End If
    If v_fin_mode = WORD_FIN_RAZOBJ Then
        '
    End If
    Word_Obj.Visible = False
    Word_Fusionner = P_OK
    Exit Function

lab_fin_visible:
    Call FICH_EffacerFichier(v_nomdata, False)
    Word_Obj.Visible = True
    Word_EstActif = False
    Word_Fusionner = P_OK
    Exit Function
    
lab_fin_err2:
    Close #fd
    Call FICH_EffacerFichier(v_nomdata, False)
lab_fin_err1:
    Call Word_Doc.Close(savechanges:=wdDoNotSaveChanges)
    Word_Obj.Visible = False
    Word_Fusionner = P_ERREUR
    Exit Function

err_sav_dest:
    MsgBox "Impossible de sauvegarder le fichier dans " & v_nomdest & vbCr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    GoTo lab_fin_err1

err_word1:
    MsgBox "Erreur word " & Err.Number & " " & Err.Description, vbInformation + vbOKOnly, "Fusion"
'Resume Next
    GoTo lab_fin_err1
    
err_open_fus:
    MsgBox "Impossible d'ouvrir le fichier de données " & v_nomdata & vbCr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    GoTo lab_fin_err1
    
err_word2:
    MsgBox "Erreur word " & Err.Number & " " & Err.Description, vbInformation + vbOKOnly, "Fusion"
'Resume Next
    GoTo lab_fin_err2
    
lab_fin_err:
    Call MsgBox("Erreur détectée au cours de la fusion", vbOKOnly, "Fusion")
    GoTo lab_fin_err2

err_fax:
    MsgBox "Impossible d'effectuer l'envoi par fax.", vbInformation + vbOKOnly, "Fusion"
    GoTo lab_fin_err2
    
End Function

Public Sub Word_Imprimer(ByVal v_nomdoc As String, _
                         ByVal v_passwd As String, _
                         ByVal v_nbex As Integer, _
                         ByVal v_deb_mode As Integer)
        
    If v_deb_mode = WORD_DEB_CROBJ Then
        If Word_Init() = P_ERREUR Then
            Exit Sub
        End If
    End If
    
    If v_deb_mode <> WORD_DEB_RIEN Then
        If Word_OuvrirDoc(v_nomdoc, False, v_passwd, Word_Doc) = P_ERREUR Then
            Exit Sub
        End If
    End If
    
    Word_Obj.ActivePrinter = Printer.DeviceName
    Word_Obj.ActiveDocument.PrintOut Background:=False, Copies:=v_nbex
    
    Call Word_Doc.Close(savechanges:=wdDoNotSaveChanges)
    
End Sub

Public Function Word_Init()

    If Word_EstActif Then
        On Error GoTo lab_plus_actif
        If Word_Obj.Visible Then
            '''
        End If
        On Error GoTo 0
    End If
    
    If Not Word_EstActif Then
        Word_EstActif = True
        On Error GoTo err_create_obj
        Set Word_Obj = CreateObject("word.application")
        On Error GoTo 0
    End If
    
    Word_Init = P_OK
    Exit Function

err_create_obj:
    MsgBox "Impossible de créer l'objet WORD." & vbCrLf & "Err:" & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, ""
    Word_Init = P_ERREUR
    Exit Function

lab_plus_actif:
    Word_EstActif = False
    Resume Next
    
End Function

Public Function Word_OuvrirDoc(ByVal v_nomdoc As String, _
                               ByVal v_readonly As Boolean, _
                               ByVal v_passwd As String, _
                               ByRef r_doc As Document) As Integer

    On Error GoTo err_open_ficr
    Set r_doc = Word_Obj.Documents.Open(FileName:=v_nomdoc, _
                                        ReadOnly:=v_readonly, _
                                        passworddocument:=v_passwd, _
                                        addtorecentfiles:=False)
    On Error GoTo 0
    
    Word_OuvrirDoc = P_OK
    Exit Function
    
err_open_ficr:
    MsgBox "Impossible d'ouvrir le fichier " & v_nomdoc & vbCr & vbLf & Err.Description, vbCritical + vbOKOnly, "Fusion"
    Word_OuvrirDoc = P_ERREUR
    Exit Function
    
End Function

Public Sub Word_Quitter(ByVal v_mode As Integer)

    Word_EstActif = False
    
    On Error GoTo err_quit
    
    If v_mode = WORD_FIN_FERMDOC Then
        Call Word_Doc.Close(savechanges:=wdDoNotSaveChanges)
    End If
    Word_Obj.Application.Quit
    Set Word_Obj = Nothing
    
    On Error GoTo 0
    Exit Sub

err_quit:
    Exit Sub
    
End Sub

Private Function w_add_bookmark(ByVal v_nom As String, _
                                ByVal v_arange As range) As Integer

    On Error GoTo err_add_book
    Word_Doc.Bookmarks.Add Name:=v_nom, range:=v_arange
    On Error GoTo 0
    w_add_bookmark = P_OK
    Exit Function
    
err_add_book:
    On Error GoTo 0
    Call MsgBox("Erreur w_add_bookmark " & v_nom & vbCr & vbLf & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "")
    w_add_bookmark = P_ERREUR
    Exit Function
    
End Function

Private Function w_ajouter_bkglob(ByVal v_nombk As String, _
                                  ByRef r_inheadfoot As Boolean) As Boolean

    Dim i As Integer, j As Integer, ntab As Integer
    
    Word_Doc.Bookmarks(v_nombk).Select
    ' Le bookmark n'est pas dans un tableau
    If Not Word_Obj.Selection.Information(wdWithInTable) Then
        w_ajouter_bkglob = False
        Exit Function
    End If
    
    r_inheadfoot = Word_Obj.Selection.Information(wdInHeaderFooter)
    Word_Obj.Selection.Tables(1).Select
    For j = 1 To Word_Obj.Selection.Bookmarks.Count
        If w_est_champ_tableau(Word_Obj.Selection.Bookmarks(j).Name) Then
            If left$(Word_Obj.Selection.Bookmarks(j).Name, 3) <> left$(v_nombk, 3) Then
                w_ajouter_bkglob = False
                Exit Function
            End If
        End If
    Next j
    
    w_ajouter_bkglob = True
    Exit Function
                
End Function

Private Sub w_bk_dans_tableau(ByVal v_nombk As String, _
                              ByRef r_dochf As Variant, _
                              ByRef r_itab As Integer, _
                              ByRef r_lig As Integer)

    Dim i As Integer, j As Integer, itab As Integer
    Dim atable As Table
    
    r_lig = -1
    
    Word_Doc.Bookmarks(v_nombk).Select
    If Not Word_Obj.Selection.Information(wdWithInTable) Then
        Exit Sub
    End If
    
    ' On cherche dans le corps du document
    For itab = 1 To Word_Doc.Tables.Count
        Word_Doc.Tables(itab).Select
        For j = 1 To Word_Obj.Selection.Bookmarks.Count
            If Word_Obj.Selection.Bookmarks(j).Name = v_nombk Then
                Set r_dochf = Word_Doc
                r_itab = itab
                Word_Doc.Bookmarks(v_nombk).Select
                r_lig = Word_Obj.Selection.Information(wdStartOfRangeRowNumber)
                Exit Sub
            End If
        Next j
    Next itab
    
    ' On cherche dans entete et pied
    Word_Doc.Bookmarks(v_nombk).Select
    If Not Word_Obj.Selection.Information(wdInHeaderFooter) Then
        Exit Sub
    End If
    
    ' Entete
    For i = 1 To Word_Doc.Sections(1).Headers.Count
        Word_Doc.Sections(1).Headers(i).range.Select
        For itab = 1 To Word_Obj.Selection.Tables.Count
            Set atable = Word_Obj.Selection.Tables(itab)
            atable.Select
            For j = 1 To Word_Obj.Selection.Bookmarks.Count
                If Word_Obj.Selection.Bookmarks(j).Name = v_nombk Then
                    r_dochf = Word_Doc.Sections(1).Headers(i)
                    r_itab = itab
                    Word_Obj.Selection.Bookmarks(j).Select
                    r_lig = Word_Obj.Selection.Information(wdStartOfRangeRowNumber)
                    Exit Sub
                End If
            Next j
            Word_Doc.Sections(1).Headers(i).range.Select
        Next itab
    Next i
    
    ' Pied de page
    For i = 1 To Word_Doc.Sections(1).Footers.Count
        Word_Doc.Sections(1).Footers(i).range.Select
        For itab = 1 To Word_Obj.Selection.Tables.Count
            Set atable = Word_Obj.Selection.Tables(itab)
            atable.Select
            For j = 1 To Word_Obj.Selection.Bookmarks.Count
                If Word_Obj.Selection.Bookmarks(j).Name = v_nombk Then
                    r_dochf = Word_Doc.Sections(1).Footers(i)
                    r_itab = itab
                    Word_Obj.Selection.Bookmarks(j).Select
                    r_lig = Word_Obj.Selection.Information(wdStartOfRangeRowNumber)
                    Exit Sub
                End If
            Next j
            Word_Doc.Sections(1).Footers(i).range.Select
        Next itab
    Next i
    
End Sub

Private Sub w_conv_champ_en_signet(ByVal v_range As range)

    Dim slibchp As String, champ As String, nombk As String
    Dim trouve As Boolean
    Dim nbfields As Integer, ifield As Integer, ibk As Integer, isignet As Integer
    Dim i As Integer, j As Integer, isig As Integer, pos As Integer
    Dim arange As range
    
    nbfields = v_range.Fields.Count
    ifield = 1
    For i = 1 To nbfields
        champ = v_range.Fields(ifield).code
        If InStr(champ, "CHAMPFUSION") > 0 Then
            slibchp = "CHAMPFUSION"
        ElseIf InStr(champ, "MERGEFIELD") > 0 Then
            slibchp = "MERGEFIELD"
        Else
            ifield = ifield + 1
            GoTo lab_chpe_suiv
        End If
        champ = Mid$(champ, Len(slibchp) + 3, Len(champ) - Len(slibchp) - 3)
        ' Champ tableau sous la forme Txx_Nomtableau_NomChamp
        If w_est_champ_tableau(champ) = P_OUI Then
            ' Supprime Txx_ du champ
            'champ = Right(champ, Len(champ) - 4)
        End If
        trouve = False
        For isig = 1 To Word_nbsignet
            If Word_tblsignet(isig).nom = champ Then
                ibk = Word_tblsignet(isig).indice
                trouve = True
            ElseIf trouve Then
                Exit For
            End If
        Next isig
        Word_nbsignet = Word_nbsignet + 1
        ReDim Preserve Word_tblsignet(1 To Word_nbsignet) As WORD_SSIGNET
        If Not trouve Then
            ibk = 1
            pos = Word_nbsignet
        Else
            ibk = ibk + 1
            pos = isig
            For j = Word_nbsignet To pos + 1 Step -1
                Word_tblsignet(j) = Word_tblsignet(j - 1)
            Next j
        End If
        Word_tblsignet(pos).nom = champ
        Word_tblsignet(pos).indice = ibk
        nombk = champ & "_" & ibk
        v_range.Fields(ifield).Select
        Set arange = Word_Obj.Selection.range
        If w_add_bookmark(nombk, arange) = P_ERREUR Then
            Exit Sub
        End If
        Word_Obj.Selection.range.InsertBefore champ
        v_range.Fields(ifield).Delete
lab_chpe_suiv:
    Next i
    
End Sub

Private Sub w_copier_corps(ByRef v_docsrc As Document, _
                           ByRef v_docdest As Document)

    Dim nomdoc_src As String
    Dim nsect As Integer, nbsect As Integer
    Dim arange As range
    Dim head_foot As HeaderFooter
    Dim doc_src As Document
    
    nomdoc_src = p_chemin_appli + "\tmp\" + p_CodeUtil + ".doc"
    nbsect = v_docsrc.Sections.Count
    nsect = 1
    If nbsect > 1 Then
        Call v_docsrc.SaveAs(FileName:=nomdoc_src, Password:="")
    Else
        Set doc_src = v_docsrc
    End If
lab_deb_sect:
    If nbsect > 1 Then
        Set doc_src = Word_Obj.Documents.Open(nomdoc_src)
    End If
    ' mettre mode page si ce n'est pas deja le cas
    If doc_src.ActiveWindow.ActivePane.View.type = wdNormalView Or _
       doc_src.ActiveWindow.ActivePane.View.type = wdOutlineView Or _
       doc_src.ActiveWindow.ActivePane.View.type = wdMasterView Then
            doc_src.ActiveWindow.ActivePane.View.type = wdPageView
    End If
        
    If nbsect > 1 Then
        If nsect > 1 Then
            Set arange = doc_src.Sections(1).range
            If nsect > 2 Then
                arange.MoveEnd Unit:=wdSection, Count:=nsect - 2
            End If
            arange.Select
            Word_Obj.Selection.Delete
        End If
        If nsect < nbsect Then
            Set arange = doc_src.Sections(1).range
            arange.Collapse wdCollapseEnd
            arange.MoveEnd Unit:=wdCharacter, Count:=-1
            arange.Select
            Word_Obj.Selection.MoveEnd Unit:=wdStory, Count:=1
            Word_Obj.Selection.Delete
        End If
    End If
    
    On Error Resume Next
    For Each head_foot In doc_src.Sections(nsect).Headers
        head_foot.range.Delete
    Next head_foot
    For Each head_foot In doc_src.Sections(nsect).Footers
        head_foot.range.Delete
    Next head_foot
    
    ' Sélectionne tout le document à recopier
    On Error GoTo err_word
    doc_src.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    doc_src.StoryRanges(wdMainTextStory).Copy
    
    ' Se positionne à la fin du document source et recopie
    Set arange = v_docdest.Content
    arange.Collapse wdCollapseEnd
    If nsect > 1 Then
        arange.InsertAfter vbFormFeed
        arange.Collapse wdCollapseEnd
    End If
    arange.Paste
    Call doc_src.Close(savechanges:=wdDoNotSaveChanges)
    If nsect < nbsect Then
        nsect = nsect + 1
        GoTo lab_deb_sect
    End If
    On Error GoTo 0
    
    If nbsect > 1 Then
        Call FICH_EffacerFichier(nomdoc_src, False)
    End If
    Exit Sub
    
err_word:
    MsgBox "Erreur word " & Err.Number & " " & Err.Description, vbInformation + vbOKOnly, "CopierModele"
    Exit Sub
    
End Sub

Private Function w_est_champ_tableau(ByVal v_champ As String) As Integer

    If Len(v_champ) > 5 Then
        If left(v_champ, 1) = "T" And Mid(v_champ, 4, 1) = "_" Then
            If InStr(left$(v_champ, 5), "_") > 0 Then
                w_est_champ_tableau = P_OUI
                Exit Function
            End If
        End If
    End If
    w_est_champ_tableau = P_NON

End Function

Private Function w_est_champ_tableau_global(ByVal v_champ As String) As Integer

    Dim s As String
    Dim pos As Integer
    
    If Len(v_champ) > 5 Then
        If left(v_champ, 1) = "T" And Mid(v_champ, 4, 1) = "_" Then
            s = Mid$(v_champ, 5)
            pos = InStr(s, "_")
            If pos = 0 Then
                w_est_champ_tableau_global = P_OUI
                Exit Function
            End If
            s = Mid$(s, pos + 1)
            If s = "" Then
                w_est_champ_tableau_global = P_OUI
                Exit Function
            End If
            If IsNumeric(s) Then
                w_est_champ_tableau_global = P_OUI
                Exit Function
            End If
        End If
    End If
    w_est_champ_tableau_global = P_NON
    
End Function

Private Function w_get_txtp(ByVal v_deb As Long, _
                            ByVal v_fin As Long, _
                            ByRef v_buf As Variant) As Integer

    Dim arange As range
    
    On Error GoTo err_get_txt
    Set arange = Word_Doc.range(v_deb, v_fin)
    v_buf = arange.Text
    On Error GoTo 0
    w_get_txtp = P_OK
    Exit Function
    
err_get_txt:
    On Error GoTo 0
    Call MsgBox("Erreur w_get_txtp " & v_deb & " " & v_fin & vbCr & vbLf & Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "")
    w_get_txtp = P_ERREUR
    
End Function

Private Sub w_init_tblsignet()

    Dim nom As String, s As String
    Dim encore As Boolean, trouve As Boolean
    Dim i As Integer, j As Integer, pos As Integer, ind  As Integer
    Dim un_signet As WORD_SSIGNET
Dim v As Variant

    Word_nbsignet = Word_Doc.Bookmarks.Count
    
    If Word_nbsignet = 0 Then Exit Sub
    
    ReDim Word_tblsignet(1 To Word_Doc.Bookmarks.Count) As WORD_SSIGNET
    For i = 1 To Word_Doc.Bookmarks.Count
        nom = Word_Doc.Bookmarks(i).Name
        pos = InStrRev(nom, "_")
        trouve = False
        If pos > 0 Then
            If IsNumeric(Mid$(nom, pos + 1)) Then
                trouve = True
                Word_tblsignet(i).nom = left$(nom, pos - 1)
                Word_tblsignet(i).indice = Mid$(nom, pos + 1)
            End If
        End If
        If Not trouve Then
            Word_tblsignet(i).nom = nom
            Word_tblsignet(i).indice = 0
        End If
    Next i
    ' Tri
    Do
        encore = False
        For i = 1 To UBound(Word_tblsignet) - 1
            For j = i + 1 To UBound(Word_tblsignet)
                If Word_tblsignet(i).nom = Word_tblsignet(j).nom Then
                    If j > i + 1 Then
                        If Word_tblsignet(i + 1).nom <> Word_tblsignet(j).nom Then
                            un_signet = Word_tblsignet(i + 1)
                            Word_tblsignet(j) = Word_tblsignet(i + 1)
                            Word_tblsignet(i + 1) = un_signet
                            encore = True
                        End If
                    ElseIf Word_tblsignet(i).indice > Word_tblsignet(j).indice Then
                        un_signet = Word_tblsignet(i)
                        Word_tblsignet(j) = Word_tblsignet(i)
                        Word_tblsignet(i) = un_signet
                        encore = True
                    End If
                End If
            Next j
        Next i
    Loop Until encore = False
'v = "Après tri" & vbCrLf
'For i = 1 To UBound(word_tblsignet)
'v = v & word_tblsignet(i).nom & " " & word_tblsignet(i).indice & vbCrLf
'Next i
'MsgBox v

    ' Renommage
    nom = ""
    i = 1
    While i <= UBound(Word_tblsignet)
        If Word_tblsignet(i).nom <> nom Then
            nom = Word_tblsignet(i).nom
            ind = 1
        End If
        j = i
        While j <= UBound(Word_tblsignet)
            If Word_tblsignet(j).nom = nom Then
                i = i + 1
                If Word_tblsignet(j).indice <> ind Then
                    Call w_renommer_signet(Word_tblsignet(j).nom, Word_tblsignet(j).indice, ind)
                    Word_tblsignet(j).indice = ind
                End If
            End If
            ind = ind + 1
            j = j + 1
        Wend
    Wend
'v = "Après renomme" & vbCrLf
'For i = 1 To UBound(word_tblsignet)
'v = v & word_tblsignet(i).nom & " " & word_tblsignet(i).indice & vbCrLf
'Next i
'MsgBox v

End Sub

Private Function w_lire_fich(ByVal v_fd As Integer, _
                             ByRef a_ligne As Variant) As Integer

    On Error GoTo fin_fichier
    Line Input #v_fd, a_ligne
    On Error GoTo 0
    w_lire_fich = P_OUI
    Exit Function

fin_fichier:
    On Error GoTo 0
    w_lire_fich = P_NON

End Function

Private Function w_put_txtbk(ByVal v_str As String, _
                             ByVal v_nombk As String) As Integer

    Dim str As String, sparam As String, nomimg As String
    Dim n As Integer, n2 As Integer, i As Integer, j As Integer
    Dim arange As range
    Dim shp As Shape
'    Dim ctrl As Object, obj_shape As Object
    
    If left$(v_str, 1) = "ê" Then
        sparam = STR_GetChamp(Mid$(v_str, 2), "ê", 0)
        str = STR_GetChamp(Mid$(v_str, 2), "ê", 1)
    Else
        sparam = ""
        str = v_str
    End If
    On Error GoTo err_range
    Set arange = Word_Doc.Bookmarks(v_nombk).range
    If InStr(v_nombk, "HyperLien") > 0 Then
'        If arange.Hyperlinks.Count > 0 Then arange.Hyperlinks(1).Delete
        arange.Text = " "
        If Len(str) > 1 Then
            arange.Text = "Accès au document"
            On Error GoTo err_add_hyp
            Call Word_Doc.Hyperlinks.Add(Anchor:=arange, Address:=str, SubAddress:="")
            On Error GoTo 0
        End If
' Cas de gestion d'un label
'        arange.text = " "
'        Set obj_shape = g_doc.InlineShapes.AddOLEControl("KaliDocCtrl.KalidocCmd", _
'                                                         arange)
'        Set ctrl = obj_shape.OLEFormat.object
'        ctrl.hNotify = Documentation.txtWord.hWnd
    Else
        On Error GoTo err_put_txt
        arange.Text = str
        On Error GoTo 0
    End If
    If sparam <> "" Then
        n = STR_GetNbchamp(sparam, "|")
        For i = 0 To n - 1
            str = STR_GetChamp(sparam, "|", i)
            If left$(str, 4) = "lien" Then
                str = Mid$(str, 6)
                On Error GoTo err_add_hyp
                Call Word_Doc.Hyperlinks.Add(Anchor:=arange, Address:=str, SubAddress:="")
                On Error GoTo 0
            ElseIf left$(str, 3) = "img" Then
                str = Mid$(str, 5)
                n2 = STR_GetNbchamp(str, "$")
                For j = 0 To n2 - 1
                    On Error Resume Next
                    nomimg = STR_GetChamp(str, "$", j)
                    Call Word_Doc.InlineShapes.AddPicture(FileName:=nomimg, LinkToFile:=False, SaveWithDocument:=True, range:=arange)
                    On Error GoTo 0
                Next j
            End If
        Next i
    End If
    
    ' On rajoute le bookmark si pas tableau
    If g_garder_bookmark Then
        w_put_txtbk = w_add_bookmark(v_nombk, arange)
        Exit Function
    End If
    
    w_put_txtbk = P_OK
    Exit Function
    
err_range:
    On Error GoTo 0
    Call MsgBox("Erreur w_put_txtbk : Erreur range " & v_nombk & vbCr & vbLf & Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "")
    w_put_txtbk = P_ERREUR
    Exit Function
    
err_range_hyp:
    On Error GoTo 0
    Call MsgBox("Erreur w_put_txtbk : Erreur range hyperline" & vbCr & vbLf & Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "")
    w_put_txtbk = P_ERREUR
    Exit Function
    
err_put_txt:
    On Error GoTo 0
    Call MsgBox("Erreur w_put_txtbk " & v_str & " " & v_nombk & vbCr & vbLf & Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "")
    w_put_txtbk = P_ERREUR
    Exit Function
    
err_add_hyp:
    On Error GoTo 0
    Call MsgBox("Erreur add hyperlink " & v_str & " " & v_nombk & vbCr & vbLf & Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "")
    w_put_txtbk = P_ERREUR
    
End Function

Private Function w_range(ByVal v_deb As Long, _
                         ByVal v_fin As Long, _
                         ByRef r_range As range)

    On Error GoTo err_range
    Set r_range = Word_Doc.range(v_deb, v_fin)
    On Error GoTo 0
    w_range = P_OK
    Exit Function
    
err_range:
    On Error GoTo 0
    Call MsgBox("Erreur w_range " & v_deb & " " & v_fin & vbCr & vbLf & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "")
    w_range = P_ERREUR
    Exit Function
    
End Function

Private Function w_rangeb(ByVal v_deb As Long, _
                         ByVal v_fin As Long, _
                         ByRef r_range As range)

    On Error GoTo err_range
    Set r_range = Word_Doc.range(v_deb, v_fin)
    On Error GoTo 0
    w_rangeb = P_OK
    Exit Function
    
err_range:
    On Error GoTo 0
    w_rangeb = P_ERREUR
    Exit Function
    
End Function

Private Sub w_renommer_signet(ByVal v_nom As String, _
                              ByVal v_old_ind As Integer, _
                              ByVal v_new_ind As Integer)

    Dim nom As String
    Dim arange As range
    
    nom = v_nom
    If v_old_ind > 0 Then
        nom = nom & "_" & v_old_ind
    End If
    Set arange = Word_Doc.Bookmarks(nom).range
    Word_Doc.Bookmarks(nom).Delete
    nom = v_nom & "_" & v_new_ind
    Call w_add_bookmark(nom, arange)
    
End Sub

Private Sub w_suppr_bk_doublon(ByRef v_doc1 As Document, _
                               ByRef v_doc2 As Document)
                                 
    Dim i As Integer
    
    ' Suppression des bookmarks de doc_modele qui sont déjà dans doc
    For i = 1 To v_doc2.Bookmarks.Count
        If v_doc1.Bookmarks.Exists(v_doc2.Bookmarks(i).Name) Then
            v_doc1.Bookmarks(v_doc2.Bookmarks(i).Name).Delete
        End If
    Next i

End Sub

Private Function w_ya_bktbl_dans_sel(ByVal v_nombk_tbl As String) As Boolean

    Dim i As Integer, lenv As Integer
    
    lenv = Len(v_nombk_tbl)
    For i = 1 To Word_Obj.Selection.Bookmarks.Count
        If Len(Word_Obj.Selection.Bookmarks(i).Name) > lenv Then
            If left$(Word_Obj.Selection.Bookmarks(i).Name, lenv) = v_nombk_tbl Then
                w_ya_bktbl_dans_sel = True
                Exit Function
            End If
        End If
    Next i
    
    w_ya_bktbl_dans_sel = False
    
End Function


