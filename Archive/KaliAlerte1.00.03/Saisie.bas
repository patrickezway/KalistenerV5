Attribute VB_Name = "MSaisie"
Option Explicit

'Type de saisie gérée
Public Const SAIS_TYP_TOUT_CAR = 0
Public Const SAIS_TYP_JOUR_SEMAINE = 1
Public Const SAIS_TYP_HEURE = 2
Public Const SAIS_TYP_DATE = 3
Public Const SAIS_TYP_ENTIER = 4
Public Const SAIS_TYP_LETTRE = 5
Public Const SAIS_TYP_LETTRE_PONCT = 6
Public Const SAIS_TYP_ENTIER_NEG = 7
Public Const SAIS_TYP_DATNAIS = 8
Public Const SAIS_TYP_CAR_PARTICULIER = 9
Public Const SAIS_TYP_PRIX = 10
Public Const SAIS_TYP_PERIODE = 11
Public Const SAIS_TYP_CODE = 12

'Conversions possibles
Public Const SAIS_CONV_MINUSCULE = 1
Public Const SAIS_CONV_MAJUSCULE = 2
Public Const SAIS_CONV_SECRET = 3

'Retour possible
Public Const SAIS_RET_NOMODIF = -1
Public Const SAIS_RET_MODIF = 0

Public Type SAIS_SPRMFRM
    titre As String
    nomhelp As String
    x As Integer
    y As Integer
    max_nbcar_visible As Integer    '0 => la zone texte = à la taille du texte le plus grand
    reste_chargée As Boolean
End Type

'Structure permettant l'appel à la form FSAIS_
Public Type SAIS_SCHAMP
    libelle As String
    len As Integer  'Longueur du texte à saisir
    type As Integer
    chaine_type As String
    facu As Boolean 'False => si OK ce champ doit être rempli
    conversion As Integer
    sval As String  'Contenu de la zone texte au retour
    validationdirecte As Boolean
End Type

Public Type SAIS_SBOUTON
    libelle As String
    image As String
    raccourci_alt As Integer
    raccourci_touche As Integer
    largeur As Long
End Type

Public Type SAIS_SSAISIE
    prmfrm As SAIS_SPRMFRM
    champs() As SAIS_SCHAMP
    boutons() As SAIS_SBOUTON
    retour As Integer
End Type

Public SAIS_Saisie As SAIS_SSAISIE

Private Function ctrl_date(ByRef vr_str As String) As Boolean

    Dim stmp As String, siècle_en_cours As String, sdater As String, s As String
    Dim jj As Integer, mm As Integer, aa As Integer, pos As Integer
    Dim nbj As Integer

    If left$(vr_str, 1) = "j" Or left$(vr_str, 1) = "J" Then
        If Len(vr_str) = 1 Then
            nbj = 0
        ElseIf Mid$(vr_str, 2, 1) = "-" Then
            If STR_EstEntierPos(Mid$(vr_str, 3)) Then
                nbj = -(CInt(Mid$(vr_str, 3)))
            Else
                ctrl_date = False
                Exit Function
            End If
        ElseIf Mid$(vr_str, 2, 1) = "+" Then
            If STR_EstEntierPos(Mid$(vr_str, 3)) Then
                nbj = CInt(Mid$(vr_str, 3))
            Else
                ctrl_date = False
                Exit Function
            End If
        End If
        vr_str = Format(Date + nbj, "dd/mm/yyyy")
        ctrl_date = True
        Exit Function
    End If

    If left$(vr_str, 1) = "m" Or left$(vr_str, 1) = "M" Then
        If Len(vr_str) = 1 Then
            s = Date
            s = "01/" & Mid$(s, 4)
        ElseIf Mid$(vr_str, 2, 1) = "-" Then
            If STR_EstEntierPos(Mid$(vr_str, 3)) Then
                ' A FAIRE ...
                ctrl_date = False
                Exit Function
            Else
                ctrl_date = False
                Exit Function
            End If
        ElseIf Mid$(vr_str, 2, 1) = "+" Then
            If STR_EstEntierPos(Mid$(vr_str, 3)) Then
                ' A FAIRE ...
                ctrl_date = False
                Exit Function
            Else
                ctrl_date = False
                Exit Function
            End If
        Else
            ctrl_date = False
            Exit Function
        End If
        vr_str = Format(CDate(s), "dd/mm/yyyy")
        ctrl_date = True
        Exit Function
    End If

    stmp = Format(Date, "dd/mm/yyyy")
    siècle_en_cours = Mid(stmp, 7, 2)
    
    If STR_EstEntierPos(vr_str) Then
        If Len(vr_str) = 6 Then
            sdater = left$(vr_str, 2) + "/" + Mid$(vr_str, 3, 2) + "/" + siècle_en_cours + Right$(vr_str, 2)
        ElseIf Len(vr_str) = 8 Then
            sdater = left$(vr_str, 2) + "/" + Mid$(vr_str, 3, 2) + "/" + Right$(vr_str, 4)
        Else
            sdater = ""
        End If
    Else
        If Not IsDate(vr_str) Then
            sdater = ""
            GoTo lab_fin_date
        End If
        pos = InStr(vr_str, "/")
        If pos <= 0 Then
            sdater = ""
            GoTo lab_fin_date
        End If
        jj = CInt(Mid$(vr_str, 1, pos - 1))
        vr_str = Mid$(vr_str, pos + 1)
        pos = InStr(vr_str, "/")
        If pos <= 0 Then
            sdater = ""
            GoTo lab_fin_date
        End If
        mm = CInt(Mid$(vr_str, 1, pos - 1))
        pos = InStr(vr_str, "/")
        If pos <= 0 Then
            sdater = ""
            GoTo lab_fin_date
        End If
        aa = CInt(Mid$(vr_str, pos + 1))
        sdater = Format(jj, "00") + "/" + Format(mm, "00") + "/"
        If aa < 100 Then
            sdater = sdater + siècle_en_cours + Format(aa, "00")
        Else
            sdater = sdater + Format(aa, "0000")
        End If
    End If

lab_fin_date:
    If sdater = "" Or Not IsDate(sdater) Then
        MsgBox "La saisie ne correspond pas à une date.", vbOKOnly + vbExclamation, "SAIS_ Erronnée"
        ctrl_date = False
        Exit Function
    End If

    vr_str = sdater
    ctrl_date = True

End Function

Private Function ctrl_entier_pos(ByVal v_str As String) As String

    If Not STR_EstEntierPos(v_str) Then
        MsgBox "La saisie ne correspond pas à un nombre positif.", vbOKOnly + vbExclamation, "SAIS_ Erronnée"
        ctrl_entier_pos = False
        Exit Function
    End If
    ctrl_entier_pos = True

End Function

Private Function ctrl_heure(ByRef vr_str As String) As Boolean

    Dim hh As Integer, mm As Integer, pos As Integer
    Dim s As String

    If vr_str = "h" Or vr_str = "H" Then
        vr_str = Format(Time, "hh:mm")
        ctrl_heure = True
        Exit Function
    End If

    hh = -1
    If STR_EstEntierPos(vr_str) And Len(vr_str) <= 4 Then
        If Len(vr_str) <= 2 Then
            hh = val(vr_str)
            mm = 0
        Else
            hh = val(vr_str) / 100
            mm = val(vr_str) Mod 100
        End If
    Else
        pos = InStr(vr_str, ":")
        If pos > 0 Then
            s = Mid$(vr_str, pos + 1)
            If InStr(s, ":") <= 0 Then
                hh = val(left$(vr_str, pos - 1))
                mm = val(Mid$(vr_str, pos + 1))
            End If
        End If
    End If
    If hh > 24 Then
        hh = -1
    ElseIf mm > 59 Then
        hh = -1
    ElseIf (hh * 100) + mm > 2400 Then
        hh = -1
    End If
    If hh >= 0 Then
        s = ""
        If hh < 10 Then
            s = "0"
        End If
        s = s + Trim$(str(hh)) + ":"
        If mm < 10 Then
            s = s + "0"
        End If
        vr_str = s + Trim$(str(mm))
        ctrl_heure = True
        Exit Function
    End If

    MsgBox "La saisie ne correspond pas à une heure.", vbOKOnly + vbExclamation, "SAIS_ Erronnée"
    ctrl_heure = False

End Function

Public Sub SAIS_AddBouton(ByVal v_libelle As String, _
                          ByVal v_image As String, _
                          ByVal v_rcalt As Integer, _
                          ByVal v_rctouche As Integer, _
                          ByVal v_largeur As Integer)

    Dim n As Integer

    n = -1
    On Error Resume Next
    n = UBound(SAIS_Saisie.boutons)
    On Error GoTo 0

    n = n + 1
    ReDim Preserve SAIS_Saisie.boutons(n)
    SAIS_Saisie.boutons(n).libelle = v_libelle
    SAIS_Saisie.boutons(n).image = v_image
    SAIS_Saisie.boutons(n).raccourci_alt = v_rcalt
    SAIS_Saisie.boutons(n).raccourci_touche = v_rctouche
    SAIS_Saisie.boutons(n).largeur = v_largeur

End Sub

Public Sub SAIS_AddChamp(ByVal v_libelle As String, _
                         ByVal v_len As Integer, _
                         ByVal v_type As Integer, _
                         ByVal v_facu As Boolean, _
                         Optional v_sval As Variant)

    Dim n As Integer

    n = -1
    On Error Resume Next
    n = UBound(SAIS_Saisie.champs)
    On Error GoTo 0

    n = n + 1
    ReDim Preserve SAIS_Saisie.champs(n)
    SAIS_Saisie.champs(n).libelle = v_libelle
    SAIS_Saisie.champs(n).len = v_len
    SAIS_Saisie.champs(n).type = v_type
    SAIS_Saisie.champs(n).facu = v_facu
    If Not IsMissing(v_sval) Then
        SAIS_Saisie.champs(n).sval = v_sval
    Else
        SAIS_Saisie.champs(n).sval = ""
    End If
    SAIS_Saisie.champs(n).validationdirecte = False

End Sub

Public Sub SAIS_AddChampComplet(ByVal v_libelle As String, _
                                ByVal v_len As Integer, _
                                ByVal v_type As Integer, _
                                ByVal v_str_type As String, _
                                ByVal v_facu As Boolean, _
                                ByVal v_conv As Integer, _
                                ByVal v_valid_direct As Boolean, _
                                Optional v_sval As Variant)

    Dim n As Integer

    n = -1
    On Error Resume Next
    n = UBound(SAIS_Saisie.champs)
    On Error GoTo 0

    n = n + 1
    ReDim Preserve SAIS_Saisie.champs(n)
    SAIS_Saisie.champs(n).libelle = v_libelle
    SAIS_Saisie.champs(n).len = v_len
    SAIS_Saisie.champs(n).type = v_type
    SAIS_Saisie.champs(n).chaine_type = v_str_type
    SAIS_Saisie.champs(n).facu = v_facu
    SAIS_Saisie.champs(n).conversion = v_conv
    SAIS_Saisie.champs(n).validationdirecte = v_valid_direct
    If Not IsMissing(v_sval) Then
        SAIS_Saisie.champs(n).sval = v_sval
    Else
        SAIS_Saisie.champs(n).sval = ""
    End If

End Sub

Public Function SAIS_CtrlChamp(ByRef vr_str As String, _
                               ByVal v_typchamp As Integer) As Boolean

    Dim s As String, s2 As String, sc As String
    Dim fok As Boolean
    Dim pos As Integer, n As Integer, i As Integer

    Select Case v_typchamp
    Case SAIS_TYP_JOUR_SEMAINE
        Select Case LCase(left$(vr_str, 1))
        Case "l"
            vr_str = "lundi"
            SAIS_CtrlChamp = True
        Case "ma"
            vr_str = "mardi"
            SAIS_CtrlChamp = True
        Case "me"
            vr_str = "mercredi"
            SAIS_CtrlChamp = True
        Case "j"
            vr_str = "jeudi"
            SAIS_CtrlChamp = True
        Case "v"
            vr_str = "vendredi"
            SAIS_CtrlChamp = True
        Case Else
            MsgBox "La saisie ne correspond pas à un jour de la semaine", vbOKOnly + vbExclamation, "Saisie Erronnée"
            SAIS_CtrlChamp = False
        End Select
        Exit Function
    Case SAIS_TYP_PERIODE
        vr_str = UCase(vr_str)
        s = Right$(vr_str, 1)
        If s = "J" Or s = "S" Or s = "M" Or s = "A" Then
            s2 = left$(vr_str, Len(vr_str) - 1)
            If STR_EstEntierPos(s2) Then
                n = CInt(s2)
                vr_str = n & s
                SAIS_CtrlChamp = True
                Exit Function
            End If
        End If
        MsgBox "La saisie ne correspond pas à une période : nombre suivi de J(ours)/S(emaines)/M(ois)/A(nnées).", _
                vbOKOnly + vbExclamation, "Saisie Erronnée"
        SAIS_CtrlChamp = False
        Exit Function
    Case SAIS_TYP_HEURE
        SAIS_CtrlChamp = ctrl_heure(vr_str)
        Exit Function
    Case SAIS_TYP_DATE
        SAIS_CtrlChamp = ctrl_date(vr_str)
        Exit Function
    Case SAIS_TYP_ENTIER_NEG
        If InStr(vr_str, "-") > 1 Then
            MsgBox "La saisie ne correspond pas à un nombre signé.", vbOKOnly + vbExclamation, "Saisie Erronnée"
            SAIS_CtrlChamp = False
            Exit Function
        End If
        If left$(vr_str, 1) = "-" And InStr(Mid$(vr_str, 2), "-") > 0 Then
            MsgBox "La saisie ne correspond pas à un nombre signé.", vbOKOnly + vbExclamation, "Saisie Erronnée"
            SAIS_CtrlChamp = False
            Exit Function
        End If
        SAIS_CtrlChamp = True
        Exit Function
    Case SAIS_TYP_ENTIER
        SAIS_CtrlChamp = ctrl_entier_pos(vr_str)
        Exit Function
    Case SAIS_TYP_DATNAIS
        SAIS_CtrlChamp = ctrl_ddn(vr_str)
        Exit Function
    Case SAIS_TYP_PRIX
        SAIS_CtrlChamp = ctrl_prix(vr_str)
        Exit Function
    Case SAIS_TYP_CODE
        For i = 1 To Len(vr_str)
            sc = Mid$(vr_str, i, 1)
            If sc >= "A" And sc <= "Z" Then
                fok = True
            ElseIf sc >= "a" And sc <= "z" Then
                fok = True
            ElseIf sc >= "0" And sc <= "9" Then
                fok = True
            ElseIf sc = "-" Or sc = "_" Or sc = "." Then
                fok = True
            Else
                fok = False
            End If
            If Not fok Then
                MsgBox "Seuls les chiffres, lettres (sauf les caractères accentués) et - _ . sont autorisés.", _
                        vbOKOnly + vbExclamation, "Saisie Erronnée"
                SAIS_CtrlChamp = False
                Exit Function
            End If
        Next i
        SAIS_CtrlChamp = True
        Exit Function
    Case Else
        SAIS_CtrlChamp = True
    End Select

End Function

Private Function ctrl_ddn(ByRef vr_str As String) As Boolean

    Dim stmp As String, siècle_en_cours As String, sdater As String
    Dim ddn As Date

    stmp = Format(Date, "dd/mm/yyyy")
    siècle_en_cours = Mid$(stmp, 7, 2)

    If STR_EstEntierPos(vr_str) Then
        If Len(vr_str) = 6 Then
            sdater = left$(vr_str, 2) + "/" + Mid$(vr_str, 3, 2) + "/" + siècle_en_cours + Right$(vr_str, 2)
        ElseIf Len(vr_str) = 8 Then
            sdater = left$(vr_str, 2) + "/" + Mid$(vr_str, 3, 2) + "/" + Right$(vr_str, 4)
        Else
            sdater = ""
        End If
    ElseIf STR_EstEntierPos(left$(vr_str, 2)) And STR_EstEntierPos(Mid$(vr_str, 4, 2)) _
       And STR_EstEntierPos(Mid$(vr_str, 7)) And Mid$(vr_str, 3, 1) = "/" And Mid$(vr_str, 6, 1) = "/" Then
        If Len(vr_str) = 8 Then
            sdater = left$(vr_str, 6) + siècle_en_cours + Right$(vr_str, 2)
        ElseIf Len(vr_str) = 10 Then
            sdater = vr_str
        Else
            sdater = ""
        End If
    End If

    If sdater = "" Or Not IsDate(sdater) Then
        MsgBox "La saisie ne correspond pas à une date.", vbOKOnly + vbExclamation, "SAIS_ Erronée"
        ctrl_ddn = False
        Exit Function
    End If

    ddn = CDate(sdater)
    If ddn > Date Then
        MsgBox "Ce malade n'est pas encore né.", vbOKOnly + vbExclamation, "SAIS_ Erronée"
        ctrl_ddn = False
        Exit Function
    End If

    vr_str = sdater
    ctrl_ddn = True

End Function

Private Function ctrl_prix(ByRef vr_str As String) As Boolean

    Dim prix As Double

    On Error GoTo err_prix
    prix = CDbl(vr_str)
    On Error GoTo 0
    vr_str = STR_Prix(vr_str)
    ctrl_prix = True
    Exit Function

err_prix:
    MsgBox "La saisie ne correspond pas à un prix.", vbOKOnly + vbExclamation, "SAIS_ Erronée"
    ctrl_prix = False

End Function

Public Sub SAIS_Init()

    SAIS_Saisie.prmfrm.titre = ""
    SAIS_Saisie.prmfrm.nomhelp = ""
    SAIS_Saisie.prmfrm.x = 0
    SAIS_Saisie.prmfrm.y = 0
    SAIS_Saisie.prmfrm.max_nbcar_visible = 50
    SAIS_Saisie.prmfrm.reste_chargée = False

    Erase SAIS_Saisie.champs()

    Erase SAIS_Saisie.boutons()

End Sub

Public Sub SAIS_InitPos(ByVal v_posx As Long, _
                        ByVal v_posy As Long)

    SAIS_Saisie.prmfrm.x = v_posx
    SAIS_Saisie.prmfrm.y = v_posy

End Sub

Public Sub SAIS_InitResteChargée(ByVal v_restec As Boolean)

    SAIS_Saisie.prmfrm.reste_chargée = v_restec

End Sub

Public Sub SAIS_InitTitreHelp(ByVal v_nomtitre As String, _
                              ByVal v_nomhelp As String)

    SAIS_Saisie.prmfrm.titre = v_nomtitre
    SAIS_Saisie.prmfrm.nomhelp = v_nomhelp

End Sub



