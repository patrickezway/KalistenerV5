Attribute VB_Name = "MChoixListe"
Option Explicit

Public Type CL_SPRMFRM
    titre As String
    nomhelp As String
    X As Integer
    y As Integer
    hauteur_max As Integer
    largeur_max As Integer
    labels() As String
    gerer_tous_rien As Boolean
    multi_select As Boolean
    renvoyer_courant As Boolean
    reste_cachée As Boolean
End Type

Public Type CL_SBOUTON
    libelle As String
    image As String
    raccourci_alt As Integer
    raccourci_touche As Integer
    largeur As Long
End Type

Public Type CL_SLIGNE
    texte As String
    num As Long
    tag As Variant
    selected As Boolean
    fmodif As Boolean
End Type

Public Type CL_SLISTE
    prmfrm As CL_SPRMFRM
    boutons() As CL_SBOUTON
    lignes() As CL_SLIGNE
    pointeur As Integer
    retour As Integer
End Type

Public CL_liste As CL_SLISTE

Public Sub CL_AddBouton(ByVal v_libelle As String, _
                        ByVal v_image As String, _
                        ByVal v_rcalt As Integer, _
                        ByVal v_rctouche As Integer, _
                        ByVal v_largeur As Integer)

    Dim n As Integer

    n = -1
    On Error Resume Next
    n = UBound(CL_liste.boutons)
    On Error GoTo 0

    n = n + 1
    ReDim Preserve CL_liste.boutons(n)
    CL_liste.boutons(n).libelle = v_libelle
    If v_image <> "" Then
        CL_liste.boutons(n).libelle = ""
        CL_liste.boutons(n).image = v_image
    End If
    CL_liste.boutons(n).raccourci_alt = v_rcalt
    CL_liste.boutons(n).raccourci_touche = v_rctouche
    CL_liste.boutons(n).largeur = v_largeur

End Sub

Public Sub CL_AddLabel(ByVal v_label As String)

    Dim n As Integer

    n = -1
    On Error Resume Next
    n = UBound(CL_liste.prmfrm.labels)
    On Error GoTo 0

    n = n + 1
    ReDim Preserve CL_liste.prmfrm.labels(n)
    CL_liste.prmfrm.labels(n) = v_label

End Sub

Public Sub CL_AddLigne(ByVal v_texte As String, _
                       ByVal v_num As Long, _
                       ByVal v_tag As Variant, _
                       ByVal v_selected As Boolean, _
                       Optional v_fmodif As Variant)

    Dim n As Integer

    n = -1
    On Error Resume Next
    n = UBound(CL_liste.lignes)
    On Error GoTo 0

    n = n + 1
    ReDim Preserve CL_liste.lignes(n)
    CL_liste.lignes(n).texte = v_texte
    CL_liste.lignes(n).num = v_num
    CL_liste.lignes(n).tag = v_tag
    CL_liste.lignes(n).selected = v_selected
    If Not IsMissing(v_fmodif) Then
        CL_liste.lignes(n).fmodif = v_fmodif
    Else
        CL_liste.lignes(n).fmodif = True
    End If

End Sub

Public Sub CL_AffiSelFirst()

    Dim pos As Integer, lig As Integer, lig2 As Integer, siz As Integer
    Dim cl_ligne As CL_SLIGNE

    siz = -1
    On Error Resume Next
    siz = UBound(CL_liste.lignes())

    pos = 0
    For lig = 0 To siz
        If CL_liste.lignes(lig).selected Then
            cl_ligne = CL_liste.lignes(lig)
            For lig2 = lig To pos + 1 Step -1
                CL_liste.lignes(lig2) = CL_liste.lignes(lig2 - 1)
            Next lig2
            CL_liste.lignes(pos) = cl_ligne
            pos = pos + 1
        End If
    Next lig

End Sub

Public Sub CL_Init()

    CL_liste.prmfrm.titre = ""
    CL_liste.prmfrm.nomhelp = ""
    CL_liste.prmfrm.X = 0
    CL_liste.prmfrm.y = 0
    CL_liste.prmfrm.largeur_max = 0
    CL_liste.prmfrm.hauteur_max = -10
    Erase CL_liste.prmfrm.labels
    CL_liste.prmfrm.gerer_tous_rien = False
    CL_liste.prmfrm.multi_select = False
    CL_liste.prmfrm.reste_cachée = False

    Erase CL_liste.boutons

    Erase CL_liste.lignes

    CL_liste.pointeur = 0

End Sub

Public Sub CL_InitGererTousRien(ByVal v_btr As Boolean)

    CL_liste.prmfrm.gerer_tous_rien = v_btr

End Sub

Public Sub CL_InitMultiSelect(ByVal v_multisel As Boolean, _
                              ByVal v_renvoyer_courant As Boolean)

    CL_liste.prmfrm.multi_select = v_multisel
    CL_liste.prmfrm.renvoyer_courant = v_renvoyer_courant

End Sub

Public Sub CL_InitPointeur(ByVal v_numlig As Integer)

    CL_liste.pointeur = v_numlig

End Sub

Public Sub CL_InitPos(ByVal v_posx As Long, _
                      ByVal v_posy As Long)

    CL_liste.prmfrm.X = v_posx
    CL_liste.prmfrm.y = v_posy

End Sub

Public Sub CL_InitResteCachée(ByVal v_restec As Boolean)

    CL_liste.prmfrm.reste_cachée = v_restec

End Sub

Public Sub CL_InitTaille(ByVal v_largmax As Long, _
                         ByVal v_hautmax As Long)

    CL_liste.prmfrm.largeur_max = v_largmax
    CL_liste.prmfrm.hauteur_max = v_hautmax

End Sub

Public Sub CL_InitTitreHelp(ByVal v_nomtitre As String, _
                            ByVal v_nomhelp As String)

    CL_liste.prmfrm.titre = v_nomtitre
    CL_liste.prmfrm.nomhelp = v_nomhelp

End Sub

Public Sub CL_Tri(ByVal v_first_lig As Integer)

    Dim i As Integer, pos As Integer
    Dim cl_ligne As CL_SLIGNE

    pos = v_first_lig
    While pos < UBound(CL_liste.lignes())
        For i = pos + 1 To UBound(CL_liste.lignes())
            If UCase(CL_liste.lignes(i).texte) < UCase(CL_liste.lignes(pos).texte) Then
                cl_ligne = CL_liste.lignes(i)
                CL_liste.lignes(i) = CL_liste.lignes(pos)
                CL_liste.lignes(pos) = cl_ligne
            End If
        Next i
        pos = pos + 1
    Wend

End Sub


