Attribute VB_Name = "MPrinter"
Option Explicit

Private Const PR_GAUCHE = 0
Private Const PR_DROITE = 1

Private g_printer As Printer

Private g_ya_cadre As Boolean
Private g_nopage As Integer
Private g_stitre As String
Private g_posx As Integer, g_posy As Integer
Private g_prhauteur As Integer

Private Type PRCASE
    texte As String
    largeur As Integer
    cadrage As Integer
End Type

Private g_prcase() As PRCASE

Private Sub imprimer_entete()

    Dim s As String
    Dim font_italic As Boolean
    Dim i As Integer, posx As Integer, posy As Integer
    Dim font_size As Integer

    g_posx = 100
    g_posy = 300

    font_size = Printer.FontSize
    font_italic = Printer.FontItalic

    Printer.FontSize = 14
    Printer.FontBold = True
    Printer.FontItalic = False
    Printer.CurrentY = g_posy
    Printer.CurrentX = g_posx
    Printer.Print g_stitre;
    Printer.FontSize = 10
    Printer.FontBold = False
    s = "Page " & g_nopage
    Printer.CurrentX = Printer.width - 1000 - Printer.TextWidth(s)
    Printer.Print s
    Printer.Print ""
    g_posx = 100
    g_posy = Printer.CurrentY

    posx = 100
    posy = g_posy
    For i = 0 To UBound(g_prcase)
        Printer.CurrentY = g_posy + 50
        Printer.CurrentX = g_posx + (g_prcase(i).largeur - Printer.TextWidth(g_prcase(i).texte)) / 2
        Printer.FontBold = True
        Printer.Print g_prcase(i).texte
        Printer.FontBold = False
        If g_ya_cadre Then
            Printer.FillStyle = 2
            Printer.Print ""
            Printer.FillStyle = vbFSTransparent
            Printer.Line (g_posx, g_posy)-(g_posx + g_prcase(i).largeur, g_posy + g_prhauteur), , B
        End If
        g_posx = g_posx + g_prcase(i).largeur
    Next i
    g_posx = posx
    g_posy = posy + g_prhauteur

    Printer.FontItalic = font_italic
    Printer.FontSize = font_size

End Sub

Public Function PR_ChoixImp(ByVal v_yafax As Boolean, _
                            ByVal v_yabidon As Boolean, _
                            ByRef r_yafax As Boolean, _
                            ByRef r_yabidon As Boolean) As Boolean

    Dim n As Integer, nopr As Integer, lig_crt As Integer
    Dim pr As Printer

    Set g_printer = Printer

    Call CL_Init
    Call CL_InitTitreHelp("Choix de l'imprimante", "")
    Call CL_InitTaille(0, -10)
    Call CL_AddBouton("", p_chemin_appli + "\btnok.gif", vbKeyO, vbKeyF1, 0)
    Call CL_AddBouton("", p_chemin_appli + "\btnporte.gif", 0, vbKeyEscape, 0)
    n = 0
    nopr = 0
    lig_crt = 0
    For Each pr In Printers
        If pr.DeviceName <> "Rendering Subsystem" And _
           (pr.DeviceName <> "Microsoft Fax" Or Not v_yafax) Then
            Call CL_AddLigne(pr.DeviceName, nopr, "", False)
            If pr.DeviceName = Printer.DeviceName Then
                lig_crt = n
            End If
            n = n + 1
        End If
        nopr = nopr + 1
    Next pr
    If v_yafax Then
        Call CL_AddLigne("Fax", -1, "", False)
        n = n + 1
    End If
    If v_yabidon Then
        Call CL_AddLigne("Impression Fictive", -2, "", False)
        n = n + 1
    End If

    If n = 0 Then
        MsgBox "Aucune imprimante n'est paramétrée", vbOKOnly + vbExclamation, ""
        PR_ChoixImp = False
        Exit Function
    End If
    Call CL_InitPointeur(lig_crt)
    ChoixListe.Show 1

    If CL_liste.retour = 0 Then
        If CL_liste.lignes(CL_liste.pointeur).num = -1 Then
            r_yafax = True
            r_yabidon = False
        Else
            r_yafax = False
            If CL_liste.lignes(CL_liste.pointeur).num = -2 Then
                r_yabidon = True
            Else
                r_yabidon = False
                nopr = 0
                For Each pr In Printers
                    If nopr = CL_liste.lignes(CL_liste.pointeur).num Then
                        Set Printer = pr
                        Exit For
                    End If
                    nopr = nopr + 1
                Next pr
            End If
        End If
        PR_ChoixImp = True
    Else
        PR_ChoixImp = False
    End If

End Function

Public Sub PR_ImpLigne(ByRef v_textes() As String)

    Dim i As Integer, posx As Integer, posy As Integer, pos As Integer
    Dim str As String
    Dim vcol As Variant

    posx = g_posx
    posy = g_posy
    For i = 0 To UBound(v_textes)
        str = v_textes(i)
        If g_prcase(i).cadrage = PR_GAUCHE Then
            pos = g_posx + 50
            If Printer.TextWidth(str) + 100 > g_prcase(i).largeur Then
                While Printer.TextWidth(str) + 100 > g_prcase(i).largeur
                    str = left$(str, Len(str) - 1)
                Wend
            End If
        Else
            pos = g_posx + g_prcase(i).largeur - 50 - Printer.TextWidth(str)
        End If
        Printer.CurrentX = pos
        Printer.CurrentY = g_posy + (g_prhauteur - Printer.TextHeight("M")) / 2
        Printer.Print str
        If g_ya_cadre Then
            Printer.FillStyle = 2
            Printer.Print ""
            Printer.FillStyle = vbFSTransparent
            Printer.Line (g_posx, g_posy)-(g_posx + g_prcase(i).largeur, g_posy + g_prhauteur), , B
        End If
        g_posx = g_posx + g_prcase(i).largeur
    Next i
    g_posx = posx
    g_posy = posy + g_prhauteur

    If g_posy >= Printer.Height - 2000 Then
        Printer.NewPage
        g_nopage = g_nopage + 1
        g_posx = 100
        g_posy = 500
        Call imprimer_entete
    End If

End Sub

Public Sub PR_InitFormat(ByVal v_binit As Boolean, _
                         ByVal v_stitre As String, _
                         ByVal v_yacadre As Boolean, _
                         ByVal v_scadrage As String, _
                         ByRef v_stextes() As String)

    Dim n As Integer, i As Integer, lg_tot As Integer, width As Integer
    Dim s As String

    g_ya_cadre = v_yacadre
    
    For n = 0 To UBound(v_stextes)
        ReDim Preserve g_prcase(n) As PRCASE
        g_prcase(n).texte = v_stextes(n)
        g_prcase(n).largeur = 0
        g_prcase(n).cadrage = PR_GAUCHE
        g_prhauteur = Printer.TextHeight(v_stextes(n)) + 100
    Next n

    lg_tot = 0
    width = Printer.width - 1000
    For i = 0 To n - 1
        s = STR_GetChamp(v_scadrage, ";", i)
        If left$(s, 1) = "d" Then g_prcase(i).cadrage = PR_DROITE
        If Len(s) > 1 Then
            g_prcase(i).largeur = (width * Mid$(s, 2)) / 100
            lg_tot = lg_tot + g_prcase(i).largeur
        Else
            g_prcase(i).largeur = width - lg_tot
        End If
    Next i

    Printer.FontItalic = False
    Printer.FontSize = 10
    If v_binit Then
        g_stitre = v_stitre
        g_nopage = 1
        Call imprimer_entete
    End If

End Sub

Public Function PR_InitImp(ByVal v_devicename As String) As Boolean

    Dim pr As Printer

    Set g_printer = Printer
    
    For Each pr In Printers
        If pr.DeviceName = v_devicename Then
            Set Printer = pr
            PR_InitImp = True
            Exit Function
        End If
    Next pr

    PR_InitImp = False

End Function

Public Sub PR_RestoreImp()

    Dim pr As Printer

    For Each pr In Printers
        If pr.DeviceName = g_printer.DeviceName Then
            Set Printer = pr
            Exit For
        End If
    Next

End Sub

Public Sub PR_Cadre(ByVal posdx As Integer, ByVal posdy As Integer, ByVal posfx As Integer, ByVal posfy As Integer)

    Printer.Line (posdx, posdy)-(posfx, posfy), , B

End Sub

Public Sub PR_SetCadre(ByVal cadre As Boolean)

    g_ya_cadre = cadre

End Sub

