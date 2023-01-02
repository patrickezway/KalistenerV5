Attribute VB_Name = "CodeBarre"
Option Explicit

Private g_type_codebarre As Integer
Private g_largeur_codebarre As Double
Private g_hauteur_codebarre As Double
Private g_WN_ratio As Double

Public Const CB_CODE_39 = 1
Public Const CB_CODE_MONARCH = 2
Public Const CB_CODE_25E = 3

Private Sub CB_Imprimer39(ByVal scode As String)

    Dim i As Integer
    Dim s As String
    Dim OldScaleMode As Integer
    
    OldScaleMode = Printer.ScaleMode
    ' en mm
    Printer.ScaleMode = 6
    'start code = *
    Call imprimer_barres("nwnnwnwnn")
    
    For i = 1 To Len(scode)
        Select Case Mid$(scode, i, 1)
        Case "0"
            Call imprimer_barres("nnnwwnwnn")
        Case "1"
            Call imprimer_barres("wnnwnnnnw")
        Case "2"
            Call imprimer_barres("nnwwnnnnw")
        Case "3"
            Call imprimer_barres("wnwwnnnnn")
        Case "4"
            Call imprimer_barres("nnnwwnnnw")
        Case "5"
            Call imprimer_barres("wnnwwnnnn")
        Case "6"
            Call imprimer_barres("nnwwwnnnn")
        Case "7"
            Call imprimer_barres("nnnwnnwnw")
        Case "8"
            Call imprimer_barres("wnnwnnwnn")
        Case "9"
            Call imprimer_barres("nnwwnnwnn")
        Case "A"
            Call imprimer_barres("wnnnnwnnw")
        Case "B"
            Call imprimer_barres("nnwnnwnnw")
        Case "C"
            Call imprimer_barres("wnwnnwnnn")
        Case "D"
            Call imprimer_barres("nnnnwwnnw")
        Case "E"
            Call imprimer_barres("wnnnwwnnn")
        Case "F"
            Call imprimer_barres("nnwnwwnnn")
        Case "G"
            Call imprimer_barres("nnnnnwwnw")
        Case "H"
            Call imprimer_barres("wnnnnwwnn")
        Case "I"
            Call imprimer_barres("nnwnnwwnn")
        Case "J"
            Call imprimer_barres("nnnnwwwnn")
        Case "K"
            Call imprimer_barres("wnnnnnnww")
        Case "L"
            Call imprimer_barres("nnwnnnnww")
        Case "M"
            Call imprimer_barres("wnwnnnnwn")
        Case "N"
            Call imprimer_barres("nnnnwnnww")
        Case "O"
            Call imprimer_barres("wnnnwnnwn")
        Case "P"
            Call imprimer_barres("nnwnwnnwn")
        Case "Q"
            Call imprimer_barres("nnnnnnwww")
        Case "R"
            Call imprimer_barres("wnnnnnwwn")
        Case "S"
            Call imprimer_barres("nnwnnnwwn")
        Case "T"
            Call imprimer_barres("nnnnwnwwn")
        Case "U"
            Call imprimer_barres("wwnnnnnnw")
        Case "V"
            Call imprimer_barres("nwwnnnnnw")
        Case "W"
            Call imprimer_barres("wwwnnnnnn")
        Case "X"
            Call imprimer_barres("nwnnwnnnw")
        Case "Y"
            Call imprimer_barres("wwnnwnnnn")
        Case "Z"
            Call imprimer_barres("nwwnwnnnn")
        Case "-"
            Call imprimer_barres("nwnnnnwnw")
        Case "."
            Call imprimer_barres("wwnnnnwnn")
        Case " "
            Call imprimer_barres("nwwnnnwnn")
        Case "*"
            Call imprimer_barres("nwnnwnwnn")
        Case "$"
            Call imprimer_barres("nwnwnwnnn")
        Case "/"
            Call imprimer_barres("nwnwnnnwn")
        Case "+"
            Call imprimer_barres("nwnnnwnwn")
        Case "%"
            Call imprimer_barres("nnnwnwnwn")
        End Select
    Next i
    
    ' stop code *
    Call imprimer_barres("nwnnwnwnn")
    
    Printer.ScaleMode = OldScaleMode
    
End Sub

Private Sub CB_Imprimer27(ByVal scode As String)

    Dim i As Integer
    Dim s As String
    Dim OldScaleMode As Integer
    
    OldScaleMode = Printer.ScaleMode
    Printer.ScaleMode = 6
    'start code = t
    Call imprimer_barres("nnwwnwn")
    
    For i = 1 To Len(scode)
        Select Case Mid$(scode, i, 1)
        Case "0"
            Call imprimer_barres("nnnnnww")
        Case "1"
            Call imprimer_barres("nnnnwwn")
        Case "2"
            Call imprimer_barres("nnnwnnw")
        Case "3"
            Call imprimer_barres("wwnnnnn")
        Case "4"
            Call imprimer_barres("nnwnnwn")
        Case "5"
            Call imprimer_barres("wnnnnwn")
        Case "6"
            Call imprimer_barres("nwnnnnw")
        Case "7"
            Call imprimer_barres("nwnnwnn")
        Case "8"
            Call imprimer_barres("nwwnnnn")
        Case "9"
            Call imprimer_barres("wnnwnnn")
        Case "$"
            Call imprimer_barres("nnwwnnn")
        End Select
    Next i
    
    ' stop code = t
    Call imprimer_barres("nnwwnwn")
    
    Printer.ScaleMode = OldScaleMode
    
End Sub

' sert pour le code 39 et le code 27=monarch
Private Sub imprimer_barres(ByVal scb As String)
    
    Dim s As String
    Dim i As Integer
    Dim largeur As Double
    Dim hauteur As Double
    
    hauteur = g_hauteur_codebarre
    For i = 1 To Len(scb)
        s = Mid$(scb, i, 1)
        If s = "w" Then
            largeur = g_WN_ratio * g_largeur_codebarre
        Else
            largeur = g_largeur_codebarre
        End If
        If i Mod 2 = 1 Then
            Printer.Line Step(0, 0)-Step(largeur, hauteur), QBColor(0), BF
            Printer.CurrentY = Printer.CurrentY - hauteur
        Else
            Printer.CurrentX = Printer.CurrentX + largeur
        End If
    Next i
    ' espace inter caractere
    Printer.CurrentX = Printer.CurrentX + g_largeur_codebarre
    
End Sub

Private Function barres25(ByVal s As String) As String

    Select Case s
    Case "0"
        barres25 = "nnwwn"
    Case "1"
        barres25 = "wnnnw"
    Case "2"
        barres25 = "nwnnw"
    Case "3"
        barres25 = "wwnnn"
    Case "4"
        barres25 = "nnwnw"
    Case "5"
        barres25 = "wnwnn"
    Case "6"
        barres25 = "nwwnn"
    Case "7"
        barres25 = "nnnww"
    Case "8"
        barres25 = "wnnwn"
    Case "9"
        barres25 = "nwnwn"
    End Select
End Function

Private Sub CB_Imprimer25E(ByVal scode As String)

    Dim i As Integer
    Dim s1 As String, s2 As String
    Dim OldScaleMode As Integer
    
    If Len(scode) Mod 2 <> 0 Then
        MsgBox "En code 2/5E le nombre de chiffres doit etre pair" & vbCr & _
                "Ce qui n'est pas le cas pour " & scode
    End If
    OldScaleMode = Printer.ScaleMode
    Printer.ScaleMode = 6
    'start code
    Call imprimer_barres25("nn", "nn")
    
    For i = 1 To Len(scode) Step 2
        s1 = barres25(Mid$(scode, i, 1))
        s2 = barres25(Mid$(scode, i + 1, 1))
        Call imprimer_barres25(s1, s2)
    Next i
    
    ' stop code
    Call imprimer_barres25("wn", "nn")
    
    Printer.ScaleMode = OldScaleMode
    
End Sub

Private Sub imprimer_barres25(ByVal scb1 As String, ByVal scb2 As String)

    Dim s As String
    Dim i As Integer
    Dim largeur As Double
    Dim hauteur As Double
    
    hauteur = g_hauteur_codebarre
    For i = 1 To Len(scb1)
        s = Mid$(scb1, i, 1)
        If s = "w" Then
            largeur = g_WN_ratio * g_largeur_codebarre
        Else
            largeur = g_largeur_codebarre
        End If
        Printer.Line Step(0, 0)-Step(largeur, hauteur), QBColor(0), BF
        Printer.CurrentY = Printer.CurrentY - hauteur
        
        s = Mid$(scb2, i, 1)
        If s = "w" Then
            largeur = g_WN_ratio * g_largeur_codebarre
        Else
            largeur = g_largeur_codebarre
        End If
        Printer.CurrentX = Printer.CurrentX + largeur
    Next i
End Sub

Public Sub CB_Imprimer(ByVal scode As String)
    
    If g_type_codebarre = CB_CODE_39 Then
        Call CB_Imprimer39(scode)
    ElseIf g_type_codebarre = CB_CODE_MONARCH Then
        Call CB_Imprimer27(scode)
    ElseIf g_type_codebarre = CB_CODE_25E Then
        Call CB_Imprimer25E(scode)
    End If
End Sub

Public Sub CB_InitPrm(ByVal type_code As Integer, _
                      ByVal largeur As Double, _
                      ByVal hauteur As Double, _
                      ByVal ratio As Double)
                         
    g_type_codebarre = type_code
    g_largeur_codebarre = largeur
    g_hauteur_codebarre = hauteur
    g_WN_ratio = ratio
    
End Sub
