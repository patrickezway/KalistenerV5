VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ce projet avec son code a été téléchargé de
' http://www.i-france.com/estdev
' e-mail : mstoll@i-France.com
' si vous avez une amélioration pour ce code ou si vous
' trouvez un bug, écrivez moi. Merci
' Copyright © 1998 EstDéveloppement - Coulon Richard
'

Option Explicit

Private cs_init As Boolean
Private cs_mode As Integer  ' 1=vertical 2=horizontal 3=les 2
Private cs_largeur As Long
Private cs_hauteur As Long
Private cs_largeur_init As Long
Private cs_hauteur_init As Long
Private cs_pere As Object
Private cs_form As Object
Private cs_objets() As Object
Private Type SIZE_OBJ
    coord_left As Long
    coord_top As Long
    largeur As Long
    hauteur As Long
    tblcolwidth() As Long
End Type
Private cs_prmobj() As SIZE_OBJ

Private Sub Class_Initialize()

    ' defaut pour etirement hori et verti
    cs_mode = 3
    ' initialiser Width et Height
    cs_largeur = -1
    cs_hauteur = -1
    ' flag signalant si on a un client (objet) valide
    cs_init = False
    
End Sub

Property Get Mode() As Integer
    
    Mode = cs_mode

End Property

Property Let Mode(ByVal v_mode As Integer)
    
    cs_mode = v_mode

End Property

Property Get Pere() As Object

    Pere = cs_pere
    
End Property

Property Set Pere(ByVal v_pere As Object)
    
    Set cs_pere = v_pere
    
    ReDim cs_objets(0 To 0)
    
    cs_init = True
    
    cs_largeur = cs_pere.Width
    cs_hauteur = cs_pere.Height

End Property

Property Set LaForm(ByVal v_form As Object)
    
    Set cs_form = v_form
    cs_largeur_init = v_form.Width
    cs_hauteur_init = v_form.Height
    
End Property

Public Sub AddChildren(ByRef v_tblobj As Variant)
    
    Dim i As Integer, lbas  As Integer, lhaut As Integer
    
    If (VarType(v_tblobj) And vbObject) = 0 And Not IsObject(v_tblobj) Then
        Exit Sub
    End If
    
    If (VarType(v_tblobj) And vbArray) = 0 Then
        ReDim cs_objets(1 To 1)
        Set cs_objets(1) = v_tblobj
        ReDim cs_prmobj(1)
        cs_prmobj(i).coord_left = v_tblobj(i).Left
        cs_prmobj(i).coord_top = v_tblobj(i).Top
        cs_prmobj(i).largeur = v_tblobj(i).Width
        cs_prmobj(i).hauteur = v_tblobj(i).Height
    Else
        lbas = LBound(v_tblobj)
        lhaut = UBound(v_tblobj)
        ReDim cs_objets(1 To (lhaut - lbas + 1))
        ReDim cs_prmobj(1 To (lhaut - lbas + 1))
        For i = lbas To lhaut
            Set cs_objets(i - lbas + 1) = v_tblobj(i)
            cs_prmobj(i - lbas + 1).coord_left = v_tblobj(i).Left
            cs_prmobj(i - lbas + 1).coord_top = v_tblobj(i).Top
            cs_prmobj(i - lbas + 1).largeur = v_tblobj(i).Width
            cs_prmobj(i - lbas + 1).hauteur = v_tblobj(i).Height
            If TypeOf cs_objets(i - lbas + 1) Is MSFlexGrid Then
                ReDim cs_prmobj(i - lbas + 1).tblcolwidth(cs_objets(i - lbas + 1))
            End If
        Next i
    End If

End Sub

Public Sub Stretch()
    
    Dim f_deplh As Boolean, f_deplv As Boolean
    Dim i As Integer, lbas As Integer, lhaut As Integer, col As Integer
    Dim coord_left As Long, coord_top As Long, largeur As Long, hauteur As Long
    Dim size As Long
    Dim deltaX As Double, deltaY As Double
    Dim save_font As Object
    
    ' si pas client valide, sortir
    If Not cs_init Then
        Exit Sub
    End If
    
    ' si hauteur et largeur ne sont pas initialisees,
    ' les definir et sortir
'    If cs_largeur = -1 Or cs_hauteur = -1 Then
'        cs_largeur = cs_pere.Width
'        cs_hauteur = cs_pere.Height
'        Exit Sub
'    End If
    
    ' facteur pour etirer horizontalement
    deltaX = cs_pere.Width / cs_largeur_init
    ' facteur pour etirer verticalement
    deltaY = cs_pere.Height / cs_hauteur_init

    ' definir les variables hauteur/largeur vers
    ' la hauteur et largeur courante
    cs_largeur = cs_pere.Width
    cs_hauteur = cs_pere.Height
    
    lbas = LBound(cs_objets)
    lhaut = UBound(cs_objets)
    If lbas = 0 And lhaut = 0 Then
        Exit Sub
    End If
    
    For i = lbas To lhaut
        f_deplv = False
        f_deplh = False
        coord_left = cs_prmobj(i).coord_left
        coord_top = cs_prmobj(i).coord_top
        largeur = cs_prmobj(i).largeur
        hauteur = cs_prmobj(i).hauteur
        ' Déplacement horizontal
        If (cs_mode = 2 Or cs_mode = 3) And (deltaX <> 1) Then
            f_deplh = True
            coord_left = coord_left * deltaX
            ' Cas du label
            If TypeOf cs_objets(i) Is Label Then
                Set save_font = cs_form.Font
                Set cs_form.Font = cs_objets(i).Font
                size = cs_form.TextWidth(cs_objets(i).Caption)
                If largeur * deltaX >= size Then
                    largeur = largeur * deltaX
                End If
                Set cs_form.Font = save_font
            Else
                largeur = largeur * deltaX
            End If
        End If
        ' Déplacement vertical
        If (cs_mode = 1 Or cs_mode = 3) And (deltaY <> 1) Then
            f_deplv = True
            coord_top = coord_top * deltaY
            ' Changement hauteur sauf label et textbox
            If Not (TypeOf cs_objets(i) Is Label) And _
               Not (TypeOf cs_objets(i) Is TextBox) Then
                hauteur = hauteur * deltaY
            End If
        End If
        ' Changement taille ou position
        If f_deplv Or f_deplh Then
'            On Error Resume Next
'            cs_objets(i).FontSize = cs_objets(i).FontSize * (deltaX + deltaY) / 2
'            On Error GoTo 0
            If TypeOf cs_objets(i) Is MSFlexGrid Then
                If f_deplh And deltaX <> 1 Then
                    For col = 0 To cs_objets(i).Cols - 1
                        cs_objets(i).ColWidth(col) = cs_objets(i).ColWidth(col) * deltaX
                    Next col
                End If
            End If
            cs_objets(i).Move coord_left, coord_top, largeur, hauteur
        End If
    Next i
    
End Sub

