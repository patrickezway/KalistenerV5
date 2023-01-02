Attribute VB_Name = "MFormResize"
Option Explicit

Private g_FormResize() As New FormResize

Public Sub FormResize_Init(ByVal v_form As Form)

    Dim ctrl As Control
    Dim i As Integer, nbs As Integer, nbobj As Integer, iobj As Integer
    Dim tblobj() As Object, tblresize() As Object
    
    nbs = 0
    
    nbobj = 0
    ReDim tblobj(0) As Object
    Set tblobj(0) = v_form
    
    iobj = 0
    While iobj <= nbobj
        ReDim Preserve g_FormResize(nbs)
        ' étirer horizontalement
        Set g_FormResize(nbs).LaForm = v_form
        g_FormResize(nbs).Mode = 3
        ' etirer la feuille courante
        Set g_FormResize(nbs).Pere = tblobj(iobj)
        ' faire liste des controles  a etirer
        i = 1
        For Each ctrl In v_form.Controls
'            If ctrl.Container.hWnd = tblobj(iobj).hWnd Then
                ReDim Preserve tblresize(1 To i)
                Set tblresize(i) = ctrl
'                If TypeOf ctrl Is Frame Then
'                    nbobj = nbobj + 1
'                    ReDim Preserve tblobj(nbobj)
'                    Set tblobj(nbobj) = ctrl
'                End If
                i = i + 1
'            End If
        Next ctrl
        ' quel controle etirer
        g_FormResize(nbs).AddChildren tblresize()
        nbs = nbs + 1
        iobj = iobj + 1
    Wend
    
End Sub

Public Sub FormResize_Resize()

    Dim i As Integer
    
    ' etirer/bouger les controles lorsque la taille de la feuille a changer
    For i = 0 To UBound(g_FormResize)
        g_FormResize(i).Stretch
    Next i

End Sub

