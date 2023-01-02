Attribute VB_Name = "MTreeView"
Option Explicit

Public Function TV_NodeExiste(ByVal v_tv As TreeView, _
                              ByVal v_key As String, _
                              ByRef r_nd As Node) As Integer

    On Error GoTo lab_no_key
    Set r_nd = v_tv.Nodes(v_key)
    On Error GoTo 0
    TV_NodeExiste = P_OUI
    Exit Function

lab_no_key:
    TV_NodeExiste = P_NON
    Exit Function

End Function

Public Function TV_ChildNextParent(ByRef vr_nd As Node, _
                                   ByVal v_ndr As Node) As Boolean

    Dim s As String
    Dim ya_suiv As Boolean
    Dim nd As Node

    If vr_nd.Children > 0 Then
        Set vr_nd = vr_nd.Child
        TV_ChildNextParent = True
        Exit Function
    Else
        If vr_nd.Index = v_ndr.Index Then
            TV_ChildNextParent = False
            Exit Function
        End If
        ya_suiv = True
        Set nd = vr_nd.Next
        On Error GoTo lab_no_next
        s = nd.Text
        On Error GoTo 0
        If ya_suiv Then
            Set vr_nd = nd
            TV_ChildNextParent = True
            Exit Function
        End If
        Set vr_nd = vr_nd.Parent
        While vr_nd.Index <> v_ndr.Index
            Set nd = vr_nd.Next
            ya_suiv = True
            On Error GoTo lab_no_next
            s = nd.Text
            On Error GoTo 0
            If ya_suiv Then
                Set vr_nd = nd
                TV_ChildNextParent = True
                Exit Function
            End If
            Set vr_nd = vr_nd.Parent
        Wend
        TV_ChildNextParent = False
        Exit Function
    End If

lab_no_next:
    ya_suiv = False
    Resume Next

End Function

Public Function TV_ChildNextParent_Niveau(ByRef vr_nd As Node, _
                                          ByVal v_ndr As Node, _
                                          ByRef r_decal As Integer) As Boolean

    Dim s As String
    Dim ya_suiv As Boolean
    Dim nd As Node

    If vr_nd.Children > 0 Then
        Set vr_nd = vr_nd.Child
        r_decal = 1
        TV_ChildNextParent_Niveau = True
        Exit Function
    Else
        If vr_nd.Index = v_ndr.Index Then
            TV_ChildNextParent_Niveau = False
            Exit Function
        End If
        ya_suiv = True
        Set nd = vr_nd.Next
        On Error GoTo lab_no_next
        s = nd.Text
        On Error GoTo 0
        If ya_suiv Then
            Set vr_nd = nd
            r_decal = 0
            TV_ChildNextParent_Niveau = True
            Exit Function
        End If
        Set vr_nd = vr_nd.Parent
        r_decal = -1
        While vr_nd.Index <> v_ndr.Index
            Set nd = vr_nd.Next
            ya_suiv = True
            On Error GoTo lab_no_next
            s = nd.Text
            On Error GoTo 0
            If ya_suiv Then
                Set vr_nd = nd
                TV_ChildNextParent_Niveau = True
                Exit Function
            End If
            Set vr_nd = vr_nd.Parent
            r_decal = r_decal - 1
        Wend
        TV_ChildNextParent_Niveau = False
        Exit Function
    End If

lab_no_next:
    ya_suiv = False
    Resume Next

End Function

Public Sub TV_FirstParent(ByVal v_nd As Node, _
                          ByRef vr_ndp As Node)

    Dim s As String
    Dim encore As Boolean
    Dim nd As Node, ndp As Node

    Set nd = v_nd
    Do
        encore = True
        Set ndp = nd.Parent
        On Error GoTo lab_no_prev
        s = ndp.Text
        On Error GoTo 0
        If encore Then
            Set nd = ndp
        End If
    Loop Until Not encore

    Set vr_ndp = nd
    Exit Sub

lab_no_prev:
    encore = False
    Resume Next

End Sub

Public Sub TV_SupprimerAvecPeresVides(ByVal v_tv As TreeView, _
                                      ByVal v_nd As Node)

    Dim encore As Boolean
    Dim nbenf As Integer
    Dim nd As Node, ndp As Node

    Set nd = v_nd
    encore = True
    While encore
        Set ndp = nd.Parent
        Call v_tv.Nodes.Remove(nd.Index)
        On Error GoTo lab_no_parent
        nbenf = ndp.Children
        On Error GoTo 0
        If nbenf > 0 Then encore = False
        Set nd = ndp
    Wend
    Exit Sub

lab_no_parent:
    nbenf = 1
    Resume Next

End Sub

Public Function TV_NodeNext(ByRef vr_nd As Node) As Boolean

    Dim s As String
    Dim nd As Node

    Set nd = vr_nd.Next
    On Error GoTo lab_no_node
    s = nd.tag
    On Error GoTo 0
    Set vr_nd = nd
    TV_NodeNext = True
    Exit Function

lab_no_node:
    TV_NodeNext = False
    Exit Function

End Function

Public Function TV_NodeParent(ByRef vr_nd As Node) As Boolean
' ****************************************************************************
' Trouver le noeud père s'il existe le retourner + True, sinon retourner False
' ****************************************************************************
    Dim s As String
    Dim nd As Node

    Set nd = vr_nd.Parent
    On Error GoTo lab_no_node
    s = nd.tag
    On Error GoTo 0
    Set vr_nd = nd
    TV_NodeParent = True
    Exit Function

lab_no_node:
    TV_NodeParent = False
    Exit Function

End Function

Public Function TV_NodePrev(ByRef vr_nd As Node) As Boolean

    Dim s As String
    Dim nd As Node

    Set nd = vr_nd.Previous
    On Error GoTo lab_no_node
    s = nd.tag
    On Error GoTo 0
    Set vr_nd = nd
    TV_NodePrev = True
    Exit Function

lab_no_node:
    TV_NodePrev = False
    Exit Function

End Function

