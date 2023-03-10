Attribute VB_Name = "MRegistre"

Public Sub RunAtStartUp(nom As String, chemin As String)
    
    'Ecriture dans la Base de Registre de la Cl? de D?marrage
    RegEcrire 0, "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\" & nom, chemin
  
End Sub
  
Public Sub StopRunningStartUp(nom As String)
    
    'Suppression de la Cl? de D?marrage
    RegSupprimer "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\" & nom
  
End Sub
  
Public Function IsRunningOnStartup(nom As String) As Boolean

    IsRunningOnStartup = False
    On Error GoTo fin
    
    Dim Resultat As String
    
    'On lit la cl?...
    RegLire "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\" & nom, Resultat
    'On v?rifie si le chemin de la cl? est valide
    If Dir$(STR_GetChamp(Resultat, " ", 0)) <> "" Then
        IsRunningOnStartup = True
    Else
        IsRunningOnStartup = False
    End If
fin:

End Function
  
Public Sub RegLire(CheminComplet As String, Destination As String)
  
    Dim WshShell As Object
  
    'permet de lire une valeur dans la base de registre
    Set WshShell = CreateObject("Wscript.Shell")
  
    Destination = WshShell.RegRead(CheminComplet)
  
End Sub
  
Public Sub RegEcrire(StyleDeCl? As Integer, CheminComplet As String, Valeur As String)

    Dim WshShell As Object
    
    'Style de Cl? :
    '0 -> Valeur Cha?ne
    '1 -> DWord
    '2 -> Binaire
    Set WshShell = CreateObject("Wscript.Shell")
  
    If StyleDeCl? = 0 Then WshShell.RegWrite CheminComplet, Valeur
    If StyleDeCl? = 1 Then WshShell.RegWrite CheminComplet, Valeur, "REG_DWORD"
    If StyleDeCl? = 2 Then WshShell.RegWrite CheminComplet, Valeur, "REG_BINARY"
  
End Sub

Public Sub RegSupprimer(CheminComplet As String)
  
    Dim WshShell As Object
  
    'Permet d'effacer dans la base de registre tout type de valeur (valeur cha?ne, dword, binaire)
    Set WshShell = CreateObject("Wscript.Shell")
  
    WshShell.RegDelete CheminComplet
  
End Sub
