Attribute VB_Name = "MRegistre"

Public Sub RunAtStartUp(nom As String, chemin As String)
    
    'Ecriture dans la Base de Registre de la Clé de Démarrage
    RegEcrire 0, "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\" & nom, chemin
  
End Sub
  
Public Sub StopRunningStartUp(nom As String)
    
    'Suppression de la Clé de Démarrage
    RegSupprimer "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\" & nom
  
End Sub
  
Public Function IsRunningOnStartup(nom As String) As Boolean

    IsRunningOnStartup = False
    On Error GoTo fin
    
    Dim Resultat As String
    
    'On lit la clé...
    RegLire "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\" & nom, Resultat
    'On vérifie si le chemin de la clé est valide
    If Dir$(Resultat) <> "" Then
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
  
Public Sub RegEcrire(StyleDeClé As Integer, CheminComplet As String, Valeur As String)

    Dim WshShell As Object
    
    'Style de Clé :
    '0 -> Valeur Chaîne
    '1 -> DWord
    '2 -> Binaire
    Set WshShell = CreateObject("Wscript.Shell")
  
    If StyleDeClé = 0 Then WshShell.RegWrite CheminComplet, Valeur
    If StyleDeClé = 1 Then WshShell.RegWrite CheminComplet, Valeur, "REG_DWORD"
    If StyleDeClé = 2 Then WshShell.RegWrite CheminComplet, Valeur, "REG_BINARY"
  
End Sub

Public Sub RegSupprimer(CheminComplet As String)
  
    Dim WshShell As Object
  
    'Permet d'effacer dans la base de registre tout type de valeur (valeur chaîne, dword, binaire)
    Set WshShell = CreateObject("Wscript.Shell")
  
    WshShell.RegDelete CheminComplet
  
End Sub
