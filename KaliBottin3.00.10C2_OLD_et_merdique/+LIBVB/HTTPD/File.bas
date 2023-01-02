Attribute VB_Name = "File"
'API Fichiers
Public Const MOVEFILE_REPLACE_EXISTING = &H1
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_BEGIN = 0
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const OPEN_ALWAYS = 4

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

Public Const INVALID_HANDLE_VALUE = -1

Public Declare Function GetLastError Lib "kernel32" () As Long

Public Declare Function ReadFile Lib "kernel32" _
    (ByVal hFile As Long, _
     lpBuffer As Any, _
     ByVal nNumberOfBytesToRead As Long, _
     lpNumberOfBytesRead As Long, _
     ByVal lpOverlapped As Long) As Long

Public Declare Function WriteFile Lib "kernel32" _
    (ByVal hFile As Long, _
     lpBuffer As Any, _
     ByVal nNumberOfBytesToWrite As Long, _
     lpNumberOfBytesWritten As Long, _
     lpOverlapped As Any) As Long

Public Declare Function CreateFile Lib "kernel32" _
    Alias "CreateFileA" _
    (ByVal lpFileName As String, _
     ByVal dwDesiredAccess As Long, _
     ByVal dwShareMode As Long, _
     ByVal lpSecurityAttributes As Long, _
     ByVal dwCreationDisposition As Long, _
     ByVal dwFlagsAndAttributes As Long, _
     ByVal hTemplateFile As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Public Declare Function SetFilePointer Lib "kernel32" _
    (ByVal hFile As Long, _
     ByVal lDistanceToMove As Long, _
     lpDistanceToMoveHigh As Long, _
     ByVal dwMoveMethod As Long) As Long

'Public Declare Function GetFileSize Lib "kernel32" _
'    (ByVal hFile As Long, _
'    lpFileSizeHigh As Long) As Long
Private Declare Function APIGetFileSize Lib "kernel32.dll" Alias "GetFileSize" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Function GetFileSize(ByRef vsFilePath As String) As Currency
Dim hFile As Long
Dim xnFileSize(1) As Long
    hFile = CreateFile(vsFilePath, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, 0, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
        xnFileSize(0) = APIGetFileSize(hFile, xnFileSize(1))
        CopyMemory GetFileSize, xnFileSize(0), LenB(GetFileSize)
        GetFileSize = GetFileSize * 10000
        CloseHandle hFile
    End If
End Function

