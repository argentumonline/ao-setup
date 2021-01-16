Attribute VB_Name = "General"
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Const ERROR_SUCCESS = &H0
 
Public Function FileExist(ByVal File As String, ByVal fileType As VbFileAttribute) As Boolean
    FileExist = Dir(File, fileType) <> ""
End Function

Public Sub LogError(ByVal errStr As String)
    On Error GoTo ErrHandler
  
    Dim path As String
    Dim oFile As Integer

    path = App.path & "\Errores" & Year(Now) & Month(Now) & Day(Now) & ".log"
    oFile = FreeFile
    
    Open path For Append As #oFile
        Print #oFile, Time & " - " & errStr
    Close #oFile
  
    Exit Sub
  
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LogError de General.bas")
End Sub

Public Function IsWindows7SP1OrBetter() As Boolean
    Dim Version As OSVERSIONINFO
    Version.OSVSize = Len(Version)
    
    If (GetVersionEx(Version) = 1) Then
        If (Version.PlatformID >= 2) Then
            If (Version.dwVerMajor > 6 Or Version.dwVerMinor > 1 Or Version.dwBuildNumber > 7600) Then
                IsWindows7SP1OrBetter = True
            End If
        End If
    End If
End Function

Public Function IsWindows7SP0() As Boolean
    Dim Version As OSVERSIONINFO
    Version.OSVSize = Len(Version)
    
    If (GetVersionEx(Version) = 1) Then
        If (Version.PlatformID >= 2) Then
            If (Version.dwVerMajor = 6 And Version.dwVerMinor = 1 And Version.dwBuildNumber = 7600) Then
                IsWindows7SP0 = True
            End If
        End If
    End If
End Function
 
Public Function RegisterDLL(DllServerPath As String, bRegister As Boolean) As Boolean
    On Error Resume Next
 
    Dim Library As Long, Procedure As Long
    Library = LoadLibrary(DllServerPath)

    If (Library) Then
       If bRegister Then
           Procedure = GetProcAddress(Library, "DllRegisterServer")
       Else
           Procedure = GetProcAddress(Library, "DllUnregisterServer")
       End If
    
        If (Procedure) Then
            RegisterDLL = (CallWindowProc(Procedure, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&) = ERROR_SUCCESS)
        End If

        Call FreeLibrary(Library)
    End If

End Function
