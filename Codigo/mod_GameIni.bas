Attribute VB_Name = "mod_GameIni"
Option Explicit

Public Enum MODO_MEMORIA
    DX9_MEM_D = 0 'manejo de memoria por defecto
    DX9_MEM_A = 1 'manejo de memoria administrador
    DX9_MEM_S = 2 'manejo de memoria de sistema
End Enum
Public Enum API_grafica
    DX9 = 0
    DX92 = 1
    OGL = 2
    DX8 = 3
End Enum
Public Enum MODO2
    DX9_VH = 0 'usar procesamiento de vertices por hardware
    DX9_VS = 1 'usar procesamiento de vertices por software
    OGL_V = 2 'solo es una guia
End Enum

Public Enum Modo
    DX9_HARD = 0
    DX9_REF = 1
    DX9_SOF = 2
    OGL_ = 2 'solo es una guia
End Enum

Private Const INIT_PATH As String = "\INIT\"

Public Sub LoadUserConfig()
On Error GoTo ErrHandler
    
    Dim iniMan As clsIniManager
    Set iniMan = New clsIniManager
    Dim sPath As String
    
    sPath = App.path & INIT_PATH & "UserConfig.ini"
    
    If Not FileExist(sPath, vbArchive) Then
        ' Create a default configuration file if it's not present
        Call LoadDefaultUserConfig
    End If
    
    Call iniMan.Initialize(sPath)
    
    ' Load GraphicsEngine Config
    If iniMan.KeyExists("GraphicsEngine") Then
        ClientConfig.ddexConfig.api = API_grafica.OGL ' This is not used anymore, but needed. FUCK YOU LOOPZER!
        ClientConfig.ddexConfig.isDefferal = iniMan.GetValue("GraphicsEngine", "UseDeferral")
        ClientConfig.ddexConfig.memoria = iniMan.GetValue("GraphicsEngine", "MemoryMode")
        ClientConfig.ddexConfig.Modo = iniMan.GetValue("GraphicsEngine", "VideoMode")
        ClientConfig.ddexConfig.MODO2 = iniMan.GetValue("GraphicsEngine", "VertexMode")
        ClientConfig.ddexConfig.vsync = 0 ' Not used anymore, but needed.
        ClientConfig.ddexSelectedPlugin = iniMan.GetValue("GraphicsEngine", "SelectedPlugin")
        ClientConfig.ddexConfigured = True
    Else
        'Load the default graphics
        Call LoadDefaultUserConfig
    End If
    
    ' Check if there's a ddex plugin selected
    If ClientConfig.ddexSelectedPlugin = "" Then
        ClientConfig.ddexSelectedPlugin = "DDEX_DX9.dll"
    End If
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadUserConfig de mod_GameIni.bas")
End Sub

Public Sub SaveUserConfig()
On Error GoTo ErrHandler:
    Dim iniMan As clsIniManager
    Set iniMan = New clsIniManager
    Dim sPath As String
    Dim oFile As Integer
    
    sPath = App.path & INIT_PATH & "UserConfig.ini"
    
    If Not FileExist(sPath, vbArchive) Then
        ' Create an empty file if don't exists.
        oFile = FreeFile
        Open sPath For Append As #oFile
            
        Close #oFile
    End If
        
    Call iniMan.Initialize(sPath)
    
    Call iniMan.ChangeValue("GraphicsEngine", "UseDeferral", ClientConfig.ddexConfig.isDefferal)
    Call iniMan.ChangeValue("GraphicsEngine", "MemoryMode", ClientConfig.ddexConfig.memoria)
    Call iniMan.ChangeValue("GraphicsEngine", "VideoMode", ClientConfig.ddexConfig.Modo)
    Call iniMan.ChangeValue("GraphicsEngine", "VertexMode", ClientConfig.ddexConfig.MODO2)
    Call iniMan.ChangeValue("GraphicsEngine", "SelectedPlugin", ClientConfig.ddexSelectedPlugin)
    Call iniMan.DumpFile(sPath)
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveUserConfig de mod_GameIni.bas")
    
End Sub


Private Sub LoadDefaultUserConfig()
On Error GoTo ErrHandler:

    ClientConfig.ddexConfig.api = API_grafica.OGL ' This is not used anymore, but needed. FUCK YOU LOOPZER!
    ClientConfig.ddexConfig.isDefferal = 1
    ClientConfig.ddexConfig.memoria = MODO_MEMORIA.DX9_MEM_A
    ClientConfig.ddexConfig.Modo = Modo.DX9_HARD
    ClientConfig.ddexConfig.MODO2 = MODO2.DX9_VH
    ClientConfig.ddexConfig.vsync = 0 ' Not used anymore, but needed.
    ClientConfig.ddexSelectedPlugin = "DDEX_DX9.dll"
    
    'Save the default values.
    Call SaveUserConfig
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadDefaultUserConfig de mod_GameIni.bas")
  
End Sub
