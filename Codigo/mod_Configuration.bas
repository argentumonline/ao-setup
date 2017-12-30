Attribute VB_Name = "mod_Configuration"
Option Explicit

Private Const INIT_PATH As String = "\INIT\"

''' NEW CONFIG
    
    Public Type tGraphicsAmbienceConfig
        bLightsEnabled As Boolean
        bAmbientLightsEnabled As Boolean
        bUseRainWithParticles As Boolean
    End Type
    
    Public Type tGraphicsConfig
        ddexConfig As DDEXCFG
        ddexConfigured As Boolean
        ddexSelectedPlugin As String
        Ambience As tGraphicsAmbienceConfig
        
        GraphicsIndToUse As String * 13 'sGraficos
        bUseFullScreen As Boolean '!bNoRes
        
        ' Not used anymore
        bUseDynamicLoad As Boolean      'bDinamic
        bUseVideoMemory As Boolean      'bUseVideo
        MaxVideoMemory As Byte          'byMemory
        
    End Type
    
    Public Type tFragShooterConfig
        bEnabled As Boolean             'bActive
        EnemyLevelGreaterThan As Byte   'byMurderedLevel
        bShootWhenDied As Boolean       'bDie
        bShootWhenKill As Boolean       'bKill
    End Type
    
    Public Type tGuildsConfig
        bShowGuildNews As Boolean       'bGuildNews
        bShowDialogsInConsole As Boolean 'bGldMsgConsole
        MaxMessageQuantity As Byte      'bCantMsgs
    End Type
    
    Public Type tSoundsConfig
        bMusicEnabled As Boolean        '!bNoMusic
        bSoundsEnabled As Boolean       '!bNoSound
        bSoundEffectsEnabled As Boolean '!bNoSoundEffects
    End Type
    
    Public Type tExtraConfig
        bRightClickEnabled As Boolean   'rightClickActivated
        bAskForResolutionChange As Boolean
    End Type
    
    Public Type tGameConfig
        Graphics    As tGraphicsConfig
        Sounds      As tSoundsConfig
        FragShooter As tFragShooterConfig
        Guilds      As tGuildsConfig
        Extras       As tExtraConfig
    End Type
    
    Public GameConfig As tGameConfig
    
''' END NEW CONFIG

    Public Sub LoadGameConfig()
        On Error GoTo ErrHandler
        
        Dim iniMan As clsIniManager
        Set iniMan = New clsIniManager
        Dim sPath As String
        Dim bFileExists As Boolean
        
        sPath = App.path & INIT_PATH & "UserConfig.ini"
        
        bFileExists = FileExist(sPath, vbArchive)
        
        ' If the file exists, then initialize the iniMan. If it doesnt exists then
        ' We will still be using the iniMan variable so we can get the default values.
        If bFileExists Then
            ' Initialize the INI manager.
            Call iniMan.Initialize(sPath)
        End If
        
        Call LoadExtrasConfig(iniMan)
        Call LoadGraphicsConfig(iniMan)
        Call LoadSoundsConfig(iniMan)
        Call LoadFragShooterConfig(iniMan)
        Call LoadGuildConfig(iniMan)

        
        ' Save the file, because we need it.
        Call SaveGameConfig
        
                
        Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadUserConfig de mod_GameIni.bas")
    End Sub
    
    Private Sub LoadSoundsConfig(ByRef iniMan As clsIniManager)
        On Error GoTo ErrHandler

        GameConfig.Sounds.bMusicEnabled = iniMan.GetValueBoolean("Sound", "MusicEnabled", True)
        GameConfig.Sounds.bSoundEffectsEnabled = iniMan.GetValueBoolean("Sound", "SoundEffectsEnabled", True)
        GameConfig.Sounds.bSoundsEnabled = iniMan.GetValueBoolean("Sound", "SoundEnabled", True)
        
        Exit Sub
        
ErrHandler:
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadSoundsConfig de mod_Configuration.bas")
    End Sub
    
    
    Private Sub LoadExtrasConfig(ByRef iniMan As clsIniManager)
        On Error GoTo ErrHandler
    
        GameConfig.Extras.bRightClickEnabled = iniMan.GetValueBoolean("Extras", "RightClickEnabled", True)
        GameConfig.Extras.bAskForResolutionChange = iniMan.GetValueBoolean("Extras", "AskForResolutionChange", True)
        
        Exit Sub
ErrHandler:
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadExtrasConfig de mod_Configuration.bas")
    End Sub
        
    Private Sub LoadGuildConfig(ByRef iniMan As clsIniManager)
        On Error GoTo ErrHandler
    
        GameConfig.Guilds.bShowDialogsInConsole = iniMan.GetValueBoolean("Guild", "ShowDialogsInConsole", True)
        GameConfig.Guilds.bShowGuildNews = iniMan.GetValueBoolean("Guild", "ShowGuildNews", True)
        GameConfig.Guilds.MaxMessageQuantity = iniMan.GetValueByte("Guild", "MaxMessageQuantity", 5)
        
        Exit Sub
ErrHandler:
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuildConfig de mod_Configuration.bas")
    End Sub
        
    Private Sub LoadFragShooterConfig(ByRef iniMan As clsIniManager)
        On Error GoTo ErrHandler
    
        GameConfig.FragShooter.bEnabled = iniMan.GetValueBoolean("FragShooter", "FragShooterEnabled", False)
        GameConfig.FragShooter.bShootWhenDied = iniMan.GetValueBoolean("FragShooter", "ShootWhenKilled", False)
        GameConfig.FragShooter.bShootWhenKill = iniMan.GetValueBoolean("FragShooter", "ShootWhenKill", False)
        GameConfig.FragShooter.EnemyLevelGreaterThan = iniMan.GetValueByte("FragShooter", "ShootWhenEnemyLevelGreaterThan", 15)
        
        Exit Sub
ErrHandler:
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadFragShooterConfig de mod_Configuration.bas")
    End Sub
    
    
    Private Sub LoadGraphicsConfig(ByRef iniMan As clsIniManager)
        On Error GoTo ErrHandler
            
        GameConfig.Graphics.ddexConfig.api = API_grafica.DX9 ' This is not used anymore, but needed.
        GameConfig.Graphics.ddexConfig.isDefferal = iniMan.GetValueByte("GraphicsEngine", "UseDeferral", 0)
        GameConfig.Graphics.ddexConfig.memoria = iniMan.GetValueByte("GraphicsEngine", "MemoryMode", 0)
        GameConfig.Graphics.ddexConfig.Modo = iniMan.GetValueByte("GraphicsEngine", "VideoMode", 0)
        GameConfig.Graphics.ddexConfig.MODO2 = iniMan.GetValueByte("GraphicsEngine", "VertexMode", 0)
        GameConfig.Graphics.ddexConfig.vsync = 0 ' Not used anymore, but needed.
        GameConfig.Graphics.ddexSelectedPlugin = iniMan.GetValueString("GraphicsEngine", "SelectedPlugin", "DDEX_DX9.dll")
        
        GameConfig.Graphics.ddexConfigured = True
        
        GameConfig.Graphics.bUseDynamicLoad = iniMan.GetValueBoolean("GraphicsEngine", "UseDynamicLoad", True)
        GameConfig.Graphics.bUseVideoMemory = iniMan.GetValueBoolean("GraphicsEngine", "UseVideoMemory", True)
        GameConfig.Graphics.MaxVideoMemory = iniMan.GetValueByte("GraphicsEngine", "MaxVideoMemory", 40)
        
        GameConfig.Graphics.bUseFullScreen = iniMan.GetValueBoolean("GraphicsEngine", "UseFullScreen", False)
        GameConfig.Graphics.GraphicsIndToUse = iniMan.GetValueString("GraphicsEngine", "GraphicsIndToUse", "Graficos1.ind")
        
        GameConfig.Graphics.Ambience.bAmbientLightsEnabled = iniMan.GetValueBoolean("GraphicsEngine", "EnableAmbientLights", True)
        GameConfig.Graphics.Ambience.bLightsEnabled = iniMan.GetValueBoolean("GraphicsEngine", "EnableLights", True)
        GameConfig.Graphics.Ambience.bUseRainWithParticles = iniMan.GetValueBoolean("GraphicsEngine", "UseRainWithParticles", False)
        
        
        Exit Sub
ErrHandler:
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGraphicsConfig de mod_Configuration.bas")
    End Sub
    
    Private Sub SaveGraphicsConfig(ByRef iniMan As clsIniManager)
        On Error GoTo ErrHandler
        
        Call iniMan.ChangeValue("GraphicsEngine", "UseDeferral", GameConfig.Graphics.ddexConfig.isDefferal)
        Call iniMan.ChangeValue("GraphicsEngine", "MemoryMode", GameConfig.Graphics.ddexConfig.memoria)
        Call iniMan.ChangeValue("GraphicsEngine", "VideoMode", GameConfig.Graphics.ddexConfig.Modo)
        Call iniMan.ChangeValue("GraphicsEngine", "VertexMode", GameConfig.Graphics.ddexConfig.MODO2)
        Call iniMan.ChangeValue("GraphicsEngine", "SelectedPlugin", GameConfig.Graphics.ddexSelectedPlugin)
        Call iniMan.ChangeValue("GraphicsEngine", "UseDynamicLoad", BooleanToNumber(GameConfig.Graphics.bUseDynamicLoad))
        Call iniMan.ChangeValue("GraphicsEngine", "UseVideoMemory", BooleanToNumber(GameConfig.Graphics.bUseDynamicLoad))
        Call iniMan.ChangeValue("GraphicsEngine", "MaxVideoMemory", GameConfig.Graphics.MaxVideoMemory)
        
        Call iniMan.ChangeValue("GraphicsEngine", "UseFullScreen", BooleanToNumber(GameConfig.Graphics.bUseFullScreen))
        Call iniMan.ChangeValue("GraphicsEngine", "GraphicsIndToUse", GameConfig.Graphics.GraphicsIndToUse)
        
        Call iniMan.ChangeValue("GraphicsEngine", "EnableAmbientLights", BooleanToNumber(GameConfig.Graphics.Ambience.bAmbientLightsEnabled))
        Call iniMan.ChangeValue("GraphicsEngine", "EnableLights", BooleanToNumber(GameConfig.Graphics.Ambience.bLightsEnabled))
        Call iniMan.ChangeValue("GraphicsEngine", "UseRainWithParticles", BooleanToNumber(GameConfig.Graphics.Ambience.bUseRainWithParticles))
        
        Exit Sub
ErrHandler:
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveGraphicsConfig de mod_Configuration.bas")
    End Sub
    
    Private Sub SaveExtrasConfig(ByRef iniMan As clsIniManager)
        On Error GoTo ErrHandler
        
        Call iniMan.ChangeValue("Extras", "RightClickEnabled", BooleanToNumber(GameConfig.Extras.bRightClickEnabled))
        Call iniMan.ChangeValue("Extras", "AskForResolutionChange", BooleanToNumber(GameConfig.Extras.bAskForResolutionChange))
        
        Exit Sub
ErrHandler:
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveExtrasConfig de mod_Configuration.bas")
    End Sub
    
    Private Sub SaveFragshooterConfig(ByRef iniMan As clsIniManager)
        On Error GoTo ErrHandler
        
        Call iniMan.ChangeValue("FragShooter", "FragShooterEnabled", BooleanToNumber(GameConfig.FragShooter.bEnabled))
        Call iniMan.ChangeValue("FragShooter", "ShootWhenKilled", BooleanToNumber(GameConfig.FragShooter.bShootWhenDied))
        Call iniMan.ChangeValue("FragShooter", "ShootWhenKill", BooleanToNumber(GameConfig.FragShooter.bShootWhenKill))
        Call iniMan.ChangeValue("FragShooter", "ShootWhenEnemyLevelGreaterThan", GameConfig.FragShooter.EnemyLevelGreaterThan)
        
        Exit Sub
ErrHandler:
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveFragshooterConfig de mod_Configuration.bas")
    End Sub

    Private Sub SaveGuildConfig(ByRef iniMan As clsIniManager)
        On Error GoTo ErrHandler
                
        Call iniMan.ChangeValue("Guild", "ShowDialogsInConsole", BooleanToNumber(GameConfig.Guilds.bShowDialogsInConsole))
        Call iniMan.ChangeValue("Guild", "ShowGuildNews", BooleanToNumber(GameConfig.Guilds.bShowGuildNews))
        Call iniMan.ChangeValue("Guild", "MaxMessageQuantity", GameConfig.Guilds.MaxMessageQuantity)
        
        Exit Sub
ErrHandler:
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveGuildConfig de mod_Configuration.bas")
    End Sub
    
    Private Sub SaveSoundsConfig(ByRef iniMan As clsIniManager)
        On Error GoTo ErrHandler

        Call iniMan.ChangeValue("Sound", "MusicEnabled", BooleanToNumber(GameConfig.Sounds.bMusicEnabled))
        Call iniMan.ChangeValue("Sound", "SoundEffectsEnabled", BooleanToNumber(GameConfig.Sounds.bSoundEffectsEnabled))
        Call iniMan.ChangeValue("Sound", "SoundEnabled", BooleanToNumber(GameConfig.Sounds.bSoundsEnabled))
        
        Exit Sub
ErrHandler:
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveSoundsConfig de mod_Configuration.bas")
    End Sub
   
    
    Private Function BooleanToNumber(ByVal boolValue As Boolean) As Byte
        BooleanToNumber = IIf(boolValue = True, 1, 0)
    End Function
    
    
    Public Sub SaveGameConfig()
        On Error GoTo ErrHandler
        Dim sPath As String
        Dim oFile As Integer
        Dim iniMan As clsIniManager
        Set iniMan = New clsIniManager
        
        sPath = App.path & INIT_PATH & "UserConfig.ini"
    
        If Not FileExist(sPath, vbArchive) Then
            ' Create an empty file if don't exists.
            oFile = FreeFile
            Open sPath For Append As #oFile
                
            Close #oFile
        End If
        
        ' Initialize the INI manager.
        Call iniMan.Initialize(sPath)
        
        Call SaveExtrasConfig(iniMan)
        Call SaveGraphicsConfig(iniMan)
        Call SaveSoundsConfig(iniMan)
        Call SaveFragshooterConfig(iniMan)
        Call SaveGuildConfig(iniMan)
        
        Call iniMan.SaveOpenFile
                
        Exit Sub
ErrHandler:
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveGameConfig de mod_Configuration.bas")
        
    End Sub


Public Function OldConfigExists() As Boolean
    OldConfigExists = FileExist(App.path & "\Init\AO.dat", vbNormal)
End Function

Public Sub RemoveOldConfigFile()
On Error GoTo ErrHandler:

    Call Kill(App.path & "\Init\AO.dat")
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub RemoveOldConfigFile de mod_GameIni.bas")
End Sub

Public Function LoadOldConfigFile() As tSetupMods
On Error GoTo ErrHandler:

    
    ' Check if the INIT folder exists
    If Not FileExist(App.path & "\INIT\", vbDirectory) Then
        Call MkDir(App.path & "\INIT\")
    End If
    
    Dim handle As Integer
    handle = FreeFile
    
    ' Get a reference to the file and load the content in a binary way.
    Open App.path & "\Init\AO.dat" For Binary As handle
        Get handle, , LoadOldConfigFile
    Close handle
    
    Exit Function
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadOldConfigFile de mod_GameIni.bas")
    
End Function

Public Sub MigrateOldConfigFormat()
    On Error GoTo ErrHandler:
                
    Dim oldConfig As tSetupMods
    oldConfig = LoadOldConfigFile()
    
    ' Load the new config file so we can get the default values for all the properties
    ' that are not included in the old file
    Call LoadGameConfig
    
    ' Extras
    GameConfig.Extras.bAskForResolutionChange = oldConfig.bNoRes
    GameConfig.Extras.bRightClickEnabled = Not oldConfig.bRightClick
    
    ' Sound
    GameConfig.Sounds.bMusicEnabled = Not oldConfig.bNoMusic
    GameConfig.Sounds.bSoundsEnabled = Not oldConfig.bNoMusic
    GameConfig.Sounds.bSoundEffectsEnabled = Not oldConfig.bNoSound
    
    ' Graphics
    GameConfig.Graphics.bUseDynamicLoad = oldConfig.bDinamic
    GameConfig.Graphics.bUseVideoMemory = oldConfig.bUseVideo
    GameConfig.Graphics.MaxVideoMemory = oldConfig.byMemory
    GameConfig.Graphics.ddexSelectedPlugin = oldConfig.ddexSelectedPlugin
    GameConfig.Graphics.GraphicsIndToUse = oldConfig.sGraficos
    ' Graphics -> DDEx
    GameConfig.Graphics.ddexConfig.api = oldConfig.ddexConfig.api
    GameConfig.Graphics.ddexConfig.isDefferal = oldConfig.ddexConfig.isDefferal
    GameConfig.Graphics.ddexConfig.memoria = oldConfig.ddexConfig.memoria
    GameConfig.Graphics.ddexConfig.Modo = oldConfig.ddexConfig.Modo
    GameConfig.Graphics.ddexConfig.MODO2 = oldConfig.ddexConfig.MODO2
    GameConfig.Graphics.ddexConfig.vsync = oldConfig.ddexConfig.vsync

    ' Guilds
    GameConfig.Guilds.bShowDialogsInConsole = oldConfig.bGldMsgConsole
    GameConfig.Guilds.bShowGuildNews = oldConfig.bGuildNews
    GameConfig.Guilds.MaxMessageQuantity = oldConfig.bCantMsgs
    
    ' FragShooter
    GameConfig.FragShooter.bEnabled = oldConfig.bActive
    GameConfig.FragShooter.bShootWhenDied = oldConfig.bDie
    GameConfig.FragShooter.bShootWhenKill = oldConfig.bKill
    GameConfig.FragShooter.EnemyLevelGreaterThan = oldConfig.byMurderedLevel

    Call SaveGameConfig
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MigrateOldConfigFormat de mod_GameIni.bas")
End Sub
