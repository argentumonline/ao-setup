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
        Ambience As tGraphicsAmbienceConfig
        
        GraphicsIndToUse As String * 13 'sGraficos
        bUseFullScreen As Boolean '!bNoRes
        bUseVerticalSync As Boolean
        bUseCompatibleMode As Boolean
    End Type
    
    Public Type tSoundsConfig
        bMusicEnabled As Boolean        '!bNoMusic
        MusicVolume As Byte
        bSoundsEnabled As Boolean       '!bNoSound
        SoundsVolume As Byte
        bSoundEffectsEnabled As Boolean '!bNoSoundEffects
    End Type
    
    Public Type tGuildsConfig
        bShowGuildNews As Boolean        'bGuildNews
        bShowDialogsInConsole As Boolean 'bGldMsgConsole
        MaxMessageQuantity As Byte       'bCantMsgs
    End Type
    
    Public Type tExtraConfig
        Name As String
        NameStyle As Byte               ' Nombres
        bRightClickEnabled As Boolean   'rightClickActivated
        bAskForResolutionChange As Boolean
    End Type
    
    Public Type tGameConfig
        Graphics    As tGraphicsConfig
        Sounds      As tSoundsConfig
        Guilds      As tGuildsConfig
        Extras      As tExtraConfig
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
        Call LoadGuildConfig(iniMan)

        If Not bFileExists Then
            ' Save the file, because we need it.
            Call SaveGameConfig
        End If
           
        Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadUserConfig de mod_GameIni.bas")
    End Sub
    
    Private Sub LoadSoundsConfig(ByRef iniMan As clsIniManager)
        On Error GoTo ErrHandler

        GameConfig.Sounds.bMusicEnabled = iniMan.GetValueBoolean("Sound", "MusicEnabled", True)
        GameConfig.Sounds.bSoundEffectsEnabled = iniMan.GetValueBoolean("Sound", "SoundEffectsEnabled", True)
        GameConfig.Sounds.bSoundsEnabled = iniMan.GetValueBoolean("Sound", "SoundEnabled", True)
        GameConfig.Sounds.MusicVolume = iniMan.GetValueByte("Sound", "MusicVolume", 100)
        GameConfig.Sounds.SoundsVolume = iniMan.GetValueByte("Sound", "SoundsVolume", 100)
        
        Exit Sub
        
ErrHandler:
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadSoundsConfig de mod_Configuration.bas")
    End Sub
    
    
    Private Sub LoadExtrasConfig(ByRef iniMan As clsIniManager)
        On Error GoTo ErrHandler
        
        GameConfig.Extras.Name = iniMan.GetValueString("Extras", "Name", vbNullString)
        GameConfig.Extras.NameStyle = iniMan.GetValueByte("Extras", "NameStyle", 2)
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

    Private Sub LoadGraphicsConfig(ByRef iniMan As clsIniManager)
        On Error GoTo ErrHandler
            
        GameConfig.Graphics.bUseFullScreen = iniMan.GetValueBoolean("GraphicsEngine", "UseFullScreen", False)
        GameConfig.Graphics.bUseVerticalSync = iniMan.GetValueBoolean("GraphicsEngine", "UseVerticalSync", False)
        GameConfig.Graphics.bUseCompatibleMode = iniMan.GetValueBoolean("GraphicsEngine", "UseCompatibleMode", False)
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
                
        Call iniMan.ChangeValue("GraphicsEngine", "UseFullScreen", BooleanToNumber(GameConfig.Graphics.bUseFullScreen))
        Call iniMan.ChangeValue("GraphicsEngine", "UseVerticalSync", BooleanToNumber(GameConfig.Graphics.bUseVerticalSync))
        Call iniMan.ChangeValue("GraphicsEngine", "UseCompatibleMode", BooleanToNumber(GameConfig.Graphics.bUseCompatibleMode))
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
        
        Call iniMan.ChangeValue("Extras", "Name", GameConfig.Extras.Name)
        Call iniMan.ChangeValue("Extras", "NameStyle", GameConfig.Extras.NameStyle)
        Call iniMan.ChangeValue("Extras", "RightClickEnabled", BooleanToNumber(GameConfig.Extras.bRightClickEnabled))
        Call iniMan.ChangeValue("Extras", "AskForResolutionChange", BooleanToNumber(GameConfig.Extras.bAskForResolutionChange))
        
        Exit Sub
ErrHandler:
        Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveExtrasConfig de mod_Configuration.bas")
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
        
        Call iniMan.ChangeValue("Sound", "MusicVolume", GameConfig.Sounds.MusicVolume)
        Call iniMan.ChangeValue("Sound", "SoundsVolume", GameConfig.Sounds.SoundsVolume)
        
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
                
    'Dim oldConfig As tSetupMods
    'oldConfig = LoadOldConfigFile()
    
    ' Load the new config file so we can get the default values for all the properties
    ' that are not included in the old file
    'Call LoadGameConfig
    
    ' Extras
    'GameConfig.Extras.bAskForResolutionChange = oldConfig.bNoRes
    'GameConfig.Extras.bRightClickEnabled = Not oldConfig.bRightClick
    
    ' Sound
    'GameConfig.Sounds.bMusicEnabled = Not oldConfig.bNoMusic
    'GameConfig.Sounds.bSoundsEnabled = Not oldConfig.bNoMusic
    'GameConfig.Sounds.bSoundEffectsEnabled = Not oldConfig.bNoSound
    
    ' Graphics
    'GameConfig.Graphics.bUseDynamicLoad = oldConfig.bDinamic
    'GameConfig.Graphics.bUseVideoMemory = oldConfig.bUseVideo
    'GameConfig.Graphics.MaxVideoMemory = oldConfig.byMemory
    'GameConfig.Graphics.ddexSelectedPlugin = oldConfig.ddexSelectedPlugin
    'GameConfig.Graphics.GraphicsIndToUse = oldConfig.sGraficos
    
    ' Guilds
    'GameConfig.Guilds.bShowDialogsInConsole = oldConfig.bGldMsgConsole
    'GameConfig.Guilds.bShowGuildNews = oldConfig.bGuildNews
    'GameConfig.Guilds.MaxMessageQuantity = oldConfig.bCantMsgs
    

    'Call SaveGameConfig
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub MigrateOldConfigFormat de mod_GameIni.bas")
End Sub
