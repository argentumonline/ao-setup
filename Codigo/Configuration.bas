Attribute VB_Name = "Configuration"
Option Explicit

Private Const INIT_PATH As String = "\INIT\"

''' NEW CONFIG
  
Public Type tGraphicsAmbienceConfig
    bLightsEnabled        As Boolean
    bAmbientLightsEnabled As Boolean
    bUseRainWithParticles As Boolean
End Type
    
Public Type tGraphicsConfig
    Ambience           As tGraphicsAmbienceConfig
        
    bUseFullScreen     As Boolean
    bUseVerticalSync   As Boolean
    bUseCompatibleMode As Boolean
End Type
    
Public Type tSoundsConfig
    bMusicEnabled        As Boolean
    MusicVolume          As Byte
    bSoundsEnabled       As Boolean
    SoundsVolume         As Byte
    bSoundEffectsEnabled As Boolean
End Type
    
Public Type tGuildsConfig
    bShowGuildNews        As Boolean
    bShowDialogsInConsole As Boolean
    MaxMessageQuantity    As Byte
End Type
    
Public Type tExtraConfig
    Name                    As String
    NameStyle               As Byte
    bRightClickEnabled      As Boolean
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
        
    Dim iniMan As IniFile
    Set iniMan = New IniFile
    Dim sPath As String
    Dim bFileExists As Boolean
        
    sPath = App.path & INIT_PATH & "UserConfig.ini"
        
    bFileExists = FileExist(sPath, vbArchive)
        
    ' If the file exists, then initialize the iniMan. If it doesnt exists then
    ' We will still be using the iniMan variable so we can get the default values.
    If bFileExists Then
        Call iniMan.Initialize(sPath)
    End If
        
    Call LoadExtrasConfig(iniMan)
    Call LoadGraphicsConfig(iniMan)
    Call LoadSoundsConfig(iniMan)
    Call LoadGuildConfig(iniMan)

    If Not bFileExists Then
        Call SaveGameConfig
    End If
           
    Exit Sub
    
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadUserConfig de mod_GameIni.bas")
End Sub
    
Private Sub LoadSoundsConfig(ByVal iniMan As IniFile)
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
    
Private Sub LoadExtrasConfig(ByVal iniMan As IniFile)
    On Error GoTo ErrHandler
        
    GameConfig.Extras.Name = iniMan.GetValueString("Extras", "Name", vbNullString)
    GameConfig.Extras.NameStyle = iniMan.GetValueByte("Extras", "NameStyle", 2)
    GameConfig.Extras.bRightClickEnabled = iniMan.GetValueBoolean("Extras", "RightClickEnabled", True)
    GameConfig.Extras.bAskForResolutionChange = iniMan.GetValueBoolean("Extras", "AskForResolutionChange", True)
        
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadExtrasConfig de mod_Configuration.bas")
End Sub
        
Private Sub LoadGuildConfig(ByVal iniMan As IniFile)
    On Error GoTo ErrHandler
    
    GameConfig.Guilds.bShowDialogsInConsole = iniMan.GetValueBoolean("Guild", "ShowDialogsInConsole", True)
    GameConfig.Guilds.bShowGuildNews = iniMan.GetValueBoolean("Guild", "ShowGuildNews", True)
    GameConfig.Guilds.MaxMessageQuantity = iniMan.GetValueByte("Guild", "MaxMessageQuantity", 5)
        
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGuildConfig de mod_Configuration.bas")
End Sub

Private Sub LoadGraphicsConfig(ByVal iniMan As IniFile)
    On Error GoTo ErrHandler
            
    GameConfig.Graphics.bUseFullScreen = iniMan.GetValueBoolean("GraphicsEngine", "UseFullScreen", False)
    GameConfig.Graphics.bUseVerticalSync = iniMan.GetValueBoolean("GraphicsEngine", "UseVerticalSync", False)
    GameConfig.Graphics.bUseCompatibleMode = iniMan.GetValueBoolean("GraphicsEngine", "UseCompatibleMode", False)
 
    GameConfig.Graphics.Ambience.bAmbientLightsEnabled = iniMan.GetValueBoolean("GraphicsEngine", "EnableAmbientLights", True)
    GameConfig.Graphics.Ambience.bLightsEnabled = iniMan.GetValueBoolean("GraphicsEngine", "EnableLights", True)
    GameConfig.Graphics.Ambience.bUseRainWithParticles = iniMan.GetValueBoolean("GraphicsEngine", "UseRainWithParticles", False)
 
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LoadGraphicsConfig de mod_Configuration.bas")
End Sub
    
Private Sub SaveGraphicsConfig(ByVal iniMan As IniFile)
    On Error GoTo ErrHandler
                
    Call iniMan.ChangeValue("GraphicsEngine", "UseFullScreen", BooleanToNumber(GameConfig.Graphics.bUseFullScreen))
    Call iniMan.ChangeValue("GraphicsEngine", "UseVerticalSync", BooleanToNumber(GameConfig.Graphics.bUseVerticalSync))
    Call iniMan.ChangeValue("GraphicsEngine", "UseCompatibleMode", BooleanToNumber(GameConfig.Graphics.bUseCompatibleMode))

    Call iniMan.ChangeValue("GraphicsEngine", "EnableAmbientLights", BooleanToNumber(GameConfig.Graphics.Ambience.bAmbientLightsEnabled))
    Call iniMan.ChangeValue("GraphicsEngine", "EnableLights", BooleanToNumber(GameConfig.Graphics.Ambience.bLightsEnabled))
    Call iniMan.ChangeValue("GraphicsEngine", "UseRainWithParticles", BooleanToNumber(GameConfig.Graphics.Ambience.bUseRainWithParticles))
        
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveGraphicsConfig de mod_Configuration.bas")
End Sub
    
Private Sub SaveExtrasConfig(ByVal iniMan As IniFile)
    On Error GoTo ErrHandler
        
    Call iniMan.ChangeValue("Extras", "Name", GameConfig.Extras.Name)
    Call iniMan.ChangeValue("Extras", "NameStyle", GameConfig.Extras.NameStyle)
    Call iniMan.ChangeValue("Extras", "RightClickEnabled", BooleanToNumber(GameConfig.Extras.bRightClickEnabled))
    Call iniMan.ChangeValue("Extras", "AskForResolutionChange", BooleanToNumber(GameConfig.Extras.bAskForResolutionChange))
        
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveExtrasConfig de mod_Configuration.bas")
End Sub
    
Private Sub SaveGuildConfig(ByVal iniMan As IniFile)
    On Error GoTo ErrHandler
                
    Call iniMan.ChangeValue("Guild", "ShowDialogsInConsole", BooleanToNumber(GameConfig.Guilds.bShowDialogsInConsole))
    Call iniMan.ChangeValue("Guild", "ShowGuildNews", BooleanToNumber(GameConfig.Guilds.bShowGuildNews))
    Call iniMan.ChangeValue("Guild", "MaxMessageQuantity", GameConfig.Guilds.MaxMessageQuantity)
        
    Exit Sub
    
ErrHandler:
    Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub SaveGuildConfig de mod_Configuration.bas")
End Sub
    
Private Sub SaveSoundsConfig(ByVal iniMan As IniFile)
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
   
Public Function BooleanToNumber(ByVal boolValue As Boolean) As Byte
    BooleanToNumber = IIf(boolValue = True, 1, 0)
End Function
    
Public Sub SaveGameConfig()
    On Error GoTo ErrHandler
    
    Dim sPath As String
    Dim oFile As Integer
    Dim iniMan As IniFile
    Set iniMan = New IniFile
        
    If Not FileExist(App.path & INIT_PATH, vbDirectory) Then
        Call MkDir(App.path & INIT_PATH)
    End If
    
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

