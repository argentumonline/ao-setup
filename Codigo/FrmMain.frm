VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Argentum Online (Setup)"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6660
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   358
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox cEjecutar 
      Appearance      =   0  'Flat
      Caption         =   "Ejecutar el juego al Aceptar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2100
      TabIndex        =   23
      Top             =   4995
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4680
      TabIndex        =   22
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pruebas de Motor"
      Height          =   1095
      Left            =   2640
      TabIndex        =   16
      Top             =   1560
      Width           =   3975
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Aurora.Multimedia"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Aurora.Network"
         Top             =   720
         Width           =   1695
      End
      Begin VB.Line Line5 
         X1              =   960
         X2              =   3120
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblMultimediaOK 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EXITO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblNetworkOK 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EXITO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   720
         Width           =   735
      End
      Begin VB.Line Line2 
         X1              =   960
         X2              =   3120
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Noticias del clan"
      Height          =   735
      Left            =   2640
      TabIndex        =   13
      Top             =   2760
      Width           =   3975
      Begin VB.OptionButton optMostrarNoticias 
         Caption         =   "Mostrar noticias al conectarse"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   315
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton optNoMostrar 
         Caption         =   "No mostrarlas"
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Diálogos de clan"
      Height          =   735
      Left            =   2640
      TabIndex        =   8
      Top             =   3960
      Width           =   2895
      Begin VB.OptionButton optConsola 
         Caption         =   "En consola"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "En pantalla"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   450
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtCantMsgs 
         Height          =   285
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   9
         Text            =   "5"
         Top             =   400
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "mensajes"
         Height          =   195
         Left            =   1920
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones de Sonido"
      Height          =   1575
      Left            =   75
      TabIndex        =   4
      Top             =   3120
      Width           =   2490
      Begin VB.CheckBox chkSonido 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "&Sonido Activado"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkMusica 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "&Música Activada"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkEfectos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "&Efectos de sonido"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Opciones de Video"
      Height          =   1575
      Left            =   75
      TabIndex        =   0
      Top             =   1560
      Width           =   2490
      Begin VB.CheckBox chkVSync 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "VSync"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkCompatible 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Compatibilidad"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkPantallaCompleta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Pantalla Completa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   432
      Y1              =   320
      Y2              =   320
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1500
      Left            =   0
      Picture         =   "FrmMain.frx":0442
      Top             =   0
      Width           =   6675
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub btnAceptar_Click()
    ' Sounds
    GameConfig.Sounds.bSoundsEnabled = CBool(Me.chkSonido.Value)
    GameConfig.Sounds.bMusicEnabled = CBool(Me.chkMusica.Value)
    GameConfig.Sounds.bSoundEffectsEnabled = CBool(Me.chkEfectos.Value)
    
    ' Graphics
    GameConfig.Graphics.bUseFullScreen = CBool(Me.chkPantallaCompleta.Value)
    GameConfig.Graphics.bUseVerticalSync = CBool(Me.chkVSync.Value)
    GameConfig.Graphics.bUseCompatibleMode = CBool(Me.chkCompatible.Value)

    GameConfig.Guilds.MaxMessageQuantity = Val(txtCantMsgs.text)

    Call SaveGameConfig

    If cEjecutar.Value = 1 Then
        If FileExist(App.path & "\Argentum.exe", vbArchive) = True Then
            Call Shell(App.path & "\Argentum.exe")
        End If
    End If

    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    On Error Resume Next

    Call Configuration.LoadGameConfig
    Call LeerSetup
    
    Call CheckEngineLibraries
    
End Sub

Private Sub LeerSetup()

    On Error Resume Next

    ' Sounds
    chkPantallaCompleta.Value = BooleanToNumber(GameConfig.Graphics.bUseFullScreen)
    chkCompatible.Value = BooleanToNumber(GameConfig.Graphics.bUseCompatibleMode)
    chkVSync.Value = BooleanToNumber(GameConfig.Graphics.bUseVerticalSync)
    
    ' Graphics
    chkMusica.Value = BooleanToNumber(GameConfig.Sounds.bMusicEnabled)
    chkSonido.Value = BooleanToNumber(GameConfig.Sounds.bSoundsEnabled)
    chkEfectos.Value = BooleanToNumber(GameConfig.Sounds.bSoundEffectsEnabled)

    If GameConfig.Guilds.bShowGuildNews Then
        optMostrarNoticias.Value = True
        optNoMostrar.Value = False
    Else
        optMostrarNoticias.Value = False
        optNoMostrar.Value = True
    End If
    
    If GameConfig.Guilds.MaxMessageQuantity = 0 Then
        GameConfig.Guilds.MaxMessageQuantity = 5
    End If

    optConsola.Value = GameConfig.Guilds.bShowDialogsInConsole
    txtCantMsgs.text = GameConfig.Guilds.MaxMessageQuantity
End Sub

Private Sub CheckEngineLibraries()

    If (Not RegisterDLL("Aurora.Multimedia.DLL", True)) Then
        lblMultimediaOK.Caption = "ERROR"
        lblMultimediaOK.ForeColor = &HFF&
    End If
  
    If (Not RegisterDLL("Aurora.Network.DLL", True)) Then
        lblNetworkOK.Caption = "ERROR"
        lblNetworkOK.ForeColor = &HFF&
    End If

    If (Not IsWindows7SP1OrBetter()) Then
        If (IsWindows7SP0()) Then
            MsgBox "Para poder jugar Argentum Online en Windows 7, necesitas actualizarlo a SP1. A continuacion te abriremos el link de tal update"
            
            Call ShellExecute(0, "open", "https://www.microsoft.com/en-us/download/details.aspx?id=36805", 0, 0, 0)
        Else
            MsgBox "Necesitas al menos Windows 7 SP1 para jugar Argentum Online"
        End If
    End If

End Sub

