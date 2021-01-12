VERSION 5.00
Begin VB.Form frmAOSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Argentum Online Setup"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "frmAOSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Di�logos de clan"
      Height          =   735
      Left            =   4200
      TabIndex        =   30
      Top             =   5760
      Width           =   2535
      Begin VB.TextBox txtCantMsgs 
         Height          =   285
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   33
         Text            =   "5"
         Top             =   400
         Width           =   375
      End
      Begin VB.OptionButton optPantalla 
         Caption         =   "En pantalla,"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   450
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optConsola 
         Caption         =   "En consola"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "mensajes"
         Height          =   195
         Left            =   1750
         TabIndex        =   34
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Noticias del clan"
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   5760
      Width           =   3975
      Begin VB.OptionButton optNoMostrar 
         Caption         =   "No mostrarlas"
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   315
         Width           =   1275
      End
      Begin VB.OptionButton optMostrarNoticias 
         Caption         =   "Mostrar noticias al conectarse"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   315
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tipo de Arboles"
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   5040
      Width           =   6615
      Begin VB.OptionButton optBig 
         Caption         =   "Grandes"
         Height          =   255
         Left            =   5400
         TabIndex        =   26
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptAverage 
         Caption         =   "Medianos"
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optSmall 
         Caption         =   "Peque�os"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pruebas de DirectX"
      Height          =   3270
      Left            =   2640
      TabIndex        =   5
      Top             =   1680
      Width           =   4095
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
         TabIndex        =   10
         Text            =   "DirectX 7"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text3 
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
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "DirectSound"
         Top             =   900
         Width           =   1215
      End
      Begin VB.PictureBox fondoVersion 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   3795
         TabIndex        =   6
         Top             =   2715
         Width           =   3855
         Begin VB.Label lVersionFondo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Versi�n detectada:"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   45
            Width           =   1335
         End
         Begin VB.Label lDirectX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   1500
            TabIndex        =   7
            Top             =   45
            Width           =   135
         End
      End
      Begin AOSetup.chameleonButton bProbarSonido 
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "S&onido"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAOSetup.frx":0442
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame4 
         Caption         =   "Probar"
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label lblDX 
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblDS 
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   940
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   2280
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   2280
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin VB.CheckBox cEjecutar 
      Appearance      =   0  'Flat
      Caption         =   "Ejecutar el juego al Aceptar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   7200
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones de Sonido"
      Height          =   1350
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   2415
      Begin VB.CheckBox chkEfectos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "&Efectos de sonido"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   810
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkMusica 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "&M�sica Activada"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   525
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkSonido 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "&Sonido Activado"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Video"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   2415
      Begin VB.CheckBox chkPantallaCompleta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Pantalla Completa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkCompatible 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Compatibilidad"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   680
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkVSync 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "VSync"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   970
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin AOSetup.chameleonButton bCancelar 
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   7080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAOSetup.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton bAceptar 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   7080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Aceptar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632319
      BCOLO           =   12632319
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421631
      MPTR            =   1
      MICON           =   "frmAOSetup.frx":047A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton cLibrerias 
      Height          =   375
      Left            =   105
      TabIndex        =   15
      Top             =   6540
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Verificar &Librerias"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAOSetup.frx":0496
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AOSetup.chameleonButton cCreditos 
      Height          =   255
      Left            =   6360
      TabIndex        =   20
      ToolTipText     =   "Creditos"
      Top             =   1320
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BTYPE           =   5
      TX              =   "?"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAOSetup.frx":04B2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6720
      Y1              =   6975
      Y2              =   6975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6720
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1500
      Left            =   120
      Picture         =   "frmAOSetup.frx":04CE
      Top             =   120
      Width           =   6675
   End
End
Attribute VB_Name = "frmAOSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit

' sonido
Dim m_dsBuffer As DirectSoundBuffer
Dim m_bLoaded As Boolean
' video
Private Const SW_SHOWNORMAL = 1
Dim Primary As DirectDrawSurface7
Dim BackBuffer As DirectDrawSurface7
Dim Clipper As DirectDrawClipper
Dim ddsCharacter As DirectDrawSurface7
Dim ddsd As DDSURFACEDESC2
Dim ddsdback As DDSURFACEDESC2
Dim destRect As RECT
Dim srcRect As RECT
Dim chanRect As RECT
Dim CharWidth As Integer
Dim CharHight As Integer
Dim PostionX As Integer
Dim postionY As Integer
Dim running As Boolean

Private Sub bAceptar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 11/19/09
'11/19/09: Pato - Now is optional show the frmGuildNews form in the client
'*************************************************
    Dim sFile As String
    
    ' Sounds
    GameConfig.Sounds.bSoundsEnabled = CBool(Me.chkSonido.Value)
    GameConfig.Sounds.bMusicEnabled = CBool(Me.chkMusica.Value)
    GameConfig.Sounds.bSoundEffectsEnabled = CBool(Me.chkEfectos.Value)
    
    ' Graphics
    GameConfig.Graphics.bUseFullScreen = CBool(Me.chkPantallaCompleta.Value)
    GameConfig.Graphics.bUseVerticalSync = CBool(Me.chkVSync.Value)
    GameConfig.Graphics.bUseCompatibleMode = CBool(Me.chkCompatible.Value)
   
    If optBig.Value Then
        sFile = "Graficos3.ind"
    ElseIf OptAverage.Value Then
        sFile = "Graficos2.ind"
    Else
        sFile = "Graficos1.ind"
    End If
    
    GameConfig.Graphics.GraphicsIndToUse = sFile
    
    GameConfig.Guilds.MaxMessageQuantity = Val(txtCantMsgs.text)
    
    DoEvents
    
    Call SaveGameConfig
    
    'Dim handle As Integer
    'handle = FreeFile
    'Open App.path & "\Init\AO.DAT" For Binary As handle
    '    Put handle, , ClientConfig
    'Close handle
    'DoEvents
    
    If cEjecutar.Value = 1 Then
        If FileExist(App.path & "\Argentum.exe", vbArchive) = True Then _
            Call Shell(App.path & "\Argentum.exe")
        DoEvents
    End If
    
    Unload Me
End Sub

Private Sub bCancelar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
    Unload Me
End Sub

Private Sub bProbarSonido_Click()
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: 24/06/06
'10/03/06: ^[GS]^ - Agregue una revision de la existencia del archivo de sonido
'24/06/06: ^[GS]^ - Una correccion mas mejorada.
'*************************************************
On Error Resume Next
    
    If bProbarSonido.Value = True Then
        ' [GS]
        Dim sonido As String
        sonido = App.path & "\wav\18.wav"
        
        If FileExist(sonido, vbArchive) = False Then
            MsgBox "No se puede probar el sonido porque falta el archivo de pruebas.", vbCritical
            bProbarSonido.Value = False ' 24/06/06 - ^[GS]^
            Exit Sub
        End If
        ' [/GS]
        
        DirectSound.SetCooperativeLevel Me.hwnd, DSSCL_NORMAL
        
        If m_bLoaded = False Then
            m_bLoaded = True
            LoadWave 0, sonido
        End If
        Dim flag As Long
        flag = 0
        m_dsBuffer.Play flag
        
        If Err.Number <> 0 Then
            MsgBox "Problemas de DirectSound, Reinstale DIRECTX.", vbOKOnly, "Argentum Online Setup"
        End If
    Else
        If m_dsBuffer Is Nothing Then Exit Sub
        m_dsBuffer.Stop
        m_dsBuffer.SetCurrentPosition 0
    End If
End Sub

Sub LoadWave(i As Integer, sFile As String)
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: 10/03/06
'10/03/06: ^[GS]^ - Borre un codigo al final que no se utilizaba
'*************************************************

    Dim bufferDesc As DSBUFFERDESC  'a new object that when filled in is passed to the DS object to describe
    Dim waveFormat As WAVEFORMATEX
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2    '2 channels
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16  '16 bit rather than 8 bit
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    Set m_dsBuffer = DirectSound.CreateSoundBufferFromFile(sFile, bufferDesc, waveFormat)
    
    If Err.Number <> 0 Then
        MsgBox "Error en " + sFile
        End
    End If
End Sub

Private Sub cCreditos_Click()
'*************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last modified: 15/03/06
'*************************************************
    frmAbout.Show vbModal, Me
End Sub

Private Sub chkCompatible_Click()
    GameConfig.Graphics.bUseCompatibleMode = CBool(chkCompatible.Value)
End Sub

Private Sub chkPantallaCompleta_Click()
    GameConfig.Graphics.bUseFullScreen = CBool(chkPantallaCompleta.Value)
End Sub

Private Sub chkVSync_Click()
    GameConfig.Graphics.bUseVerticalSync = CBool(chkVSync.Value)
End Sub

Private Sub cLibrerias_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
frmLibrerias.Show
End Sub

Private Sub Form_Load()
'*************************************************
'Author: ^[GS]^
'Last modified: 09/27/2010
'History:
' 09/27/2010: C4b3z0n - Ahora la version del directx la brinda otra funcion y asignamos la version del DirectX al label, en vez de acceder desde el sub al label.
' 10/03/06: Era el last modified anterior (^[GS]^).
'*************************************************
On Error Resume Next
    Me.Show
    
    DoEvents
    
    Call mod_Configuration.LoadGameConfig
    Call LeerSetup
    'Call mod_GameIni.LoadUserConfig

    
    Call mod_DirectX.ProbarDirectX
    lDirectX.Caption = mod_DirectX.GetVersion()
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last modified: 11/03/06
'*************************************************
    If FileExist("C:\DXTest.txt", vbArchive) Then _
        Kill "C:\DXTest.txt"
    End
End Sub

Private Sub optConsola_Click()
    GameConfig.Guilds.bShowDialogsInConsole = True
End Sub

Private Sub optMostrarNoticias_Click()
    GameConfig.Guilds.bShowGuildNews = True
End Sub

Private Sub optNoMostrar_Click()
    GameConfig.Guilds.bShowGuildNews = False
End Sub

Private Sub optPantalla_Click()
    GameConfig.Guilds.bShowDialogsInConsole = False
End Sub

