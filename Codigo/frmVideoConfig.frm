VERSION 5.00
Begin VB.Form frmVideoConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Video"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmVideoConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3375
      Begin VB.ComboBox cboApiGrafica 
         Height          =   315
         ItemData        =   "frmVideoConfig.frx":0442
         Left            =   120
         List            =   "frmVideoConfig.frx":0449
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   2895
      End
      Begin VB.ComboBox cboModoVideo 
         Height          =   315
         ItemData        =   "frmVideoConfig.frx":0458
         Left            =   120
         List            =   "frmVideoConfig.frx":0465
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   2895
      End
      Begin VB.ComboBox cboModoVertex 
         Height          =   315
         ItemData        =   "frmVideoConfig.frx":0498
         Left            =   120
         List            =   "frmVideoConfig.frx":04A2
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1920
         Width           =   2895
      End
      Begin VB.ComboBox cboModoMemoria 
         Height          =   315
         ItemData        =   "frmVideoConfig.frx":04BA
         Left            =   120
         List            =   "frmVideoConfig.frx":04C7
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2640
         Width           =   2895
      End
      Begin VB.CheckBox chkVerticalSync 
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Modo de memoria"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Modo vertex"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Modo de video"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Librería Gráfica"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Vertical Sync"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   975
      End
   End
   Begin AOSetup.chameleonButton bCancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   3375
      _ExtentX        =   5953
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmVideoConfig.frx":04EF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmVideoConfig"
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

Private Sub bCancelar_Click()
    Call SaveOptions
    Unload Me
End Sub

Private Sub Form_Load()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 04/06/15
    '*************************************************
    Me.Show
    DoEvents
    Call LoadOptions
End Sub


Private Sub LoadOptions()
    '*************************************************
    'Author: Nightw
    'Last modified: 04/06/15
    '*************************************************
    If setupMod.ddexConfigured = False Then
        cboApiGrafica.ListIndex = 0
        cboModoMemoria.ListIndex = 0
        cboModoVideo.ListIndex = 0
        cboModoVertex.ListIndex = setupMod.ddexConfig.MODO2
        
    
        chkVerticalSync.Value = 1
    Else
        cboApiGrafica.ListIndex = setupMod.ddexConfig.api
        cboModoMemoria.ListIndex = setupMod.ddexConfig.memoria
        cboModoVideo.ListIndex = setupMod.ddexConfig.Modo
        cboModoVertex.ListIndex = setupMod.ddexConfig.MODO2
        
    
        chkVerticalSync.Value = IIf(setupMod.ddexConfig.vsync = 1, 1, 0)
    End If
    
End Sub


Private Sub SaveOptions()
    '*************************************************
    'Author: Nightw
    'Last modified: 04/06/15
    '*************************************************
    setupMod.ddexConfig.api = cboApiGrafica.ListIndex
    setupMod.ddexConfig.memoria = cboModoMemoria.ListIndex
    setupMod.ddexConfig.Modo = cboModoVideo.ListIndex
    setupMod.ddexConfig.MODO2 = cboModoVertex.ListIndex
    setupMod.ddexConfig.vsync = IIf(chkVerticalSync.Value, 1, 0)
    setupMod.ddexConfigured = True
    
End Sub

