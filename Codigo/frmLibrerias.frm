VERSION 5.00
Begin VB.Form frmLibrerias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Librerias"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   Icon            =   "frmLibrerias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox LibName 
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
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "MSCOMCTL.OCX"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox LibName 
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
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "RICHTX32.OCX"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox LibName 
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
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "AAM532.DLL"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox LibName 
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
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "MSINET.OCX"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuraci�n de Proxy para Descargas"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   3255
      Begin VB.TextBox txtProxy 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox ChkProxy 
         Caption         =   "Usar servidor"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin AOSetup.chameleonButton bCancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   3255
      _extentx        =   5741
      _extenty        =   661
      btype           =   3
      tx              =   "&Aceptar"
      enab            =   -1  'True
      font            =   "frmLibrerias.frx":0442
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   12648384
      bcolo           =   12648384
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmLibrerias.frx":046E
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin AOSetup.chameleonButton cSolucion 
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      btype           =   3
      tx              =   ""
      enab            =   -1  'True
      font            =   "frmLibrerias.frx":048C
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   12648384
      bcolo           =   12648384
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmLibrerias.frx":04B8
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin AOSetup.chameleonButton cVerificar 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   3255
      _extentx        =   5741
      _extenty        =   661
      btype           =   3
      tx              =   "&Verificar nuevamente"
      enab            =   -1  'True
      font            =   "frmLibrerias.frx":04D6
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   12648384
      bcolo           =   12648384
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmLibrerias.frx":0502
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin AOSetup.chameleonButton cSolucion 
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      btype           =   3
      tx              =   ""
      enab            =   -1  'True
      font            =   "frmLibrerias.frx":0520
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   12648384
      bcolo           =   12648384
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmLibrerias.frx":054C
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin AOSetup.chameleonButton cSolucion 
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      btype           =   3
      tx              =   ""
      enab            =   -1  'True
      font            =   "frmLibrerias.frx":056A
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   12648384
      bcolo           =   12648384
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmLibrerias.frx":0596
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin AOSetup.chameleonButton cSolucion 
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      btype           =   3
      tx              =   ""
      enab            =   -1  'True
      font            =   "frmLibrerias.frx":05B4
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   12648384
      bcolo           =   12648384
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmLibrerias.frx":05E0
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Label lblOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   1560
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   2280
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1560
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lblOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lblOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   2280
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2280
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   2280
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmLibrerias"
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

Private Const URL_DOWNLOAD As String = "http://argentumonline.3dgames.com.ar/autoupdate/"
' MD5 de Referencia:
' RICHTX32.OCX 722435ba4d18f1704b43e823a12e489a
' CSWSK32.OCX 5181704b2772e050e4a8331e15ee4bb4
' MSINET.OCX 40d81470a19269d88bf44e766be7f84a
' MSWINSCK.OCX 3d8fd62d17a44221e07d5c535950449b

Private Const MD5_1 As String = "cefd956a1ef122cda4d53007bab6c694"
Private Const MD5_2 As String = "045a16822822426c305ea7280270a3d6"
Private Const MD5_3 As String = "5181704b2772e050e4a8331e15ee4bb4"

Public descargando As Boolean


Private Sub bCancelar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
Unload Me
End Sub

Sub LibError(ByVal index As Byte, ByVal Solucion As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
    lblOK(index).Caption = "ERROR"
    lblOK(index).ForeColor = RGB(255, 0, 0)
    lblOK(index).Visible = True
    cSolucion(index).Caption = Solucion
    cSolucion(index).Visible = True
    LibName(index).BackColor = lblOK(index).ForeColor
End Sub

Sub LibOK(ByVal index As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
    lblOK(index).Caption = "OK"
    lblOK(index).ForeColor = &H8000&
    lblOK(index).Visible = True
    cSolucion(index).Visible = False
    LibName(index).BackColor = lblOK(index).ForeColor
End Sub

Private Sub cSolucion_Click(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 10/01/07
'Last Modified by: Lucas Tavolaro Ortiz (Tavo)
'De ahora en mas se utiliza la funcion LibraryExist()
'*************************************************
    Dim sName As String
    
    Select Case index
        Case 0  ' inet
            sName = "MSINET.OCX"
        
        Case 1 'AA
            sName = "aamd532.dll"
            
        Case 2  ' Rich
            sName = "RICHTX32.OCX"
            
        Case 3 'MSCOMCTL
            sName = "MSCOMCTL.OCX"
    End Select
        
    If cSolucion(index).Caption = "&Registrar" Then
        ' registrar
        Dim fsoObject As FileSystemObject
        
        If Not LibraryExist(sName, vbNormal) Then
            MsgBox "ERROR, el archivo " & sName & " descargado tiene que ser copiado a este directorio.", vbCritical, "Argentum Online Setup"
        Else
            Set fsoObject = New FileSystemObject
            
            fsoObject.CopyFile sName, fsoObject.GetSpecialFolder(SystemFolder) & "\", True
            If Err Then MsgBox Err.Description
            Shell "regsvr32 /s " & fsoObject.GetSpecialFolder(SystemFolder) & sName
            MsgBox "Copia y registro realizados con �xito.", vbOKOnly, "Argentum Online Setup"
        
            Set fsoObject = Nothing
        End If
        
        DoEvents
        Call cVerificar_Click
    Else
        ' descargar
        If descargando = True Then
            MsgBox "Debes esperar a que se termine la descarga actual", vbCritical
            Exit Sub
        End If
        
        Dim rta As VbMsgBoxResult
        
        If index = 0 Then 'El inet es un caso especial, si no lo tenemos es medio dif�cil usarlo para bajarse a si mismo :P
            rta = MsgBox("Necesita descargar el archivo " & sName & "." & vbCrLf & _
                "Es necesario que este archivo sea descargando manualmente y colocado en el directorio del juego, si esta de acuerdo presione S�", vbInformation + vbYesNo, "Soluci�n al problema")
            
            If rta = vbYes Then
                Call ShellExecute(hwnd, "open", URL_DOWNLOAD & sName, vbNullString, vbNullString, SW_SHOWNORMAL)
            End If
        Else
            rta = MsgBox("Necesita descargar el archivo " & sName & "." & vbCrLf & _
                "Si desea descargarlo y registrarlo automaticamente precione Si.", vbYesNo, "Soluci�n al problema")
            
            If rta = vbYes Then
                'Bajarlo
                descargando = True
                
                If ChkProxy.Value = 1 Then
                    Call DownloadForm.DownloadFile(URL_DOWNLOAD & sName, sName, , , 2, txtProxy.text)
                Else
                    Call DownloadForm.DownloadFile(URL_DOWNLOAD & sName, sName)
                End If
                
                If (Not DownloadForm.DownloadSuccess) Or (DownloadForm.BotonCancel = True) Then
                   Unload DownloadForm
                   MsgBox "Descarga cancelada", vbInformation, "Error no solucionado"
                   Exit Sub
                Else
                   Unload DownloadForm
                End If
                
                descargando = False
                
                If FileExist(sName, vbNormal) Then
                    If mod_MD5.MD5File(sName) <> getMD5OriginalFile(index) Then
                        MsgBox "No se puede comprobar la originalidad del archivo descargado, no se instalara.", vbCritical, "Error en MD5"
                        Exit Sub
                    Else
                        DoEvents
                        Call cVerificar_Click
                    End If
                Else
                    MsgBox "No se pudo descargar el archivo", vbInformation, "Falta archivo"
                End If
            End If
        End If
    End If
End Sub

Private Sub cVerificar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 04/11/08
'Last Modified by: NicoNZ
'Busca la existencia de la libreria "mscomctl.ocx"
'*************************************************
On Error Resume Next
    Err.Clear
    
    Load DownloadForm
    If Err Then
        If Not LibraryExist("mscomctl.ocx", vbNormal) Then
            Call LibError(5, "&Explorar")
        Else
            Call LibError(5, "&Registrar")
        End If
    Else
        Call LibOK(5)
    End If
            
    Err.Clear

    If Err Then
        If Not LibraryExist("msinet.ocx", vbNormal) Then
            Call LibError(0, "&Explorar")
        Else
            Call LibError(0, "&Registrar")
        End If
    Else
        Call LibOK(0)
    End If
    
    If Not LibraryExist("aamd532.dll", vbNormal) Then
        Call LibError(1, "&Descargar")
    Else
        Call LibOK(1)
    End If
    
    Err.Clear

    If Err Then
        If Not LibraryExist("richtx32.ocx", vbNormal) Then
            Call LibError(2, "&Descargar")
        Else
            Call LibError(2, "&Registrar")
        End If
    Else
        Call LibOK(2)
    End If
    
End Sub

Private Sub Form_Load()
'*************************************************
'Author: ^[GS]^
'Last modified: 10/03/06
'*************************************************
Me.Show
DoEvents
Call cVerificar_Click
End Sub

Private Function getMD5OriginalFile(ByVal index As Byte) As String
Select Case index
    Case 1
        getMD5OriginalFile = MD5_1
    Case 2
        getMD5OriginalFile = MD5_2
    Case 3
        getMD5OriginalFile = MD5_3
End Select
End Function
