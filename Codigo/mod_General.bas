Attribute VB_Name = "mod_General"
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

Public Type DDEXCFG
    vsync As Byte
    api As Byte
    Modo As Byte
    MODO2 As Byte
    memoria As Byte
    isDefferal As Byte
End Type

Public Type tSetupMods
    bDinamic    As Boolean
    byMemory    As Byte
    bUseVideo   As Boolean
    bNoMusic    As Boolean
    bNoSound    As Boolean
    bNoRes      As Boolean ' 24/06/2006 - ^[GS]^
    bNoSoundEffects As Boolean
    sGraficos   As String * 13
    bGuildNews  As Boolean ' 11/19/09
    bDie        As Boolean ' 11/23/09 - FragShooter
    bKill       As Boolean ' 11/23/09 - FragShooter
    byMurderedLevel As Byte ' 11/23/09 - FragShooter
    bActive     As Boolean
    bGldMsgConsole As Boolean
    bCantMsgs   As Byte
    bRightClick As Boolean
    ddexConfig As DDEXCFG
    ddexConfigured As Boolean
    ddexSelectedPlugin As String
End Type

Public ClientConfig As tSetupMods

Public Const SW_SHOWNORMAL As Long = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function FileExist(ByVal File As String, ByVal fileType As VbFileAttribute) As Boolean
'*************************************************
'Author: Ivan Leoni y Fernando Costa
'Last modified: ?/?/?
'Se fija si existe el archivo
'*************************************************
    FileExist = Dir(File, fileType) <> ""
End Function

Public Sub LeerSetup()
'*************************************************
'Author: ^[GS]^
'Last modified: 03/11/10
'11/19/09: Pato - Now is optional show the frmGuildNews form in the client
'*************************************************
On Error Resume Next
    If Not FileExist(App.path & "\INIT\", vbDirectory) Then
        Call MkDir(App.path & "\INIT\")
    End If
    
    Dim handle As Integer
    handle = FreeFile
    
    Open App.path & "\Init\AO.dat" For Binary As handle
        Get handle, , ClientConfig
    Close handle
    
    If ClientConfig.bDinamic Then
        frmAOSetup.chkDinamico.value = True
        frmAOSetup.lCuantoVideo.ForeColor = vbBlack
        frmAOSetup.pMemoria.EnabledSlider = True
        frmAOSetup.pMemoria.picFillColor = &H8080FF
        frmAOSetup.pMemoria.picForeColor = &H80FF80
    Else
        frmAOSetup.chkDinamico.value = False
        frmAOSetup.lCuantoVideo.ForeColor = &H808080
        frmAOSetup.pMemoria.EnabledSlider = False
        frmAOSetup.pMemoria.picFillColor = &H808080
        frmAOSetup.pMemoria.picForeColor = &HC0C0C0
    End If
    
    If ClientConfig.byMemory >= 4 And ClientConfig.byMemory <= 40 Then
        frmAOSetup.pMemoria.value = ClientConfig.byMemory
    End If
    
    frmAOSetup.chkPantallaCompleta.value = Not ClientConfig.bNoRes ' 24/06/2006 - ^[GS]^
    
    frmAOSetup.chkUserVideo = ClientConfig.bUseVideo
    
    frmAOSetup.chkMusica.value = Not ClientConfig.bNoMusic
    
    frmAOSetup.chkSonido.value = Not ClientConfig.bNoSound
    
    frmAOSetup.chkEfectos.value = Not ClientConfig.bNoSoundEffects
    
    If ClientConfig.sGraficos <> vbNullString Then
        If ClientConfig.sGraficos = "Graficos1.ind" Then
            frmAOSetup.optSmall.value = True
        ElseIf ClientConfig.sGraficos = "Graficos2.ind" Then
            frmAOSetup.OptAverage.value = True
        End If
    End If
    
    ClientConfig.bGuildNews = Not ClientConfig.bGuildNews
    
    If ClientConfig.bGuildNews Then
        frmAOSetup.optMostrarNoticias.value = True
        frmAOSetup.optNoMostrar.value = False
    Else
        frmAOSetup.optMostrarNoticias.value = False
        frmAOSetup.optNoMostrar.value = True
    End If
    
    If ClientConfig.bCantMsgs = 0 Then ClientConfig.bCantMsgs = 5

    frmAOSetup.optConsola.value = ClientConfig.bGldMsgConsole
    frmAOSetup.txtCantMsgs.text = ClientConfig.bCantMsgs
End Sub

Public Function LibraryExist(ByVal File As String, ByVal fileType As VbFileAttribute) As Boolean
'*************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last modified: 10/01/07
'Esta funcion chequea en la propia carpeta y en el directorio de windows. Ademas
'llama para que se registren las librerias (Si estan registradas no pasa nada
'igual)
'*************************************************
'Chequeo progresivo a mano, primero se fija en el mismo path
LibraryExist = True

If FileExist(File, fileType) Then
    Shell "regsvr32 /s " & File
    Exit Function
End If

If FileExist("C:\WINDOWS\SYSTEM32\" & File, fileType) Then
    Shell "regsvr32 /s " & File
    Exit Function
End If

Dim fsoObject As FileSystemObject

Set fsoObject = New FileSystemObject

If fsoObject.FileExists(File) Then
    Shell "regsvr32 /s " & File
    
    Set fsoObject = Nothing
    Exit Function
End If

LibraryExist = False
Set fsoObject = Nothing

MsgBox fsoObject.GetAbsolutePathName(vbNullString)
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
