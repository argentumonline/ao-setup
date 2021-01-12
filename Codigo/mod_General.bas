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
'Last modified: 30/12/2017
'11/19/09: Pato - Now is optional show the frmGuildNews form in the client
'*************************************************
On Error Resume Next

    'Check if the old config file exists
    If OldConfigExists() Then
        ' Migrate the old config file to the new format
        'Call MigrateOldConfigFormat
        'Remove the old file
       ' Call RemoveOldConfigFile
    End If

    
    frmAOSetup.chkPantallaCompleta.Value = GameConfig.Graphics.bUseFullScreen ' 24/06/2006 - ^[GS]^
        
    frmAOSetup.chkCompatible.Value = GameConfig.Graphics.bUseCompatibleMode
    
    frmAOSetup.chkVSync.Value = GameConfig.Graphics.bUseVerticalSync
    
    frmAOSetup.chkMusica.Value = GameConfig.Sounds.bMusicEnabled
    
    frmAOSetup.chkSonido.Value = GameConfig.Sounds.bSoundsEnabled
    
    frmAOSetup.chkEfectos.Value = GameConfig.Sounds.bSoundEffectsEnabled
    
    If GameConfig.Graphics.GraphicsIndToUse <> vbNullString Then
        If GameConfig.Graphics.GraphicsIndToUse = "Graficos1.ind" Then
            frmAOSetup.optSmall.Value = True
        ElseIf GameConfig.Graphics.GraphicsIndToUse = "Graficos2.ind" Then
            frmAOSetup.OptAverage.Value = True
        End If
    End If
    
    If GameConfig.Guilds.bShowGuildNews Then
        frmAOSetup.optMostrarNoticias.Value = True
        frmAOSetup.optNoMostrar.Value = False
    Else
        frmAOSetup.optMostrarNoticias.Value = False
        frmAOSetup.optNoMostrar.Value = True
    End If
    
    If GameConfig.Guilds.MaxMessageQuantity = 0 Then GameConfig.Guilds.MaxMessageQuantity = 5

    frmAOSetup.optConsola.Value = GameConfig.Guilds.bShowDialogsInConsole
    frmAOSetup.txtCantMsgs.text = GameConfig.Guilds.MaxMessageQuantity
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
    
    Debug.Print (errStr)
    
    path = App.path & "\Errores" & Year(Now) & Month(Now) & Day(Now) & ".log"
    oFile = FreeFile
    
    Open path For Append As #oFile
        Print #oFile, Time & " - " & errStr
    Close #oFile
  
    
  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub LogError de General.bas")
End Sub
