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

    frmAOSetup.chkPantallaCompleta.Value = BooleanToNumber(GameConfig.Graphics.bUseFullScreen)
        
    frmAOSetup.chkCompatible.Value = BooleanToNumber(GameConfig.Graphics.bUseCompatibleMode)
    
    frmAOSetup.chkVSync.Value = BooleanToNumber(GameConfig.Graphics.bUseVerticalSync)
    
    frmAOSetup.chkMusica.Value = BooleanToNumber(GameConfig.Sounds.bMusicEnabled)
    
    frmAOSetup.chkSonido.Value = BooleanToNumber(GameConfig.Sounds.bSoundsEnabled)
    
    frmAOSetup.chkEfectos.Value = BooleanToNumber(GameConfig.Sounds.bSoundEffectsEnabled)

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
