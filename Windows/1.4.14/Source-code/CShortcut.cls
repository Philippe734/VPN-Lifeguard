VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'-----------------------------------------------------
'    VPN Lifeguard - Reconnecter son VPN tout en bloquant ses logiciels
'    Copyright 2010 philippe734
'    http://sourceforge.net/projects/vpnlifeguard/
'
'    VPN Lifeguard is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    VPN Lifeguard is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program. If not, write to the
'    Free Software Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
'-----------------------------------------------------

'---------------------------------------------------------------------------------------
' Module    : CShortcut
' Author    : philippe734
' Date      : 22/05/2010
' Purpose   : créer ou supprime un raccourci
'---------------------------------------------------------------------------------------

Option Explicit

'répertoires de windows
'public Enum CSIDL
'    CSIDL_DESKTOP = &H0
'    CSIDL_PROGRAMS = &H2
'    CSIDL_CONTROLS = &H3
'    CSIDL_PRINTERS = &H4
'    CSIDL_PERSONAL = &H5
'    CSIDL_FAVORITES = &H6
'    CSIDL_STARTUP = &H7
'    CSIDL_RECENT = &H8
'    CSIDL_SENDTO = &H9
'    CSIDL_BITBUCKET = &HA
'    CSIDL_STARTMENU = &HB
'    CSIDL_DESKTOPDIRECTORY = &H10
'    CSIDL_DRIVES = &H11
'    CSIDL_NETWORK = &H12
'    CSIDL_NETHOOD = &H13
'    CSIDL_FONTS = &H14
'    CSIDL_TEMPLATES = &H15
'End Enum

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long


Private Function GetSpecialfolder(subCSIDL As Long) As String
    Dim r As Long
    Dim IDL As ITEMIDLIST
    Dim Path As String

    'Get the special folder
    r = SHGetSpecialFolderLocation(100, subCSIDL, IDL)
    If r = 0 Then
        'Create a buffer
        Path = Space$(512)
        'Get the path from the IDList
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        'Remove the unnecessary vbnullchar's
        GetSpecialfolder = Left$(Path, InStr(Path, vbNullChar) - 1)
    Else
        GetSpecialfolder = vbNullString
    End If
End Function

Public Function CreateShortcut(ByVal sFilePath As String, ByVal strLinkName As String, ByVal CSIDL_FolderDestination As Long) As Boolean
    Dim oShell As New WshShell
    Dim oShortCut As WshShortcut
    Dim sDestination As String

    On Error GoTo err:

    sDestination = GetSpecialfolder(CSIDL_FolderDestination)

    Set oShortCut = oShell.CreateShortcut(sDestination & "\" & strLinkName & ".lnk")

    With oShortCut
        .TargetPath = sFilePath
        .Save
    End With

    CreateShortcut = True

fin:
    Set oShell = Nothing
    Set oShortCut = Nothing

    Exit Function

err:
    CreateShortcut = False
    Resume fin
End Function


Public Function DeleteShortcut(ByVal strLinkName As String, ByVal CSIDL_FolderDestination As Long) As Boolean
    Dim sParent As String
    Dim sLinkPath As String

    sParent = GetSpecialfolder(CSIDL_FolderDestination)
    sLinkPath = sParent & "\" & strLinkName & ".lnk"
    'Debug.Print timer, sLinkPath

    On Error Resume Next
    If Dir(sLinkPath) <> vbNullString Then
        Kill sLinkPath
        DoEvents
        DeleteShortcut = True
    Else
        DeleteShortcut = False
    End If
    On Error GoTo 0
End Function






