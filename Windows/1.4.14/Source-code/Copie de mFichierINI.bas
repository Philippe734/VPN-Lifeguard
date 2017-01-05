Attribute VB_Name = "modFichierINI"

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
' Module     : modFichierINI
' Author     : Nix
' Date       : 15/05/1999
' Internet   : http://www.vbfrance.com/codes/LIRE-ECRIRE-DANS-FICHIER-INI_32.aspx
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'


' Pour l'executer ex :
'EcrireINI("MonEntete", "MaVariable", "MaValeur")
'LireINI("MonEntete", "MaVariable")


Public Function LireINI(Entete As String, Variable As String, Optional Fichier As String = vbNullString) As String
    Dim Retour As String

    On Error Resume Next

    If Fichier = vbNullString Then
        Fichier = App.Path & "\" & App.Title & ".ini"
    End If
    Retour = String(255, Chr(0))
    LireINI = Left$(Retour, GetPrivateProfileString(Entete, ByVal Variable, "", Retour, Len(Retour), Fichier))
    
    On Error GoTo 0
End Function

Public Function EcrireINI(Entete As String, Variable As String, Valeur As String) As String
    Dim Fichier As String

    On Error Resume Next

    Fichier = App.Path & "\" & App.Title & ".ini"
    EcrireINI = WritePrivateProfileString(Entete, Variable, Valeur, Fichier)
    
    On Error GoTo 0
End Function

' Log simpliste pour faire des tests
Public Function oldEcrireLog(ByVal Valeur As String) As String
    Dim Fichier As String
    Dim sNow As String

    On Error Resume Next

    Fichier = App.Path & "\" & "TheLog" & ".log"
    sNow = Now
    oldEcrireLog = WritePrivateProfileString("Log info", sNow, Valeur, Fichier)
    DoEvents
    
    On Error GoTo 0
End Function

' Log
Public Function EcrireLog(ByVal NomCourtFichierLog As String, ByVal Valeur As String) As String
    Dim Fichier As String
    Dim sNow As String
    On Error Resume Next

    ' l'entete est la date et l'heure
    sNow = Now
    
    Fichier = App.Path & "\" & NomCourtFichierLog & ".log"
    
    EcrireLog = WritePrivateProfileString("Log info", sNow, Valeur, Fichier)
    
    DoEvents
    On Error GoTo 0
End Function

