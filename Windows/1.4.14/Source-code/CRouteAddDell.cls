VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRouteAddDell"
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
' Module    : CRouteAddDell
' Author    : philippe734
' Date      : 22/05/2010
' Purpose   : exécute un script pour
' soit effacer la route d'une adresse IP (del route)
' soit restaurer la route (add route)
' dans l'objectif d'être sûr de passer par le VPN via le scrip del route
'---------------------------------------------------------------------------------------


Option Explicit

Public Function DelRoute(ByVal IP As String) As Boolean
    Dim oShell As WshShell
    Dim i As Byte

10  On Error GoTo DelRoute_Error

    For i = 1 To 2

20      Set oShell = New WshShell

30      oShell.Run "route delete 0.0.0.0 " & IP, WshHide

        DoEvents

40      Set oShell = Nothing

    Next i

50  DelRoute = True

60  On Error GoTo 0
70  Exit Function

DelRoute_Error:

80  Set oShell = Nothing
90  DelRoute = False
    'MsgBox "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in DelRoute of CRouteAddDell", vbCritical

End Function


Public Function AddRoute(ByVal IP As String, ByVal Index As Long) As Boolean
    Dim oShell As WshShell
    Dim sIP As String
    Dim iIndex As Long
    Dim command As String
    Dim i As Byte

10  On Error GoTo AddRoute_Error

    sIP = IP
    iIndex = Index




    '    ' ******** Only my profil / admin
    '    Dim bBug As Boolean
    '    If sIP <> "192.168.1.1" Then bBug = True
    '    If bBug = True Then
    '        MsgBox "Problème avec Add route :" & vbNewLine & _
             '               "IP Passerelle = " & sIP & vbNewLine & _
             '               "Index = " & iIndex & vbNewLine & vbCritical
    '    End If
    '    sIP = "192.168.1.1"
    '     'iIndex = 2
    '    ' ************************************





20  If Index > 0 Then

        For i = 1 To 2

30          Set oShell = New WshShell

40          command = "route add 0.0.0.0 mask 0.0.0.0 " & sIP & " IF " & iIndex

50          oShell.Run command, WshHide

            DoEvents

60          Set oShell = Nothing

        Next i

70      AddRoute = True

80  Else

90      AddRoute = False
100     Debug.Print Timer, "Index de la passerelle = False"

110 End If


120 On Error GoTo 0
130 Exit Function

AddRoute_Error:

140 Set oShell = Nothing
150 AddRoute = False
    'MsgBox "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in AddRoute of CRouteAddDell", vbCritical

End Function

Public Function ShowIPconfig() As Boolean
    Dim oShell As WshShell
    Dim command As String

10  On Error GoTo err

20  Set oShell = New WshShell

    ' exécute un script qui exécute IPCONFIG tout en faisant une pause mais en continuant le programme
30  command = "cmd /c ipconfig & PAUSE"

40  oShell.Run command, 1, False

50  ShowIPconfig = True

60  On Error GoTo 0
70  Exit Function

err:

80  ShowIPconfig = False
90  MsgBox "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in ShowIPconfig of CRouteAddDell", vbCritical

End Function

