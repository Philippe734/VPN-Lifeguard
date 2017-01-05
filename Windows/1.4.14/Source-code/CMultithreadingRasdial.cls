VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMultithreadingRasdial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

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
' Module    : CMultithreadingRasdial
' Author    : philippe734
' Date      : 22/03/2010
' Purpose   : effectue une connexion RasDial
'---------------------------------------------------------------------------------------

Option Explicit

Public Event Result(ByVal HandleRasConn As Long, ByVal ReturnCode As Long)
Public Event Error(ByVal Message As String)


Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, ByVal lpString2 As String) As Long
Private Declare Function RasGetEntryDialParams Lib "rasapi32.dll" Alias "RasGetEntryDialParamsA" (ByVal lpcstr As String, ByRef lprasdialparamsa As Any, ByRef lpbool As Long) As Long
Private Declare Function RasDial Lib "rasapi32.dll" Alias "RasDialA" (lpRasDialExtensions As Any, ByVal lpszPhonebook As String, lprasdialparams As Any, ByVal dwNotifierType As Long, lpvNotifier As Long, lphRasConn As Long) As Long
'

Public Sub RasDialThreaded(ByVal iThreadIndex As Long, ByVal ConnName As String)
    Dim l_hRasConn As Long
    Dim l_RetCode As Long
    Dim l_iThreadIndex As Long
    Dim msg As String

10  On Error GoTo RasDialThreaded_Error

20  ptRasdialParamsThreaded.dwSize = LenB(ptRasdialParamsThreaded)
30  lstrcpy ptRasdialParamsThreaded.szEntryName(0), ConnName
40  RasGetEntryDialParams vbNullString, ptRasdialParamsThreaded, 0

    'synchronous dial
50  l_RetCode = RasDial(ByVal &H0, vbNullString, ptRasdialParamsThreaded, &H0, ByVal &H0, l_hRasConn)

    ' si on veux récupèrer l'index de la thread
    l_iThreadIndex = iThreadIndex

    'envoie les résultas
60  RaiseEvent Result(l_hRasConn, l_RetCode)


70  On Error GoTo 0
80  Exit Sub

RasDialThreaded_Error:

90  msg = "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in RasDialThreaded of CMultithreadingRasdial"
100 RaiseEvent Error(msg)

End Sub

