Attribute VB_Name = "modPingICMP2"

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
' Module     : modPingICMP2
' Author     : KPD-Team
' Date       : 2000
' Internet   : http://www.allapi.net/
'---------------------------------------------------------------------------------------

Option Explicit

Const SOCKET_ERROR = 0
Private Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    imaxsockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Private Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Private Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type
Private Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    Data As Long
    Options As IP_OPTION_INFORMATION
End Type

Private Declare Function GetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal HostName As String) As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAdata As WSAdata) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
'private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Private Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean

Public Function DoPing2(HostName As String, Optional sResultat As String) As Boolean
'KPD-Team 2000
'URL: http://www.allapi.net/
'E-Mail: KPDTeam@Allapi.net
'Const HostName = "192.168.1.23"
    Dim hfile As Long, lpWSAdata As WSAdata
    Dim hHostent As Hostent, AddrList As Long
    Dim Address As Long, rIP As String
    Dim OptInfo As IP_OPTION_INFORMATION
    Dim EchoReply As IP_ECHO_REPLY
    Const iTimeOut As Integer = 200    'ms

    On Error GoTo err:


10  Call WSAStartup(&H101, lpWSAdata)

20  If GetHostByName(HostName + String(64 - Len(HostName), 0)) <> SOCKET_ERROR Then
30      CopyMemory hHostent.h_name, ByVal GetHostByName(HostName + String(64 - Len(HostName), 0)), Len(hHostent)
40      CopyMemory AddrList, ByVal hHostent.h_addr_list, 4
50      CopyMemory Address, ByVal AddrList, 4
60  End If

70  hfile = IcmpCreateFile()
80  If hfile = 0 Then
90      MsgBox MLSGetString("0053") ' MLS-> "Impossible de créer le Handle du Ping"
100     Exit Function
110 End If
120 OptInfo.TTL = 255
130 If IcmpSendEcho(hfile, Address, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, iTimeOut) Then
140     rIP = CStr(EchoReply.Address(0)) + "." + CStr(EchoReply.Address(1)) + "." + CStr(EchoReply.Address(2)) + "." + CStr(EchoReply.Address(3))
150 Else
        ' Timeout
        err.Raise 1234, "Time out", "Ping failed..."
160 End If
170 If EchoReply.Status = 0 Then
        'Ping OK
        'MsgBox "Reply from " + HostName + " (" + rIP + ") recieved after " + Trim$(CStr(EchoReply.RoundTripTime)) + "ms"
180     sResultat = rIP
190     DoPing2 = True
200 Else
        'MsgBox "Failure ..."
210     DoPing2 = False
220 End If
230 Call IcmpCloseHandle(hfile)
240 Call WSACleanup
250 Exit Function

err:
    DoPing2 = False
    err.Clear
270 Call IcmpCloseHandle(hfile)
280 Call WSACleanup
End Function



