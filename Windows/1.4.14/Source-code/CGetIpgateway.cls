VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGetIpgateway"
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
' Module            : CGetIpgateway
' Original idea     : IPhelper
' Date              : 3/3/2004
' Internet          : www.vbrad.com/article.aspx?id=32
'---------------------------------------------------------------------------------------

Option Explicit
Option Base 0

'-----------------------------------------------------------------
'
'récupère les adresses IP des passerelles réseaux
'
'-----------------------------------------------------------------

Private Const WSA_DESCRIPTIONLEN = 256
Private Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
Private Const WSA_SYS_STATUS_LEN = 128
Private Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1


Private Const API_TRUE = 1
Private Const ADAPTER_ADDRESS_DELIMITER = "-"
Private Const STRING_ZERO = "0"


Private Enum MAXLEN
    MAXLEN_IFDESCR = 256
    MAXLEN_PHYSADDR = 8
End Enum


Private Enum WinsockVersion
    SOCKET_VERSION_11 = &H101
    SOCKET_VERSION_22 = &H202
End Enum


Private Type MIB_IPNETROW
    dwIndex As Long
    dwPhysAddrLen As Long
    bPhysAddr(0 To MAXLEN_PHYSADDR - 1) As Byte
    dwAddr As Long
    dwType As Long
End Type


Private Type MIB_IPNETTABLE
    dwNumEntries As Long
    table() As MIB_IPNETROW
End Type

Private Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    imaxsockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type


Private Declare Function GetIpNetTable Lib "iphlpapi.dll" (ByRef pIpNetTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long    'Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAdata As WSADataType) As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAdata As WSADataType) As Long
'Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal source As Long, ByVal Size As Long)
Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
'



'récupère l'index d'une adresse IP
Public Function GetIndexIP(ByVal IP As String) As Long
    Dim AddRow() As MIB_IPNETROW
    Dim i As Integer
    Dim tempIP As String
    Dim tempIndex As Long

    ' Create the network tables class

    ' Get the rows of the IP Address Table
10  On Error GoTo GetIpGateway_Error

20  AddRow = IpNetTable

    ' For each row
30  For i = 0 To UBound(AddRow)
40      tempIP = getascip(AddRow(i).dwAddr)
50      'Debug.Print timer, "tempip = "; tempIP
60      If tempIP = IP Then
70          tempIndex = AddRow(i).dwIndex
            'Debug.Print timer, "temp Index = "; tempIndex
80          Exit For
90      End If
100 Next i

110 If tempIndex > 0 Then
120     GetIndexIP = tempIndex
130 Else
140     GetIndexIP = -1
150 End If

160 On Error GoTo 0
170 Exit Function

GetIpGateway_Error:
180 GetIndexIP = -1
    'MsgBox "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in GetIpGateway of mGetIpGateway", vbCritical

End Function


'récupère les adresses IP des passerelles réseaux
' l'option permet de récuper toutes la table d'IP plutot que seulement celles des passerelles
Public Function GetIpGateway(Optional FullIpTable As Boolean = False) As Collection
    Dim AddRow() As MIB_IPNETROW
    Dim colIPgateway As New Collection
    Dim i As Integer
    Dim sTempPhysicalAdress As String

    ' Create the network tables class

    ' Get the rows of the IP Address Table
10  On Error GoTo GetIpGateway_Error

20  AddRow = IpNetTable

30  On Error Resume Next
40  For i = 0 To UBound(AddRow)
50      sTempPhysicalAdress = ConvAddress(AddRow(i).bPhysAddr, AddRow(i).dwPhysAddrLen)
60      If Mid$(sTempPhysicalAdress, 1, 2) = "00" Then
70          colIPgateway.Add getascip(AddRow(i).dwAddr)
80      ElseIf FullIpTable = True Then
90          colIPgateway.Add getascip(AddRow(i).dwAddr)
100     End If
110 Next i

120 Set GetIpGateway = colIPgateway

130 On Error GoTo 0
140 Exit Function

GetIpGateway_Error:
150 Set GetIpGateway = Nothing
    'MsgBox "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in GetIpGateway of mGetIpGateway", vbCritical

End Function


Private Function getascip(ByVal inn As Long) As String
' Convert a network address to a application IP address
    Call StartWinsock
    Dim lpStr&, nStr&, retString$
    getascip = "255.255.255.255"
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr <> 0 Then
        nStr = lstrlen(lpStr)
        If nStr > 32 Then nStr = 32
        CopyMemory ByVal retString, ByVal lpStr, nStr
        retString = Left(retString, nStr)
        getascip = retString
    End If
    Call EndWinsock
End Function

Private Function StartWinsock() As Boolean
' Start Winsock
2:     Dim StartupData As WSADataType, RetVal As Long
3:     RetVal = WSAStartup(WinsockVersion.SOCKET_VERSION_22, StartupData)
End Function

Private Function EndWinsock() As Long
' Stop Winsock
2:     Dim RetVal As Long
3:     RetVal = WSACleanup
End Function


Private Property Get IpNetTable() As MIB_IPNETROW()
    Dim lngRet As Long, tableSize As Long
    Dim mIpNetTable As MIB_IPNETTABLE
    Dim byteArr() As Byte
    Dim ArrSize As Long

    lngRet = GetIpNetTable(ByVal 0&, tableSize, API_TRUE)
    If tableSize > 0 Then
        ReDim byteArr(0 To tableSize - 1) As Byte
        lngRet = GetIpNetTable(ByVal VarPtr(byteArr(0)), tableSize, API_TRUE)

        Call CopyMemory(ByVal VarPtr(mIpNetTable.dwNumEntries), ByVal VarPtr(byteArr(0)), Len(mIpNetTable.dwNumEntries))
        If mIpNetTable.dwNumEntries > 0 Then
            ReDim mIpNetTable.table(0 To mIpNetTable.dwNumEntries - 1) As MIB_IPNETROW
            ArrSize = mIpNetTable.dwNumEntries * Len(mIpNetTable.table(0))
            Call CopyMemory(ByVal VarPtr(mIpNetTable.table(0)), ByVal VarPtr(byteArr(4)), ArrSize)
        End If
        Erase byteArr
    End If
    IpNetTable = mIpNetTable.table
End Property

Private Function ConvAddress(ByRef Addr() As Byte, ByVal Length As Long) As String
' This converts a byte array to a MAC address.
    Dim strdef As String, i As Integer

    For i = 0 To Length - 1
        If i = 0 Then
            strdef = strdef & Pad(Hex(Addr(i)))
        Else
            strdef = strdef & ADAPTER_ADDRESS_DELIMITER & Pad(Hex(Addr(i)))
        End If
    Next i
    ConvAddress = strdef
End Function

Private Function Pad(ByVal hexStr As String) As String
' This pads a character with a preceding 0 if it requires one
    If Len(hexStr) = 1 Then
        Pad = STRING_ZERO & hexStr
    ElseIf Len(hexStr) = 2 Then
        Pad = hexStr
    End If
End Function



