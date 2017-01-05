VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGetOthersIp"
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

'
'récupère les adresses IP locales autres que localhost, tel qu'un vpn
'

Private Const API_TRUE = 1
Private Const WSA_DESCRIPTIONLEN = 256
Private Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
Private Const WSA_SYS_STATUS_LEN = 128
Private Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Private Enum WinsockVersion
    SOCKET_VERSION_11 = &H101
    SOCKET_VERSION_22 = &H202
End Enum

Private Type MIB_IPADDRROW
    dwAddr As Long
    dwIndex As Long
    dwMask As Long
    dwBCastAddr As Long
    dwReasmSize As Long
    unused1 As Integer
    unused2 As Integer
End Type

Private Type MIB_IPADDRTABLE
    dwNumEntries As Long
    table() As MIB_IPADDRROW
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

Private Declare Function GetIpAddrTable Lib "iphlpapi.dll" (ByRef pIpAddrTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long

Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long

Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long

'Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

'private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal source As Long, ByVal Size As Long)

Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
'

'récupère les adresses IP locales autres que localhost, tel qu'un vpn
Public Function Get_Locals_IPs() As Collection

    Dim AddRow() As MIB_IPADDRROW

    'tableau des adresses IP
    'max lignes avec 2 colonnes
    'colonne 0 = adresses
    'colonne 1 = index
    Dim subIPtables()

    Dim sTempIpAdress As String
    Dim col As New Collection
    Dim i As Integer
    Dim r As Byte    'ligne du tableau subIPtables

    ' Get the rows of the IP Address Table
10  On Error GoTo Get_Locals_IPs_Error

    ' Create the network tables class
20  AddRow = IpAddressTable

30  ReDim subIPtables(0 To 2, 0 To UBound(AddRow))

    ' For each row
40  r = 0
50  For i = 0 To UBound(AddRow)
60      sTempIpAdress = getascip(AddRow(i).dwAddr)
70      If sTempIpAdress <> "127.0.0.1" And sTempIpAdress <> "0.0.0.0" Then
80          subIPtables(0, r) = sTempIpAdress
90          subIPtables(1, r) = AddRow(i).dwIndex
            'ligne suivante
100         r = r + 1
110     End If
120 Next i

    'prépare le tableau pour être trié = une ligne et une colonne en +
130 ReDim Preserve subIPtables(0 To 2, 0 To r)

    Dim cls As CSortArray2Col
140 Set cls = New CSortArray2Col
    Dim ret As Boolean

    '-----------------------------------------------------
    '150 Debug.Print timer, Time
    '160 Debug.Print timer, "avant sort :"
    '170 For i = 0 To UBound(subIPtables, 2) - 1
    '180     Debug.Print timer, "subIPtables(0,i) = "; subIPtables(0, i)
    '190 Next i
    '-----------------------------------------------------

    'tri du tableau suivant l'index des adresses = suivant la colonne 1
200 ret = cls.SortArray(subIPtables, 1)
210 Set cls = Nothing

    '220 Debug.Print timer, "après sort :"
    'boucle pour remplir la collection
230 For i = 0 To UBound(subIPtables, 2) - 1
240     col.Add subIPtables(0, i)
250     'Debug.Print timer, "subIPtables(0,i) = "; subIPtables(0, i)
260 Next i

270 Set Get_Locals_IPs = col

280 On Error GoTo 0
290 Exit Function

Get_Locals_IPs_Error:

    Set cls = Nothing
    'MsgBox "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in Get_Locals_IPs of CGet_Locals_IPs", vbCritical

End Function

Public Function GetIndexIP(ByVal IP As String) As Long

    Dim AddRow() As MIB_IPADDRROW

    Dim sTempIpAdress As String
    Dim iIndex As Long
    Dim i As Integer

    On Error Resume Next

    ' Get the rows of the IP Address Table


    ' Create the network tables class
    AddRow = IpAddressTable

    ' init Index
    iIndex = -1

    For i = 0 To UBound(AddRow)
        sTempIpAdress = getascip(AddRow(i).dwAddr)
        If sTempIpAdress = IP Then
            iIndex = AddRow(i).dwIndex
        End If
    Next i

    GetIndexIP = iIndex

    On Error GoTo 0
End Function



Private Property Get IpAddressTable() As MIB_IPADDRROW()
    Dim lngRet As Long, tableSize As Long
    Dim mIpAddrTable As MIB_IPADDRTABLE
    Dim byteArr() As Byte
    Dim ArrSize As Long

10  On Error GoTo IpAddressTable_Error

20  lngRet = GetIpAddrTable(ByVal 0&, tableSize, API_TRUE)
30  If tableSize > 0 Then
40      ReDim byteArr(0 To tableSize - 1) As Byte
50      lngRet = GetIpAddrTable(ByVal VarPtr(byteArr(0)), tableSize, API_TRUE)

60      Call CopyMemory(ByVal VarPtr(mIpAddrTable.dwNumEntries), ByVal VarPtr(byteArr(0)), Len(mIpAddrTable.dwNumEntries))
70      If mIpAddrTable.dwNumEntries > 0 Then
80          ReDim mIpAddrTable.table(0 To mIpAddrTable.dwNumEntries - 1) As MIB_IPADDRROW
90          ArrSize = mIpAddrTable.dwNumEntries * Len(mIpAddrTable.table(0))
100         Call CopyMemory(ByVal VarPtr(mIpAddrTable.table(0)), ByVal VarPtr(byteArr(4)), ArrSize)
110     End If
120     Erase byteArr
130 End If
140 IpAddressTable = mIpAddrTable.table

150 On Error GoTo 0
160 Exit Property

IpAddressTable_Error:

    'MsgBox "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in IpAddressTable of CGetOthersIp", vbCritical

End Property


Private Function getascip(ByVal inn As Long) As String
' Convert a network address to a application IP address
    Dim lpStr&, nStr&, retString$

10  On Error GoTo getascip_Error

20  Call StartWinsock
30  getascip = "255.255.255.255"
40  retString = String(32, 0)
50  lpStr = inet_ntoa(inn)
60  If lpStr <> 0 Then
70      nStr = lstrlen(lpStr)
80      If nStr > 32 Then nStr = 32
90      CopyMemory ByVal retString, ByVal lpStr, nStr
100     retString = Left(retString, nStr)
110     getascip = retString
120 End If
130 Call EndWinsock

140 On Error GoTo 0
150 Exit Function

getascip_Error:
    'MsgBox "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in getascip of CGetOthersIp", vbCritical

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





