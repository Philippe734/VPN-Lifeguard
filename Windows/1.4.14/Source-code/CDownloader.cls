VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDownloader"
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
' Module    : CDownloader
' Author    : Renfield
' Date      : 14/04/2010
' Internet  : www.vbfrance.com/codes/MODULE-TELECHARGEMENT-FICHIER-BASIQUE-SYNCHRONE-ASYNCHRONE_51425.aspx
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function InternetCrackUrl Lib "wininet.dll" Alias "InternetCrackUrlA" (ByVal lpszURL As String, ByVal dwUrlLength As Long, ByVal dwFlags As Long, lpUrlComponents As URL_COMPONENTS) As Long
Private Declare Function InternetCanonicalizeUrl Lib "wininet.dll" Alias "InternetCanonicalizeUrlA" (ByVal lpszURL As String, ByVal lpszBuffer As String, lpdwBufferLength As Long, ByVal dwFlags As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long

Private Const KEY_QUERY_VALUE As Long = &H1
Private Const STANDARD_RIGHTS_READ = &H20000
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const SYNCHRONIZE = &H100000
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const HKEY_CURRENT_USER = &H80000001

Private Const REG_SZ = 1
Private Const REG_DWORD = 4

Private Const ICU_BROWSER_MODE = &H2000000
Private Const ICU_ESCAPE = &H80000000

Private Type PathType
    Folder As String  '# Inclue le '\' final
    FileName As String
    FileExt As String    '# Inclue le '.'
End Type

Private Type URL_COMPONENTS
    StructSize As Long
    Scheme As String
    SchemeLength As Long
    nScheme As Long
    HostName As String
    HostNameLength As Long
    nPort As Long
    UserName As String
    UserNameLength As Long
    Password As String
    PasswordLength As Long
    URLPath As String
    UrlPathLength As Long
    ExtraInfo As String
    ExtraInfoLength As Long
End Type

Private WithEvents m_clsHttp As WinHttpRequest
Attribute m_clsHttp.VB_VarHelpID = -1
Private moStream As Object
Private miContentLength As Long
Private miAlreadyLength As Long
Private msTarget As String
Private msURL As String
Private msMessage As String
Public psProxy As String

Public Event Error(ByVal vsURL As String, ByVal vsMessage As String, ByVal vnCode As Long)
Public Event Finished(ByVal vsURL As String, ByVal vsTarget As String)
Public Event Begin(ByVal vsURL As String, ByVal vsTarget As String, ByVal vnContentLength As Long)
Public Event DataReceived(ByVal vsURL As String, ByVal vsTarget As String, ByVal vnContentLength As Long, ByVal vnStepLength As Long)

Public Property Get Target() As String
    Target = msTarget
End Property

Public Property Get URL() As String
    URL = msURL
End Property

Public Property Get ContentLength() As Long
    ContentLength = miContentLength
End Property

Public Sub GetFile(ByVal vsURL As String, Optional ByVal vsTarget As String, Optional ByVal vbbAsynchronous As Boolean = True)
    Dim tURL As URL_COMPONENTS
    Dim tTarget As PathType
    Dim sBuffer As String
    Dim nLength As Long
    Dim iFile As Integer
10  On Error GoTo GetFile_Error

20  sBuffer = Space$(2048)
30  nLength = 2048
40  If InternetCanonicalizeUrl(vsURL, sBuffer, nLength, ICU_BROWSER_MODE) Then
50      msURL = Strings.Left$(sBuffer, nLength)
60      With tURL
70          .StructSize = LenB(tURL)
80          .Scheme = Space$(32)
90          .SchemeLength = 32
100         .nScheme = -1
110         .HostName = Space$(256)
120         .HostNameLength = 256
130         .URLPath = Space$(256)
140         .UrlPathLength = 256
150         .ExtraInfoLength = 256
160         .ExtraInfo = Space$(256)
170         .Password = Space$(256)
180         .PasswordLength = 256
190         .UserName = Space$(256)
200         .UserNameLength = 256

210         InternetCrackUrl msURL, Len(msURL), ICU_ESCAPE, tURL

220         tURL.URLPath = Left$(tURL.URLPath, tURL.UrlPathLength)
230     End With

        '# On va reconstituer le nom du fichier de sortie.
240     tTarget = CrackPath(vsTarget)
250     If LenB(tTarget.Folder) = 0 Then
            '# L'utilisateur n'a pas spécifié de repertoire de sortie.
            '# On va écrire dans le répertoire courant
260         tTarget.Folder = CurDir
270     End If
280     If Right$(tTarget.Folder, 1) <> "\" Then
290         tTarget.Folder = tTarget.Folder & "\"
300     End If
        '# L'utilisateur n'a pas spécifié de nom, nous allons reprendre le nom du fichier distant.
310     If LenB(tTarget.FileName) = 0 Then
320         tTarget.FileExt = vbNullString
330         nLength = InStrRev(tURL.URLPath, "/")
340         tTarget.FileName = Mid$(tURL.URLPath, nLength + 1)
350     End If
360     msTarget = tTarget.Folder & tTarget.FileName & tTarget.FileExt

370     miContentLength = 0
380     miAlreadyLength = 0

390     iFile = FreeFile
400     Open msTarget For Output As #iFile: Close #iFile
410     Set moStream = CreateObject("Adodb.Stream")
420     moStream.Open
430     moStream.Type = 1

440     Set m_clsHttp = New WinHttp.WinHttpRequest
450     With m_clsHttp
460         If LenB(psProxy) Then
470             .SetProxy 2, psProxy
480         End If
490         .Open "GET", msURL, True
500         .Send

510         If Not vbbAsynchronous Then
520             Do Until Nothing Is moStream
530                 DoEvents
540             Loop
550         End If
560     End With
570 End If

580 On Error GoTo 0
590 Exit Sub

GetFile_Error:

    msMessage = "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in GetFile of CDownloader"
    Call ErrorCodeLine

End Sub

Public Sub Abort()
    If Not Nothing Is m_clsHttp Then
        m_clsHttp.Abort
        RaiseEvent Error(msURL, "Aborted by user", 400)
    End If
End Sub

Private Sub ErrorCodeLine()
    If Not Nothing Is m_clsHttp Then
        m_clsHttp.Abort
        Debug.Print Timer, "m_clsHttp.Abort"
    End If
    RaiseEvent Error(msURL, msMessage, 0)
End Sub

Private Function CrackPath(ByVal vsInput As String) As PathType
    Dim nPos As Long
10  On Error GoTo CrackPath_Error

20  If LenB(vsInput) Then
30      nPos = InStrRev(vsInput, "\")
40      If nPos Then
50          CrackPath.Folder = Left$(vsInput, nPos)
60          vsInput = Mid$(vsInput, nPos + 1)
70      End If

80      nPos = InStrRev(vsInput, ".")
90      If nPos Then
100         CrackPath.FileExt = Mid$(vsInput, nPos)
110         CrackPath.FileName = Left$(vsInput, nPos - 1)
120     Else
130         CrackPath.FileName = vsInput
140     End If
150 End If

160 On Error GoTo 0
170 Exit Function

CrackPath_Error:
180 msMessage = "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in CrackPath of CDownloader"
190 Call ErrorCodeLine
End Function

Private Sub Class_Initialize()
    psProxy = GetProxy
End Sub

Private Sub m_clsHttp_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    Debug.Print Timer, "m_clsHttp_OnError = stop"
    msMessage = "Internet connection error"
    Call ErrorCodeLine
End Sub

Private Sub m_clsHttp_OnResponseDataAvailable(Data() As Byte)
    Dim nStepLength As Long
    '# Si status OK ...
    If m_clsHttp.Status = 200 Then
        nStepLength = 1 + UBound(Data)
        miAlreadyLength = miAlreadyLength + nStepLength
        moStream.Write Data
        RaiseEvent DataReceived(msURL, msTarget, miContentLength, nStepLength)
    End If
End Sub

Private Sub m_clsHttp_OnResponseFinished()
    If m_clsHttp.Status = 200 Then
        moStream.SaveToFile msTarget, 2
    End If
    moStream.Close
    Set moStream = Nothing
    If m_clsHttp.Status = 200 Then
        RaiseEvent Finished(msURL, msTarget)
    Else
        RaiseEvent Error(msURL, m_clsHttp.StatusText, m_clsHttp.Status)
    End If
End Sub

Private Sub m_clsHttp_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
    On Local Error Resume Next
    miContentLength = Val(m_clsHttp.GetResponseHeader("Content-Length"))
    If m_clsHttp.Status = 200 Then
        RaiseEvent Begin(msURL, msTarget, miContentLength)
    End If
End Sub

Private Property Get RegKeyValue(ByVal veRootKey As Long, ByRef vsKeyName As String, ByRef vsValueName As String) As Variant
    Dim hKey As Long
    Dim sBuffer As String
    Dim nBuffer As Long
    Dim nLength As Long
    Dim eValueType As Long
    If 0 = RegOpenKeyEx(veRootKey, vsKeyName, 0&, KEY_READ, hKey) Then
        RegQueryValueEx hKey, vsValueName, 0, eValueType, ByVal 0&, nLength

        Select Case eValueType
        Case REG_SZ
            sBuffer = Space$(nLength)
            RegQueryValueEx hKey, vsValueName, 0, eValueType, ByVal sBuffer, nLength
            RegKeyValue = Left$(sBuffer, nLength - 1)
        Case REG_DWORD
            RegQueryValueEx hKey, vsValueName, 0, eValueType, nBuffer, 4
            RegKeyValue = nBuffer
        End Select

        RegCloseKey hKey
    End If
End Property

Private Function GetProxy() As String
    Dim sBuffer As String
    Dim nPos As Long
    Const IE_SETTINGS = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\"
10  On Error GoTo GetProxy_Error

20  sBuffer = RegKeyValue(HKEY_CURRENT_USER, IE_SETTINGS, "AutoConfigURL")
30  If LenB(sBuffer) Then
40      If StrComp(Left$(sBuffer, 7), "http://", vbTextCompare) Then
50          sBuffer = "http://" & sBuffer
60      End If

70      If LenB(sBuffer) Then
80          nPos = InStr(1, sBuffer, "psProxy ")
90          sBuffer = Mid$(sBuffer, nPos + 6)
100         nPos = InStr(1, sBuffer, """")
110         sBuffer = Trim(Left$(sBuffer, nPos - 1))
120         If Right$(sBuffer, 1) = ";" Then
130             sBuffer = Left$(sBuffer, Len(sBuffer) - 1)
140         End If
150         GetProxy = sBuffer
160     End If
170 ElseIf RegKeyValue(HKEY_CURRENT_USER, IE_SETTINGS, "ProxyEnable") = 1 Then
180     sBuffer = RegKeyValue(HKEY_CURRENT_USER, IE_SETTINGS, "ProxyServer")
190     If LenB(sBuffer) Then
200         nPos = InStr(1, sBuffer, "http=")
210         If nPos Then
220             sBuffer = Mid$(sBuffer, nPos + 5)
230             nPos = InStr(1, sBuffer, ";")
240             sBuffer = Left$(sBuffer, nPos - 1)
250         End If

260         If StrComp(Left$(sBuffer, 7), "http://", vbTextCompare) Then
270             GetProxy = "http://" & sBuffer
280         Else
290             GetProxy = sBuffer
300         End If
310     End If
320 End If

330 On Error GoTo 0
340 Exit Function

GetProxy_Error:

350 msMessage = "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in GetProxy of CDownloader"
360 Call ErrorCodeLine

End Function
