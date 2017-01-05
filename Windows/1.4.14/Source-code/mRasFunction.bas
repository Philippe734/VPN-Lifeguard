Attribute VB_Name = "modRasFunction"

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
' Module     : modRasFunction
' Author     : MatMoul
' Date       : 29/01/2003
' Internet   : http://www.vbfrance.com/codes/CONNECTION-VPN-MODEM-AVEC-API-RASDIAL_5815.aspx
'---------------------------------------------------------------------------------------

Option Explicit

'Module de démonstration de l'API RasDial réalisé suite à une question sur le forum de VBFrance
'Vous pouvez utiliser et redistribuer ce code.
'Bug connu : ne renvoye pas le handle sous 98.
'Dans la mesure du possible, merci de laisser un petit clin d'oeil à l'auteur...
'J'me suis bien cassé la tête sur ce code

'Testé sur Win 98 SE; Win 2000 Pro,Server; Win XP Pro

'Pseudo : MatMoul
'Site   : www.matmoul.ch
'Mail   : mat@matmoul.ch

'Remarque concernant les constantes :
'Les valeurs des constantes ont été extraite des fichiers RAS.H, LMCONS.H
'Autre fichier à consulter RASERROR.H

'Remarque concerant les variables commentées du type RASDIALPARAMS :
'En activant ces deux lignes, la taille du buffer vas passer de 1052 à 1060.
'Cette taille n'est pas supportée par Windows 98.

Public Const RAS_MaxEntryName = 256
Public Const RAS_MaxPhoneNumber = 128
Public Const RAS_MaxCallbackNumber = 128
Public Const UNLEN = 256
Public Const DNLEN = 15
Public Const PWLEN = 256

Private m_VarhRasConn As Long    'handle rasconn





Public Enum RASCONNSTATE
    RASCS_OpenPort
    RASCS_PortOpened
    RASCS_ConnectDevice
    RASCS_DeviceConnected
    RASCS_AllDevicesConnected
    RASCS_Authenticate
    RASCS_AuthNotify
    RASCS_AuthRetry
    RASCS_AuthCallback
    RASCS_AuthChangePassword
    RASCS_AuthProject
    RASCS_AuthLinkSpeed
    RASCS_AuthAck
    RASCS_ReAuthenticate
    RASCS_Authenticated
    RASCS_PrepareForCallback
    RASCS_WaitForModemReset
    RASCS_WaitForCallback
    RASCS_Projected
    RASCS_SubEntryConnected
    RASCS_SubEntryDisconnected
    RASCS_ApplySettings
    RASCS_Interactive
    RASCS_RetryAuthentication
    RASCS_CallbackSetByCaller
    RASCS_PasswordExpired
    RASCS_InvokeEapUI
    RASCS_Connected = 8192
    RASCS_Disconnected = 0
End Enum


Private Type RASEntryName
    dwSize As Long
    szEntryName(RAS_MaxEntryName) As Byte
End Type

Public Type RASDIALPARAMS
    dwSize As Long
    szEntryName(RAS_MaxEntryName) As Byte
    szPhoneNumber(RAS_MaxPhoneNumber) As Byte
    szCallbackNumber(RAS_MaxCallbackNumber) As Byte
    szUserName(UNLEN) As Byte
    szPassword(PWLEN) As Byte
    szDomain(DNLEN) As Byte
    'dwSubEntry As Long    '2K Only
    'dwCallbackId As Long  '2K Only
End Type

Private Type RasConn
    dwSize As Long
    hRasConn As Long
    szEntryName(0 To 256) As Byte
    szDeviceType(0 To 16) As Byte
    szDeviceName(0 To 128) As Byte
    Pad As Byte
End Type

Private Type RASconnStatus
    dwSize As Long
    RASCONNSTATE As RASCONNSTATE
    dwError As Long
    szDeviceType(0 To 16) As Byte
    szDeviceName(0 To 128) As Byte
    'szPhoneNumber(RAS_MaxPhoneNumber) As Byte
End Type

Private Type RAS_STATS
    dwSize As Long
    dwBytesXmited As Long
    dwBytesRcved As Long
    dwFramesXmited As Long
    dwFramesRcved As Long
    dwCrcErr As Long
    dwTimeoutErr As Long
    dwAlignmentErr As Long
    dwHardwareOverrunErr As Long
    dwFramingErr As Long
    dwBufferOverrunErr As Long
    dwCompressionRatioIn As Long
    dwCompressionRatioOut As Long
    dwBps As Long
    dwConnectDuration As Long
End Type

'l'api rasdial est dans le module modRasDialThreaded

Private Declare Function RasGetConnectStatus Lib "rasapi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasConn As Long, lpRasconnstatus As RASconnStatus) As Long
'Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, ByVal lpString2 As String) As Long
Private Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal lphRasConn As Long) As Long
Private Declare Function RasEnumEntries Lib "rasapi32.dll" Alias "RasEnumEntriesA" (ByVal Reserved As String, ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcentries As Long) As Long
Private Declare Function RasEnumConnections Lib "rasapi32" Alias "RasEnumConnectionsA" (ByVal lpRasConn As Long, ByVal lpcb As Long, ByVal lpcConnections As Long) As Long
Private Declare Function RasGetConnectionStatistics Lib "rasapi32" (ByVal hRasConn As Long, ByVal lpStatistics As Long) As Long
'Private Declare Function RasGetEntryDialParams Lib "rasapi32.dll" Alias "RasGetEntryDialParamsA" (ByVal lpcstr As String, ByRef lprasdialparamsa As Any, ByRef lpbool As Long) As Long

'Retourne la liste des connexions disponnibles dans le dossier Connexions réseaux et Accès à distance
' Separator : Permet de définir le caractère de séparation des éléments
Public Function GetConnectionList(Optional ByVal Separator As String = ";") As String
    Dim lpcb As Long, lpce As Long, i As Long
    Dim ConName As String
    Dim vRAS(255) As RASEntryName
    vRAS(0).dwSize = LenB(vRAS(0))
    lpcb = 256 * vRAS(0).dwSize
    Call RasEnumEntries(vbNullString, vbNullString, vRAS(0), lpcb, lpce)
    DoEvents
    For i = 0 To lpce - 1
        ConName = StrConv(vRAS(i).szEntryName(), vbUnicode)
        If InStr(1, ConName, Chr(0)) > 0 Then ConName = Left(ConName, InStr(1, ConName, Chr(0)) - 1)
        'Debug.Print timer, Trim(Len(ConName)) + " : " + ConName
        GetConnectionList = GetConnectionList + ConName + Separator
    Next i
End Function

'-----------------------------------------------------
''Appelle une connexion existante dans Connexion Réseau et accès distant.
''Retourne le handle de la connexion ou 0 si erreur
'' ConName  : Nom de la connexion
'' UserName : Nom d'utilisateur
'' Password : Mot de passe
'' Domain   : Nom du domaine (Optionel)
'' Number   : Numéro de tél ou ip (Optionel)
'Public Function Dial(ByVal ConName As String, ByVal UserName As String, ByVal Password As String, Optional ByVal Domain As String = vbNullString, Optional ByVal Number As String = vbNullString) As Long
'    Dim lngRetCode As Long
'    Dim hRasConn As Long
'    Dim lprasdialparams As RASDIALPARAMS
'
'    lprasdialparams.dwSize = LenB(lprasdialparams)
'    lstrcpy lprasdialparams.szEntryName(0), ConName
'    lstrcpy lprasdialparams.szUserName(0), UserName
'    lstrcpy lprasdialparams.szPassword(0), Password
'    lstrcpy lprasdialparams.szDomain(0), Domain
'    lstrcpy lprasdialparams.szPhoneNumber(0), Number
'    lstrcpy lprasdialparams.szCallbackNumber(0), ""
'    lngRetCode = RasDial(ByVal &H0, vbNullString, lprasdialparams, &H0, ByVal &H0, hRasConn)
'    DoEvents
'    If lngRetCode = 0 Then
'        Dial = hRasConn
'    Else
'        If Not hRasConn = 0 Then HangUp hRasConn
'        Dial = 0
'    End If
'End Function
'-----------------------------------------------------

'-----------------------------------------------------
''Appelle une connexion existante dans Connexion Réseau et accès distant en utilisant les paramètres enregistrés.
''Retourne le handle de la connexion ou 0 si erreur
'' ConName  : Nom de la connexion
'Public Function AutoDial(ByVal ConName As String) As Long
'    Dim lngRetCode As Long
'    Dim hRasConn As Long
'    Dim lprasdialparams As RASDIALPARAMS
'    lprasdialparams.dwSize = LenB(lprasdialparams)
'    lstrcpy lprasdialparams.szEntryName(0), ConName
'    RasGetEntryDialParams vbNullString, lprasdialparams, 0
'    lngRetCode = RasDial(ByVal &H0, vbNullString, lprasdialparams, &H0, ByVal &H0, hRasConn)
'    DoEvents
'    If lngRetCode = 0 Then
'        AutoDial = hRasConn
'        m_VarhRasConn = hRasConn
'    Else
'        If Not hRasConn = 0 Then HangUp hRasConn
'        AutoDial = 0
'    End If
'End Function
'-----------------------------------------------------

'Appelle une connexion existante dans Connexion Réseau et accès distant en utilisant les paramètres enregistrés.
'Retourne le handle de la connexion ou 0 si erreur
' ConName  : Nom de la connexion
Public Function AutoDialThreaded(ByVal ConName As String) As Long
    Dim StatusRAScon As RASCONNSTATE

10  On Error GoTo AutoDialThreaded_Error

    ' copie public afin d'avoir
    ' le nom de la connexion vpn dans
    ' la procédure rasdial lors du multithreading
20  psConnectionName = ConName

30  StatusRAScon = GetStatusRasconn(GetCurrentHandleRasConn)

40  If StatusRAScon <> RASCS_Connected Then
50      If pbThreadEnCours = False Then
            ' test si une thread est déjà en cours
            ' afin d'avoir qu'une connexion en cours
60          pbThreadEnCours = True

            ' désactive les boutons durant le multithreading
            ' afin d'éviter des anomalies
70          frmMain.cmdAbout.Enabled = False
80          'frmMain.cmdArreter.Enabled = False
90          frmMain.cmdConfigurer.Enabled = False
100         frmMain.cmdQuitter.Enabled = False


            'initialise l'indicateur de fin de thread
110         pbFinThread = False

            ' La connexion vpn s'effectue via l'api rasdial
            ' Mais durant l'étape de connexion en cours
            ' le programme est figé tant que l'api n'envois pas
            ' son résultat. C'est pour cela que j'ai choisis
            ' d'exécuter l'api rasdial dans une thread suivant
            ' le principe du multithreading activeX

            ' méthode multithreading activeX
120         Call frmMain.CreateThreadPourRasdial

            'attend que la thread soit terminée pour continuer
            'la suite de cette procédure
130         Do
140             DoEvents
150         Loop Until pbFinThread = True

            ' active les boutons car le multithreading est terminé
160         frmMain.cmdAbout.Enabled = True
170         'frmMain.cmdArreter.Enabled = True
180         frmMain.cmdConfigurer.Enabled = True
190         frmMain.cmdQuitter.Enabled = True

200         pbThreadEnCours = False
210     End If    'de If pbThreadEnCours = False


220     If plngRetCodeThreaded = 0 Then
230         AutoDialThreaded = phRasConnThreaded
240         m_VarhRasConn = phRasConnThreaded
            'MsgBox "dans ret=0 ; AutoDialThreaded = " & AutoDialThreaded
250     Else
260         If Not phRasConnThreaded = 0 Then HangUp phRasConnThreaded
270         AutoDialThreaded = 0
            'MsgBox "dans ret<>0 ; AutoDialThreaded = " & AutoDialThreaded
280     End If

290 Else

        'VPN déjà connecté
300     m_VarhRasConn = GetCurrentHandleRasConn
310     AutoDialThreaded = m_VarhRasConn

320 End If    'de If StatusRAScon <> RASCS_Connected


330 On Error GoTo 0
340 Exit Function

AutoDialThreaded_Error:

350 Call GestionErreur(frmMain, "AutoDialThreaded", pcErreurMessageSupplementaire)
360 frmMain.Show
370 frmMain.SetFocus
380 If pbEnvoyerFormulaire = True Then
390     frmMain.WindowState = vbMinimized
400     Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
410 End If
End Function

'Déconnecte une connexion
'Retourne True si réussi
' hRasConn : handle de la connexion à déconnecter
Public Function HangUp(ByVal hRasConn As Long) As Boolean
    HangUp = (RasHangUp(hRasConn) = 0)
    DoEvents
End Function


'retourne le handle de la connexion active
'sinon retourne 0
Public Function GetCurrentHandleRasConn() As Long
    Dim conn As RasConn
    Dim stat As RAS_STATS
    Dim y As Long
    Dim iNRasConn As Long    'nombre de connexion active

    conn.dwSize = Len(conn)
    y = conn.dwSize

    If RasEnumConnections(VarPtr(conn), VarPtr(y), VarPtr(iNRasConn)) = 0 Then
        GetCurrentHandleRasConn = conn.hRasConn
        stat.dwSize = Len(stat)
        'Debug.Print timer, "nombre de connexions active : " & iNRasConn
        'Debug.Print timer, "conn :"
        'Debug.Print timer, conn.dwSize
        'Debug.Print timer, conn.hRasConn
        'Debug.Print timer, conn.pad
        'Debug.Print timer, conn.szDeviceName
        'Debug.Print timer, conn.szDeviceType
        'Debug.Print timer, conn.szEntryName
        If RasGetConnectionStatistics(conn.hRasConn, VarPtr(stat)) = 0 Then
            'Debug.Print timer, "stat :"
            'Debug.Print timer, stat.dwAlignmentErr
            'Debug.Print timer, stat.dwBps
            'Debug.Print timer, stat.dwBufferOverrunErr
            'Debug.Print timer, stat.dwBytesRcved
            'Debug.Print timer, stat.dwBytesXmited
            'Debug.Print timer, stat.dwCompressionRatioIn
            'Debug.Print timer, stat.dwCompressionRatioOut
            'Debug.Print timer, stat.dwConnectDuration
            'Debug.Print timer, stat.dwCrcErr
            'Debug.Print timer, stat.dwFramesRcved
            'Debug.Print timer, stat.dwFramesXmited
            'Debug.Print timer, stat.dwFramingErr
            'Debug.Print timer, stat.dwTimeoutErr
        End If
    Else
        GetCurrentHandleRasConn = 0
    End If
End Function

Public Function GetStatusRasconn(ByVal hRasconns As Long) As RASCONNSTATE
    Dim bufferStatus As RASconnStatus
    Dim ret As Long
    'Dim hwd As Long

    bufferStatus.dwSize = 160
    'hwd = GetCurrentHandleRasConn
    ret = RasGetConnectStatus(hRasconns, bufferStatus)
    'Debug.Print timer, "ret = " & ret
    'Debug.Print timer, "bufferstatus = " & bufferStatus.rasconnstate
    GetStatusRasconn = bufferStatus.RASCONNSTATE
End Function

Public Property Get GetCurrentHandleRasConnFromAutodial() As Long
'utilisé lors de la lecture de la valeur de la propriété, du coté droit de l'instruction.
'Syntax: Debug.Print timer, X.FileName
    GetCurrentHandleRasConnFromAutodial = m_VarhRasConn
End Property

