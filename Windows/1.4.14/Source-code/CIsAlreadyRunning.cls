VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIsAlreadyRunning"
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
' Module     : CIsAlreadyRunning
' Author     : Renfield
' Date       : 31/08/2006
' Internet   : http://www.vbfrance.com/codes/INSTANCE-UNIQUE-VOTRE-APPLICATION_39361.aspx
'---------------------------------------------------------------------------------------


Option Explicit

Private Const INVALID_HANDLE_VALUE As Long = &HFFFFFFFF
Private Const PAGE_READWRITE As Long = &H4
Private Const SECTION_MAP_READ As Long = &H4
Private Const SECTION_MAP_WRITE As Long = &H2
Private Const FILE_MAP_READ As Long = SECTION_MAP_READ
Private Const FILE_MAP_WRITE As Long = SECTION_MAP_WRITE
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOW As Long = 5
Private Const SW_RESTORE As Long = 9
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWpsHOWWINDOW = &H40
Private Const GW_HWNDNEXT As Long = 2

'Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingA" (ByVal hfile As Long, ByRef lpFileMappigAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Private Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function OpenFileMapping Lib "kernel32.dll" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByRef lpBaseAddress As Any) As Long
'private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef source As Any, ByVal Length As Long)
Private Declare Function GetParent Lib "user32" (ByVal HWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal HWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal HWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal HWnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal HWnd As Long) As Long

'# Indique si le programme tourne dans l'IDE ou en mode compilé
Private mbIsInIDE As Boolean
Private mbHasBeenQueriedOnce As Boolean
'

'# Cette fonction permet d'initialiser la variables mbIsInIDE, qui indique si nous sommes dans l'IDE
Private Function TestEnvironment() As Boolean
'# Si cette fonction est appelée, c'est que nous sommes dans l'IDE
    mbIsInIDE = True
    '# On renvoie True, pour ne pas que l'instruction Debug.Assert ne génère un point d'arrêt
    TestEnvironment = True
End Function


Private Function InstanceToWindowHandle(ByVal target_pid As Long) As Long
    Dim test_hwnd As Long
    Dim test_pid As Long
    Dim test_thread_id As Long

    On Error Resume Next

    'Find the first window
    test_hwnd = FindWindow(ByVal 0&, ByVal 0&)

    Do While test_hwnd <> 0

        'Check if the window isn't a child
        If GetParent(test_hwnd) = 0 Then

            'Get the window's thread
            test_thread_id = GetWindowThreadProcessId(test_hwnd, test_pid)

            If test_pid = target_pid Then
                InstanceToWindowHandle = test_hwnd
                Exit Do
            End If

        End If

        'retrieve the next window
        test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

    Loop
End Function

Public Function IsAlreadyRunning() As Boolean
    Dim hMap As Long
    Dim pMem As Long
    Dim hFrm As Long
    Dim FrmHwnd As Long

    '# Cette ligne de code sera supprimée à la compilation
    Debug.Assert TestEnvironment
    '# La fonction renvoie toujours False, si nous sommes en mode IDE.
    If Not mbIsInIDE Then
        '# Nous tentons ici d'acceder au mappage (précedemment créé ?)
        hMap = OpenFileMapping(FILE_MAP_READ, 0, App.Title)
        If hMap <> 0 Then
            '# L'application est déjà lancée.
            pMem = MapViewOfFile(hMap, FILE_MAP_READ, 0, 0, 0)
            If pMem <> 0 Then
                '# On récupère le handle vers la précédente fenêtre
                CopyMemory hFrm, ByVal pMem, 4
                If hFrm Then
                    On Error Resume Next
                    '# On active l'instance précedente
                    'AppActivate hPid
                    SetForegroundWindow hFrm
                    DoEvents
                    ShowWindow hFrm, SW_RESTORE
                    DoEvents
                    BringWindowToTop hFrm
                    DoEvents
                    On Error GoTo 0
                End If
                UnmapViewOfFile pMem
            End If
            '# On libère le handle hmap
            CloseHandle hMap
            '# et on prévient l'appelant que l'application avait dejà été lancée.
            IsAlreadyRunning = True
        Else
            '# Nous sommes dans la première instance de l'application.
            '# Nous allons laisser une marque en mémoire, pour l'indiquer
            hMap = CreateFileMapping(INVALID_HANDLE_VALUE, ByVal 0&, PAGE_READWRITE, 0, 4, App.Title)
            If hMap <> 0 Then
                '# On ouvre le 'fichier' en écriture
                pMem = MapViewOfFile(hMap, FILE_MAP_WRITE, 0, 0, 0)
                If pMem <> 0 Then
                    
                    ' méthode originelle avec le PID
'                    '# On y écrit l'ID du process courant
'                    CopyMemory ByVal pMem, GetCurrentProcessId(), 4
                    
                    ' On y écrit l'handle de la form principale
                    CopyMemory ByVal pMem, frmMain.HWnd, 4
                    
                    UnmapViewOfFile pMem
                End If
                '# Pas de CloseHandle hMap ici, sous peine de détruire le mappage lui-même...
            End If
        End If
        'ElseIf mbHasBeenQueriedOnce Then
        '# Conceptuellement, il n'est pas logique d'appeler cette fonction plus d'une fois.
        '# En mode IDE, nous avertissons l'utilisateur.
        'err.Raise vbObjectError, "IsAlreadyRunning", "Tentative d'appeler la fonction IsAlreadyRunning plus d'une fois au cours de l'éxecution." & vbNewLine & "Vous sortez du cadre logique de l'utilisation de cette fonction."
    Else
        mbHasBeenQueriedOnce = True
    End If
End Function




