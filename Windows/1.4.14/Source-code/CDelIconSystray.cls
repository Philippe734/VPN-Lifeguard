VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDelIconSystray"
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
' Module     : CDelIconSystray
' Author     : sur une idée originale de Dragonmaster
' Date       : 21/06/2010
' Internet   : http://www.vbfrance.com/codes/SYSTRAY-VIRER-IMPORTE-QUELLE-ICONE_39241.aspx
'---------------------------------------------------------------------------------------

Option Explicit

Private Const NIM_DELETE As Long = &H2

'début des messages utilisateurs
Private Const WM_USER As Long = &H400
'renvoie des infos sur un bouton d'une barre d'outils
Private Const TB_GETBUTTON As Long = (WM_USER + 23)
'renvoie le nombre de boutons d'une barre d'outils
Private Const TB_BUTTONCOUNT As Long = (WM_USER + 24)

'autorise à obtenir des infos sur une processus
Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)
'autorise les opérations sur la mémoire virtuelle d'un processus
Private Const PROCESS_VmoPERATION As Long = (&H8)
'autorise à lire la mémoire virtuelle d'un processus
Private Const PROCESS_VM_READ As Long = (&H10)
'autorise à écrire la mémoire virtuelle d'un processus
Private Const PROCESS_VM_WRITE As Long = (&H20)
'autorise toutes les opérations précédantes
Private Const PROCESS_VM As Long = PROCESS_QUERY_INFORMATION Or PROCESS_VmoPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE    'Or PROCESS_DUphANDLE

'alloue de la mémoire
Private Const MEM_COMMIT As Long = &H1000
'en lecture écriture
Private Const PAGE_READWRITE As Long = &H4

'libère de la mémoire
Private Const MEM_RELEASE As Long = &H8000

'contient des infos sur un bouton de barre d'outils
Private Type TBBUTTON
    iBitmap As Long    'index (dans l'ImageList) de l'image affichée
    idCommand As Long    'identifiant de la commande actionnée par le bouton
    fsState As Byte    'état du bouton
    fsStyle As Byte    'style du bouton
    bReserved(1 To 2) As Byte    'padding
    dwData As Long    'pointeur vers des données définies par l'utilisateur
    iString As Long    'texte de la ToolTip
End Type

'contient des infos sur une icone du "tray"
'pour plus d'infos voir NOTIFYICONDATA et TBBUTTON
Private Type Tray
    HWnd As Long    'handle de la fenêtre
    uID As Long    'id de l'icone
    uCallbackMessage As Long    'message envoyé à la fenêtre
    Unknown1(1) As Long    'interprétation inconnue
    hIcon As Long    'handle de l'icone affichée
    Unknown2(2) As Long    'interprétation inconnue
    uniPath As String    'chemin et nom du fichier du processus qui a cette icone dans le tray
    sTip As String    'ToolTip de l'icone
    iBitmap As Long    'index de l'icone affichée
    idCommand As Long    'id de la commande
    fsState As Byte    'état de l'icone
    fsStyle As Byte    'style de l'icone (TBSTYLE_BUTTON)
    dwData As Long    'pointeur vers les données du tray
    iString As Long    'pointeur vers le Tip
    PID_CustomAdd As Long    ' ajout perso pour le PID du processus
End Type

'contient des infos sur les icones dans le Tray
Private Type NOTIFYICONDATA
    cbSize As Long    'taille de cette structure
    HWnd As Long    'handle de la fenêtre propriétaire
    uID As Long    'ID de l'icone
    uFlags As Long    'drapeau
    uCallbackMessage As Long    'message envoyé à l'application
    hIcon As Long    'handle de l'icone
    szTip As String * 64    'texte affiché lors du survol de l'icone
    dwState As Long
    szInfo As String * 256
    uTimeoutAndVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

Private NID As NOTIFYICONDATA

Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef lpcbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long


'trouve une fenêtre par sa classe et/ou par son titre
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'trouve une fenêtre par sa classe et/ou par son titre
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hwnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'ferme un handle
'Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

'obtient un handle sur un processus
'Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

'envoie un message
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'alloue de la mémoire dans le processus spécifié
Private Declare Function VirtualAllocEx Lib "kernel32.dll" (ByVal hProcess As Long, lpAddress As Any, ByRef dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
'libère de la mémoire allouée dans le processus spécifié
Private Declare Function VirtualFreeEx Lib "kernel32.dll" (ByVal hProcess As Long, lpAddress As Any, ByRef dwSize As Long, ByVal dwFreeType As Long) As Long

'renvoie l'identifiant de processus de la fenêtre spécifiée
'Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

'lit une zone mémoire dans le processus spécifié
Private Declare Function ReadProcessMemory Lib "kernel32.dll" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'


'renvoie le Hwnd du tray
Private Function GethWndTray() As Long

    Dim HWnd As Long, hwnd2 As Long    'handle

'on descend dans la hiérarchie parent-enfant
    HWnd = FindWindow("Shell_TrayWnd", vbNullString)
    HWnd = FindWindowEx(HWnd, ByVal 0&, "TrayNotifyWnd", vbNullString)
    hwnd2 = FindWindowEx(HWnd, ByVal 0&, "SysPager", vbNullString)    'uniquement XP
    If (hwnd2 = 0) Then hwnd2 = HWnd    ' ME,2000
    'jusqu'à obtenir le handle de la barre d'outils faisant de tray
    hwnd2 = FindWindowEx(hwnd2, ByVal 0&, "ToolbarWindow32", vbNullString)    ' ME, 2000, XP...
    If (hwnd2 = 0) Then
        GethWndTray = HWnd    ' 95,98
    Else
        GethWndTray = hwnd2  ' ME, 2000, XP...
    End If
End Function


'renvoie la liste des icones présentes dans le tray
'comme les infos se trouvent dans Explorer.exe et non dans notre processus
'on doit allouée de la mémoire dans Explorer.exe pour pouvoir fournir un pointeur valide vers un TBBUTTON à SendMessage
'cette fonction ne fonctionne que sous NT/2000/XP (car VirtualAllocEx n'est pas présent sous 9x)
Private Function GetAllIconsFromSystray() As Tray()

    Dim hProc As Long    'handle de Explorer.exe
    Dim hTray As Long    'hWnd du tray
    Dim PID As Long    'identifiant de Explorer.exe
    Dim ptrVM As Long    'pointeur vers la zone mémoire pour TBBUTTON
    Dim ret As Long    'valeur de retour
    Dim ptrTip As Long    'pointeur vers une chaîne Unicode (Tip ou Path)
    Dim tbut As TBBUTTON    'les infos sur le bouton
    Dim TChar As Integer    'contient un caractère unicode
    Dim x As Long    'compteur
    Dim nIcons As Long    'nombre d'icones dans le tray
    Dim Icons() As Tray    'tableau temporaire des icones dans le tray

'on demande le hWnd du tray
    hTray = GethWndTray

    'si erreur
    If hTray = 0 Then Exit Function

    'on demande l'identifiant de Explorer.exe
    GetWindowThreadProcessId hTray, PID

    'on demande un handle de Explorer.exe
    hProc = OpenProcess(PROCESS_VM, 0, PID)

    'si erreur
    If hProc = -1 Then err.Raise &H80070000 & err.LastDllError

    'on alloue de la mémoire pour TBBUTTON
    ptrVM = VirtualAllocEx(hProc, ByVal 0&, Len(tbut), MEM_COMMIT, PAGE_READWRITE)

    'si erreur
    'If ptrVM = 0 Then err.Raise &H80070000 & err.LastDllError

    'on demande le nombre de boutons dans le tray
    nIcons = SendMessage(hTray, TB_BUTTONCOUNT, ByVal 0&, ByVal 0&)

    'on redimensionne le tableau au nombre d'icones
    ReDim Icons(nIcons - 1)

    'pour chaque icone
    For x = 0 To nIcons - 1
        'on demande les infos sur le bouton
        SendMessage hTray, TB_GETBUTTON, x, ByVal ptrVM
        'on lit les infos sur le bouton
        ReadProcessMemory hProc, ByVal ptrVM, tbut, Len(tbut), ret

        'on extrait les données utiles
        Icons(x).dwData = tbut.dwData
        Icons(x).fsState = tbut.fsState
        Icons(x).fsStyle = tbut.fsStyle
        Icons(x).iBitmap = tbut.iBitmap
        Icons(x).idCommand = tbut.idCommand
        Icons(x).iString = tbut.iString

        'la Tip
        ptrTip = tbut.iString
        Do
            'on lit chaque caractère unicode
            ReadProcessMemory hProc, ByVal ptrTip, TChar, 2&, ret
            Icons(x).sTip = Icons(x).sTip & ChrW$(TChar)
            ptrTip = ptrTip + 2
            'jusqu'au NULL
        Loop While TChar

        'on lit les données sur l'icone (Path...)
        ReadProcessMemory hProc, ByVal tbut.dwData, Icons(x).HWnd, 36&, ret

        'on demande le PID du processus
        GetWindowThreadProcessId Icons(x).HWnd, PID

        ' ajout perso pour stocker le PID du processus dans l'icone
        Icons(x).PID_CustomAdd = PID

        'on lit le chemin et le nom de l'application propriétaire de l'icone
        GetProcessNameFromPID_NT PID, Icons(x).uniPath
    Next x

    'on renvoie le tout
    GetAllIconsFromSystray = Icons

Cleanup:
    'on libère la mémoire allouée
    VirtualFreeEx hProc, ByVal ptrVM, 0&, MEM_RELEASE
    'on ferme le handle de Explorer.exe
    CloseHandle hProc
End Function

'renvoie le nom du processu du PID spécifié
'sous NT/2000/XP..., il faut utliser PSAPI
'=================================================
'IN pid : identifiant du processus
'OUT lpName : nom et chemin du processus
Private Sub GetProcessNameFromPID_NT(PID As Long, lpName As String)

    Dim hP As Long    'handle du processus
'on demande un handle du processus
    hP = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, False, PID)
    'si succès
    If (hP) Then
        Dim hMod As Long    'handle de l'exe
        Dim ret As Long
        'on demande le handle de l'exe de ce processus
        EnumProcessModules hP, hMod, 4&, ret
        'si succès
        If (hMod) Then
            'on alloue de l'espace mémoire
            lpName = Space(255)
            'on demande le nom et le chemin de l'exe
            GetModuleFileNameExA hP, hMod, lpName, 255
            'on supprime le NULL
            lpName = Left$(lpName, InStr(lpName, vbNullChar) - 1)
        End If
        'on ferme de handle du processus
        CloseHandle hP
    End If
End Sub

'supprime un icone du systray à l'aide du PID de son processus
Public Function DelIconByPID(ByVal PIDicon As Long) As Boolean
    Dim ret As Long
    Dim i As Long
    Dim AllTrayIcons() As Tray

    'on récupère toutes les infos des icones du systray
    AllTrayIcons = GetAllIconsFromSystray

    'pour tester
    'PIDicon = 3536

    'boucle pour trouver l'icone parmis tous à supprimer en fonction du PID de son processus
    For i = 0 To UBound(AllTrayIcons)

        If AllTrayIcons(i).PID_CustomAdd = PIDicon Then

            ' PID trouvé alors on peut le supprimer
            With NID
                .HWnd = AllTrayIcons(i).HWnd     'ces 4 paramètres sont nécessaires
                .uID = AllTrayIcons(i).uID       'pout effacer une icône
                .dwState = AllTrayIcons(i).uID  'oui dwState : uID
                .uCallbackMessage = AllTrayIcons(i).uCallbackMessage
            End With

            'suppression de l'icone
            ret = Shell_NotifyIcon(NIM_DELETE, NID)
            'Debug.Print timer, "ret ShellDelIcon ="; Ret

            'signale que l'opération est ok
            DelIconByPID = True

            'sortie de la boucle car on à supprimé l'icone qu'on cherchait
            Exit For

        Else
            DelIconByPID = False
        End If
    Next i
End Function
