Attribute VB_Name = "modKillProcess"

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
' Module     : modKillProcess
' Author     : Satirik
' Date       : 02/07/2002
' Internet   : http://www.vbfrance.com/codes/KILL-PROCESS-WIN2K-WINXP_3886.aspx
'---------------------------------------------------------------------------------------

Option Explicit

Private Const MAX_PATH As Integer = 260
Private Const TH32CS_SNAPPROCESS As Long = 2&

'Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type

'-------- Les deux collections à appeler -----------------------------------------
'
Public pcoListeProcessPID As New Collection
Public pcoListeProcessString As New Collection
'
'---------------------------------------------------------------------------------

'-----------------------------------------------------
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'-----------------------------------------------------

Public Function ProcessTerminate(Optional lProcessID As Long, Optional lHwndWindow As Long) As Boolean
    Dim lhwndProcess As Long
    Dim lExitCode As Long
    Dim lRetVal As Long
    Dim lhThisProc As Long
    Dim lhTokenHandle As Long
    Dim tLuid As LUID
    Dim tTokenPriv As TOKEN_PRIVILEGES, tTokenPrivNew As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long

    Const TOKEN_ADJUST_PRIVILEGES = &H20
    Const PROCESS_TERMINAT = &H1
    Const TOKEN_QUERY = &H8, SE_DEBUG_NAME As String = "SeDebugPrivilege"
    Const SE_PRIVILEGE_ENABLED = &H2

    On Error Resume Next
    If lHwndWindow Then
        'Get the process ID from the window handle
        lRetVal = GetWindowThreadProcessId(lHwndWindow, lProcessID)
    End If

    If lProcessID Then
        'Give Kill permissions to this process
        lhThisProc = GetCurrentProcess

        OpenProcessToken lhThisProc, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lhTokenHandle
        LookupPrivilegeValue "", SE_DEBUG_NAME, tLuid
        'Set the number of privileges to be change
        tTokenPriv.PrivilegeCount = 1
        tTokenPriv.TheLuid = tLuid
        tTokenPriv.Attributes = SE_PRIVILEGE_ENABLED
        'Enable the kill privilege in the access token of this process
        AdjustTokenPrivileges lhTokenHandle, False, tTokenPriv, Len(tTokenPrivNew), tTokenPrivNew, lBufferNeeded

        'Open the process to kill
        lhwndProcess = OpenProcess(PROCESS_TERMINAT, 0, lProcessID)

        If lhwndProcess Then
            'Obtained process handle, kill the process
            ProcessTerminate = CBool(TerminateProcess(lhwndProcess, lExitCode))
            Call CloseHandle(lhwndProcess)
        End If
    End If
    On Error GoTo 0
End Function

Public Sub ListerProcess()
    Dim hSnapshot As Long
    Dim uProcess As PROCESSENTRY32
    Dim r As Long

    'les deux listes public des processus sont :
    '- pcoListeProcessString
    '- pcoListeProcessPID
    Set pcoListeProcessString = Nothing
    Set pcoListeProcessPID = Nothing
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot = 0 Then Exit Sub
    uProcess.dwSize = Len(uProcess)
    r = ProcessFirst(hSnapshot, uProcess)
    Do While r
        pcoListeProcessString.Add Mid$(uProcess.szexeFile, 1, InStr(1, uProcess.szexeFile, vbNullChar) - 1)
        pcoListeProcessPID.Add uProcess.th32ProcessID
        r = ProcessNext(hSnapshot, uProcess)
    Loop
End Sub

