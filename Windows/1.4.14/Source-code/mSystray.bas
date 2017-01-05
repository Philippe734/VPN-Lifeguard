Attribute VB_Name = "modSystray"

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
' Module     : modSystray
' Author     : Mark Mokoski
' Date       : 6-NOV-2004
' Internet   : www.cmtelephone.com
'
'Put App in SysTray, remove App from SysTray, Form on top, Balloon ToolTip code
'
'Also see Microsoft Knowledge base http://support.microsoft.com/default.aspx?scid=kb;en-us;149276
'for more information.
'
'This code is based on the Microsoft Knowledge Base code.
'---------------------------------------------------------------------------------------

Option Explicit

'API Functions used in this module
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal HWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    HWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

'Private blnClick As Boolean
Public vbTray As NOTIFYICONDATA

Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const wFlags As Long = SWP_NOMOVE Or SWP_NOSIZE
Public Const WM_MOUSEMOVE As Long = &H200    ' posté à une fenêtre lorsque le curseur bouge
Public Const WM_LBUTTONDBLCLK As Long = &H203  ' Double-clic gauche
Public Const WM_LBUTTONDOWN As Long = &H201    ' Bouton gauche enfoncé
Public Const WM_LBUTTONUP As Long = &H202    ' Bouton gauche relâché
Public Const WM_RBUTTONDBLCLK As Long = &H206  ' Double-clic droit
Public Const WM_RBUTTONDOWN As Long = &H204    ' Bouton droit enfoncé
Public Const WM_RBUTTONUP As Long = &H205    ' Bouton droit relâché
Public Const NIM_ADD As Long = &H0
Private Const NIM_DELETE As Long = &H2
Private Const NIF_ICON As Long = &H2
Private Const NIF_MESSAGE As Long = &H1
Private Const NIM_MODIFY As Long = &H1
Private Const NIF_TIP As Long = &H4
Private Const NIF_INFO As Long = &H10
'Private Const NIS_HIDDEN As Long = &H1
'Private Const NIS_SHAREDICON As Long = &H2
Private Const NIIF_NONE As Long = &H0
'Private Const NIIF_WARNING As Long = &H2
'Private Const NIIF_ERROR As Long = &H3
Private Const NIIF_INFO As Long = &H1
'Private Const NIIF_GUID As Long = &H4
'Private Const HWND_NOTOPMOST As Long = -2
Private Const HWND_TOPMOST As Long = -1

' Utilisation -----------------------------------------
' Dans le Form_Load :
'
'    systrayon me, texte
'
'---------------------
' Dans le Unload :
'
'    SystrayOff me
'
'---------------------
' Dans le Form_MouseMove :
'
'    Static rec As Boolean, msg As Long
'    msg = X / Screen.TwipsPerPixelX
'    If rec = False Then
'        rec = True
'        Select Case msg
'            Case WM_LBUTTONDBLCLK
'               Call mnuServiceAfficherApplication_Click
'            Case WM_RBUTTONUP
'                PopupMenu mnuzService, , , , mnuServiceAfficherApplication
'        End Select
'        rec = False
'    End If
'
'------------------------------------------------------------------------------------


Public Sub SystrayOn(Frm As Form, IconTooltipText As String)

'Adds Icon to SysTray

    With vbTray
        .cbSize = Len(vbTray)
        .HWnd = Frm.HWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .szTip = Trim(IconTooltipText$) & vbNullChar
        .hIcon = Frm.Icon
    End With

    Call Shell_NotifyIcon(NIM_ADD, vbTray)
    App.TaskVisible = False

End Sub

Public Sub SystrayOff(Frm As Form)

'Removes Icon from SysTray

    With vbTray
        .cbSize = Len(vbTray)
        .HWnd = Frm.HWnd
        .uID = vbNull
    End With

    Call Shell_NotifyIcon(NIM_DELETE, vbTray)

End Sub

Public Sub ChangeSystrayToolTip(Frm As Form, IconTooltipText As String)

'Changes the SysTray Balloon Tool Tip Text

    With vbTray
        .cbSize = Len(vbTray)
        .HWnd = Frm.HWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .szTip = Trim(IconTooltipText$) & vbNullChar
        .hIcon = Frm.Icon
    End With

    Call Shell_NotifyIcon(NIM_MODIFY, vbTray)

End Sub

Public Sub FormOnTop(Frm As Form)

'Puts your form ontop of all the other windows!
    Call SetWindowPos(Frm.HWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, wFlags)

End Sub

Public Sub PopupBalloon(Frm As Form, Message As String, Title As String)

'Set a Balloon tip on Systray

'Call RemoveBalloon(frm), This removes any current Balloon Tip that is active.
'If you want Balloon Tips to "Stack up" and display in sequence
'after each times out (or you click on the Balloon Tip to clear it),
'comment out the Call below.

    Call RemoveBalloon(Frm)

    With vbTray
        .cbSize = Len(vbTray)
        .HWnd = Frm.HWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY    'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Frm.Icon
        .dwState = 0
        .dwStateMask = 0
        .szInfo = Message & Chr(0)
        .szInfoTitle = Title & Chr(0)
        'Choose the message icon below, NIIF_NONE, NIIF_WARNING, NIIF_ERROR, NIIF_INFO
        .dwInfoFlags = NIIF_INFO
    End With

    Call Shell_NotifyIcon(NIM_MODIFY, vbTray)

End Sub

Public Sub RemoveBalloon(Frm As Form)

'Kill any current Balloon tip on screen for referenced form

    On Error Resume Next

    With vbTray
        .cbSize = Len(vbTray)
        .HWnd = Frm.HWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Frm.Icon
        .dwState = 0
        .dwStateMask = 0
        .szInfo = Chr(0)
        .szInfoTitle = Chr(0)
        .dwInfoFlags = NIIF_NONE
    End With

    Call Shell_NotifyIcon(NIM_MODIFY, vbTray)

    On Error GoTo 0
End Sub
