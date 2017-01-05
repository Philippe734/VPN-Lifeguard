Attribute VB_Name = "modGestionErreur"

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
' Module    : modGestionErreur
' Author    : philippe734
' Date      : 22/03/2010
' Purpose   : gestion d'erreur lié au site internet de VPN Lifeguard
'
' ce module est complémentaire de la form affichant le message d'erreur : frmMessageErreur
'---------------------------------------------------------------------------------------


Option Explicit

Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Const win95 As String = "Windows 95"
Private Const win98 As String = "Windows 98"
Private Const win2000 As String = "Windows 200"
Private Const XP As String = "XP"
Private Const Vista As String = "Vista"
Private Const win2003 As String = "Windows 2003"
Private Const Seven As String = "Seven"

'message d'erreur
Public psMessageEnteteErreur As String
Public psMessageErreurCorps As String
Public psMessageErreurComplet As String    'entete + erreur

'flag pour la fermeture du message d'erreur et l'envoie du formulaire
Public pbFermetureMessageErreur As Boolean
Public pbEnvoyerFormulaire As Boolean
'

Public Sub GestionErreur(Frm As Form, sNomProcedure As String, Optional sMessageIntroSupplementaire As String)
'gestion des erreurs affichant le nom de la procédure, la form et du NUMERO DE LIGNE
    Dim sMessage As String
    Dim sTEMP As String

    'On Error Resume Next
    'initialisation des flags
    pbFermetureMessageErreur = False
    pbEnvoyerFormulaire = False

    If sMessageIntroSupplementaire <> vbNullString Then
        psMessageEnteteErreur = sMessageIntroSupplementaire & vbCrLf & vbCrLf
    End If

    ' donne la version du programme en fonction qu'il soit portable ou pas
    If Dir$(App.Path & "\msvbvm60.dll") = vbNullString Then
        sTEMP = App.Major & "." & App.Minor & "." & App.Revision
    Else
        sTEMP = App.Major & "." & App.Minor & "." & App.Revision & " P"
    End If

    'mise en forme du message d'erreur
    sMessage = getVersion & " - " & sTEMP

    sMessage = sMessage & " - L'erreur " & err.Number & " s'est produite dans la fenêtre " & TypeName(Frm)
    sMessage = sMessage & " de la procédure " & sNomProcedure & " à la ligne " & IIf(Erl = 0, "(non spécifiée) ", Erl)
    sMessage = sMessage & " : " & err.Description & vbCrLf

    psMessageErreurCorps = sMessage
    psMessageErreurComplet = psMessageEnteteErreur + psMessageErreurCorps

    ' je désactive l'erreur de form modal car elle me prend la tête à débugger
    If err.Number = 401 Then
        Debug.Print Timer, "---"
        Debug.Print Timer, "Erreur de form modal, msg = "; psMessageErreurCorps
        Debug.Print Timer, "---"
        err.Clear
        Exit Sub
    End If

    err.Clear
    frmMessageErreur.Show vbModal
End Sub

Private Function getVersion(Optional ByRef AtLeastXP As Boolean) As String
    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer

    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)

    With osinfo
        Select Case .dwMajorVersion

        Case 4
            Select Case .dwMinorVersion
            Case 0
                getVersion = win95
                AtLeastXP = False
            Case 10
                getVersion = win98
                AtLeastXP = False
            End Select

        Case 5
            Select Case .dwMinorVersion
            Case 0
                getVersion = win2000
                AtLeastXP = False
            Case 1
                getVersion = XP
                AtLeastXP = True
            Case 2
                getVersion = win2003
                AtLeastXP = True
            End Select

        Case 6
            Select Case .dwMinorVersion
            Case 0
                getVersion = Vista
                AtLeastXP = True
            Case 1
                getVersion = Seven
                AtLeastXP = True
            End Select

        Case Else
            getVersion = "Failed"
            AtLeastXP = False
        End Select

    End With
End Function



