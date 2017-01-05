VERSION 5.00
Begin VB.Form frmMessageErreur 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6690
   Icon            =   "frmMessageErreur.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timerDemarrage 
      Enabled         =   0   'False
      Left            =   240
      Top             =   840
   End
   Begin VB.TextBox txtEntete 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmMessageErreur.frx":000C
      Top             =   240
      Width           =   5655
   End
   Begin VB.CommandButton cmdQuitter 
      Caption         =   "OK"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Renvoie au formulaire de http://vpnlifeguard.blogspot.com en copiant l'erreur"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtErreur 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmMessageErreur.frx":00B8
      Top             =   960
      Width           =   5655
   End
   Begin VB.Image imageIcon 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   240
      Picture         =   "frmMessageErreur.frx":0146
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmMessageErreur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
' Module    : frmMessageErreur
' Author    : philippe734
' Date      : 22/04/2010
' Purpose   : affiche un message d'erreur et renvoie au formulaire de contact du site internet
' ce module est complémentaire de modGestionErreur
'---------------------------------------------------------------------------------------

Option Explicit

Private Sub cmdQuitter_Click()

    On Error Resume Next

    'copie le texte
    Clipboard.Clear
    Clipboard.SetText Me.txtErreur.Text & vbCrLf & vbCrLf

    'renvoie au formulaire
    ShellExecute 0&, vbNullString, pcURLformulaire, vbNullString, vbNullString, vbNormalFocus
    DoEvents
    pbEnvoyerFormulaire = True
    Unload Me
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()

    MLSLoadLanguage Me '<- Add by Multi-Languages Support Add-in

    On Error Resume Next

    'attribution des messages aux textbox
    Me.txtEntete.Text = psMessageEnteteErreur
    Me.txtErreur.Text = psMessageErreurCorps

    Me.Caption = App.Title
    Me.timerDemarrage.Interval = 1
    Me.timerDemarrage.Enabled = True
    Set Me.Icon = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error Resume Next

    If UnloadMode = 0 Then
        If MsgBox(MLSGetString("0051") & vbCrLf & MLSGetString("0052"), vbOKCancel + vbQuestion) = vbCancel Then ' MLS-> "Si vous ne voulez pas m'envoyer le rapport d'erreur, alors cliquez sur ok." ' MLS-> "Pour me l'envoyer, cliquez sur annuler."
            Cancel = True
        Else
            Cancel = False
        End If
    End If

    If Cancel = True Then
        Me.SetFocus
    Else
        pbFermetureMessageErreur = True
    End If
End Sub

Private Sub timerDemarrage_Timer()

    On Error Resume Next

    Me.timerDemarrage.Enabled = False
    Me.cmdQuitter.SetFocus
End Sub


Private Sub txtEntete_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next

    Me.cmdQuitter.SetFocus
End Sub

Private Sub txtErreur_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next

    Me.cmdQuitter.SetFocus
End Sub
