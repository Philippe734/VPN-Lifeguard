VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "À propos de..."
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSiteWeb 
      Caption         =   "Site web"
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Tag             =   "http://vpnlifeguard.blogspot.com"
      ToolTipText     =   "http://vpnlifeguard.blogspot.com"
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   300
      Left            =   3000
      TabIndex        =   1
      Top             =   3720
      Width           =   1005
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Open Source GNU/GPL"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1740
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "©2010 philippe734"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1350
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":08CA
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label lblNotice 
      BackStyle       =   0  'Transparent
      Caption         =   "Notice"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2760
      Width           =   1665
   End
   Begin VB.Line Line 
      X1              =   -480
      X2              =   4320
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "Faire un don"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   1
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Tag             =   "http://vpnlifeguard.blogspot.com/p/faire-un-don.html"
      ToolTipText     =   "Par CB ou Paypal"
      Top             =   3000
      Width           =   3675
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "Dernière version et code source"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   0
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3240
      Width           =   3795
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "VPN Lifeguard 1.X.YY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2280
   End
End
Attribute VB_Name = "frmAbout"
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
' Module    : frmAbout
' Author    : philippe734
' Date      : 22/04/2010
' Purpose   : affiche une fenêtre "à propos" de VPN Lifeguard
'---------------------------------------------------------------------------------------

Option Explicit

Private Sub cmdQuit_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cmdSiteWeb_Click()
    On Error Resume Next
    ShellExecute 0&, vbNullString, Me.cmdSiteWeb.Tag, vbNullString, vbNullString, vbNormalFocus
    DoEvents
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()

    MLSLoadLanguage Me '<- Add by Multi-Languages Support Add-in
    On Error Resume Next

    ' définit le nom et la version à afficher, en fonction qu'il soit portable ou pas
    If Dir$(App.Path & "\msvbvm60.dll") = vbNullString Then
        Me.lblTitle.Caption = App.Title & " - " & App.Major & "." & App.Minor & "." & App.Revision
    Else
        Me.lblTitle.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & " portable"
    End If

    ' Assigne le lien du site web au label dernière version
    Me.lblURL(0).Tag = pcSiteInternet
    Me.lblURL(0).ToolTipText = pcSiteInternet

    ' Flag son ouverture
    pbFormComplemantaireOuverte = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim item As Control

    ' Ne souligne plus les liens
    For Each item In frmAbout.Controls
        If TypeOf item Is Label Then
            item.FontUnderline = False
        End If
    Next item

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' Flag sa fermeture
    pbFormComplemantaireOuverte = False
End Sub


Private Sub lblNotice_Click()
    On Error Resume Next
    ShellExecute 0&, vbNullString, App.Path & "\Lisez-moi.rtf", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub lblNotice_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' Souligne le lien pointé par la souris
    Me.lblNotice.FontUnderline = True

    ' Change la souris en main
    MousePointerHand

End Sub

Private Sub lblURL_Click(Index As Integer)
    On Error Resume Next
    ShellExecute 0&, vbNullString, Me.lblURL(Index).Tag, vbNullString, vbNullString, vbNormalFocus
    DoEvents
End Sub

Private Sub lblURL_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

' Souligne le lien pointé par la souris
    Me.lblURL(Index).FontUnderline = True

    ' Change la souris en main
    MousePointerHand

End Sub

