VERSION 5.00
Begin VB.Form frmWidget 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2355
   ClientLeft      =   7935
   ClientTop       =   8955
   ClientWidth     =   5430
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBlack 
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   1440
      Picture         =   "frmWidget.frx":0000
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   2
      ToolTipText     =   "Connecté"
      Top             =   1200
      Width           =   810
   End
   Begin VB.PictureBox picRed 
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   2760
      Picture         =   "frmWidget.frx":0E6A
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   1
      ToolTipText     =   "Déconnecté"
      Top             =   1200
      Width           =   810
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quitter"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
End
Attribute VB_Name = "frmWidget"
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
' Module    : frmWidget
' Author    : philippe734
' Date      : 22/07/2010
' Purpose   : affiche un widget notifiant l'état du vpn
'---------------------------------------------------------------------------------------

Option Explicit

' définit l'objet de la class widget
' private car toute la durée de vie de la form
Private Widget As CWidget

' couleur du masque à rendre transparente
Private ColorMask As Long
'

Private Sub HideWidget()
    Unload Me
End Sub

Private Sub ShowApp()
'menu popup pour afficher l'interface
    On Error GoTo err:

10  Call SetForegroundWindow(frmMain.HWnd)
20  Call RemoveBalloon(frmMain)    's'il y a infobulle alors on la supprime

    'affiche l'interface
30  frmMain.WindowState = vbNormal
40  frmMain.Show
50  frmMain.SetFocus
60  Exit Sub

err:

    ' on n'affiche pas l'erreur si c'est l'affichage du form lorsqu'il y a une form modal affichée
70  If err.Number = 401 Then
80      err.Clear
90      Exit Sub
100 End If

110 Call GestionErreur(Me, "ShowApp", pcErreurMessageSupplementaire)
120 frmMain.Show
130 frmMain.SetFocus
140 If pbEnvoyerFormulaire = True Then
150     frmMain.WindowState = vbMinimized
160     Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
170 End If

End Sub


Private Sub QuitterProgramme()

    Dim item
    For Each item In frmMain.Controls
        If TypeOf item Is Timer Then
            item.Enabled = False
        End If
    Next item

    'ferme aussi la fenetre about si elle était ouverte
    Unload frmAbout
    Unload frmMain

End Sub

Private Sub cmdQuit_Click()
    HideWidget
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()

    MLSLoadLanguage Me    '<- Add by Multi-Languages Support Add-in
10  On Error GoTo Form_Load_Error

20  Set Widget = New CWidget

    ' définit la couleur de notre image qui sera transparente = le masque
30  ColorMask = &HFF4C&

    ' définit le niveau d'opacité de 0 à 255 = la transparence
40  Dim OpacityLevel As Byte: OpacityLevel = 200

    ' affecte la transparence et la forme à la fenêtre
50  Widget.DoShape Me, Me.picRed, ColorMask, OpacityLevel

    ' définit l'infobulle de la fenêtre
60  Widget.ToolTipCreate Me.HWnd, psWidgetToolTip

    ' crée le menu popup
70  Widget.MenuPopupAdd MLSGetString("0074"), MF_STRING    ' ret = 1 ' MLS-> "Masquer"
    ' Widget.MenuPopupAdd "-", MF_SEPARATOR
72  Widget.MenuPopupAdd "Show interface", MF_STRING    ' ret = 2
73  Widget.MenuPopupAdd "Reset VPN", MF_STRING    ' ret =3
74  Widget.MenuPopupAdd "Stop VPN", MF_STRING    ' ret =4
79  Widget.MenuPopupAdd "Quit", MF_STRING    ' ret = 5

    On Error Resume Next
    ' positionne le widget
80  Me.Top = IIf(LireINI("Divers", "Widget Top") <> "", LireINI("Divers", "Widget Top"), Me.Top)
90  Me.Left = IIf(LireINI("Divers", "Widget Left") <> "", LireINI("Divers", "Widget Left"), Me.Left)
    On Error GoTo Form_Load_Error

100 pbWidgetDisplayed = True

110 On Error GoTo 0
120 Exit Sub

Form_Load_Error:

130 Call GestionErreur(frmWidget, "Form_Load", pcErreurMessageSupplementaire)
140 If pbEnvoyerFormulaire = True Then
150     Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
160 End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ' Annule le clic et le remplace par la simulation d'un clic sur sa barre de titre
        ' donc ça permet de déplacer la fenêtre
        Widget.Moving Me
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then

        ' affiche le menu popup
        Dim RetourChoix As Long
        Widget.MenuPopup RetourChoix

        Select Case RetourChoix
        Case 1: Call HideWidget    ' masquer le widget
        Case 2: Call ShowApp    ' montrer l'interface
        Case 3: Call frmMain.ResetVPN
        Case 4: Call frmMain.StopVPN
        Case 5: Call QuitterProgramme
            '        Case 3
            '            ' permute l'image du widget
            '            ' affecte la transparence et la forme à la fenêtre
            '            Widget.DoShape Me, Me.picBlack, ColorMask, 200
            '            ' change l'infobulle
            '            Widget.ToolTipRemove
            '            Widget.ToolTipCreate Me.HWnd, "banana activated"
            '        Case 4
            '            ' permute l'image du widget
            '            ' affecte la transparence et la forme à la fenêtre
            '            Widget.DoShape Me, Me.picRed, ColorMask, 200
            '            Widget.ToolTipRemove
            '            Widget.ToolTipCreate Me.HWnd, "padlock activated"
            '        Case Else: HideWidget
        End Select

    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' pas besoin de restaurer l'aspect originale de notre fenêtre
' mais faut libérer la class
10  On Error GoTo Form_QueryUnload_Error

20  Set Widget = Nothing

    ' sauvegarde la position du widget
30  EcrireINI "Divers", "Widget Top", Me.Top
40  EcrireINI "Divers", "Widget Left", Me.Left

50  pbWidgetDisplayed = False

60  On Error GoTo 0
70  Exit Sub

Form_QueryUnload_Error:

80  Call GestionErreur(frmWidget, "Form_QueryUnload", pcErreurMessageSupplementaire)
90  If pbEnvoyerFormulaire = True Then
100     Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
110 End If
End Sub

Public Sub ChangeState(ByVal vpnConnected As Boolean)
' permute l'image du widget
' affecte la transparence et la forme à la fenêtre
' change l'infobulle

10  On Error GoTo ChangeState_Error

20  If vpnConnected = True Then
30      Widget.DoShape Me, Me.picBlack, ColorMask, 200
50      Widget.ToolTipCreate Me.HWnd, MLSGetString("0075")    ' MLS-> "Connecté"
60  Else
70      Widget.DoShape Me, Me.picRed, ColorMask, 200
90      Widget.ToolTipCreate Me.HWnd, MLSGetString("0076")    ' MLS-> "Déconnecté"
100 End If

110 On Error GoTo 0
120 Exit Sub

ChangeState_Error:

130 Call GestionErreur(frmWidget, "ChangeState", pcErreurMessageSupplementaire)
160 If pbEnvoyerFormulaire = True Then
180     Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
190 End If
End Sub
