VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VPN Lifeguard"
   ClientHeight    =   3960
   ClientLeft      =   3795
   ClientTop       =   8940
   ClientWidth     =   12255
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   12255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timerInitialisation 
      Enabled         =   0   'False
      Left            =   9960
      Top             =   1920
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "Check IP"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   32
      Top             =   1680
      Width           =   975
   End
   Begin VB.Timer timerReDelRoute 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11160
      Top             =   1920
   End
   Begin VB.Frame Frame4 
      Caption         =   "Test"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6720
      TabIndex        =   29
      Top             =   3000
      Width           =   5175
      Begin VB.ComboBox comboDebugInfo 
         Height          =   330
         Left            =   3360
         TabIndex        =   31
         Text            =   "Debug Info"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtDebugInfo 
         Height          =   315
         Left            =   240
         TabIndex        =   30
         Text            =   "Debug Info"
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdTest4 
      Caption         =   "Test 4"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   28
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdTest3 
      Caption         =   "Add route"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   23
      Top             =   600
      Width           =   975
   End
   Begin VB.Timer timerPing 
      Enabled         =   0   'False
      Left            =   9360
      Top             =   1920
   End
   Begin VB.CheckBox chkDecoOFF 
      BackColor       =   &H80000000&
      Caption         =   "OFF/quit"
      Height          =   255
      Left            =   4500
      TabIndex        =   22
      ToolTipText     =   $"frmMain.frx":08CA
      Top             =   3650
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton cmdTest2 
      Caption         =   "Del route"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   24
      Top             =   1200
      Width           =   975
   End
   Begin VB.Timer timerTest 
      Enabled         =   0   'False
      Left            =   10560
      Top             =   1920
   End
   Begin VB.CommandButton cmdTest1 
      Caption         =   "Test 1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      TabIndex        =   27
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdConfigurer 
      Caption         =   "Config."
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      ToolTipText     =   "Configure les adresses IP et la liste des applications à gérer"
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "?"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      ToolTipText     =   "À propos de..."
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton cmdQuitter 
      Caption         =   "Quitter"
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      ToolTipText     =   "Déconnecte le VPN puis quitte"
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdDemarrer 
      Caption         =   "Démarrer"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      ToolTipText     =   "Démarre la connexion du VPN sélectionné dans la liste"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdArreter 
      Caption         =   "Arrêter"
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      ToolTipText     =   "Arrête la surveillance du VPN et le déconnecte"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Tag             =   "utorrent.exe"
      Top             =   1680
      Width           =   5415
      Begin VB.TextBox txtPremierLancement 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         ForeColor       =   &H80000017&
         Height          =   315
         Left            =   2640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "frmMain.frx":09A4
         Top             =   120
         Width           =   1575
      End
      Begin VB.CheckBox chkApplicationSupporte 
         Caption         =   "Libre2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   9
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkApplicationSupporte 
         Caption         =   "Libre3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkApplicationSupporte 
         Caption         =   "Libre5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3840
         TabIndex        =   12
         Top             =   720
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkApplicationSupporte 
         Caption         =   "Libre4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   11
         Top             =   720
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkApplicationSupporte 
         Caption         =   "Libre1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   8
         Top             =   480
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkApplicationSupporte 
         Caption         =   "Libre0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Applications à gérer"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   1665
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "IP local du VPN"
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   9360
      TabIndex        =   25
      Top             =   840
      Width           =   1935
      Begin VB.ComboBox comboListingIP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   26
         Text            =   "comboListingIP"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.ListBox listConnexionsReseaux 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Connexions VPN disponibles. Sélectionnez celle du VPN"
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Liste des réseaux"
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.Image ImageIconNormal 
      Height          =   480
      Left            =   8040
      Picture         =   "frmMain.frx":09B4
      Top             =   240
      Width           =   480
   End
   Begin VB.Image ImageIconDeconnected 
      Height          =   480
      Left            =   7320
      Picture         =   "frmMain.frx":127E
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblIpVpn 
      AutoSize        =   -1  'True
      Caption         =   "IP locale du VPN : xxx.xxx.xxx.xxx"
      Height          =   210
      Left            =   1680
      TabIndex        =   3
      Top             =   1395
      Width           =   2655
   End
   Begin VB.Label lblOperationEnCours 
      Caption         =   "labelStatus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   3360
      Width           =   1875
   End
   Begin VB.Label lblStatut 
      BackColor       =   &H80000000&
      Caption         =   "Service arrêté"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   3600
      Width           =   3735
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   -120
      TabIndex        =   21
      Top             =   3600
      Width           =   5895
   End
   Begin VB.Label lblDateDeconnexion 
      Caption         =   "labelDateDeco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   19
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Menu mnuPopupTrucs 
      Caption         =   "Menu Popup Del route"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupDelRoute 
         Caption         =   "Rendre le VPN exclusif"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Menu Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupAbout 
         Caption         =   "À propos"
      End
      Begin VB.Menu mnuPopupSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupArreter 
         Caption         =   "Arrêter VPN"
      End
      Begin VB.Menu mnuPopupResetVPN 
         Caption         =   "Réinitialiser VPN"
      End
      Begin VB.Menu mnuPopupLog 
         Caption         =   "Consulter le log"
      End
      Begin VB.Menu mnuPopupCheckUpdate 
         Caption         =   "Vérifier les mises à jour"
      End
      Begin VB.Menu mnuPopupWidget 
         Caption         =   "Afficher le widget"
      End
      Begin VB.Menu mnuPopupShowMe 
         Caption         =   "Afficher l'interface"
      End
      Begin VB.Menu mnuPopupSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupQuit 
         Caption         =   "Quitter"
      End
   End
End
Attribute VB_Name = "frmMain"
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
' Module    : frmMain
' Author    : philippe734
' Date      : 22/06/2010
' Purpose   : interface principale de VPN Lifeguard
' Caractéristiques :
' - connexion du VPN
' - surveille la connexion du VPN par ping
' - ferme les logiciels configurés à gérer en cas de déconnexion
' - reconnexion du VPN
' - recharge les logiciels fermés après reconnexion
' - connexion VPN exclusif = être sûr de passer par le VPN
'---------------------------------------------------------------------------------------


Option Explicit

'bouton croix de la form pour gérer sa réduction dans le systray
Private pbBoutonCroix As Boolean

Private WithEvents m_clsRasdial As CMultithreadingRasdial
Attribute m_clsRasdial.VB_VarHelpID = -1
'


Public Sub CreateThreadPourRasdial()
    Dim clsThread As CMultithreadingSet
    Dim iThreadIndex As Long

    ' création d'une thread en multithreading
    ' pour exécuter la connexion vpn via l'api rasdial
    ' afin de ne pas geler le programme durant connexion en cours

    On Error GoTo err

    Set clsThread = CreateObject("vbpVpnLifeguard.CMultithreadingSet")

    Set m_clsRasdial = CreateObject("vbpVpnLifeguard.CMultithreadingRasdial")

suite:

    Call clsThread.SetThread(m_clsRasdial, "RasDialThreaded", iThreadIndex, psConnectionName)

    Set clsThread = Nothing

    On Error GoTo 0
    Exit Sub

err:

    ' exécute les opérations en mode normale plutot qu'en multithreading
    Set clsThread = New CMultithreadingSet

    Set m_clsRasdial = New CMultithreadingRasdial

    Resume suite

End Sub



Private Sub chkApplicationSupporte_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

' change la souris en main
    MousePointerHand

End Sub

Private Sub cmdArreter_Click()

    Call StopByUser
    ' bouton arrêter

End Sub

Private Sub cmdPing_Click()

10  On Error GoTo err

    Dim mess As String, ret As Boolean

20  ret = DoPing2(psAdresseIPpasserelle)

30  mess = MLSGetString("0001") & vbNewLine & psAdresseIPpasserelle & " = " & ret & vbNewLine    ' MLS-> "Adresses IP contrôlées :"

40  ret = DoPing2(psAdressIPvpn)

50  mess = mess & psAdressIPvpn & " = " & ret

60  MsgBox mess, vbInformation

70  On Error GoTo 0
80  Exit Sub

err:
90  MsgBox "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in cmdPing_Click of frmMain", vbCritical

End Sub

Private Sub Form_Activate()
    Static Flag As Boolean
    Dim varTemp As String

10  On Error Resume Next

20  If Flag = True Then Exit Sub
30  Flag = True

    ' on définit la position de la fenetre
40  Debug.Print Timer, "Set form position"
50  varTemp = LireINI("Paramètres de la fenêtre", "Left")
60  If Val(varTemp) > 0 Then
70      frmMain.Left = CSng(varTemp)
80  End If
90  varTemp = LireINI("Paramètres de la fenêtre", "Top")
100 If Val(varTemp) > 0 Then
110     frmMain.Top = CSng(varTemp)
120 End If

130 Me.Refresh

140 DoEvents

150 On Error GoTo err

    'réduit la fenetre si l'option a été cochée
160 If pbAutoReduire = True Then Me.WindowState = vbMinimized

    'crée son icone dans le systray
170 Call SystrayOn(Me, App.Title & vbCrLf & MLSGetString("0032"))    ' MLS-> "Service arrêté"

180 Me.Icon = Me.ImageIconDeconnected

    'efface la liste des connexions réseaux précédente
190 Call RemplirListeReseaux

    'sélectionne la 1ère connexion réseau
200 If Me.listConnexionsReseaux.ListCount > 0 Then
210     Me.listConnexionsReseaux.ListIndex = 0
220     pbConnexionReseauPresente = True
230 Else
240     pbConnexionReseauPresente = False
250 End If

    ' Premier démarrage
260 Call Initialization

270 Exit Sub

err:
280 Call GestionErreur(Me, "Form_Activate", pcErreurMessageSupplementaire)
290 Me.Show
300 Me.SetFocus
310 If pbEnvoyerFormulaire = True Then
320     Me.WindowState = vbMinimized
330     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
340 End If
End Sub

Private Sub lblIpVpn_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        ' Affiche le menu popup pour choisir de rendre exclusif le VPN
        PopupMenu mnuPopupTrucs
    End If
End Sub

Private Sub m_clsRasdial_Error(ByVal Message As String)

    On Error GoTo m_clsRasdial_Error_Error

    ' la thread est terminée alors on libère la classe
    Set m_clsRasdial = Nothing

    err.Raise 9999, , Message

    pbFinThread = True

    On Error GoTo 0
    Exit Sub

m_clsRasdial_Error_Error:

    Call GestionErreur(frmMain, "m_clsRasdial_Error", pcErreurMessageSupplementaire)
    If pbEnvoyerFormulaire = True Then
        Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
    End If

End Sub

Private Sub m_clsRasdial_Result(ByVal HandleRasConn As Long, ByVal ReturnCode As Long)
' event en retour de la connexion vpn via rasdial
' initialement exécuté en multithreading

' la thread est terminée alors on libère la classe
    Set m_clsRasdial = Nothing

    ' copie public
    phRasConnThreaded = HandleRasConn
    plngRetCodeThreaded = ReturnCode

    ' signale la fin de la thread
    pbFinThread = True

    ' la suite de l'exécution du programme se trouve
    ' dans la procédure AutoDialThreaded juste après
    ' la boucle qui attendait la fin de la thread

End Sub

Private Sub cmdAbout_Click()
'affiche la fenetre à propos

'aucune action si une thread est en cours
'c'est à dire, si l'opération en cours est "connexion en cours"
    If pbThreadEnCours = True Then Exit Sub

    On Error Resume Next
    frmAbout.Show vbModal
    On Error GoTo 0

End Sub

Private Sub cmdConfigurer_Click()
    On Error GoTo err:

'aucune action si une thread est en cours
'c'est à dire, si l'opération en cours est "connexion en cours"
10  If pbThreadEnCours = True Then Exit Sub

    ' si on vient juste démarrer et que l'autodémarrage est actif
    ' et qu'internet n'est pas accessible alors on désactive le démarrage
20  If pbJustLoaded = True And pbAutoDemarrer = True Then
30      Call ArreterVPN
40  End If


    '--------------------------------------------
    'fin de la sécurisation du tunnel du vpn
    ' add route passerelle afin de récupérer la passerelle
    ' add route uniquement si le vpn n'est pas connecté
    If pbConnexionVPN = False Then
50      If pbConfigSecurisationDuTunnel = True And psAdresseIPpasserelle <> vbLocalHost Then
60          Call modGlobal.SecurisationDuTunnelVPN(desactiver)
70      End If
    End If

80  Call ConfigurerAppli

90  Exit Sub

err:
100 Call GestionErreur(Me, "cmdConfigurer_Click", pcErreurMessageSupplementaire)
110 Me.Show
120 Me.SetFocus
130 If pbEnvoyerFormulaire = True Then
140     Me.WindowState = vbMinimized
150     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
160 End If
End Sub

Private Sub cmdTest1_Click()

'bouton de test, pour tester des trucs justement...

10  On Error GoTo cmdTest1_Click_Error



70  On Error GoTo 0
80  Exit Sub

cmdTest1_Click_Error:
90  MsgBox "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in cmdTest1_Click of frmMain", vbCritical
End Sub

Private Sub cmdTest3_Click()


    Exit Sub

10  On Error GoTo cmdTest3_Click_Error

    '-----------------------------------------------------
    '    ' add route
    '-----------------------------------------------------

    Dim clsN As CGetIpgateway
    Dim sAdresseTemp As String
    Dim iIndexIP As Long

    sAdresseTemp = "192.168.1.1"

90  Set clsN = New CGetIpgateway

100 iIndexIP = clsN.GetIndexIP(sAdresseTemp)

110 Set clsN = Nothing

130 piIndexPasserelle = iIndexIP
140 psAdresseIPpasserelle = sAdresseTemp

150 Call modGlobal.SecurisationDuTunnelVPN(desactiver)

160 On Error GoTo 0
170 Exit Sub

cmdTest3_Click_Error:
180 MsgBox "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in cmdTest3_Click of frmMain", vbCritical
End Sub

Private Sub cmdTest2_Click()
    Exit Sub

10  On Error GoTo cmdTest2_Click_Error

    ' del route

    Dim sAdresseTemp As String

    sAdresseTemp = "192.168.1.1"

90  psAdresseIPpasserelle = sAdresseTemp

110 Call modGlobal.SecurisationDuTunnelVPN(activer)

120 On Error GoTo 0
130 Exit Sub

cmdTest2_Click_Error:
140 MsgBox "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in cmdTest2_Click of frmMain", vbCritical
End Sub

Private Sub cmdTest4_Click()
    Exit Sub
10  On Error GoTo cmdTest4_Click_Error







170 On Error GoTo 0
180 Exit Sub
cmdTest4_Click_Error:
190 MsgBox "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in cmdTest4_Click of frmMain", vbCritical
End Sub

Private Sub Form_Initialize()

'donne le style de windows au programme
    InitCommonControls

    ' active ou désactive le log
    pbLogActif = False

    'MyLog "in form initialize"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static lngMsg As Long    'récupère le signal du curseur
    Static bFlag As Boolean    'indicateur pour éviter les conneries

    On Error GoTo err:

'aucune action à faire si une autre form est déjà ouverte
10  If pbFormComplemantaireOuverte = True Then Exit Sub

    'le flag sert à limiter l'action
30  If bFlag = True Then Exit Sub
40  bFlag = True

20  lngMsg = x / Screen.TwipsPerPixelX

50  Select Case lngMsg

    Case WM_RBUTTONDOWN    'clic droit pour afficher le menu popup
60      Call SetForegroundWindow(Me.HWnd)
70      Call RemoveBalloon(Me)    's'il y a infobulle alors on la supprime
        'affiche le menu popup, configuré dans le menu editor...
80      PopupMenu Me.mnuPopup

        'MyLog "in mousemove, juste après popup menu"

90  Case WM_LBUTTONDBLCLK  'double clic gauche, affiche la fenetre
        'utiliser la ligne ci-dessous si on veut enlever l'icone du systray
        'Call SystrayOff(Me)
100     Call SetForegroundWindow(Me.HWnd)
110     Call RemoveBalloon(Me)    's'il y a infobulle alors on la supprime

        'affiche la fenetre
        On Error Resume Next
120     Me.WindowState = vbNormal
130     Me.Show
140     Me.SetFocus

150 End Select
160 bFlag = False

170 Exit Sub

err:
180 Call GestionErreur(Me, "Form_MouseMove", pcErreurMessageSupplementaire)
190 Me.Show
200 Me.SetFocus
210 If pbEnvoyerFormulaire = True Then
220     Me.WindowState = vbMinimized
230     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
240 End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Byte

    On Error Resume Next

    'récupère la manière de quitter
    ' 0 = bouton croix
    ' 1 = par l'interface comme un bouton
    ' 2 = éteindre ou redémarrer le PC
    Select Case UnloadMode
    Case 0
        ' on annule la fermeture que si l'option réduire dans le systray
        ' en quittant a été cochée
        Cancel = IIf(pbOptionReduireSystrayQuit, True, False)
    Case 1, 2
        ' demande de quitter via l'interface
        Cancel = False
    End Select

    ' si une thread est en cours alors on annule la fermeture du programme
    If pbThreadEnCours = True Then
        Cancel = True
    End If

    'sauvegarde les données
    If Cancel = False Then

        'MyLog "in queryunload, quand cancel=false"

        ' signale la fermeture du programme
        Me.lblOperationEnCours.Caption = MLSGetString("0002")    ' MLS-> "Fermeture en cours..."

        ' fermeture des applications gérées
        ' si l'option a été cochée
        If pbCloseAppliOnQuit = True Then
            Call FermerApplications
        End If

        'sauvegarde les coches des cases à cocher
        For i = 0 To Me.chkApplicationSupporte.Count - 1
            'sauvegarde dans le fichier ini
            EcrireINI "Paramètres des applications", "Checkbox" & i, Me.chkApplicationSupporte(i).value
        Next i

        'sauvegarde la position de la fenetre
        If Me.Left > 0 And Me.Top > 0 Then
            EcrireINI "Paramètres de la fenêtre", "Left", Me.Left
            EcrireINI "Paramètres de la fenêtre", "Top", Me.Top
        End If

        'sauvegarde l'IP pingé dans le fichier ini
        EcrireINI "Paramètres du VPN", "Adresse IP", psAdressIPvpn

        ' si redémarrage ou extinction du PC, alors pas necessaire de déconnecter VPN
        If UnloadMode <> 2 Then
            'déconnexion vpn et déblocage globale en quittant
            If Me.chkDecoOFF.value = 1 Then
                Call ArreterVPN
            Else
                'attention : ne restaure pas la passerelle
                'donc impossible de reconnecter le vpn après
            End If
        End If

        Me.lblOperationEnCours.Caption = MLSGetString("0003")    ' MLS-> "Fermeture en cours..."
        Me.lblOperationEnCours.Refresh

        'MyLog "in queryunload, juste avant unload autreform"

        Dim AutreForm As Form
        'fermeture des autres fenetres
        For Each AutreForm In Forms
            If AutreForm.Name <> Me.Name Then Unload AutreForm
        Next AutreForm

        'MyLog "in queryunload, juste après unload autreform"

        ' enlève l'icone du systray
        Call SystrayOff(Me)

        Debug.Print Timer, vbNewLine & "--- Fin   ----------------------------------------------------"

    Else    'Cancel = true
        ' fermeture annulée, on réduit la fenetre

        pbBoutonCroix = True
        Me.WindowState = vbMinimized
    End If

End Sub

Private Sub Form_Resize()

'aucune action si une thread est en cours
'c'est à dire, si l'opération en cours est "connexion en cours"
    If pbThreadEnCours = True Then Exit Sub

    'MyLog "in form resize"

    If Me.WindowState = vbMinimized Then
        'c'est pour avoir l'animation de réduction
        'qu'on utilise vbMinimized

        If pbBoutonCroix = True Then
            Me.Hide
            pbBoutonCroix = False
        End If
    End If
End Sub

Private Sub mnuPopupArreter_Click()

    Call StopByUser
    ' menue pop arrêter

End Sub

Private Sub mnuPopupAbout_Click()
'affiche la fenetre A propos
    On Error GoTo err:
    frmAbout.Show
    Exit Sub

err:
    Call GestionErreur(Me, "mnuPopupAbout_Click", pcErreurMessageSupplementaire)
End Sub

Private Sub mnuPopupCheckUpdate_Click()

    On Error GoTo mnuPopupCheckUpdate_Click_Error

    frmCheckUpdate.Show

    On Error GoTo 0
    Exit Sub

mnuPopupCheckUpdate_Click_Error:

    Call GestionErreur(frmMain, "mnuPopupCheckUpdate_Click", pcErreurMessageSupplementaire)

End Sub

Private Sub mnuPopupDelRoute_Click()

    If psAdresseIPpasserelle <> vbLocalHost Then
        ' Rend le VPN exclusif en supprimant les autres routes
        ' Del route de la passerelle
        Call modGlobal.SecurisationDuTunnelVPN(activer)
    End If

End Sub

Private Sub mnuPopupLog_Click()
    Dim Fich As String

    On Error Resume Next

    Fich = App.Path & "\" & pcFichLog & ".log"

    If IsFileExist(Fich) = True Then
        ' ouvre le log de déconnexions
        ShellExecute Me.HWnd, "Open", Fich, vbNullString, vbNullString, vbNormalFocus
    Else
        MsgBox MLSGetString("0004"), vbInformation, App.Title    ' MLS-> "Aucun Log présent. Vous pouvez le créer en l'activant dans les options."
    End If

    On Error GoTo 0
End Sub

Private Sub mnuPopupShowMe_Click()
'menu popup Afficher la fenetre
    On Error GoTo err:

10  Call SetForegroundWindow(Me.HWnd)
20  Call RemoveBalloon(Me)    's'il y a infobulle alors on la supprime

    'affiche la fenetre
30  Me.WindowState = vbNormal
40  Me.Show
50  Me.SetFocus
60  Exit Sub

err:

    ' on n'affiche pas l'erreur si c'est l'affichage du form lorsqu'il y a une form modal affichée
70  If err.Number = 401 Then
80      err.Clear
90      Exit Sub
100 End If

110 Call GestionErreur(Me, "mnuPopupShowMe_Click", pcErreurMessageSupplementaire)
120 Me.Show
130 Me.SetFocus
140 If pbEnvoyerFormulaire = True Then
150     Me.WindowState = vbMinimized
160     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
170 End If
End Sub

Private Sub mnuPopupQuit_Click()

    Dim item
    For Each item In frmMain.Controls
        If TypeOf item Is Timer Then
            item.Enabled = False
        End If
    Next item

    'ferme aussi la fenetre about si elle était ouverte
    Unload frmAbout
    Unload Me
End Sub

Private Sub mnuPopupWidget_Click()
' affiche le widget de notification d'état du vpn
    frmWidget.Show
    frmWidget.ChangeState pbConnexionVPN
End Sub

Private Sub mnuPopupResetVPN_Click()

' Déconnecte puis reconnecte le VPN
    Call ResetVPN

End Sub

Private Sub timerInitialisation_Timer()
    Me.timerInitialisation.Enabled = False
    Call Initialization
End Sub

Private Sub timerPing_Timer()
    Dim i As Integer
    Dim sTEMP As String
    Static bFlag As Boolean

    On Error GoTo err:

10  If bFlag = True Then Exit Sub
20  bFlag = True


    'Debug.Print timer, "timerPing cycle en cours..."

    'si le vpn est connecté alors on récupère son IP puis on ping
    'sinon on va direct au ping = mauvais
30  If pbConnexionVPN = True Then

        'change le temps du timer pour accélerer le ping
40      If Me.timerPing.Interval <> pcTempsPing Then
50          Me.timerPing.Interval = pcTempsPing


            '------------------------------------------------------
            'on récupère OBLIGATOIREMENT la nouvelle IP du vpn

60          Debug.Print Timer, "retrouve l'IP du vpn en cours..."

70          Me.comboListingIP.Locked = False

            'boucle pour récupérer l'IP afin de contrer les reconnexions/déconnexions rapides
            Dim bGetIpVpn As Boolean
            Dim RAScom As RASCONNSTATE
80          Do
90              Me.comboListingIP.Clear
100             bGetIpVpn = modGlobal.GetIpVpn(Me.comboListingIP)
110             RAScom = modRasFunction.GetStatusRasconn(GetCurrentHandleRasConnFromAutodial)
120             DoEvents
130             Sleep 100
140             Debug.Print Timer, "comboListingIP.ListCount = "; Me.comboListingIP.ListCount
150             Debug.Print Timer, "bGetIpVpn = "; bGetIpVpn
                'Debug.Print timer, "RAScom = "; RAScom
160         Loop Until bGetIpVpn = True Or RAScom = RASCS_Disconnected

            'pour tester =
            'Me.comboListingIP.Clear

170         psAdressIPvpn = Me.comboListingIP
180         Debug.Print Timer, "IP vpn = " & psAdressIPvpn
190         Me.comboListingIP.Locked = True
            'End If
            'Fin de récupération de l'adresse IP du vpn
            '---------------------------------------------------



            '---------------------------------------------------
            '
            ' CONNEXION VPN ET IP CORRECTES
            '
200         If pbConnexionVPN = True And bGetIpVpn = True Then

210             Me.lblIpVpn.Caption = MLSGetString("0005") & " " & psAdressIPvpn   ' MLS-> "IP locale du VPN : "

                'sauvegarde l'IP pingé dans le fichier ini
220             EcrireINI "Paramètres du VPN", "Adresse IP", psAdressIPvpn

                'si le VPN a été reconnecté alors...
                'efface l'infobulle de la déconnexion
230             Call RemoveBalloon(Me)

                'actualise le statut
240             If piNombreDeconnexion < 2 Then
250                 sTEMP = piNombreDeconnexion & " " & MLSGetString("0006")    ' MLS-> " déconnexion"
260             Else
270                 sTEMP = piNombreDeconnexion & " " & MLSGetString("0007")   ' MLS-> " déconnexions"
280             End If
290
                'charge l'icone normale
300             Me.Icon = Me.ImageIconNormal

                ' change l'image du widget
310             If pbWidgetDisplayed = True Then
320                 frmWidget.ChangeState True
330             End If

340             Me.lblOperationEnCours.Caption = sTEMP
                'modifie l'infobulle de l'icone dans le systray
350             Call ChangeSystrayToolTip(Me, App.Title & vbCrLf & sTEMP)

                ' écrit log connexion
360             If pbLogDeconnexion = True Then
370                 EcrireLogDeconnexion False
380             End If


390         End If    'de If pbConnexionVPN = True And bGetIpVpn = True
400         Debug.Print Timer, "Ping OK à "; Time
410     End If    'de If Me.timerPing.Interval = ...
420 End If    'de If pbConnexionVPN = true


    'pour éviter les erreurs de ping si l'adresse ip vpn n'est pas une adresse
430 If Len(psAdressIPvpn) > Len("0.0.0.") Then
        Dim ret As Boolean
        'ping
440     For i = 1 To 3
450         ret = ret Or DoPing2(psAdressIPvpn)
460     Next i
470 Else
480     ret = False
490 End If
500
510 If ret = True Then
        'ping OK
        'Debug.Print timer, "ping OK à " & Time
520     pbCycleDeconnexionReconnexionEnCours = False


        ' controles supplémentaires avant de relancer les applications
530     If pbFlagAppliAutoDemarrer = True And pbConnexionVPN = True And pbPopUpArretEnCours = False Then
540         pbFlagAppliAutoDemarrer = False
550         Call RelancerAppli
560     End If


        'VPN toujours connecté donc ne rien faire sauf si faut relancer les appli
570     If pbRelancerApplications = True Then

            'il y a eu déconnexion puis reconnexion donc relancer les applications
580         pbRelancerApplications = False
590         sTEMP = MLSGetString("0008")    ' MLS-> "Chargement des appli..."

            'actualise le statut
600         Me.lblOperationEnCours.Caption = sTEMP
610         Call ChangeSystrayToolTip(Me, App.Title & vbCrLf & sTEMP)

            ' controles supplémentaires avant de relancer les applications
            If pbConnexionVPN = True And pbPopUpArretEnCours = False Then
620             Call RelancerAppli
            End If

            'actualise le statut
630         If piNombreDeconnexion < 2 Then
640             sTEMP = piNombreDeconnexion & " " & MLSGetString("0009")   ' MLS-> " déconnexion"
650         Else
660             sTEMP = piNombreDeconnexion & " " & MLSGetString("0010")   ' MLS-> " déconnexions"
670         End If
680
690         Me.lblOperationEnCours.Caption = sTEMP
700         Call ChangeSystrayToolTip(Me, App.Title & vbCrLf & sTEMP)
710     End If
720 Else

        '----------------------------------------------------------
        '
        '   PING MAUVAIS = déconnexion du VPN
        '
730     Debug.Print Timer, "Ping mauvais"

        ' chronomètre
        piChrono = Timer

740     Me.timerPing.Enabled = False

750     pbCycleDeconnexionReconnexionEnCours = True

        'déconnexion = fermeture des applications
760     Call FermerApplications

        'change l'icone du systray
770     Me.Icon = Me.ImageIconDeconnected

        ' notifie le widget
780     If pbWidgetDisplayed = True Then
790         frmWidget.ChangeState False
800     End If

        'actualise le statut avec infobulle de la déconnexion
810     Call PopupBalloon(Me, App.Title, MLSGetString("0011"))    ' MLS-> "Déconnexion mais reconnexion en cours..."

        'actualise le statut
820     sTEMP = MLSGetString("0012")    ' MLS-> "Fermeture des appli..."
830     Me.lblOperationEnCours.Caption = sTEMP
840     Call ChangeSystrayToolTip(Me, App.Title & vbCrLf & sTEMP)

        'si le VPN était connecté alors mise à jour du statut de la déconnexion
850     If pbConnexionVPN = True Then
860         piNombreDeconnexion = piNombreDeconnexion + 1
            '870         psDerniereDeconnexion = "Dernière déconnexion à " & FormatDateTime(Time, vbShortTime) & " le " & FormatDateTime(Date, vbShortDate)
870         psDerniereDeconnexion = MLSGetString("0077") & " " & FormatDateTime(Time, vbShortTime) & " " & MLSGetString("0078") & " " & FormatDateTime(Date, vbShortDate)
880         Me.lblDateDeconnexion = psDerniereDeconnexion
890     End If

        'déconnexion rasdial pour débloquer son port
900     Call DeconnexionDuVpn

        '--------------------------------------------
        'fin de la sécurisation du tunnel du vpn
        ' add route passerelle
910     If pbConfigSecurisationDuTunnel = True And psAdresseIPpasserelle <> vbLocalHost Then
920         Call modGlobal.SecurisationDuTunnelVPN(desactiver)
930     End If

        ' écrit dans le log
940     If pbLogDeconnexion = True Then
950         EcrireLogDeconnexion True
960     End If

        ' temps mis pour fermer les applications en ms
        piChrono = 1000 * (Timer - piChrono)

        ' txtDebugInfo.Text = "durée = " & piChrono & " ms"

        'reconnecter le VPN = recommencer
970     cmdDemarrer_Click
980     Debug.Print Timer, "relance les opérations avec cmdDemarrer_Click"
990 End If    'test du ping

1000 bFlag = False

1010 Exit Sub

err:
1020 Call GestionErreur(Me, "timerPing_Timer", pcErreurMessageSupplementaire)
1030 Me.Show
1040 Me.SetFocus
1050 If pbEnvoyerFormulaire = True Then
1060    Me.WindowState = vbMinimized
1070    Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
1080 End If
End Sub

Private Sub ConnexionVPN()
    Dim sTEMP As String
    Dim RASconnStatus As RASCONNSTATE

    On Error GoTo err:

    'actualise le statut
10  sTEMP = MLSGetString("0013")    ' MLS-> "Connexion en cours..."
20  Me.lblOperationEnCours.Caption = sTEMP
30  Call ChangeSystrayToolTip(Me, App.Title & vbCrLf & sTEMP)
40  Me.Refresh
50  Me.lblIpVpn.Caption = MLSGetString("0014")    ' MLS-> "IP locale du VPN :  . . ."
60  DoEvents


    'fermeture des autres fenetres afin d'éviter les crash en multi-threading
    Dim AutreForm As Form
    'fermeture des autres fenetres
70  For Each AutreForm In Forms
80      If (AutreForm.Name <> Me.Name) And (AutreForm.Name <> frmWidget.Name) Then Unload AutreForm
90  Next AutreForm


    'reconnexion du VPN
100 Debug.Print Timer, "connexion en cours..."

    'RasDial sans multi-threading
    'If AutoDial(listConnexionsReseaux.List(listConnexionsReseaux.ListIndex)) <> 0 Then

    'RasDial en multi-threading
110 If AutoDialThreaded(listConnexionsReseaux.List(listConnexionsReseaux.ListIndex)) <> 0 Then

120     Debug.Print Timer, "connexion ok"


130     Do
140         RASconnStatus = modRasFunction.GetStatusRasconn(modRasFunction.GetCurrentHandleRasConnFromAutodial)
            'Debug.Print timer, "RASconnStatus = " & RASconnStatus
            Dim iTimer As Single
150         iTimer = Timer
            Do
                DoEvents
160         Loop Until Timer - iTimer >= 1
170     Loop Until RASconnStatus = RASCS_Connected Or RASconnStatus = RASCS_Disconnected

180     Select Case RASconnStatus
        Case RASCS_Connected: pbConnexionVPN = True
190     Case RASCS_Disconnected: pbConnexionVPN = False: psAdressIPvpn = vbNullString
200     End Select


210     If pbConnexionVPN = True Then
            '--------------------------------------------
            'sécurisation du tunnel du vpn
            ' del route passerelle
220         If pbConfigSecurisationDuTunnel = True And psAdresseIPpasserelle <> vbLocalHost Then
230             Call modGlobal.SecurisationDuTunnelVPN(activer)
240         End If

            'actualise le statut
250         Me.lblOperationEnCours.Caption = MLSGetString("0015")    ' MLS-> "Connexion effectuée"

            ' active un timer qui refait del route toutes les heures
            Me.timerReDelRoute.Enabled = True

260     End If

        'flag l'action
270     pbInitialisation = False

280 Else

        'MsgBox "dans if AutoDialThreaded<>0 = false", vbInformation

        'flag l'action
290     pbConnexionVPN = False
300     Debug.Print Timer, "Connexion au vpn impossible"
310 End If

320 Me.Refresh
330 DoEvents

340 Exit Sub

err:
350 Call GestionErreur(Me, "ConnexionVPN", pcErreurMessageSupplementaire)
360 Me.Show
370 Me.SetFocus
380 If pbEnvoyerFormulaire = True Then
390     Me.WindowState = vbMinimized
400     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
410 End If
End Sub

Private Sub cmdDemarrer_Click()    'Bouton Démarrer
    Dim sTmp As String

    On Error GoTo err:


    'MyLog "au début de cmddémarrer"



10  If Me.listConnexionsReseaux.ListCount = 0 Then
        'si aucune connexion VPN trouvée alors on quitte
20      MsgBox MLSGetString("0016") _
             & Chr(34) & MLSGetString("0017") & Chr(34) & " " & MLSGetString("0018"), vbExclamation, MLSGetString("0019")   ' MLS-> " de windows mais pas avec celle de type OpenVPN. Donc, si vous avez une connexion VPN, alors elle est peut-être de ce type." ' MLS-> "Aie ! Connexion VPN absente"
        'MLS->"Aucune connexion VPN présente. Le programme ne fonctionne qu'avec les connexions VPN affichées dans "
        'MLS->"connexion réseau"

30      Exit Sub
40  End If



    'vérifie si l'IP de la passerelle est accessible
    'en faisant un ping
50  If Len(psAdresseIPpasserelle) > Len("0.0.0.0") Then
        Dim bPing As Boolean
60      bPing = DoPing2(psAdresseIPpasserelle)
70      Debug.Print Timer, "Ping passerelle = "; bPing

        ' option spéciale pour la version portable
        If Dir$(App.Path & "\msvbvm60.dll") <> vbNullString Then
            ' on va supposer qu'on a toujours accès au web
            bPing = True
        End If

80      If bPing = False Then
            'la passerelle n'est pas accessible alors on le signal
90          Call ArreterVPN
100         Me.lblStatut.Caption = MLSGetString("0020")    ' MLS-> "Internet inaccessible"


            ' deux cas possibles :
            ' soit c'est le lancement du programme
            ' soit c'est un cycle déconnexion / reconnexion

            ' si c'est le lancement du programme alors
            ' on affiche la fenetre au bout de 3 heures
110         If pbJustLoaded = True And pbAutoDemarrer = True Then

                Static iTentativeLancementAuto As Long
120             iTentativeLancementAuto = iTentativeLancementAuto + 1
130             If iTentativeLancementAuto > 2700 Then    ' 3 heures
140                 pbInternetInaccessible = True
150             End If

160         Else
                ' alors on est dans un cycle déco / reco
                ' donc on laisse tourner le programme
                ' sans limiter le nombre de tentatives
170         End If

            ' comme le ping est mauvais alors on
            'actualise le statut
180         sTmp = MLSGetString("0021")    ' MLS-> "En attente d'Internet..."
190         Me.lblOperationEnCours.Caption = sTmp
200         Call ChangeSystrayToolTip(Me, App.Title & vbCrLf & sTmp)

            ' affiche une info bulle
            ' seulement pour la première tentative de reconnexion
            Static bPremierEssai As Boolean
210         If bPremierEssai = False Then
220             bPremierEssai = True
230             Call PopupBalloon(Me, App.Title, MLSGetString("0022"))    ' MLS-> "En attente d'Internet..."
240         End If

            ' réactive l'initialisation afin de
            ' faire une autre tentative de connexion
250         Me.timerInitialisation.Interval = 4000    ' 2700x4 = 3 heures
260         Me.timerInitialisation.Enabled = True

            ' on sort et on recommance
270         Exit Sub

280     End If
290 End If


    '----------------------------------------------
    ' PASSERELLE ACCESSIBLE ---- sPing = True
    '----------------------------------------------

    ' on remet la route de la passerelle
    ' au cas ou il y aurait eu un crash
300 If pbJustLoaded = True Then
        '--------------------------------------------
        'sécurisation du tunnel du vpn
        ' add route passerelle
310     If psAdresseIPpasserelle <> vbLocalHost Then
320         Call modGlobal.SecurisationDuTunnelVPN(desactiver)
330     End If
340 End If

    'change l'indicateur de premier lancement
350 pbJustLoaded = False

    'ré-initialise le compteur de tentative
360 iTentativeLancementAuto = 0

    'internet n'est donc pas inaccessible
370 pbInternetInaccessible = False

    ' on sauvegarde l'index de la listebox des réseaux
380 EcrireINI "Paramètres du VPN", "N° connexion choisis ; listConnexionsReseaux.ListIndex", listConnexionsReseaux.ListIndex

    'désactive le bouton démarrer pour éviter les conneries
390 Me.cmdDemarrer.Enabled = False

    'efface les variables
400 Call InitialisationVar


    'Tentative de connexion au VPN à l'aide des paramètres sauvegardés dans windows
410 Call ConnexionVPN


420 If pbDemandeArretEnCours = True Then
        ' l'utilisateur à cliqué sur arreter
        ' donc on stop les tentatives de reconnexion
        ' sans lancer le timerping
440     Call ArreterVPN
470     Exit Sub
480 End If

    'si c'est le premier démarrage et qu'il n'y a pas de VPN alors on le dit
490 If pbInitialisation = True And pbConnexionVPN = False Then
500     Me.lblStatut.Caption = MLSGetString("0023")    ' MLS-> "Service arrêté"
510     Me.lblOperationEnCours.Caption = vbNullString
520     If pbAutoDemarrer = True Then
            'démarrage automatique, donc on réessaye de connecter le vpn
530         Me.lblOperationEnCours = MLSGetString("0024")    ' MLS-> "Connexion en cours..."

            Me.timerInitialisation.Interval = 5000
            Me.timerInitialisation.Enabled = True

560     Else

            ' connexion impossible
            ' faut quand meme faire déconnecter
570         Call ArreterVPN

            ' on fait 3 tentatives puis message box
            Static iTentativeCo As Byte
580         iTentativeCo = iTentativeCo + 1
590         If iTentativeCo >= 3 Then
600             iTentativeCo = 0
610             MsgBox MLSGetString("0025"), vbInformation    ' MLS-> "Connexion impossible, veuillez réessayer plus tard."
620         Else
630             Call cmdDemarrer_Click
640         End If
650     End If
660 Else

        'ce n'est pas le premier lancement,
        'donc on est dans le cycle déconnexion/reconnexion VPN

        'actualise le statut
670     Me.lblStatut.Caption = MLSGetString("0026")    ' MLS-> "Actif"

        ' attend plusieurs secondes avant de refaire une tentative de connexion VPN
680     Me.timerPing.Interval = 10000
690     Me.timerPing.Enabled = True

700     Debug.Print Timer, "timerPing activé avec temporisation : Actualisation..."

        'MyLog "in démarrer, juste après timerping ON"

710     Me.lblOperationEnCours.Caption = MLSGetString("0027")    ' MLS-> "Actualisation..."

720 End If


730 Exit Sub

err:
740 Call GestionErreur(Me, "cmdDemarrer_Click", pcErreurMessageSupplementaire)
750 Me.Show
760 Me.SetFocus
770 If pbEnvoyerFormulaire = True Then
780     Me.WindowState = vbMinimized
790     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
800 End If
End Sub

Private Sub ArreterVPN()

    On Error GoTo err:


'si on tente de connecter le vpn, alors on signale la demande d'arrêt
10  If pbThreadEnCours = True Then
20      pbDemandeArretEnCours = True

30      Me.cmdArreter.Enabled = False
40      Me.lblOperationEnCours.Caption = MLSGetString("0028")    ' MLS-> "Arrêt en cours..."

50      Me.Refresh
60      Exit Sub
70  Else
80      pbDemandeArretEnCours = False
90  End If

    DoEvents

    ' stop les timers en cours
    Dim item
100 For Each item In frmMain.Controls
110     If TypeOf item Is Timer Then
120         item.Enabled = False
130     End If
140 Next

    'sauvegarde l'IP pingé dans le fichier ini
150 EcrireINI "Paramètres du VPN", "Adresse IP", psAdressIPvpn


160 Call InitialisationVar
170 Me.cmdDemarrer.Enabled = True
180 pbCycleDeconnexionReconnexionEnCours = False


    'actualise le statut
190 Me.lblOperationEnCours.Caption = MLSGetString("0029")    ' MLS-> "Déconnexion en cours..."
200 Me.Refresh

    'déconnexion du VPN
210 Call DeconnexionDuVpn

    'fin de la sécurisation du tunnel du vpn
    ' add route passerelle
220 If pbConfigSecurisationDuTunnel = True And psAdresseIPpasserelle <> vbLocalHost Then
230     Call modGlobal.SecurisationDuTunnelVPN(desactiver)
240 End If

250 On Error Resume Next

    'efface le statut
260 Me.lblOperationEnCours.Caption = MLSGetString("0030")    ' MLS-> "Déconnecté"

    'change l'icone du systray
270 Me.Icon = Me.ImageIconDeconnected

    ' notifie le widget
280 If pbWidgetDisplayed = True Then
290     frmWidget.ChangeState False
300 End If

    'actualise le statut
310 Call ChangeSystrayToolTip(Me, App.Title & vbCrLf & MLSGetString("0031"))    ' MLS-> "Service arrêté"

    'efface l'infobulle, s'il y en avait une
320 Call RemoveBalloon(Me)

330 Me.cmdArreter.Enabled = True


340 Exit Sub

err:
350 Call GestionErreur(Me, "ArreterVPN", pcErreurMessageSupplementaire)
360 Me.Show
370 Me.SetFocus
380 If pbEnvoyerFormulaire = True Then
390     Me.WindowState = vbMinimized
400     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
410 End If
End Sub

Private Sub cmdQuitter_Click()    'Bouton Quitter

'aucune action si une thread est en cours
'c'est à dire, si l'opération en cours est "connexion en cours"
    If pbThreadEnCours = True Then Exit Sub

    'MyLog "in quitter avant timer OFF"

    Dim item
    For Each item In frmMain.Controls
        If TypeOf item Is Timer Then
            item.Enabled = False
        End If
    Next item

    'MyLog "in quitter après timer OFF"

    Unload Me

    'MyLog "in quitter après unload me"

End Sub

Private Sub Form_Load()

2   MLSFillMenuLanguages    '<- Add by Multi-Languages Support Add-in
3   MLSLoadLanguage Me   '<- Add by Multi-Languages Support Add-in
    Dim i As Long
    Dim clsA As CIsAlreadyRunning

10  Debug.Print Timer, "--- Begin ----------------------------------------------------" & vbNewLine

    'MyLog "in form load au début"

    On Error GoTo err:

30  If pbConfigurationEnCours = False Then
        'n'autorise qu'une seule instance du programme
        'si une autre instance est ouverte alors on là met au premier plan
        'et on ferme celle là
40      Set clsA = New CIsAlreadyRunning
50      If clsA.IsAlreadyRunning = True Then
60          Set clsA = Nothing
70          End
80      End If
90  End If

    ' check qu'on est dans l'IDE
    Dim clsIDE As CIsInIde
100 Set clsIDE = New CIsInIde
110 If clsIDE.IsInIde = True Then
        ' on est dans l'IDE
        'Debug.Print timer, "on est dans l'IDE"
120 Else
        ' on n'est pas dans l'IDE
        'Debug.Print timer, "cette ligne ne sera jamais écrite"
130 End If
140 Set clsIDE = Nothing

    ' cache le menu popup
    Me.mnuPopup.Visible = False

    'vide le combo de l'ip du vpn
150 Me.comboListingIP.Locked = False
160 Me.comboListingIP.Text = vbNullString
170 Me.comboListingIP.Locked = True

    Me.timerReDelRoute.Enabled = False
    Me.timerReDelRoute.Interval = 1000

    'liste les processus pour mettre en mémoire la procedure
    'histoire d'aller plus vite au moment de fermer les applications
210 Call ListerProcess

    'masquage des objets suivants sinon crash
220 Me.txtPremierLancement.Visible = False
230 For i = 0 To Me.chkApplicationSupporte.Count - 1
240     Me.chkApplicationSupporte(i).Visible = False
250 Next i

    'donne le nom du programme à sa fenêtre en fonction qu'il soit portable ou pas
    If Dir$(App.Path & "\msvbvm60.dll") = vbNullString Then
330     Me.Caption = App.Title & " - " & App.Major & "." & App.Minor & "." & App.Revision
    Else
335     Me.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & " portable"
    End If

    'efface et initialise les variables
340 Call InitialisationVar

    'Lit à partir du fichier INI les paramètres
350 If FichierINI(lire) = False Then
        'fichier INI absent alors on le dit
360     Me.txtPremierLancement.Text = MLSGetString("0033") & vbCrLf & MLSGetString("0034")    ' MLS-> "Puisque c'est votre 1er lancement alors" ' MLS-> "cliquez sur cette zone pour configurer"
370     Me.txtPremierLancement.Top = 360
380     Me.txtPremierLancement.Width = 4935
390     Me.txtPremierLancement.Height = 555
400     Me.txtPremierLancement.Left = 240
410     Me.txtPremierLancement.Visible = True
420 Else
        'fichier ini présent donc on peut cacher la zone de texte
430     Me.txtPremierLancement.Visible = False
440 End If

450 pbJustLoaded = True

    'redimensionne la fenetre
460 Me.Width = 5760

    'MyLog "in form load avant timerStart ON"

480 Exit Sub

err:
490 Call GestionErreur(Me, "Form_Load", pcErreurMessageSupplementaire)
500 Me.Show
510 Me.SetFocus
520 If pbEnvoyerFormulaire = True Then
530     Me.WindowState = vbMinimized
540     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
550 End If
End Sub

Private Sub InitialisationVar()
'initialise les variables

    On Error GoTo err:
10
20  If pbRelancerApplications = False Then
30      piNombreDeconnexion = 0
40      pbInitialisation = True
50      Me.lblStatut.Caption = MLSGetString("0035")    ' MLS-> "Service arrêté"
60      Me.lblOperationEnCours.Caption = vbNullString
70      Me.timerPing.Enabled = False
80      Me.lblDateDeconnexion = vbNullString
90      Me.comboListingIP.ToolTipText = MLSGetString("0036")    ' MLS-> "Cliquez sur 'Configurer' pour lister les IP et sélectionnez celle du VPN"
100     pbIPselectionne = False
110 End If
120 Exit Sub

err:
130 Call GestionErreur(Me, "InitialisationVar", pcErreurMessageSupplementaire)
140 Me.Show
150 Me.SetFocus
160 If pbEnvoyerFormulaire = True Then
170     Me.WindowState = vbMinimized
180     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
190 End If
End Sub

Private Sub FermerApplications()
    Dim i As Long, k As Long
    Dim sTEMP As String
    Dim ret As Boolean
    'Dim rfsh As CRefreshDesktop
    Dim clsT As CDelIconSystray

    On Error Resume Next
    'on désactive les erreurs car étape délicate

    'liste tout les processus en cours
    'Debug.Print timer, "Listing des processus en cours..."
    Call ListerProcess
    'Debug.Print timer, "Processus listés"

    Set clsT = New CDelIconSystray

    For i = 0 To Me.chkApplicationSupporte.Count - 1
        If Me.chkApplicationSupporte(i).Tag <> vbNullString Then
            If Me.chkApplicationSupporte(i).value = 1 Then
                For k = 1 To pcoListeProcessString.Count
                    'Debug.Print timer, "boucle dans la liste des processus"
                    'Debug.Print timer, "pcoListeProcessString(k) = " & pcoListeProcessString(k)
                    sTEMP = pcoListeProcessString(k)

                    'lit le nom du fichier avec son extension dans le Tag du check
                    'Debug.Print timer, "UCase(Me.chkApplicationSupporte(i).Tag) = " & UCase(Me.chkApplicationSupporte(i).Tag)
                    If UCase(Me.chkApplicationSupporte(i).Tag) = UCase(sTEMP) Then
                        Me.lblOperationEnCours.Caption = MLSGetString("0037") & Me.chkApplicationSupporte(i).Tag    ' MLS-> "Fermeture de "
                        'Debug.Print timer, "fermeture de " & sTEMP
                        'Debug.Print timer, "process PID de " & sTEMP & " : " & pcoListeProcessPID(k)

                        'supprime l'icone du systray
                        ret = clsT.DelIconByPID(pcoListeProcessPID(k))
                        'Debug.Print timer, "ret del icon ="; Ret

                        'fermeture du processus
                        ProcessTerminate (pcoListeProcessPID(k))

                        DoEvents
                    End If
                Next k
            End If
        End If
    Next i
    Set clsT = Nothing

    pbRelancerApplications = True

    On Error GoTo 0
End Sub

Private Function FichierINI(operation As OPERATIONINI) As Boolean
    Dim sFichierINI As String
    Dim i As Integer

    On Error GoTo err:

    Debug.Print Timer, "Lecture du fichier INI"

    'test de la présence du fichier ini
10  sFichierINI = App.Path & "\" & App.Title & ".ini"
20  If Dir(sFichierINI) <> "" Then
        'fichier INI présent
30      pbFichierINIpresent = True
40  Else
50      pbFichierINIpresent = False
60  End If

70  Select Case operation
    Case OPERATIONINI.lire

        'Test du fichier INI
80      If pbFichierINIpresent = True Then

            'efface le contenu précédent des checkbox
90          For i = 0 To frmMain.chkApplicationSupporte.Count - 1
100             frmMain.chkApplicationSupporte(i).Caption = "Libre" & i
110             frmMain.chkApplicationSupporte(i).Tag = vbNullString
120             frmMain.chkApplicationSupporte(i).ToolTipText = vbNullString
130             frmMain.chkApplicationSupporte(i).Visible = False
140             frmMain.chkApplicationSupporte(i).value = 1
150         Next i

            'lecture du fichier ini
160         psNombreApplicationsGerees = LireINI("Paramètres des applications", "Nombre d'applications")
170         If Val(psNombreApplicationsGerees) > 0 Then

                'remplis le contenu des checkbox en lisant le INI
180             For i = 1 To Val(psNombreApplicationsGerees)
                    'récupère les infos sur les applications
190                 frmMain.chkApplicationSupporte(i - 1).Caption = GetFileName(LireINI("Paramètres des applications", "Application" & i))
200                 frmMain.chkApplicationSupporte(i - 1).Tag = GetFileNameExt(LireINI("Paramètres des applications", "Application" & i))
210                 frmMain.chkApplicationSupporte(i - 1).ToolTipText = LireINI("Paramètres des applications", "Application" & i)
220                 frmMain.chkApplicationSupporte(i - 1).Visible = True
230                 frmMain.chkApplicationSupporte(i - 1).value = Val(LireINI("Paramètres des applications", "Checkbox" & i - 1))
240             Next i
250         End If

260         On Error Resume Next
            'récupère l'adresse IP du VPN
270         psAdressIPvpn = LireINI("Paramètres du VPN", "Adresse IP")

            'vpn à IP dynamique ???
280         pbIpVpnDynamique = True
            '-----------------------------------------------------
            '         pbIpVpnDynamique = CBool(LireINI("Paramètres du VPN", "IP dynamique"))
            '-----------------------------------------------------

            'IP de la passerelle
290         psAdresseIPpasserelle = LireINI("Sécurisation du tunnel VPN", "Adresse IP de la passerelle")

            'IP du pc par défaut
300         psAdresseIPparDefaut = LireINI("Sécurisation du tunnel VPN", "Adresse IP par défaut")

            'lancement avec windows
310         pbLancerAvecWindows = CBool(LireINI("Autodémarrage", "Windows"))

            ' pour récupèrer l'ip du pc en plus
320         pbAdressIPproblem = CBool(LireINI("Paramètres du VPN", "En cas de problème"))

            ' option de réduction dans le systray du bouton croix
330         pbOptionReduireSystrayQuit = CBool(LireINI("Divers", "Bouton croix pour réduire"))

            ' option pour fermer les applications gérées en quittant
340         pbCloseAppliOnQuit = CBool(LireINI("Divers", "Fermer applications gérées en quittant"))

            'démarrer les applications au lancement
350         pbAppliAutoDemarrer = CBool(LireINI("Autodémarrage", "Applications gérées"))

            'Réduire la fenetre au démarrage
360         pbAutoReduire = CBool(LireINI("Autodémarrage", "Réduire"))

            'démarrer automatiquement
370         pbAutoDemarrer = CBool(LireINI("Autodémarrage", "Démarrer"))

            ' afficher le widget au lancement
375         pbWidgetAutoDisplayed = CBool(LireINI("Divers", "Lancer le widget au démarrage"))

            ' afficher le widget au lancement
376         pbLogDeconnexion = CBool(LireINI("Divers", "Log de déconnexions"))

            'récupère la préférence pour la sécurisation du tunnel VPN
380         pbConfigSecurisationDuTunnel = CBool(LireINI("Sécurisation du tunnel VPN", "Activer"))
            On Error GoTo err:

490         FichierINI = True
500     Else
510         FichierINI = False
520     End If

530 Case OPERATIONINI.Ecrire
        'inutile
540 End Select
550 Exit Function

err:
560 Call GestionErreur(Me, "FichierINI", pcErreurMessageSupplementaire)
570 frmMain.Show
580 frmMain.SetFocus
590 If pbEnvoyerFormulaire = True Then
600     frmMain.WindowState = vbMinimized
610     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
620 End If
End Function

Private Sub TimerReDelRoute_Timer()
    Static iCompteur As Long
    Static Flag As Boolean

    ' refait del route toutes les heures car
    ' sur certains system, la route de passerelle est rétablie toutes seule [?]
    ' anomalie que je tente de résoudre

    If Flag = True Then Exit Sub
    Flag = True

    ' chaque seconde
    iCompteur = 1 + iCompteur

    ' refait del route toutes les heures = 3600s
    If iCompteur >= 3600 Then

        iCompteur = 0

        If pbConnexionVPN = True Then
            '--------------------------------------------
            'sécurisation du tunnel du vpn
            ' del route passerelle
            If pbConfigSecurisationDuTunnel = True And psAdresseIPpasserelle <> vbLocalHost Then
                Call modGlobal.SecurisationDuTunnelVPN(activer)
            End If
        End If

    End If

    Flag = False
End Sub

Private Sub Initialization()
    Dim sTEMP As String

    On Error GoTo err:

    'MyLog "in timerstart"

    ' affiche le widget
20  If pbWidgetAutoDisplayed = True Then frmWidget.Show

    ' permet de dessiner correctement le widget
    DoEvents

60  If pbConnexionReseauPresente = True Then

70      On Error Resume Next
80      sTEMP = LireINI("Paramètres du VPN", "N° connexion choisis ; listConnexionsReseaux.ListIndex")
90      If sTEMP <> vbNullString Then
100         If Me.listConnexionsReseaux.ListCount >= CInt(sTEMP) + 1 Then
110             Me.listConnexionsReseaux.ListIndex = CInt(sTEMP)
120         Else
130             Me.listConnexionsReseaux.ListIndex = 0
140         End If
150     End If
        On Error GoTo err:

        'affecte l'ip du vpn, lu dans le fichier ini, au combo
160     If psAdressIPvpn <> vbNullString Then
170         Me.comboListingIP.Locked = False
180         Me.comboListingIP.Clear
190         Me.comboListingIP.AddItem psAdressIPvpn
200         Me.lblIpVpn.Caption = MLSGetString("0038") & " " & psAdressIPvpn      ' MLS-> "IP locale du VPN : "
210         Me.comboListingIP.ListIndex = 0
220         Me.comboListingIP.Locked = True
230     End If

        ' désactive le redémarrage si on demande l'arrêt
240     If pbDemandeArretEnCours = True Then
250         Call ArreterVPN
260         Exit Sub
270     End If


280     If pbAutoDemarrer = True Then
            'démarrer au lancement
290         Call cmdDemarrer_Click

            'démarrer les applications gérées au lancement
300         If pbAppliAutoDemarrer Then
310             pbFlagAppliAutoDemarrer = True
320         Else
330             pbFlagAppliAutoDemarrer = False
340         End If

            'réduit la fenetre si l'option a été cochée
350         If pbAutoReduire = True Then
                'et si internet n'est pas inaccessible
360             If pbInternetInaccessible = False Then
370                 pbBoutonCroix = True
380                 Me.WindowState = vbMinimized
390             End If
400         End If

410     Else
420         pbFlagAppliAutoDemarrer = False
430     End If

440 End If

450 Exit Sub

err:
460 Call GestionErreur(Me, "timerStart_Timer", pcErreurMessageSupplementaire)
470 Me.Show
480 Me.SetFocus
490 If pbEnvoyerFormulaire = True Then
500     Me.WindowState = vbMinimized
510     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
520 End If
End Sub

Private Sub timerTest_Timer()
    Me.timerTest.Enabled = False




End Sub

Private Sub txtPremierLancement_Click()
    Call ConfigurerAppli
End Sub

Private Sub ConfigurerAppli()
    Dim i As Integer
    Dim bService As Boolean
    Dim sTEMP As String

    On Error GoTo err:

10  bService = Not CBool(Me.cmdDemarrer.Enabled)

20  Load frmConfig
30  frmConfig.Show vbModal

    'fichier ini présent donc on peut cacher la zone de texte
40  Me.txtPremierLancement.Visible = False

    'efface la liste des connexions réseaux précédente
50  Call RemplirListeReseaux

    'sélectionne la 1ère connexion réseau
60  If Me.listConnexionsReseaux.ListCount > 0 Then
70      Me.listConnexionsReseaux.ListIndex = 0
80  End If


    'efface et initialise les variables
90  Call InitialisationVar

100 Call FichierINI(lire)

110 On Error Resume Next
120 sTEMP = LireINI("Paramètres du VPN", "N° connexion choisis ; listConnexionsReseaux.ListIndex")
130 If sTEMP <> vbNullString Then
140     If Me.listConnexionsReseaux.ListCount >= CInt(sTEMP) + 1 Then
150         Me.listConnexionsReseaux.ListIndex = CInt(sTEMP)
160     Else
170         Me.listConnexionsReseaux.ListIndex = 0
180     End If
190 End If
    On Error GoTo err:

    'élément lu dans le fichier ini
200 For i = 0 To Val(psNombreApplicationsGerees) - 1
210     Me.chkApplicationSupporte(i).value = 1
220 Next i

    'initialise les objets
230 Me.timerPing.Enabled = False
240 Me.cmdDemarrer.Enabled = True

    'efface le statut
250 Me.lblOperationEnCours.Caption = vbNullString

260 Me.comboListingIP.Locked = False
270 Me.comboListingIP.Clear
280 Me.comboListingIP.AddItem psAdressIPvpn
290 Me.lblIpVpn.Caption = MLSGetString("0039") & " " & psAdressIPvpn   ' MLS-> "IP locale du VPN : "
300 Me.comboListingIP.ListIndex = 0
310 Me.comboListingIP.Locked = True

    'actualise le statut
320 Call ChangeSystrayToolTip(Me, App.Title & vbCrLf & MLSGetString("0040"))    ' MLS-> "Service arrêté"

330 If bService = True Then
340     Call cmdDemarrer_Click
350 End If

360 Exit Sub

err:
370 Call GestionErreur(Me, "ConfigurerAppli", pcErreurMessageSupplementaire)
380 Me.Show
390 Me.SetFocus
400 If pbEnvoyerFormulaire = True Then
410     Me.WindowState = vbMinimized
420     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
430 End If
End Sub

Private Sub RemplirListeReseaux()
    Dim sTEMP As String
    Dim lTmp As Long

    On Error GoTo err:

10  Me.listConnexionsReseaux.Clear
    'Liste les connexions réseaux pour récupérer la connexion VPN
20  sTEMP = GetConnectionList
30  If Len(sTEMP) > 1 Then
40      Do While Not Len(sTEMP) = 0
50          lTmp = InStr(1, sTEMP, ";")
60          If lTmp = 0 Then lTmp = Len(sTEMP) + 1
70          listConnexionsReseaux.AddItem Left(sTEMP, lTmp - 1)
80          sTEMP = Mid(sTEMP, lTmp + 1)
90      Loop
100 End If
110 Exit Sub

err:
120 Call GestionErreur(Me, "RemplirListeReseaux", pcErreurMessageSupplementaire)
130 Me.Show
140 Me.SetFocus
150 If pbEnvoyerFormulaire = True Then
160     Me.WindowState = vbMinimized
170     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
180 End If
End Sub

Private Sub DeconnexionDuVpn()
    Dim iBoucleDeco As Byte
    Dim bDeconnexionActive As Boolean
    Dim k As Byte

    On Error GoTo err:

    ' compteur pour limiter les déconnexions infinies
    k = 1

    'déconnecte jusqu'à qu'il n'est y a plus d'handle de rasdial
10  Do
20      iBoucleDeco = iBoucleDeco + 1
30      bDeconnexionActive = HangUp(GetCurrentHandleRasConn)
        'VPN déconnecté OK
40      Debug.Print Timer, "déconnexion #" & iBoucleDeco & " : ok"
        DoEvents
        k = k + 1
50  Loop Until bDeconnexionActive = False Or k >= 10

    'flag son statut
60  pbConnexionVPN = False

    ' just pause one second
    Dim iTimer As Single
    iTimer = Timer
    Do
        DoEvents
    Loop Until Timer - iTimer >= 1

70  Exit Sub

err:
80  Call GestionErreur(Me, "DeconnexionDuVpn", pcErreurMessageSupplementaire)
90  Me.Show
100 Me.SetFocus
110 If pbEnvoyerFormulaire = True Then
120     Me.WindowState = vbMinimized
130     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
140 End If
End Sub

Private Sub RelancerAppli()
    Dim i As Integer

10  For i = 0 To Me.chkApplicationSupporte.Count - 1
        'Debug.Print timer, "boucle me.check.value pour relancer les applications"
20      If Me.chkApplicationSupporte(i).value = 1 Then

            'rechargement des appli
30          psFichierPath = Me.chkApplicationSupporte(i).ToolTipText

            'executer le fichier que si le tag de la case à cocher est remplie
40          If Me.chkApplicationSupporte(i).Tag <> vbNullString Then
50              ShellExecute HWnd, "Open", psFichierPath, "", App.Path, 1
60          End If

70          DoEvents
80      End If
90  Next i

100 Exit Sub

err:
110 Call GestionErreur(Me, "RelancerAppli", pcErreurMessageSupplementaire)
120 Me.Show
130 Me.SetFocus
140 If pbEnvoyerFormulaire = True Then
150     Me.WindowState = vbMinimized
160     Call PopupBalloon(Me, App.Title, pcCollerErreurFormulaire)
170 End If
End Sub

Public Sub ResetVPN()

' Déconnecte le vpn puis le reconnecte

    Call StopVPN

    cmdDemarrer_Click

End Sub

Public Sub StopVPN()

' Déconnecte le vpn

    FermerApplications

    ArreterVPN

End Sub

Private Sub StopByUser()

    DoEvents
    ' termine la procédure en cours d'exécution

    pbPopUpArretEnCours = True

    Call FermerApplications

    Call ArreterVPN

    piNombreDeconnexion = 0

    pbPopUpArretEnCours = False

End Sub

