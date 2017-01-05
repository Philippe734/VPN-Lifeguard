VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11550
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboLang 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   5445
      Width           =   1695
   End
   Begin VB.CheckBox CheckLog 
      Caption         =   "Log de déconnexion"
      Height          =   255
      Left            =   5040
      TabIndex        =   44
      ToolTipText     =   "Sauvegarde l'historique des déconnexions et connexions"
      Top             =   5860
      Width           =   2775
   End
   Begin VB.CheckBox CheckWidget 
      Caption         =   "Afficher le widget au lancement"
      Height          =   255
      Left            =   5040
      TabIndex        =   43
      ToolTipText     =   "Objet qui change de couleur en fonction de la connexion"
      Top             =   6100
      Width           =   2775
   End
   Begin VB.CheckBox CheckCloseAppliOnQuit 
      Caption         =   "Fermer les applications gérées en quittant"
      Height          =   255
      Left            =   960
      TabIndex        =   41
      Top             =   6100
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.CommandButton cmdTest1 
      Caption         =   "Test 1"
      Height          =   375
      Left            =   9480
      TabIndex        =   40
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CheckBox CheckIconQuit 
      Caption         =   "Réduire au lieu de quitter par la croix"
      Height          =   255
      Left            =   960
      TabIndex        =   38
      ToolTipText     =   "En cliquant sur le bouton croix en haut à droite, permet de réduire le programme en icône dans le systray plutôt que de le quitter"
      Top             =   5860
      Width           =   3735
   End
   Begin VB.CheckBox CheckAdressIPproblem 
      Caption         =   "En cas de problème"
      Height          =   375
      Left            =   1800
      TabIndex        =   36
      ToolTipText     =   "Cochez cette case si votre VPN connecté n'est pas listé"
      Top             =   3400
      Width           =   1095
   End
   Begin VB.ComboBox ComboNombreApplications 
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   650
      Width           =   735
   End
   Begin VB.CommandButton cmdAjouterAppli 
      Caption         =   "Parcourir..."
      Height          =   375
      Left            =   5880
      TabIndex        =   31
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame Frame 
      ForeColor       =   &H00C00000&
      Height          =   1935
      Index           =   8
      Left            =   240
      TabIndex        =   26
      Top             =   6600
      Width           =   7815
      Begin VB.CheckBox CheckSecuTunnel 
         Caption         =   "VPN exclusif"
         Height          =   255
         Left            =   3240
         TabIndex        =   27
         ToolTipText     =   "Être sûr de passer par le VPN"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Connexion VPN exclusive "
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   0
         Width           =   1875
      End
      Begin VB.Label labelInfoSecuTunnel 
         Caption         =   $"frmConfig.frx":0000
         Height          =   1095
         Left            =   360
         TabIndex        =   28
         Top             =   480
         Width           =   6975
      End
   End
   Begin VB.CheckBox CheckAutoReduire 
      Caption         =   "Réduire au lancement"
      Height          =   255
      Left            =   4680
      TabIndex        =   21
      ToolTipText     =   "Se réduit en icône au lancement du programme"
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Vérifer la dernière version"
      Height          =   375
      Left            =   600
      TabIndex        =   20
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   240
      TabIndex        =   19
      Top             =   4080
      Width           =   2775
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Mise à jour "
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   810
      End
   End
   Begin VB.CheckBox checkDynamiqueIPvpn 
      Caption         =   "VPN à IP dynamique"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9240
      TabIndex        =   18
      ToolTipText     =   "cf form load"
      Top             =   1200
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox ComboDefautIP 
      Height          =   315
      Left            =   9360
      Style           =   2  'Dropdown List
      TabIndex        =   15
      ToolTipText     =   "Correspond à l'adresse IP de votre PC"
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Frame Frame 
      ForeColor       =   &H00C00000&
      Height          =   1575
      Index           =   7
      Left            =   3240
      TabIndex        =   11
      Top             =   3480
      Width           =   4815
      Begin VB.CheckBox CheckAutoAppli 
         Caption         =   "Démarrer les applications"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         ToolTipText     =   "Démarre les applications cochées au démarrage du programme"
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CheckBox CheckAutoWin 
         Caption         =   "Démarrer au lancement"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   13
         ToolTipText     =   "Démarre la connexion du VPN à l'ouverture du programme"
         Top             =   720
         Width           =   3375
      End
      Begin VB.CheckBox CheckAutoWin 
         Caption         =   "Lancer avec windows"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   12
         ToolTipText     =   "Lance le programme au démarrage de windows sans pour autant démarrer la connexion au VPN"
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Autodémarrage "
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   35
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Frame Frame 
      ForeColor       =   &H00C00000&
      Height          =   3255
      Index           =   4
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   4815
      Begin VB.PictureBox Picture 
         BorderStyle     =   0  'None
         Height          =   2895
         Index           =   1
         Left            =   240
         ScaleHeight     =   2895
         ScaleWidth      =   4455
         TabIndex        =   8
         Top             =   240
         Width           =   4455
         Begin VB.Frame Frame 
            ForeColor       =   &H00C00000&
            Height          =   1935
            Index           =   6
            Left            =   0
            TabIndex        =   9
            Top             =   960
            Width           =   4335
            Begin VB.ListBox ListAppli 
               Height          =   1230
               Left            =   240
               TabIndex        =   10
               Top             =   360
               Width           =   3855
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Liste "
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   8
               Left            =   120
               TabIndex        =   34
               Top             =   0
               Width           =   375
            End
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Nombre : "
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   7
            Left            =   480
            TabIndex        =   33
            Top             =   315
            Width           =   690
         End
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Applications à gérer "
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   0
         Width           =   1440
      End
   End
   Begin VB.Frame Frame 
      ForeColor       =   &H00C00000&
      Height          =   3855
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.PictureBox Picture 
         BorderStyle     =   0  'None
         Height          =   3375
         Index           =   0
         Left            =   120
         ScaleHeight     =   3375
         ScaleWidth      =   2535
         TabIndex        =   1
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton cmdIpConfig 
            Caption         =   "?"
            Height          =   375
            Left            =   960
            TabIndex        =   46
            ToolTipText     =   "Ouvre IPCONFIG"
            Top             =   2930
            Width           =   375
         End
         Begin VB.CommandButton cmdListerIP 
            Caption         =   "Lister IP"
            Height          =   375
            Left            =   120
            TabIndex        =   45
            ToolTipText     =   "Récupère les adresses IP de la passerelle et du VPN connecté."
            Top             =   2930
            Width           =   735
         End
         Begin VB.Frame Frame 
            ForeColor       =   &H00C00000&
            Height          =   1215
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   1560
            Width           =   2295
            Begin VB.ComboBox ComboPasserelleIP 
               Height          =   315
               ItemData        =   "frmConfig.frx":01E3
               Left            =   240
               List            =   "frmConfig.frx":01E5
               TabIndex        =   6
               Text            =   "ComboPasserelleIP"
               ToolTipText     =   "Correspond à l'adresse IP de votre box internet. Si le programme ne parvient pas à la trouver, alors sélectionnez 127.0.0.1"
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "IP de la box "
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   24
               Top             =   0
               Width           =   885
            End
            Begin VB.Label Label 
               Caption         =   "= de la passerelle = du routeur"
               Height          =   495
               Index           =   0
               Left            =   240
               TabIndex        =   5
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame 
            ForeColor       =   &H00C00000&
            Height          =   855
            Index           =   3
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   2295
            Begin VB.ComboBox ComboVpnIP 
               Height          =   315
               Left            =   240
               TabIndex        =   3
               Text            =   "ComboVpnIP"
               ToolTipText     =   "Adresse IP locale du VPN"
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "IP locale du VPN "
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   23
               Top             =   0
               Width           =   1260
            End
         End
         Begin VB.Label labelChoixIP 
            Alignment       =   2  'Center
            Caption         =   "Sélectionnez ou confirmez les adresses IP"
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   240
            TabIndex        =   16
            Top             =   0
            Width           =   2055
         End
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Adresses IP "
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   0
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   240
      TabIndex        =   37
      Top             =   5160
      Width           =   7815
      Begin VB.Label LabelTutoURL 
         AutoSize        =   -1  'True
         Caption         =   "Consulter l'aide du site web"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3120
         TabIndex        =   42
         ToolTipText     =   "http://vpnlifeguard.blogspot.com/p/tuto-rapide.html"
         Top             =   280
         Width           =   1920
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Divers"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   39
         Top             =   0
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmConfig"
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
' Module    : frmConfig
' Author    : philippe734
' Date      : 22/04/2010
' Purpose   : affiche la fenetre de configuration de VPN Lifeguard
'---------------------------------------------------------------------------------------


Option Explicit

Private Sub CheckAdressIPproblem_Click()

    pbAdressIPproblem = IIf(Me.CheckAdressIPproblem.value = 1, True, False)
    If pbAdressIPproblem = True Then
        Call cmdListerIp_click
    End If
End Sub

Private Sub CheckAutoWin_Click(Index As Integer)
    If Me.CheckAutoWin(1).value = 0 Then
        Me.CheckAutoAppli.value = 0
        Me.CheckAutoAppli.Enabled = False
        Me.CheckAutoReduire.value = 0
        Me.CheckAutoReduire.Enabled = False
    Else
        Me.CheckAutoAppli.Enabled = True
        Me.CheckAutoReduire.Enabled = True
    End If
End Sub

Private Sub cmdAjouterAppli_Click()

    Dim dlg As New CCommonDialog
    Dim i As Integer

    On Error GoTo err:

10  Me.ListAppli.Clear

    'initialisation des variables
20  psNombreApplicationsGerees = 0
    'ouvre la boite de dialogue

30  For i = 1 To Me.ComboNombreApplications.List(Me.ComboNombreApplications.ListIndex)

        'paramètre la boite de dialogue qui permettra de choisir l'application à gérer
40      dlg.Flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST
50      dlg.Filter = MLSGetString("0054") ' MLS-> "Fichiers exécutables (*.exe)|*.exe|"
60      dlg.Parent = Me.HWnd
70      dlg.FileName = vbNullString
80      dlg.DialogTitle = MLSGetString("0055") & (Val(psNombreApplicationsGerees) + 1) ' MLS-> "Ajouter l'application n°"
90      If dlg.ShowOpen Then
100         psResultatFichierChoisis = dlg.FileName
110         psNombreApplicationsGerees = Val(psNombreApplicationsGerees) + 1
120         EcrireINI "Paramètres des applications", "Nombre d'applications", Val(psNombreApplicationsGerees)
130         EcrireINI "Paramètres des applications", "Application" & Val(psNombreApplicationsGerees), psResultatFichierChoisis
140         Me.ListAppli.AddItem psResultatFichierChoisis
150     Else
            'sélection annulée par l'utilisateur
160         Debug.Print Timer, "annulé parcourir par le user"
170         Exit For
180     End If
190     Set dlg = Nothing
200 Next i


210 Exit Sub

err:
220 Set dlg = Nothing
230 Call GestionErreur(Me, "cmdAjouterAppli_Click", pcErreurMessageSupplementaire)
240 If pbEnvoyerFormulaire = True Then
250     Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
260 End If

End Sub

Private Sub cmdFermer_Click()

    If pbThreadEnCours = True Then Exit Sub

    Unload Me

End Sub

Private Sub CreationListeIP(Optional bSansMessages As Boolean)

10  On Error GoTo err

    'efface le contenu des combobox
20  Me.ComboDefautIP.Clear
30  Me.ComboVpnIP.Clear
40  Me.ComboPasserelleIP.Clear

    '-----------------------------------------------------
    'adresse IP de la passerelle
    Dim colAdresseIpGateway As New Collection
    Dim clsGateway As CGetIpgateway

    ' on désactive les erreurs le temps de générer la liste des gateway
50  On Error Resume Next
60  Set clsGateway = New CGetIpgateway
    ' récupère la liste des passerelles
70  Set colAdresseIpGateway = clsGateway.GetIpGateway(pbAdressIPproblem)
80  Set clsGateway = Nothing


90  On Error GoTo err

    Dim i As Long
    ' remplissage du combo avec la liste des passerelles
100 For i = 1 To colAdresseIpGateway.Count
110     Me.ComboPasserelleIP.AddItem colAdresseIpGateway(i)
120 Next i
130 Set colAdresseIpGateway = Nothing

    ' on ajoute Local Host dans le cas ou il n'y aurait pas de passerelles
    ' afin que l'utilisateur puisse utiliser le programme
140 Me.ComboPasserelleIP.AddItem vbLocalHost



150 If Me.ComboPasserelleIP.ListCount = 0 Then
        ' adresse IP de la passerelle non trouvé mais on ne le dit pas
        ' puisqu'elle ne sert qu'à l'option VPN exclusif
160 Else

170     Me.ComboPasserelleIP.ListIndex = 0
180 End If


    '------------------------------------------
    'adresses IP par défaut et IP du VPN
    'rempli le combo avec toutes les IP
    Dim ret As Boolean
190 ret = modGlobal.GetIpVpn(Me.ComboVpnIP)

200 If Me.ComboVpnIP.ListCount = 0 Then
210     If bSansMessages = False Then
220         MsgBox MLSGetString("0056"), vbExclamation ' MLS-> "Démarrer le VPN, puis cliquez sur 'Lister'."
230     End If
240 End If

250 Exit Sub

err:
260 Call GestionErreur(Me, "CreationListeIP", pcErreurMessageSupplementaire)
270 If pbEnvoyerFormulaire = True Then
280     Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
290 End If
End Sub

Private Sub cmdIpConfig_Click()
    Dim cls As New CRouteAddDell

    ' ouvre IPCONFIG pour montrer les adresses IP
    Set cls = New CRouteAddDell

    Call cls.ShowIPconfig

    Set cls = Nothing

End Sub

Private Sub cmdListerIp_click()

    Call CreationListeIP(False)
    'liste IP dans les combo

End Sub

Private Sub cmdTest1_Click()

10  On Error GoTo err

20  Exit Sub

err:
30  Call GestionErreur(Me, "CreationListeIP", pcErreurMessageSupplementaire)
40  If pbEnvoyerFormulaire = True Then
50      Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
60  End If

End Sub

Private Sub cmdUpdate_Click()



    On Error GoTo cmdUpdate_Click_Error

    frmCheckUpdate.Show vbModal

    On Error GoTo 0
    Exit Sub

cmdUpdate_Click_Error:

    Call GestionErreur(frmConfig, "cmdUpdate_Click", pcErreurMessageSupplementaire)
    If pbEnvoyerFormulaire = True Then
        Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
    End If

End Sub

Private Sub ComboLang_Click()
    Dim i As Long
    gsLanguageFile = Me.ComboLang
    '/ update all loaded forms for new language
    For i = 0 To Forms.Count - 1
        MLSLoadLanguage Forms(i)
    Next i
End Sub

Private Sub ComboPasserelleIP_KeyUp(KeyCode As Integer, Shift As Integer)

10  On Error GoTo ComboPasserelleIP_KeyUp_Error

    ' si la touche du clavier est enter
20  If KeyCode = vbKeyReturn Then
        ' alors on ajoute l'ip tapée dans le combo

        ' test si la valeur tapée est une adresse ip, test vite fait
30      If Len(Me.ComboPasserelleIP.Text) >= Len("0.0.0.0") And Len(Me.ComboPasserelleIP) <= Len("255.255.255.255") And InStr(1, Me.ComboPasserelleIP, ".", vbTextCompare) > 0 Then
40          Me.ComboPasserelleIP.AddItem Me.ComboPasserelleIP.Text
50          Me.ComboPasserelleIP.ListIndex = Me.ComboPasserelleIP.ListCount - 1
60          Me.cmdListerIP.SetFocus
70      Else
            'Debug.Print timer, "c'est pas une ip"
80      End If

90  End If

100 On Error GoTo 0
110 Exit Sub

ComboPasserelleIP_KeyUp_Error:

120 Call GestionErreur(frmConfig, "ComboPasserelleIP_KeyUp", pcErreurMessageSupplementaire)
130 If pbEnvoyerFormulaire = True Then
140     Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
150 End If

End Sub

Private Sub ComboNombreApplications_Click()

    On Error GoTo err:


10  If Me.ComboNombreApplications = 0 Then
20      Me.cmdAjouterAppli.Caption = MLSGetString("0057") ' MLS-> "Effacer liste"
30  Else
40      Me.cmdAjouterAppli.Caption = MLSGetString("0058") ' MLS-> "Parcourir..."
50  End If
60  Exit Sub

err:
70  Call GestionErreur(Me, "ComboNombreApplications_Click", pcErreurMessageSupplementaire)
80  If pbEnvoyerFormulaire = True Then
90      Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
100 End If
End Sub

Private Sub ComboVpnIP_KeyUp(KeyCode As Integer, Shift As Integer)
' si la touche du clavier est enter
    If KeyCode = vbKeyReturn Then
        ' alors on ajoute l'ip tapée dans le combo

        ' test si la valeur tapée est une adresse ip, test vite fait
        If Len(Me.ComboVpnIP.Text) >= Len("0.0.0.0") And Len(Me.ComboVpnIP) <= Len("255.255.255.255") And InStr(1, Me.ComboVpnIP, ".", vbTextCompare) > 0 Then
            Me.ComboVpnIP.AddItem Me.ComboVpnIP.Text
            Me.ComboVpnIP.ListIndex = Me.ComboVpnIP.ListCount - 1
            Me.cmdListerIP.SetFocus
        Else
            'Debug.Print timer, "c'est pas une ip"
        End If

    End If

End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()

    MLSLoadLanguage Me '<- Add by Multi-Languages Support Add-in
    Dim item As Object
    Dim sTEMP As String

    On Error GoTo err:


10  Me.Width = 8355

    'flag son ouverture
40  pbFormComplemantaireOuverte = True

50  Me.Caption = App.Title & " - " & App.Major & "." & App.Minor & "." & App.Revision & " - Configuration et options"
60  Me.Icon = frmMain.ImageIconNormal

70  For Each item In Me.Controls
80      If TypeOf item Is ComboBox Then
90          item.Clear
100     End If
110 Next item

120 Me.ListAppli.Clear

130 sTEMP = MLSGetString("0059") ' MLS-> "IP dynamique = Qui change à chaque connexion. Cochez cette case si le programme ne parvient pas à rétablir la connexion en cas de déconnexion."
140 Me.checkDynamiqueIPvpn.ToolTipText = sTEMP

150 pbConfigurationEnCours = True

    'lecture du fichier ini pour récupérer les données
160 On Error Resume Next
170 psAdresseIPparDefaut = LireINI("Sécurisation du tunnel VPN", "Adresse IP par défaut")
180 psAdresseIPpasserelle = LireINI("Sécurisation du tunnel VPN", "Adresse IP de la passerelle")
190 psAdressIPvpn = LireINI("Paramètres du VPN", "Adresse IP")

200 pbIpVpnDynamique = True
    '-----------------------------------------------------
    ' pbIpVpnDynamique = CBool(LireINI("Paramètres du VPN", "IP dynamique"))
    '-----------------------------------------------------

210 psNombreApplicationsGerees = LireINI("Paramètres des applications", "Nombre d'applications")
220 pbLancerAvecWindows = CBool(LireINI("Autodémarrage", "Windows"))
230 pbAutoDemarrer = CBool(LireINI("Autodémarrage", "Démarrer"))
240 pbAdressIPproblem = CBool(LireINI("Paramètres du VPN", "En cas de problème"))
250 pbConfigSecurisationDuTunnel = CBool(LireINI("Sécurisation du tunnel VPN", "Activer"))
260 pbOptionReduireSystrayQuit = CBool(LireINI("Divers", "Bouton croix pour réduire"))
270 pbCloseAppliOnQuit = CBool(LireINI("Divers", "Fermer applications gérées en quittant"))
275 pbWidgetAutoDisplayed = CBool(LireINI("Divers", "Lancer le widget au démarrage"))
276 pbLogDeconnexion = CBool(LireINI("Divers", "Log de déconnexions"))
    On Error GoTo err:

300 Exit Sub

err:
310 Call GestionErreur(Me, "Form_Load", pcErreurMessageSupplementaire)
320 If pbEnvoyerFormulaire = True Then
330     Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
340 End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'répertoire de destination pour le raccourci = démarrage
    Const CSIDL_STARTUP = &H7
    Dim ret As Boolean

    On Error Resume Next


    If UnloadMode = 0 Then
        If MsgBox(MLSGetString("0060"), vbOKCancel + vbQuestion) = vbCancel Then ' MLS-> "Si vous voulez configurer, alors cliquez sur annuler."
            Cancel = True
        Else
            Cancel = False
        End If
    Else
        If pbConnexionReseauPresente = True Then
            If Me.ComboVpnIP.ListCount = 0 Or Me.ComboVpnIP.ListIndex < 0 Then
                MsgBox MLSGetString("0061"), vbExclamation ' MLS-> "Veuillez choisir l'IP du VPN"
                Cancel = True
            End If

            If Me.ComboPasserelleIP.ListCount = 0 Or Me.ComboPasserelleIP.ListIndex < 0 Then
                MsgBox MLSGetString("0062"), vbExclamation ' MLS-> "Veuillez choisir l'IP de la passerelle ou la saisir"
                Cancel = True
            End If

            If Me.ComboNombreApplications.ListIndex > 0 And Me.ListAppli.ListCount = 0 Then
                MsgBox MLSGetString("0063"), vbExclamation ' MLS-> "Veuillez ajouter les applications à gérer"
                Cancel = True
            End If
        End If
    End If



    If Cancel = True Then
        Me.SetFocus
    Else
        'fermeture
    End If


    EcrireINI "Paramètres des applications", "Nombre d'applications", Me.ListAppli.ListCount

    EcrireINI "Paramètres du VPN", "Adresse IP", Me.ComboVpnIP

    EcrireINI "Paramètres du VPN", "En cas de problème", Me.CheckAdressIPproblem.value

    EcrireINI "Sécurisation du tunnel VPN", "Adresse IP de la passerelle", Me.ComboPasserelleIP

    EcrireINI "Sécurisation du tunnel VPN", "Activer", Me.CheckSecuTunnel.value

    EcrireINI "Autodémarrage", "Windows", Me.CheckAutoWin(0).value

    EcrireINI "Autodémarrage", "Démarrer", Me.CheckAutoWin(1).value

    EcrireINI "Autodémarrage", "Applications gérées", Me.CheckAutoAppli.value

    EcrireINI "Autodémarrage", "Réduire", Me.CheckAutoReduire.value

    EcrireINI "Divers", "Bouton croix pour réduire", Me.CheckIconQuit.value

    EcrireINI "Divers", "Fermer applications gérées en quittant", Me.CheckCloseAppliOnQuit.value

    EcrireINI "Divers", "Lancer le widget au démarrage", Me.CheckWidget.value

    EcrireINI "Divers", "Log de déconnexions", Me.CheckLog.value
    
    ' Save language to ini
    '/ update the CurrentLanguage entry to LangSetting.ini
    MLSWriteINI App.Path & "\LangSetting.ini", "Language", "CurrentLanguage", gsLanguageFile

    '-----------------------------------------------------
    ' On Error Resume Next
    'VPN à IP dynamique ???
    pbIpVpnDynamique = True
    ' pbIpVpnDynamique = CBool(LireINI("Paramètres du VPN", "IP dynamique"))
    '    On Error GoTo err:
    '-----------------------------------------------------

    If pbTunnelSecurised = True And Me.CheckSecuTunnel.value = 0 And psAdresseIPpasserelle <> vbLocalHost Then
        'fin de la sécurisation du tunnel du vpn
        ' add route passerelle
        Call modGlobal.SecurisationDuTunnelVPN(desactiver)
    End If


    'gestion du raccourci au démarrage de windows
    Dim clsC As CShortcut
    Set clsC = New CShortcut
    If Me.CheckAutoWin(0).value = 0 Then
        'efface le raccourci au démarrage
        ret = clsC.DeleteShortcut(App.Title, CSIDL_STARTUP)
    Else
        'ajoute le raccourci au démarrage de windows
        Dim FullPathExeFile As String
        FullPathExeFile = App.Path & "\" & App.EXEName & ".exe"
        ret = clsC.CreateShortcut(FullPathExeFile, App.Title, CSIDL_STARTUP)
    End If
    'libère la class
    Set clsC = Nothing

    'flag sa fermeture
    pbFormComplemantaireOuverte = False

    On Error GoTo 0
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' enlève le soulignement du lien internet de l'aide
    LabelTutoURL.FontUnderline = False
End Sub

Private Sub LabelTutoURL_Click()
    On Error Resume Next
    ShellExecute 0&, vbNullString, Me.LabelTutoURL.ToolTipText, vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub LabelTutoURL_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' souligne le lien internet de l'aide online
    LabelTutoURL.FontUnderline = True
    
    ' change la souris en main
    MousePointerHand
    
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    Dim bAdressesDansFichierIni As Boolean
    Dim bIPpresente As Boolean

    On Error GoTo err:

20  For i = 0 To 6
30      Me.ComboNombreApplications.AddItem i
40  Next i
50  Me.ComboNombreApplications.ListIndex = 0

60  If psAdresseIPparDefaut <> vbNullString Or psAdresseIPpasserelle <> vbNullString Or psAdressIPvpn <> vbNullString Then
70      bAdressesDansFichierIni = True
80  Else
90      bAdressesDansFichierIni = False
100 End If

110 Call CreationListeIP(bAdressesDansFichierIni)

120 If bAdressesDansFichierIni = True Then
130     If psAdresseIPparDefaut <> vbNullString Then
            'restaure les ip sauvegardés
140         bIPpresente = False
150         For i = 0 To Me.ComboDefautIP.ListCount - 1
160             If Me.ComboDefautIP.List(i) = psAdresseIPparDefaut Then
170                 Me.ComboDefautIP.ListIndex = i
180                 bIPpresente = True
190             End If
200         Next i
210         If bIPpresente = False Then
                'IP absente donc on l'ajoute
220             Me.ComboDefautIP.AddItem psAdresseIPparDefaut
                'puis on la sélectionne
230             Me.ComboDefautIP.ListIndex = Me.ComboDefautIP.ListCount - 1
240         End If
250     End If
260     If psAdresseIPpasserelle <> vbNullString Then
            'restaure les ip sauvegardés
270         bIPpresente = False
280         For i = 0 To Me.ComboPasserelleIP.ListCount - 1
290             If Me.ComboPasserelleIP.List(i) = psAdresseIPpasserelle Then
300                 Me.ComboPasserelleIP.ListIndex = i
310                 bIPpresente = True
320             End If
330         Next i
340         If bIPpresente = False Then
                'IP absente donc on l'ajoute
350             Me.ComboPasserelleIP.AddItem psAdresseIPpasserelle
                'puis on la sélectionne
360             Me.ComboPasserelleIP.ListIndex = Me.ComboPasserelleIP.ListCount - 1
370         End If
380     End If
390     If psAdressIPvpn <> vbNullString Then
            'restaure les ip sauvegardés
400         bIPpresente = False
410         For i = 0 To Me.ComboVpnIP.ListCount - 1
420             If Me.ComboVpnIP.List(i) = psAdressIPvpn Then
430                 Me.ComboVpnIP.ListIndex = i
440                 bIPpresente = True
450             End If
460         Next i
470         If bIPpresente = False Then
                'IP absente donc on l'ajoute
480             Me.ComboVpnIP.AddItem psAdressIPvpn
                'puis on la sélectionne
490             Me.ComboVpnIP.ListIndex = Me.ComboVpnIP.ListCount - 1
500         End If
510     End If
520 End If

    'restitue les données sauvegardées dans le fichier INI
530 If psNombreApplicationsGerees <> vbNullString Then
540     Me.ComboNombreApplications.ListIndex = Val(psNombreApplicationsGerees)
550     For i = 1 To Val(psNombreApplicationsGerees)
560         Me.ListAppli.AddItem LireINI("Paramètres des applications", "Application" & i)
570     Next i
580 End If

    ' restitue les valeurs des checkbox sauvegardées
590 Me.CheckAutoWin(0).value = CByte(CByte(pbLancerAvecWindows) / CByte(True))
600 Me.CheckAutoWin(1).value = CByte(CByte(pbAutoDemarrer) / CByte(True))
610 Me.CheckSecuTunnel.value = CByte(CByte(pbConfigSecurisationDuTunnel) / CByte(True))
620 Me.checkDynamiqueIPvpn.value = CByte(CByte(pbIpVpnDynamique) / CByte(True))
630 Me.CheckAdressIPproblem.value = CByte(CByte(pbAdressIPproblem) / CByte(True))
640 Me.CheckIconQuit.value = CByte(CByte(pbOptionReduireSystrayQuit) / CByte(True))
650 Me.CheckCloseAppliOnQuit.value = CByte(CByte(pbCloseAppliOnQuit) / CByte(True))
660 Me.CheckWidget.value = CByte(CByte(pbWidgetAutoDisplayed) / CByte(True))
670 Me.CheckLog.value = CByte(CByte(pbLogDeconnexion) / CByte(True))

680 If pbAutoDemarrer = True Then
690     Me.CheckAutoAppli.value = CByte(CByte(pbAppliAutoDemarrer) / CByte(True))
700     Me.CheckAutoReduire.value = CByte(CByte(pbAutoReduire) / CByte(True))
710 Else
720     Me.CheckAutoAppli.Enabled = False
730     Me.CheckAutoReduire.Enabled = False
740 End If

    ' Load multi language in combo
    Dim k As Long
750 For k = 1 To listLanguage.Count
760     Me.ComboLang.AddItem listLanguage(k)
770 Next k
    
    ' Disable error if user change manually langage < bug fixed : L'erreur 383 s'est produite dans la fenêtre frmConfig de la procédure TimerStart_Timer à la ligne 780 : Propriété 'Text' en lecture seule
    On Error Resume Next
780 Me.ComboLang = gsLanguageFile
    On Error GoTo err:




790 Exit Sub

err:
800 Call GestionErreur(Me, "TimerStart_Timer", pcErreurMessageSupplementaire)
810 If pbEnvoyerFormulaire = True Then
820     Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
830 End If
End Sub

