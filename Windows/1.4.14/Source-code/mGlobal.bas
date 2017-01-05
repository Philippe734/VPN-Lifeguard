Attribute VB_Name = "modGlobal"

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
' Module    : modGlobal
' Author    : philippe734
' Date      : 22/03/2010
' Purpose   : déclarations globales du projet VPN Lifeguard
'---------------------------------------------------------------------------------------

Option Explicit

' déclarations globales des variables du projet
Public pbConnexionVPN As Boolean
Public pbInitialisation As Boolean
Public pbRelancerApplications As Boolean
Public pbIPselectionne As Boolean
Public pbConnexionReseauPresente As Boolean
Public pbFichierINIpresent As Boolean
Public pbConfigurationEnCours As Boolean
Public pbTelechargementReussis As Boolean
Public pbCycleDeconnexionReconnexionEnCours As Boolean
Public pbFormComplemantaireOuverte As Boolean
Public pbAutoReduire As Boolean
Public pbAutoDemarrer As Boolean    'récupèré à partir du fichier ini
Public pbLancerAvecWindows As Boolean    'récupèré à partir du fichier ini
Public pbAppliAutoDemarrer As Boolean    'récupèré à partir du fichier ini
Public pbAdressIPproblem As Boolean    'récupéré à partir du fichier ini
Public pbOptionReduireSystrayQuit As Boolean    ' option qui permet la réduction dans le systray en quittant
Public pbFlagAppliAutoDemarrer As Boolean    'récupèré à partir du fichier ini
Public pbConfigSecurisationDuTunnel As Boolean    'récupèré à partir du fichier ini
Public pbTunnelSecurised As Boolean    'indicateur du tunnel sécurisé
Public pbCloseAppliOnQuit As Boolean
Public pbIpVpnDynamique As Boolean    'indicateur d'IP dynamique du vpn
Public pbJustLoaded As Boolean    'indicateur du premier lancement
Public pbInternetInaccessible As Boolean    'indicateur en cas d'internet inaccessible
Public pbLogActif As Boolean    ' active le log
Public pbDemandeArretEnCours As Boolean    ' flag pour arreter lors de connexion en cours
Public pbWidgetDisplayed As Boolean    ' flag l'affichage du widget
Public pbWidgetAutoDisplayed As Boolean    ' auto affichage du widget
Public pbLogDeconnexion As Boolean    ' écrit un log lors de déconnexions
Public pbPopUpArretEnCours As Boolean    ' Flag l'arrêt en cours du menu popup

Public piNombreDeconnexion As Long    'nombre de déconnexion du VPN
Public piIndexPasserelle As Long    'Index de la passerelle pour faire Add route

Public piChrono As Single    ' Pour chronométrer des trucs

Public psNombreApplicationsGerees As String    'Nombre d'application à gérer lue dans le INI
Public psDerniereVersionOnLine As String    'numéro de la dernière version via téléchargement
Public psDerniereDeconnexion As String    'Date et heure de la dernière déco
Public psFichierPath As String    'chemin de l'application à relancer
Public psResultatFichierChoisis As String    'fichier sélectionné lors de la configuration des applications à gérer
Public psAdressIPvpn As String    'adresse à pinger
Public psAdresseIPpasserelle As String    'IP de la passerelle (gateway)
Public psAdresseIPparDefaut As String    'IP du pc par défaut
Public psConnectionName As String    ' Rasdial connection name
Public psWidgetToolTip As String    ' infobulle du widget

'pour rasdial réalisé en multithreading
Public pbFinThread As Boolean
Public pbThreadEnCours As Boolean
Public pbThreadUsed As Boolean
Public phRasConnThreaded As Long
Public plngRetCodeThreaded As Long
Public ptRasdialParamsThreaded As RASDIALPARAMS


'message d'intro en cas d'erreur
Public Const pcErreurMessageSupplementaire As String = "Ah non pas cool, une erreur s'est produite. " & "Merci de me contacter en cliquant sur le bouton. " & "Ça va COPIER l'erreur et vous n'aurez qu'à " & "faire COLLER dans mon FORMULAIRE."

'message pour coller l'erreur dans le formulaire
Public Const pcCollerErreurFormulaire As String = "Faites un clic droit puis COLLER dans le formulaire"

'Temps du timer entre chaque ping = délai entre chaque ping
Public Const pcTempsPing = 500

'lien du formulaire du site web
Public Const pcURLformulaire As String = "http://vpnlifeguard.blogspot.com/p/formulaire.html"

'adresse web de la dernière version pour vérifier la mise à jour
Public Const pcFichierIniLastVersion = "http://heanet.dl.sourceforge.net/project/vpnlifeguard/Windows/Version/version.ini"

'adresse du site internet
Public Const pcSiteInternet = "http://sourceforge.net/projects/vpnlifeguard"

' nom du fichier Log
Public Const pcFichLog As String = "Déconnexions"

' adresse IP en local
Public Const vbLocalHost As String = "127.0.0.1"

' nom du ficher exe
Public Const pcExeName As String = "VpnLifeguard"

'type d'opération sur le fichier INI
Public Enum OPERATIONINI
    lire = 1
    Ecrire = 2
End Enum

Public Enum GESTIONPASSERELLE
    activer = 1
    desactiver = 2
End Enum


' controle la thread créée via un atom
' Public car elle doit vivre toute la durée du programme
Public clsAtom As New CMultithreadingSet

' Cet API copie une variable ou une structure définie 'By reference' dans une variable de votre programme.
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

'renvoie l'identifiant de processus de la fenêtre spécifiée
Public Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal HWnd As Long, lpdwProcessId As Long) As Long

'obtient un handle sur un processus
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' curseur de la souris en forme de main
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Public Declare Function SetForegroundWindow Lib "user32" (ByVal HWnd As Long) As Long

Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

'pour donner le style de windows aux contrôles
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long    'DLL inclus dans tout les windows

' check internet connexion
Public Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszURL As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

'ferme un handle
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' test si l'utilisateur est admin
Private Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long
'



Public Sub Main()

    On Error Resume Next
    If CBool(IsNTAdmin(ByVal 0&, ByVal 0&)) = False Then
        MsgBox MLSGetString("0041") _
               & vbNewLine _
               & MLSGetString("0042"), vbExclamation, App.Title
        'MLS->"Vous devez exécuter le programme en tant qu'administrateur"
        'MLS->"Clic droit sur le programme... exécuter en tant que..."

        End
        Exit Sub
    End If
    On Error GoTo 0

    ' si le fichier exe est différent de celui d'origine
    If UCase(App.EXEName) <> UCase(pcExeName) And App.LogMode = 1 Then
        ' copie le fichier avec le bon nom s'il n'existe pas
        If Dir$(App.Path & "\" & pcExeName & ".exe") = vbNullString Then
            On Error Resume Next
            FileCopy App.Path & "\" & App.EXEName & ".exe", App.Path & "\" & pcExeName & ".exe"
            On Error GoTo 0
        End If

        MsgBox MLSGetString("0043"), vbExclamation, App.Title & MLSGetString("0044")    ' MLS-> "Le fichier exécutable a été renommé. Veuillez ne démarrer que VpnLifeguard.exe" ' MLS-> " - erreur"

        DoEvents
        End
        Exit Sub
    End If

    If clsAtom.IsFirstThread Then
        ' première thread, refus d'être instanciée
        If App.StartMode = vbSModeAutomation Then
            err.Raise 9999, , "Unable to be instantiated as a component"
        End If

        ' montre l'interface
        frmMain.Show
    Else
        ' ici le composant est instancié par cette même application
        ' à chaque création de thread via Set x = CreateObject("...") le programme passera ici
    End If
End Sub


' log perso pour faire des tests
Public Sub MyLog(ByVal Valeur As String)
    If pbLogActif = True Then
        EcrireLog "DebugLog", Valeur
    End If
End Sub

' récupère l'index d'une adresse IP à partir de son adresse IP
' dans ce programme, l'index est necéssaire pour faire add route
Public Function GetIndexFromIP(ByVal MonIp As String) As Long

    Dim clsIpLocales As CGetOthersIp
    Dim clsIpGateway As CGetIpgateway
    Dim colRet As New Collection
    Dim MonIndex As Long
    Dim i As Long

    On Error Resume Next

    ' init index
    MonIndex = -1


    Set clsIpLocales = New CGetOthersIp

    ' charge la table d'adresse IP locale dans la collection
    Set colRet = clsIpLocales.Get_Locals_IPs

    ' test chaque IP pour retrouver MonIP
    For i = 1 To colRet.Count
        If MonIp = colRet(i) Then
            ' MonIP a été trouvé donc on récupère son index
            MonIndex = clsIpLocales.GetIndexIP(MonIp)
            Exit For
        End If
    Next i

    ' release the class
    Set clsIpLocales = Nothing

    ' si MonIP n'a pas été retrouvé alors
    ' MonIndex est toujours à -1
    ' donc on change de table d'ip
    If MonIndex < 0 Then

        Set clsIpGateway = New CGetIpgateway

        ' charge les adresses IP de gateway dans une collection
        Set colRet = clsIpGateway.GetIpGateway(True)

        ' boucle pour retrouver MonIP
        For i = 1 To colRet.Count
            If MonIp = colRet(i) Then
                ' MonIP a été retrouvé donc on récupère son index
                MonIndex = clsIpGateway.GetIndexIP(MonIp)
                Exit For
            End If
        Next i

        ' release the class
        Set clsIpGateway = Nothing

    End If

    GetIndexFromIP = MonIndex

    On Error GoTo 0
End Function

Public Function GetFileName(flname As String) As String
    'Get the filename without the path or extension.
    'Input Values:
    '   flname - path and filename of file.
    'Return Value:
    '   GetFileName - name of file without the extension.

    Dim posn As Integer, i As Integer
    Dim fName As String

    posn = 0
    'find the position of the last "\" character in filename
    For i = 1 To Len(flname)
        If (Mid(flname, i, 1) = "\") Then posn = i
    Next i

    'get filename without path
    fName = Right(flname, Len(flname) - posn)

    'get filename without extension
    posn = InStr(fName, ".")
    If posn <> 0 Then
        fName = Left(fName, posn - 1)
    End If
    GetFileName = fName
End Function

Public Function GetFileNameExt(flname As String) As String
    'Get the filename + extension without the path.
    'Input Values:
    '   flname - path and filename of file.
    'Return Value:
    '   GetFileName - name of file + extension.

    Dim posn As Integer, i As Integer
    Dim fName As String

    posn = 0
    'find the position of the last "\" character in filename
    For i = 1 To Len(flname)
        If (Mid(flname, i, 1) = "\") Then posn = i
    Next i

    'get filename without path
    fName = Right(flname, Len(flname) - posn)

    GetFileNameExt = fName
End Function

Public Function GetIpVpn(ByRef oCombo As ComboBox) As Boolean
    Dim colAdressOthersIP As New Collection
    Dim clsVPN As CGetOthersIp

    ' pour récupérer la première adresse Ip local ou pas
    Dim iDebutListing As Byte

10  On Error GoTo GetIpVpn_Error

20  Set clsVPN = New CGetOthersIp

    'récupère toutes les adresses IP locales sauf localhost
    'les adresses sont classées par ordre numérique suivant leurs index
30  Set colAdressOthersIP = clsVPN.Get_Locals_IPs

    'libère la class
40  Set clsVPN = Nothing

    Dim i As Byte

    ' si l'ip du vpn est difficile à récupérer alors
    ' l'utilisateur à choché l'option en cas de problème
    ' qui permet de récupérer la première adresse IP du pc
    iDebutListing = IIf(pbAdressIPproblem, 1, 2)
    ' sinon on n'inclus pas la première adresse IP

    'remplissage du combo du VPN
50  For i = iDebutListing To colAdressOthersIP.Count
60      oCombo.AddItem colAdressOthersIP(i)
70  Next i

    'libère la collection
80  Set colAdressOthersIP = Nothing

    'si le combo n'est pas vide alors on sélectionne la dernière adresse
90  If oCombo.ListCount > 0 Then
100     oCombo.ListIndex = oCombo.ListCount - 1
        GetIpVpn = True
110 End If

120 On Error GoTo 0
130 Exit Function

GetIpVpn_Error:

    GetIpVpn = False
    Set colAdressOthersIP = Nothing
    Set clsVPN = Nothing
140 Call GestionErreur(frmMain, "GetIpVpn", pcErreurMessageSupplementaire)
170 If pbEnvoyerFormulaire = True Then
190     Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
200 End If

End Function

Public Function OLDCheckMiseAjour() As Boolean
    'compare la version actuelle avec la dernière version online
    Dim iMajorOnLine As Byte
    Dim iMinorOnLine As Byte
    Dim iRevisionOnLine As Byte
    Dim bFaireMiseAjour As Boolean

10  On Error GoTo CheckMiseAjour_Error

20  iMajorOnLine = CByte(Left(psDerniereVersionOnLine, 1))
30  iMinorOnLine = CByte(Mid(psDerniereVersionOnLine, 3, 1))
40  iRevisionOnLine = CByte(Mid(psDerniereVersionOnLine, 5))

50  If iMajorOnLine > App.Major Then
60      bFaireMiseAjour = True
70  ElseIf iMinorOnLine > App.Minor Then
80      bFaireMiseAjour = True
90  ElseIf iRevisionOnLine > App.Revision Then
100     bFaireMiseAjour = True
110 Else
120     bFaireMiseAjour = False
130 End If

140 If bFaireMiseAjour = True Then
150     Select Case MsgBox(MLSGetString("0045"), vbYesNo Or vbExclamation Or vbDefaultButton1, MLSGetString("0046"))    ' MLS-> "Une version plus récente est disponible. Voulez-vous la télécharger ?" ' MLS-> "Mise à jour"

        Case vbYes
160         On Error Resume Next
170         ShellExecute 0&, vbNullString, pcSiteInternet, vbNullString, vbNullString, vbNormalFocus
180         On Error GoTo 0

190     Case vbNo
200         Call MsgBox(MLSGetString("0047"), vbInformation Or vbDefaultButton1, MLSGetString("0048"))    ' MLS-> "Si vous rencontrez une erreur, alors pensez à faire la mise à jour." ' MLS-> "Mise à jour conseillée"

210     End Select
220 Else
230     MsgBox MLSGetString("0049"), vbInformation, MLSGetString("0050")    ' MLS-> "C'est bon, vous avez la dernière version ;-)" ' MLS-> "Mise à jour inutile"
240 End If


250 On Error GoTo 0
260 Exit Function

CheckMiseAjour_Error:

270 Call GestionErreur(frmConfig, "CheckMiseAjour", pcErreurMessageSupplementaire)
280 If pbEnvoyerFormulaire = True Then
290     Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
300 End If
End Function

Public Sub SecurisationDuTunnelVPN(Choix As GESTIONPASSERELLE, Optional MessageDebugInfo As Boolean = False)
    Dim ret As Boolean
    Dim clsT As CRouteAddDell

10  On Error GoTo err

20  Set clsT = New CRouteAddDell

30  Select Case Choix

    Case activer
        'Del route
40      ret = clsT.DelRoute(psAdresseIPpasserelle)
50      Debug.Print Timer, "Del route = " & ret
60      If MessageDebugInfo = True Then
70          MsgBox "Del route = " & ret
80      End If

90  Case desactiver
        'Add route

        piIndexPasserelle = GetIndexFromIP(psAdresseIPpasserelle)

100     ret = clsT.AddRoute(psAdresseIPpasserelle, piIndexPasserelle)
110     Debug.Print Timer, "Add route = " & ret
120     If MessageDebugInfo = True Then
130         MsgBox "Add route = " & ret
140     End If
150     ret = Not ret

160 End Select

170 pbTunnelSecurised = ret

180 Set clsT = Nothing
190 On Error GoTo 0
200 Exit Sub

err:
210 Set clsT = Nothing
220 Debug.Print Timer, "route = false"
230 On Error GoTo 0
End Sub

Public Function SplitNullChar(str As Variant) As String
    'coupe un string au nullchar

    If InStr(1, CStr(str), vbNullChar) > 0 Then
        SplitNullChar = Mid$(CStr(str), 1, InStr(1, CStr(str), vbNullChar) - 1)
    Else
        SplitNullChar = CStr(str)
    End If

End Function

Public Sub EcrireLogDeconnexion(ByVal bDéconnexion As Boolean)

    EcrireLog "Déconnexions", IIf(bDéconnexion, " déconnexion", " connexion")

End Sub

' Change le curseur de la souris en main
Public Function MousePointerHand()
    Const IDC_HAND As Long = 32649
    Dim iHandle As Long

    iHandle = LoadCursor(0, IDC_HAND)
    If (iHandle > 0) Then
        iHandle = SetCursor(iHandle)
    End If
End Function

' Test la présence d'un fichier
Public Function IsFileExist(ByVal PathFileName As String) As Boolean
    Dim iFile As Long

    iFile = FreeFile

    On Error GoTo err:

    Open PathFileName For Input As #iFile
    Close #iFile

    IsFileExist = True
    Exit Function

err:     IsFileExist = False
End Function

' test la connexion internet
Public Function CheckInternetConnection() As Boolean
    CheckInternetConnection = InternetCheckConnection("http://www.google.com", &H1, 0&)
End Function

