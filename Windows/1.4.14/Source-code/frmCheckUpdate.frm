VERSION 5.00
Begin VB.Form frmCheckUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerUnloadMe 
      Left            =   5040
      Top             =   120
   End
   Begin VB.Label LabelInfo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Connexion aux serveurs SourceForge en cours..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmCheckUpdate"
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
' Module    : frmCheckUpdate
' Author    : philippe734
' Date      : 22/04/2010
' Purpose   : affiche une fenetre durant le téléchargement et compare
' la version actuelle du programme avec la dernière
'---------------------------------------------------------------------------------------

Option Explicit

' class pour télécharger le fichier contenant le numéro de la dernière version
Private WithEvents m_clsDownload As CDownloader
Attribute m_clsDownload.VB_VarHelpID = -1
'

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub m_clsDownload_Error(ByVal vsURL As String, ByVal vsMessage As String, ByVal vnCode As Long)
    Debug.Print Timer, "m_clsDownloader_Error"

    Set m_clsDownload = Nothing

    MsgBox MLSGetString("0064") & vbNewLine & MLSGetString("0065") & vsMessage, vbExclamation ' MLS-> "Impossible de contacter les serveurs" ' MLS-> "Err type : "

    Screen.MousePointer = vbDefault

    pbTelechargementReussis = False

    Me.TimerUnloadMe.Interval = 1000
    Me.TimerUnloadMe.Enabled = True


End Sub

Private Sub m_clsDownload_Finished(ByVal vsURL As String, ByVal vsTarget As String)
    Dim clsUpdate As CUpdate

10  On Error GoTo m_clsDownload_Finished_Error

    'Debug.Print timer, "m_clsDownload_Finished"

20  Set m_clsDownload = Nothing

    'téléchargement réussis
30  pbTelechargementReussis = True

40  Me.LabelInfo.Caption = MLSGetString("0066") ' MLS-> "  Comparaison des versions..."

    'initialise la valeur en retour
50  psDerniereVersionOnLine = vbNullString

    'récupère le numéro de la dernière version online
    ' à partir du fichier ini téléchargé
60  On Error Resume Next
70  psDerniereVersionOnLine = LireINI("Dernière version", "v", vsTarget)
80  DoEvents
90  Kill vsTarget
100 On Error GoTo m_clsDownload_Finished_Error

    'pour tester
    'psDerniereVersionOnLine = "1.4.1"
    'Debug.Print timer, "Dernière version online = " & psDerniereVersionOnLine & ";"
    'Debug.Print timer, "Cette version = " & App.Major & "." & App.Minor & "." & App.Revision; ";"

110 Screen.MousePointer = vbDefault

120 Me.LabelInfo.Caption = MLSGetString("0067") & psDerniereVersionOnLine ' MLS-> "  Dernière version online = "

    'vérifie la version actuelle avec la dernière

130 Set clsUpdate = New CUpdate

140 clsUpdate.MajorOnLine = CInt(Left(psDerniereVersionOnLine, 1))
150 clsUpdate.MinorOnLine = CInt(Mid(psDerniereVersionOnLine, 3, 1))
160 clsUpdate.RevisionOnLine = CInt(Mid(psDerniereVersionOnLine, 5))

170 Call clsUpdate.CheckMiseAjour

180 Set clsUpdate = Nothing

190 Me.TimerUnloadMe.Interval = 1000
200 Me.TimerUnloadMe.Enabled = True

210 On Error GoTo 0
220 Exit Sub

m_clsDownload_Finished_Error:

230 Call GestionErreur(frmCheckUpdate, "m_clsDownload_Finished", pcErreurMessageSupplementaire)
240 If pbEnvoyerFormulaire = True Then
250     Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
260 End If

End Sub

Private Sub Form_Load()

    MLSLoadLanguage Me '<- Add by Multi-Languages Support Add-in

    Me.Caption = App.Title
    Me.Width = 4905
    Me.Icon = frmMain.ImageIconNormal

End Sub


Private Sub Form_Activate()

10  On Error GoTo popCheckUpdate_Error

30  Screen.MousePointer = vbHourglass

40  Set m_clsDownload = New CDownloader


    'télécharge le fichier de version en écrasant l'ancien
50  m_clsDownload.GetFile pcFichierIniLastVersion, App.Path & "\version.ini"

    ' la suite du programme se passe dans les events via RaiseEvent


60  On Error GoTo 0
70  Exit Sub

popCheckUpdate_Error:

80  Call GestionErreur(frmCheckUpdate, "TimerStart_Timer", pcErreurMessageSupplementaire)
90  If pbEnvoyerFormulaire = True Then
100     Call PopupBalloon(frmMain, App.Title, pcCollerErreurFormulaire)
110 End If

End Sub


Private Sub TimerUnloadMe_Timer()

    Me.TimerUnloadMe.Enabled = False

    Unload Me
End Sub
