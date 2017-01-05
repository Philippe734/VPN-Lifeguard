VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
' Module    : CUpdate
' Author    : philippe734
' Date      : 22/06/2010
' Purpose   : compare la version du programme avec la dernière version
'---------------------------------------------------------------------------------------

Option Explicit

Private miMajorOnLine As Integer
Private miMinorOnLine As Integer
Private miRevisionOnLine As Integer
'

Public Sub CheckMiseAjour()
'compare la version actuelle avec la dernière version online
    Dim bFaireMiseAjour As Boolean

    If miMajorOnLine > App.Major Then
        bFaireMiseAjour = True
        
    ElseIf miMajorOnLine < App.Major Then
        bFaireMiseAjour = False
        
    ElseIf miMinorOnLine < App.Minor Then
        bFaireMiseAjour = False
        
    ElseIf miMinorOnLine > App.Minor Then
        bFaireMiseAjour = True
        
    ElseIf miRevisionOnLine > App.Revision Then
        bFaireMiseAjour = True
        
    Else
        bFaireMiseAjour = False
    End If

    If bFaireMiseAjour = True Then
        Select Case MsgBox(MLSGetString("0068"), vbYesNo Or vbExclamation Or vbDefaultButton1, MLSGetString("0069")) ' MLS-> "Une version plus récente est disponible. Voulez-vous la télécharger ?" ' MLS-> "Mise à jour"

        Case vbYes
            On Error Resume Next
            ShellExecute 0&, vbNullString, pcSiteInternet, vbNullString, vbNullString, vbNormalFocus
            On Error GoTo 0

        Case vbNo
            Call MsgBox(MLSGetString("0070"), vbInformation Or vbDefaultButton1, MLSGetString("0071")) ' MLS-> "Si vous rencontrez une erreur, alors pensez à faire la mise à jour." ' MLS-> "Mise à jour conseillée"

        End Select
    Else
        MsgBox MLSGetString("0072"), vbInformation, MLSGetString("0073") ' MLS-> "C'est bon, vous avez la dernière version ;-)" ' MLS-> "Mise à jour inutile"
    End If


End Sub


Public Property Get MajorOnLine() As Integer

    MajorOnLine = miMajorOnLine

End Property

Public Property Let MajorOnLine(ByVal iMajorOnLine As Integer)

    miMajorOnLine = iMajorOnLine

End Property

Public Property Get MinorOnLine() As Integer

    MinorOnLine = miMinorOnLine

End Property

Public Property Let MinorOnLine(ByVal iMinorOnLine As Integer)

    miMinorOnLine = iMinorOnLine

End Property

Public Property Get RevisionOnLine() As Integer

    RevisionOnLine = miRevisionOnLine

End Property

Public Property Let RevisionOnLine(ByVal iRevisionOnLine As Integer)

    miRevisionOnLine = iRevisionOnLine

End Property
