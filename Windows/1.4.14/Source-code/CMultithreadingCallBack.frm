VERSION 5.00
Begin VB.Form CMultithreadingCallBack 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2040
      Top             =   1320
   End
End
Attribute VB_Name = "CMultithreadingCallBack"
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


' ------------------------------------------
' Titre : création d'une thread en multithreading activeX
' Auteur : philippe734
' Date : mai 2010
'
' Deux classes sont nécessaires.
' La classe multiuse pour définir la procédure à exécuter en thread
' et la form callback pour effectuer un court délai.
'
' But :
' On souhaite exécuter une procédure
' dans une thread via le multithreading.
' Principe :
' 1- Créer et tester un atom afin de ne pas lancer plusieurs instance
' du programme principale.
' 2- Création de la thread en définissant la class, sa procédure
' et ses variables à exécuter dans la thread.
' 3- Création d'un délai (callback) entre la création
' et l'exécution de la thread.
' 4- Une fois l'exécution de la procédure terminée,
' les résultats doivent être impérativement récupérés
' par du RaiseEvent.
' C'est donc indispensable que la procédure de la thread puisse
' générer des events via RaiseEvent.
' La procédure à envoyer en thread doit être dans une
' class multiuse.
' ------------------------------------------

Option Explicit



' Crée un délai entre la création de la thread
' et l'exécution de la procécure à exécuter en thread.
' Ce délai est indispensable, sinon le multithreading
' ne sera pas réalisé.

' variables pour les copies locales
Private msMethodeName As String
Private m_Argument_A As Variant
Private m_Argument_B As Variant
Private m_Argument_C As Variant
Private miThreadIndex As Long
Private moClass As Object
'


'-----------------------------------------------------
' Procédure qui crée un délai via un timer
' nécessaire et obligatoire sinon la thread ne sera pas créé en multithread mais en single thread
'-----------------------------------------------------
Public Sub DelayedCall(oClass As Object, ByVal MethodeName As String, ByRef ThreadIndex As Long, Optional ByVal Argument_A As Variant, Optional ByVal Argument_B As Variant, Optional ByVal Argument_C As Variant)


' compteur de thread qui est ByRef
' récupéré par la procédure à exécuter en thread
    ThreadIndex = ThreadIndex + 1

    ' copies locales
    msMethodeName = MethodeName
    m_Argument_A = Argument_A
    m_Argument_B = Argument_B
    m_Argument_C = Argument_C
    miThreadIndex = ThreadIndex


    ' instance locale de la class
    Set moClass = oClass


    ' démarre le timer
    ' obligatoire pour créer une thread multithreading.
    ' Une valeur trop faible aura le même effet que sans timer
    Timer1.Interval = 50
    Timer1.Enabled = True

End Sub


Private Sub Timer1_Timer()

' on avait juste besoin d'un délai
    Timer1.Enabled = False



    ' effectue le callback en fonction des arguments valides

    If Not IsMissing(m_Argument_C) Then
        CallByName moClass, msMethodeName, VbMethod, miThreadIndex, m_Argument_A, m_Argument_B, m_Argument_C
        Debug.Print Timer, "callback avec argu C"

    ElseIf Not IsMissing(m_Argument_B) Then
        CallByName moClass, msMethodeName, VbMethod, miThreadIndex, m_Argument_A, m_Argument_B
        Debug.Print Timer, "callback avec argu B et sans argu C"

    ElseIf Not IsMissing(m_Argument_A) Then
        CallByName moClass, msMethodeName, VbMethod, miThreadIndex, m_Argument_A
        Debug.Print Timer, "callback que argu A"

    ElseIf IsMissing(m_Argument_A) Then
        CallByName moClass, msMethodeName, VbMethod, miThreadIndex
        Debug.Print Timer, "callback sans argu"

    End If


    Set moClass = Nothing


    ' voilà pour le délai, merci, au revoir
    Unload Me


End Sub



Private Sub Form_Load()

    MLSLoadLanguage Me '<- Add by Multi-Languages Support Add-in
    '/ This Form_Load is add by Multi-Languages Support
End Sub
