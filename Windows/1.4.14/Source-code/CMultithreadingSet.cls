VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMultithreadingSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

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
' Deux classes sont nécessaires (+ une pour la procédure à exécuter)
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
' et l'exécution de la procédure dans la thread.
' 4- Une fois l'exécution de la procédure terminée,
' les résultats doivent être impérativement envoyés
' par RaiseEvent.
' C'est donc indispensable que la procédure de la thread puisse
' générer des events via RaiseEvent.
' La procédure à exécuter dans une thread doit être dans une
' class multiuse.
' ------------------------------------------


' Cette présente class est utilisée pour deux usages
' 1- créer et tester un atom
' 2- déclancher la création d'une thread


Option Explicit



' 1er usage : crée un atom afin de ne pas instancié la fenetre principale
' un atom est une sorte de variable globale windows

' API de gestion d'atom
Private Declare Function FindAtom Lib "kernel32" Alias "FindAtomA" (ByVal atomName As String) As Integer
Private Declare Function AddAtom Lib "kernel32" Alias "AddAtomA" (ByVal atomName As String) As Integer
Private Declare Function DeleteAtom Lib "kernel32" (ByVal atomName As Integer) As Integer

' l'identifiant de l'atom créé
Private atomID As Integer
'

' dès son premier lancement, un atom sera créé
Private Sub Class_Initialize()
    Dim atomName As String

    ' Build an atom name unique for this instance of the application.
    atomName = App.EXEName & App.hInstance

    ' Create the atom if it doesn't exist already.
    If FindAtom(atomName) = 0 Then atomID = AddAtom(atomName)
End Sub

Private Sub Class_Terminate()

' Delete the atom when this thread terminates.
    If atomID Then DeleteAtom atomID
End Sub

Public Function IsFirstThread() As Boolean

' This is the first thread if it was the one which created the atom.
    IsFirstThread = (atomID <> 0)
End Function


' 2è usage : déclancher la création de la thread
' étape intermédiaire mais obligatoire car
' pour créer une thread, il faut une class MultiUse
' or la form CallBack est une class non MultiUse
'-----------------------------------------------------

' définit les paramètres de votre classe
Public Sub SetThread(oClass As Object, ByVal MethodeName As String, ByRef ThreadIndex As Long, Optional ByVal Argument_A As Variant, Optional ByVal Argument_B As Variant, Optional ByVal Argument_C As Variant)

' nom de votre class contenant la procédure
' à exécuter en thread
    Dim l_oClass As Object

    ' class qui va créer un délai appelé callback
    Dim cbk As CMultithreadingCallBack

    Set cbk = New CMultithreadingCallBack
    Set l_oClass = oClass

    ' appel la class qui va créer un délai
    ' entre la création de la thread et l'exécution
    ' de votre procédure
    Call cbk.DelayedCall(l_oClass, MethodeName, ThreadIndex, Argument_A, Argument_B, Argument_C)

    Set cbk = Nothing
    Set l_oClass = Nothing

End Sub
