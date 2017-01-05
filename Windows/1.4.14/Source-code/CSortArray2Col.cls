VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSortArray2Col"
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
' Module    : CSortArray2Col
' Author    : unknown
' Date      : ?
' Purpose   : trie un tableau à deux dimensions
'---------------------------------------------------------------------------------------

Option Explicit

'-----------------------------------------------------
'
'ATTENTION : il faut une ligne et une colonne de plus
'au tableau pour que le trie se réalise
'
'
'   Déclaration :
'    Dim cls As CSortArray2Col
'    Set cls = New CSortArray2Col
'    Dim ret As Boolean
'
'   Lancement du trie avec un tableau TABLO suivant la colonne 0
'    ret = cls.SortArray(TABLO, 0)
'    Set cls = Nothing
'
'   Récupérer les données triées une par une
'    For R = 0 To UBound(TABLO, 2) - 1
'        For C = 0 To UBound(TABLO, 1) - 1
'            ret = TABLO(C, R)
'        Next C
'    Next R
'-----------------------------------------------------

'trie un tableau à 2 colonnes
Public Function SortArray(ByRef DArray(), ByVal Element As Byte) As Boolean

    Dim gap As Integer
    Dim acol As Integer
    Dim CNT As Integer
    Dim doneflag As Boolean
    Dim SwapArray()
    Dim Index As Integer

10  On Error GoTo SortArray_Error

20  ReDim SwapArray(2, UBound(DArray, 1), UBound(DArray, 2))
    'Gap is half the records
30  gap = Int(UBound(DArray, 2) / 2)
40  Do While gap >= 1
50      Do
60          doneflag = True
70          For Index = 0 To (UBound(DArray, 2) - (gap + 1))
                'Compare 1st 1/2 to 2nd 1/2
80              If DArray(Element, Index) > DArray(Element, (Index + gap)) Then
90                  For acol = 0 To (UBound(SwapArray, 2) - 1)
                        'Swap Values if 1st > 2nd
100                     SwapArray(0, acol, Index) = DArray(acol, Index)
110                     SwapArray(1, acol, Index) = DArray(acol, Index + gap)
120                 Next acol
130                 For acol = 0 To (UBound(SwapArray, 2) - 1)
                        'Swap Values if 1st > 2nd
140                     DArray(acol, Index) = SwapArray(1, acol, Index)
150                     DArray(acol, Index + gap) = SwapArray(0, acol, Index)
160                 Next acol
170                 CNT = CNT + 1
180                 doneflag = False
190             End If
200         Next
210     Loop Until doneflag = True
220     gap = Int(gap / 2)
230 Loop

240 SortArray = True

250 On Error GoTo 0
260 Exit Function

SortArray_Error:

270 SortArray = False
    'MsgBox "Error " & err.Number & " (" & err.Description & ") line " & IIf(Erl = 0, "(none)", Erl) & " in SortArray of CSortArray2Col", vbCritical

End Function
