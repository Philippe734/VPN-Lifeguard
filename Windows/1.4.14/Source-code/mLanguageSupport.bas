Attribute VB_Name = "modLanguageSupport"

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

' -----------------------------------------------------------------
' This module was made by Multi-Language Support Add-in for VB,
' by Giorgio Brausi (gibra)
' Contact me by e-mail: vbcorner@lycos.it or gibra@amc2000.it
' Web site: http://utenti.lycos.it/vbcorner/
' -----------------------------------------------------------------

Option Explicit

Public gsLanguageFile As String  '/language file (i.e. english.lng, italian.lng,...)
Public listLanguage As New Collection
Private bInit As Boolean

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'/ Update all controls properties with the current language
Public Sub MLSLoadLanguage(Form As Form)
    Dim obj As Object
    Dim sFileName As String, a As String
    
    On Error Resume Next
    
    ' If initialize not done, then do it
    If bInit = False Then MLSFillMenuLanguages
    
    If gsLanguageFile = vbNullString Then Exit Sub
    
    If Right(App.Path, 1) = "\" Then
        sFileName = App.Path & gsLanguageFile & ".lng"
    Else
        sFileName = App.Path & "\" & gsLanguageFile & ".lng"
    End If

    '/ Load Caption for Form, if there
    If Len(Form.Caption) > 0 Then
        Form.Caption = MLSReadINI(sFileName, CStr(Form.Name), CStr(Form.Name) & ".Caption")
    End If

    Form.ToolTipText = MLSReadINI(sFileName, CStr(Form.Name), CStr(Form.Name) & ".ToolTipText")
    Form.Tag = MLSReadINI(sFileName, CStr(Form.Name), CStr(Form.Name) & ".Tag")

    '/ Load properties for objects
    For Each obj In Form
        Dim bHasIndex As Boolean '/ to check if has Index property
        a$ = ""
        '/ If is not a matrix return a error code 343
        bHasIndex = (obj.Index >= 0)
        If err.Number = 343 Then     '/ The object is not a matrix
            bHasIndex = False
            err.Clear
        End If

        '/ Get Caption property
        If bHasIndex Then '/ This is a matrix
            a$ = MLSReadINI(sFileName, CStr(Form.Name), obj.Name & "(" & obj.Index & ").Caption")
        Else
            a$ = MLSReadINI(sFileName, CStr(Form.Name), obj.Name & ".Caption")
        End If
        If a$ <> "" Then
            obj.Caption = a$
        End If

        '/ Get ToolTipText property
        a$ = ""
        If bHasIndex Then '/ This is a matrix
            a$ = MLSReadINI(sFileName, CStr(Form.Name), obj.Name & "(" & obj.Index & ").ToolTipText")
        Else
            a$ = MLSReadINI(sFileName, CStr(Form.Name), obj.Name & ".ToolTipText")
        End If
        If a$ <> "" Then
            obj.ToolTipText = a$
        End If

        '/ Get Tag property
        a$ = ""
        If bHasIndex Then '/ This is a matrix
            a$ = MLSReadINI(sFileName, CStr(Form.Name), obj.Name & "(" & obj.Index & ").Tag")
        Else
            a$ = MLSReadINI(sFileName, CStr(Form.Name), obj.Name & ".Tag")
        End If
        If a$ <> "" Then
            obj.Tag = a$
        End If

        '/ check properties for SSTab control: this control has
        '/ Caption for each tab, named TabCaption
        If obj.Tabs Then
            If err = 0 Then
                a$ = ""
                Dim nT As Integer
                '/ find the caption for each Tab
                For nT = 0 To obj.Tabs
                     obj.TabCaption(nT) = MLSReadINI(sFileName, CStr(Form.Name), obj.Name & ".TabCaption(" & nT & ")")
                Next nT
            Else
                err.Clear
            End If
        End If
        
        DoEvents
    Next
End Sub

'/ Load a translate string from Section [Strings] of currente language
Public Function MLSGetString(KeyName As String) As String
Dim sFileName As String

    On Error Resume Next
    If gsLanguageFile = "" Then
        gsLanguageFile = MLSReadINI(App.Path & "\" & "LangSetting.ini", "Language", "CurrentLanguage")
    End If
    If Right(App.Path, 1) = "\" Then
        sFileName = App.Path & gsLanguageFile & ".lng"
    Else
        sFileName = App.Path & "\" & gsLanguageFile & ".lng"
    End If
    MLSGetString = MLSReadINI(sFileName, "Strings", KeyName$)
End Function

Public Function MLSReadINI(File$, SectionName$, KeyName$) As String
Dim value As String * 1024, i As Long

    '/ If the file INI is large than 64k, GetPrivateProfileString can fail
    '/ then use a my function to retrieve the value:
    If FileLen(File$) > 64000 Then

        '/ Use Open
        Dim numFile As Integer, sThisLine As String, sTmp As String
        numFile = FreeFile
        Open File$ For Input As #numFile
        Do While Not EOF(numFile)
            Line Input #numFile, sThisLine
            If Left(sThisLine, Len(KeyName$)) = KeyName$ Then
                sTmp = Mid(sThisLine, Len(KeyName$) + 2)
                MLSReadINI = Mid(sTmp, 2, Len(sTmp) - 2)
                Close #numFile
                Exit Function
            End If
        Loop
        Close #numFile

    Else
        i = GetPrivateProfileString(SectionName$, KeyName$, "", value, 512, File$)
        MLSReadINI = Left$(value, InStr(value, Chr$(0)) - 1)
    End If
End Function

Public Function MLSWriteINI(File$, SectionName$, KeyName$, NewValue$) As Long
    MLSWriteINI = WritePrivateProfileString(SectionName$, KeyName$, NewValue$, File$)
End Function

Public Sub MLSFillMenuLanguages()
'/ This sub is created by Multi-Language Support
'/ Search for all language file, load & fill the menu item array
'/ named: mnuLanguage
    Dim iNumLanguages As Integer, i As Integer, sFileName As String
    
    On Error Resume Next

    '/ Search for language files. Folder is App.Path
    sFileName = Dir(App.Path & "\*.lng")
    If sFileName = "" Then
        MsgBox "No language file exist", vbExclamation
        Exit Sub
    End If

    '/ Now, for each language file add a new menu item
    '/ -----------------------------------------------
    Do While sFileName <> ""
        listLanguage.Add Mid(sFileName, 1, Len(sFileName) - 4)
        iNumLanguages = iNumLanguages + 1
        sFileName = Dir
    Loop

    '/ Get the current language from the file: LangSetting.ini
    '/ Note: the file LangSetting.ini is create by MLS, but if you want
    '/ you can delete this and add this section in your custom INI file, if there.
    If iNumLanguages > 0 Then
        gsLanguageFile = MLSReadINI(App.Path & "\" & "LangSetting.ini", "Language", "CurrentLanguage")
    End If

    bInit = True

End Sub

