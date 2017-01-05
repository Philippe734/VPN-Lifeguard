VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWidget"
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
' Module    : CWidget
' Author    : philipe734
' Date      : 2010/07
' Purpose   : Créer un widget en trois parties
'               - Partie Shape     : Donner l'aspect d'une image à une fenêtre et la rendre un peu transparente
'               - Partie Tooltip   : Créer une infobulle
'               - Partie MenuPopup : Créer un menu popup par API
'---------------------------------------------------------------------------------------


Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal HWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal HWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal HWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal HWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal HWnd As Long, lpRect As RECT) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal HWnd As Long, ByVal lptpm As Any) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' Init Common Controls Lib, put in startup module and execute
' Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


'Windows API Constants
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_ERASE = &H4
Private Const RDW_FRAME = &H400
Private Const RDW_INVALIDATE = &H1
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const WS_POPUP As Long = &H80000000
Private Const WS_BORDER As Long = &H800000

'Tooltip Window Constants
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = WM_USER + 4
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = WM_USER + 12
Private Const TTmsETMAXTIPWIDTH = WM_USER + 24
Private Const TTmsETTIPBKCOLOR = WM_USER + 19
Private Const TTmsETTIPTEXTCOLOR = WM_USER + 20
Private Const TTmsETTITLE = WM_USER + 32
Private Const TTS_NOPREFIX = &H2
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10

'Tool Tip Icons
Private Const TTI_ERROR As Long = 3
Private Const TTI_INFO As Long = 1
Private Const TTI_NONE As Long = 0
Private Const TTI_WARNING As Long = 2

'Tool Tip API Class
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

' Menu Popup constants
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_RIGHTBUTTON = &H2&
Private Const TPM_RETURNCMD = &H100&

'Windows API Types
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Menu popup coordinates
Private Type POINTAPI
    x As Long
    y As Long
End Type

'Tooltip Window Types
Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    TiphWnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpszText As String
    lParam As Long
End Type

Private mToolBackColor As Long
Private toolTitle As String
Private toolForeColor As Long
Private toolParentControl As Long
Private toolIcon As toolIconType
Private toolCentered As Boolean
Private toolStyle As toolStyleEnum
Private mToolTipText As String
Private mToolEnabled As Boolean
Private toolhWnd As Long

'Private Data for Class
Private TiphWnd As Long
Private ti As TOOLINFO

' handle de notre fenêtre
Private m_hwndFrm As Long

' copie des propriétés pour KillShape
Private m_FrmWidth As Long
Private m_FrmHeight As Long
Private m_FrmBackcolor As Long

' collections pour créer le menupopup
Private colMnuCaptions As New Collection
Private colMnuStyles As New Collection

Public Enum toolIconType
    tipNoIcon = TTI_NONE            '= 0
    tipiconinfo = TTI_INFO          '= 1
    tipIconWarning = TTI_WARNING    '= 2
    tipIconError = TTI_ERROR        '= 3
End Enum

Public Enum toolStyleEnum
    styleStandard = 0
    styleBalloon = 1
End Enum

Public Enum MENUPOPUpsTYLE
    MF_CHECKED = &H8&
    MF_DISABLED = &H2&
    MF_GRAYED = &H1&
    MF_SEPARATOR = &H800&
    MF_STRING = &H0&
End Enum
'




'----------------------------------------------------------------------------
'
' Partie Shape
'
'----------------------------------------------------------------------------

'************************************************************
'
' Create uniquely shaped and partially transparent windows
'
' 2003/01
'
' http://www.vbaccelerator.com/home/vb/code/libraries/Graphics_and_GDI/Changing_Window_Shapes/Window_Shapes_Using_Layering/article.asp
'
'************************************************************

Private Sub BorderLess()
    Const WS_CAPTION = &HC00000
    Const GWL_STYLE = (-16)
    Const SWP_FRAMECHANGED = &H20
    Const SWP_NOMOVE = &H2
    Const SWP_NOZORDER = &H4
    Const SWP_NOSIZE = &H1
    Dim lStyle As Long

    ' Récupère le style en cours
    lStyle = GetWindowLong(m_hwndFrm, GWL_STYLE)

    ' Permute le style de la barre de titre
    lStyle = lStyle And Not WS_CAPTION

    ' Affecte le changement de style
    SetWindowLong m_hwndFrm, GWL_STYLE, lStyle
    SetWindowPos m_hwndFrm, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE

End Sub

Public Sub Moving(ByRef Frm As Form)
' Annule le clic
    ReleaseCapture
    ' Simule un clic sur sa barre de titre, donc la fenêtre se déplace
    SendMessageLong Frm.HWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Public Sub DoShape(ByRef Frm As Form, ByRef pict As PictureBox, ByVal ColorTransparency As Long, ByVal OpacityLevel As Byte)
    Dim lStyle As Long

    m_hwndFrm = Frm.HWnd

    ' le widget ne dois pas avoir de bordure
    'Call BorderLess

    ' sauvegarde les valeurs pour KillShape
    m_FrmHeight = Frm.Height
    m_FrmBackcolor = Frm.BackColor
    m_FrmWidth = Frm.Width

    Frm.Width = pict.ScaleWidth
    Frm.Height = pict.ScaleHeight
    '    Frm.Width = pict.Width ' utilisé par l'objet image
    '    Frm.Height = pict.Height ' utilisé par l'objet image

    ' définit l'image de notre fenêtre avec notre image
    Set Frm.Picture = pict.Picture

    ' définit la couleur du masque pour faire la transparence
    Frm.BackColor = ColorTransparency

    ' récupère un style de notre fenêtre
    lStyle = GetWindowLong(m_hwndFrm, GWL_EXSTYLE)

    ' modifie le style en ajoutant un calque
    lStyle = lStyle Or WS_EX_LAYERED

    ' redéfinit le style avec le calque
    SetWindowLong m_hwndFrm, GWL_EXSTYLE, lStyle

    ' applique les changements de forme et de transparence à notre fenêtre
    SetLayeredWindowAttributes m_hwndFrm, Frm.BackColor, OpacityLevel, LWA_COLORKEY Or LWA_ALPHA

End Sub

Public Sub KillShape(ByRef Frm As Form)

    Dim lStyle As Long
    Dim lhWnd As Long

    If m_hwndFrm <= 0 Then Exit Sub

    ' rétablie l'aspect de la fenêtre
    lhWnd = Frm.HWnd
    lStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
    lStyle = lStyle And Not WS_EX_LAYERED
    SetWindowLong lhWnd, GWL_EXSTYLE, lStyle
    RedrawWindow lhWnd, 0, 0, RDW_ERASE Or RDW_INVALIDATE Or RDW_FRAME Or RDW_ALLCHILDREN

    ' rétablie les valeurs sauvegardées
    Frm.BackColor = m_FrmBackcolor
    Frm.Width = m_FrmWidth
    Frm.Height = m_FrmHeight


End Sub



'----------------------------------------------------------------------------
'
' Partie Tooltip
'
'----------------------------------------------------------------------------

'************************************************************
'
' Tool Tip Class
'
' 17-NOV-2004
'
' Mark Mokoski
' C & M Telephone
' markm@cmtelephone.com
' www.rjillc.com
'
' Class for the creation of Rectangular and Balloon Multiline Tool Tips.
'
' See Code for details on Syntax, variables and constants.
'
' This Class Module works by sub classing the parent control.
' Works with most of the common controls, the ones it will not work
' with are controls that do not support tool tips themselves.
' There are some controls that support tool tips but complain with this
' sub classing module (only Microsoft knows why!).
' Do not use the native tool tip with the parent control and this together,
' you get overlapping tool tips!
'
' This class module is also know to work on the following OS's (all Win32)
' Windows Me
' Windows 2000
' Windows XP
' MSDN Docs state that IE 5 or higher is needed for Balloon tips
'
' It's usefull as it is written, so give it a try!
'
' As brought to my attention on 01-DEC-2004
' Parts of this class was writen by Eidos on PSC
' I found it in bits on other sites and did the usual inprovments
' and changes to for my needs.
' So, credit for the core of this class goes to Eidos at PSC
' Eidos's code can be founfd at:
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=28419&lngWId=1
'**************************************************************
'
' Public Methods:
'
' .ToolTipCreateBalloon     Create Balloon Tool Tip
' .CreateToolTip         Create Rectangular Tool Tip
' .Remove            Kills Tool Tip object
'
' Public Properties:
'
' .ToolTipEnabled            Boolean    Activate (visible)/Deactivate (hide) tool tip
' .ToolTipParent     Long       hWnd of Control that the tool tip is subclassed (displays on)to
' .ToolTipStyle             Enum Type  Tool Tip ToolTipStyle, Rectangular or Balloon
' .ToolTipCentered          Boolean    Tool Tip is ToolTipCentered on parent control when visible
' .ToolTipIcon              Enum Type  Tool Tip Icon used when tool tip has a title
' .ToolTipTitle             Text       Tool Tip ToolTipTitle text
' .ToolTipFore Color        Long       Tool Tip text color and border color if Balloon tip
' .ToolTipBackColor         Long       Tool Tip Background color
' .ToolTipText           Text       Tool Tip text
' .ToolTiGethWnd              Long       Tool Tip Windows Handle (.hWnd) READ ONLY
'
' Sample code:
'************************************************************
'    Option Explicit
'
'   'Make new tool tip object for this project
'   Dim Command1Tip              As New clsTooltips
'
'   Private Sub Form_Load()
'
'   'Make the complete Tool Tip, text, title, icon
'   Command1Tip.ToolTipCreateBalloon Command1, "I turned off all the Tool Tips " + vbCrLf + "Click to restore Tool Tips", "Tool Tips are OFF", tipIconWarning
'
'   End Sub
'
'   Put this delare in the Sub_Main module...
'       'Int Common Controls Lib
'       Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
'
'   Then call the sub in the Sub_Main
'       'Int Common Controls Lib
'       InitCommonControls
'
'************************************************************



Public Function ToolTipCreateBalloon(ByVal Parent As Object, Text As String, Optional Title As String = vbNullString, Optional Icon As Integer = TTI_NONE)

'Used to create a Balloon Tool Tip object.
'Pass needed parameters with call.
'Syntax: object.ToolTipCreateBalloon ToolTipParent, Tip Text, Title, Icon
'Title and Icon are optional, but you cant have an Icon without a Title
'Title can be just a space, just not an empty string

    toolStyle = styleBalloon
    toolParentControl = Parent.HWnd
    mToolTipText = Text

    If Title = vbNullString Then
        toolTitle = ""
    Else
        toolTitle = Title
    End If

    If Icon = Icon > TTI_ERROR Then
        toolIcon = TTI_NONE
    Else
        toolIcon = Icon
    End If

    Call CreateTTP
    ToolTipEnabled = True

End Function

' --- modif by philippe734
Public Function ToolTipCreate(ByVal hwndParent As Long, Text As String, Optional Title As String = vbNullString, Optional Icon As Integer = TTI_NONE)
'Public Function CreateToolTip(ByVal Parent As Object, Text As String, Optional Title As String = vbNullString, Optional Icon As Integer = TTI_NONE)
' --- end modif

'Used to create a Standard (rectangle) Tool Tip object.
'Pass needed parameters with call.
'Syntax: object.CreateToolTip ToolTipParent, Tip Text, Title, Icon
'Title and Icon are optional, but you cant have an Icon without a Title
'Title can be just a space, just not an empty string

    toolStyle = styleStandard
    ' --- modif by philippe734
    toolParentControl = hwndParent
    'toolParentControl = Parent.hwnd
    ' --- end modif
    mToolTipText = Text

    If Title = vbNullString Then
        toolTitle = ""
    Else
        toolTitle = Title
    End If

    If Icon = vbNull Or Icon > TTI_ERROR Then
        toolIcon = TTI_NONE
    Else
        toolIcon = Icon
    End If

    Call CreateTTP
    ToolTipEnabled = True

End Function

Private Sub CreateTTP()

'Private sub used with Create and Update subs/functions

    Dim lpRect As RECT
    Dim lWinStyle As Long

    'If Tool Tip already made, destroy it and reconstruct

    If TiphWnd <> 0 Then
        DestroyWindow TiphWnd
    End If

    lWinStyle = WS_POPUP Or WS_BORDER Or TTS_ALWAYSTIP Or TTS_NOPREFIX

    'Create baloon style if desired

    If toolStyle = styleBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON

    'The parent control has to be set first

    If toolParentControl <> &H0 Then
        TiphWnd = CreateWindowEx(0&, _
                                 TOOLTIPS_CLASSA, _
                                 vbNullString, _
                                 lWinStyle, _
                                 CW_USEDEFAULT, _
                                 CW_USEDEFAULT, _
                                 CW_USEDEFAULT, _
                                 CW_USEDEFAULT, _
                                 toolParentControl, _
                                 0&, _
                                 App.hInstance, _
                                 0&)

        toolhWnd = TiphWnd

        'Make our tooltip window a topmost window
        SetWindowPos TiphWnd, _
                     HWND_TOPMOST, _
                     0&, _
                     0&, _
                     0&, _
                     0&, _
                     SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE

        'Get the rectangle of the parent control
        GetClientRect toolParentControl, lpRect

        'Now set up our tooltip info structure

        With ti
            'If we want it centered, then set that flag

            If toolCentered Then
                .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
            Else
                .lFlags = TTF_SUBCLASS
            End If

            ''set the hwnd prop to our parent control's hwnd
            .TiphWnd = toolParentControl
            .lId = 0
            .hInstance = App.hInstance
            .lpszText = mToolTipText
            .lpRect = lpRect
        End With

        'Add the tooltip structure
        SendMessage TiphWnd, TTM_ADDTOOLA, 0&, ti

        'Set Max Width to 32 characters, and enable Multi Line Tool Tips
        'SendMessage TiphWnd, TTmsETMAXTIPWIDTH, 0&, &H20
        SendMessage TiphWnd, TTmsETMAXTIPWIDTH, 0, 300

        'If we want a title or we want an icon

        'If toolTitle <> vbNullString Or toolIcon <> tipNoIcon Then
        SendMessage TiphWnd, TTmsETTITLE, CLng(toolIcon), ByVal toolTitle
        'End If

        If toolForeColor <> Empty Then
            '0 (zero) or Null is seen by the API as the default color
            'See ToolTipForeColor property for more datails
            SendMessage TiphWnd, TTmsETTIPTEXTCOLOR, toolForeColor, 0&
        End If

        If mToolBackColor <> Empty Then
            '0 (zero) or Null is seen by the API as the default color
            'See ToolTipBackColor property for more datails
            SendMessage TiphWnd, TTmsETTIPBKCOLOR, mToolBackColor, 0&
        End If

    End If

End Sub

Private Sub UpDateTTP()

    Dim aTemp As Boolean

    'Used to update tooltip parameters that require reconfiguration of
    'subclass to envoke

    'Get current Atcive state
    aTemp = ToolTipEnabled
    'Refresh the object
    Call CreateTTP
    'Restore the ToolTipEnabled state
    ToolTipEnabled = aTemp

End Sub

Public Property Let ToolTipEnabled(ByVal tooldata As Boolean)

'If True, activate (show) tool tip, False deactivate (hide) tool tip
'Syntax: object.active= true/false
    mToolEnabled = tooldata
    SendMessage TiphWnd, TTM_ACTIVATE, CInt(mToolEnabled), ti


End Property

Public Property Get ToolTipEnabled() As Boolean

'Retrieving value of a property, Boolean responce (true/false)
'Syntax: BooleanVar = object.ToolTipEnabled
    ToolTipEnabled = mToolEnabled

End Property

Public Property Set ToolTipParent(ByVal tooldata As Object)

'Assigning an Object to the property, set to parent object(control)
'that the Tool Tip will "pop" up from
'Syntax: Set object.ToolTipParent = ParentObject
    toolParentControl = tooldata.HWnd
    UpDateTTP

End Property


Public Property Get ToolTipParent() As Long

'Retrieving value of a property, returns Long Windows Handle (hWnd)of Parent.
'Syntax: ObjectVar =  object.ToolTipParent
    ToolTipParent = toolParentControl

End Property

Public Property Let ToolTipStyle(ByVal tooldata As toolStyleEnum)

'Assigning a value to the property, set style param Standard or Balloon
'Syntax: object.ToolTipStyle = ToolTipStyle
    toolStyle = tooldata
    UpDateTTP

End Property


Public Property Get ToolTipStyle() As toolStyleEnum

'Retrieving value of a property, returns string.
'Syntax: StringVar = object.Style
    ToolTipStyle = toolStyle

End Property


Public Property Let ToolTipCentered(ByVal tooldata As Boolean)

'Assigning a value to the property, Set Boolean true/false if ToolTip
'is centered on the parent control
'Syntax: object.ToolTipCentered = true/false
    toolCentered = tooldata
    UpDateTTP

End Property


Public Property Get ToolTipCentered() As Boolean

'Retrieving value of a property, returns Boolean true/false.
'Syntax: BooleanVar =  object.Centered
    ToolTipCentered = toolCentered
    UpDateTTP

End Property

Public Property Let ToolTipIcon(ByVal tooldata As toolIconType)

'Assigning a value to the property, set icon style with type var.
'Syntax: object.ToolTipIcon = iconStyle
'Icon Styles are: INFO, WARNING and ERROR (tipNoIcom, tipIconInfo, tipIconWarning, tipIconError)
    toolIcon = tooldata

    'If tipHwnd <> 0 And toolTitle <> Empty And toolIcon <> tipNoIcon Then

    If TiphWnd <> 0 Then
        SendMessage TiphWnd, TTmsETTITLE, CLng(toolIcon), ByVal toolTitle
    End If

    UpDateTTP

End Property


Public Property Get ToolTipIcon() As toolIconType

'Retrieving value of a property, returns string.
'Syntax: StringVar =  object.Icon
    ToolTipIcon = toolIcon

End Property


Public Property Let ToolTipForeColor(ByVal tooldata As Long)

'Assigning a value to the property, set RGB value as Long.
'Syntax: object.ToolTipForeColor = RGB (as Long)
'Since 0 is Black (no RGB), and the API thinks 0 is
'the default color ("off" yellow),
'we need to "fudge" Black a bit (yes set bit "1" to "1",)
'I couldn't resist the pun!
'So, in module or form code, if setting to Black, make it "1"
'if restoring the default color, make it "0"
'Syntax: object.ToolTipForeColor = RGB(as long)

    toolForeColor = tooldata

    If TiphWnd <> 0 Then
        SendMessage TiphWnd, TTmsETTIPTEXTCOLOR, toolForeColor, 0&
    End If

    UpDateTTP

End Property


Public Property Get ToolTipForeColor() As Long

'Retrieving value of a property, returns RGB value as Long.
'Syntax: LongVar =  object.ForeColor
    ToolTipForeColor = toolForeColor

End Property



Public Property Let ToolTipTitle(ByVal tooldata As String)

'Assigning a value to the property, set as string.
'Syntax: object.ToolTipTitle = StringVar
    toolTitle = tooldata

    'If tipHwnd <> 0 And toolTitle <> Empty And toolIcon <> tipNoIcon Then

    If TiphWnd <> 0 Then
        SendMessage TiphWnd, TTmsETTITLE, CLng(toolIcon), ByVal toolTitle
    End If

    UpDateTTP

End Property


Public Property Get ToolTipTitle() As String

'Retrieving value of a property, returns string.
'Syntax: StringVar =  object.ToolTipTitle
    ToolTipTitle = toolTitle

End Property


Public Property Let ToolTipBackColor(ByVal tooldata As Long)

'Assigning a value to the property, set RGB value as Long.
'Syntax: object.ToolTipBackColor = RGB (as Long)
'Since 0 is Black (no RGB), and the API thinks 0 is
'the default color ("off" yellow),
'we need to "fudge" Black a bit (yes set bit "1" to "1",)
'I couldn't resist the pun!
'So, in module or form code, if setting to Black, make it "1"
'if restoring the default color, make it "0"

    mToolBackColor = tooldata

    If TiphWnd <> 0 Then
        SendMessage TiphWnd, TTmsETTIPBKCOLOR, mToolBackColor, 0&
    End If

    UpDateTTP

End Property


Public Property Get ToolTipBackColor() As Long

'Retrieving value of a property, returns RGB as Long.
'Syntax: LongVar =  object.ToolTipBackColor

    ToolTipBackColor = mToolBackColor

End Property



Public Property Let ToolTipText(ByVal tooldata As String)

'Assigning a value to the property, Set as String.
'Syntax: object.ToolTipText = StringVar
'Multi line Tips are enabled in the Create sub.
'To change lines, just add a vbCrLF between text
'ex. object.TipText= "Line 1 text" & vbCrLF & "Line 2 text"

    mToolTipText = tooldata
    ti.lpszText = mToolTipText

    If TiphWnd <> 0 Then
        SendMessage TiphWnd, TTM_UPDATETIPTEXTA, 0&, ti
    End If

    UpDateTTP

End Property


Public Property Get ToolTipText() As String

'Retrieving value of a property, returns string.
'Syntax: StringVar =  object.ToolTipText
    ToolTipText = mToolTipText

End Property

Public Property Get ToolTipGetHwnd() As Long

'Retrive Windows Handle of the Tool Tip
'Syntax: LongVar = object.ToolTipGetHwnd
    ToolTipGetHwnd = toolhWnd

End Property

Public Function ToolTipRemove() As Boolean

'Kills Tool Tip Object
    Tool_Tip_Terminate

End Function

Private Sub Tool_Tip_Terminate()

    If TiphWnd <> 0 Then
        DestroyWindow TiphWnd
    End If

End Sub

Private Sub Class_Terminate()
    Tool_Tip_Terminate
End Sub




'----------------------------------------------------------------------------
'
' Partie Menu Popup
'
'----------------------------------------------------------------------------

'************************************************************
'
' Create Menu Popup using API
'
'    KPD-Team 1998
'
'    http://allapi.mentalis.org
'
'************************************************************


Public Sub MenuPopupAdd(ByVal Caption As String, ByVal Style As MENUPOPUpsTYLE)

    colMnuCaptions.Add Caption
    colMnuStyles.Add Style

End Sub

Public Sub MenuPopup(ByRef ret As Long)

    Dim Pt As POINTAPI
    Dim hMnu As Long
    Dim k As Long

    hMnu = CreatePopupMenu

    For k = 1 To colMnuCaptions.Count
        AppendMenu hMnu, CLng(colMnuStyles(k)), k, CStr(colMnuCaptions(k))
    Next k

    ' récupère la position de la souris
    GetCursorPos Pt

    ' affiche le menu popup à la position de la souris
    ret = TrackPopupMenuEx(hMnu, TPM_LEFTALIGN Or TPM_RIGHTBUTTON Or TPM_RETURNCMD, Pt.x, Pt.y, m_hwndFrm, ByVal 0&)

    DestroyMenu hMnu

End Sub
