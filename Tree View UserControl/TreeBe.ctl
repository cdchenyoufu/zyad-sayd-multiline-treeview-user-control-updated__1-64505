VERSION 5.00
Begin VB.UserControl TreeBe 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   10.5
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   102
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   228
   Begin VB.VScrollBar VS1 
      CausesValidation=   0   'False
      Height          =   1452
      LargeChange     =   5
      Left            =   2160
      Max             =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   252
   End
   Begin VB.PictureBox Pic1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   972
      Left            =   240
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   0
      Top             =   120
      Width           =   1692
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   372
         Left            =   360
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   1
         Top             =   240
         Width           =   972
      End
   End
End
Attribute VB_Name = "TreeBe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'.:______________________________________:.
 '.:MultiLine TreeView User Control V1.1:.
'.:¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯:.

'.:Author:     Zyad Sayd
'.:Created:    01 March 2006
'.:Copyright:  2006-Zyad Sayd {saidseyam@hotmail.com}

'.:Here is a list of Credits, many thanks for them:-

'.:Paul Caton: Self-Subclassing
'.:Richard Mewett: Unicode
'.:Vlad Vissoultchev
'.:Gary Noble

'You are free to use this control as you like as long as you keep the
'Credits and Copyright in your About dialog.
'__________________________________________________________________________
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Updates:-
'        -04 March 2006 -Bug Fixed in line drawing.
'                       -Added tow line styles:-
'                          Normal = 1
'                          Group = 2
'                       -Double Click supported.
'        -06 March 2006 -Added four visual styles:-
'                          Normal Color
'                          Vertical Gradient
'                          Horizontal Gradient
'                          Mac Style
'                       -Version 1.1 released
'                       -Added Property: HoverBackColor
'                       -Added Property: LineStyle
'                       -Added Property: TreeStyle
'_________________________________________________________________________
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Private Const VER_PLATFORM_WIN32_NT = 2

'<<---General APIs --->>'
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Const COLOR_BTNFACE As Long = 15

Private Const WM_MOUSEWHEEL          As Long = &H20A

'<<---APIs for Draw Lines--->>'
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Long) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Private Declare Function ApiDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Private Declare Function ApiBitBlt Lib "gdi32" Alias "BitBlt" (ByVal hdcDest As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function ColorHLSToRGB Lib "shlwapi.dll" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long
Private Declare Sub ColorRGBToHLS Lib "shlwapi.dll" (ByVal clrRGB As Long, ByRef pwHue As Integer, ByRef pwLuminance As Integer, ByRef pwSaturation As Integer)


Private Const DSna As Long = &H220326

'<<---APIs for Drawing Texts--->>'
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DT_CALCRECT As Long = &H400      'Determines the height of the rectangle
Private Const DT_LEFT As Long = &H0            'Aligns text to the left.
Private Const DT_RIGHT As Long = &H2           'Aligns text to the right.
Private Const DT_WORDBREAK As Long = &H10       'Breaks words
Private Const DT_RTLREADING As Long = &H20000
Private HoverButton As Boolean

Private WithEvents M_Font As StdFont
Attribute M_Font.VB_VarHelpID = -1

Public Enum LineStyleEnum
 Normal = 1
 Group = 2
End Enum

Public Enum TreeStyleEnum
 NormalColor = 1
 VGradient = 2
 HGradient = 3
 MacStyle = 4
End Enum

Private m_PIcon As StdPicture
Private m_MIcon As StdPicture

Private m_Right As Boolean
Private m_BackColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_SelectedBackColor As OLE_COLOR
Private m_SelectedForeColor As OLE_COLOR
Private m_HoverBackColor As OLE_COLOR

Private m_Checkintegrity As Boolean

Private mRect As RECT
Private mWidth As Long
Private mNodes As ClsNodes
Private mNode As ClsNode
Private mOldSel As Long
Private mShowTreeLines As Boolean
Private mLineStyle As LineStyleEnum
Private mTreeStyle As TreeStyleEnum

Private Const Offset = 15
Private CurrBottom As Long
Private AllLines As Single
Private ExpandedLines As Single
Private Scrolled As Boolean
Private StartNode As Long
Private Xnow As Single, Ynow As Single, BtNow As Integer
Private ItemCounter As Long
Private OldeSelBClr As Long
Private OldeSelFClr As Long
Private mWindowsNT As Boolean
Private LastValue As Long

'<<--Events-->>
Public Event SelectedNode(ByVal Node As PrjTreeBe.ClsNode)
Public Event ButtonClick(ByVal Node As PrjTreeBe.ClsNode)
Public Event Collapse(ByVal Node As PrjTreeBe.ClsNode)
Public Event Expand(ByVal Node As PrjTreeBe.ClsNode)
Public Event Click(ByVal Button As Integer)


'==================================================================================================
'<<---Self-subclassing--->>'
Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum
Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Type tSubData                                                                   'Subclass data type
  hWnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type
Private sc_aSubData()                As tSubData                                        'Subclass data array

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean
Private bInCtrl                      As Boolean
'==================================================================================================

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
'Parameters:
  'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
  'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  'hWnd     - The window handle
  'uMsg     - The message number
  'wParam   - Message related data
  'lParam   - Message related data
'Notes:
  'If you really know what you're doing, it's possible to change the values of the
  'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
  'values get passed to the default handler.. and optionaly, the 'after' callback
  Static bMoving As Boolean
  
  Select Case uMsg
    Case WM_MOUSEWHEEL
    If wParam > 0 Then
      If VS1.Value > 0 Then VS1.Value = VS1.Value - 1
    Else
      If VS1.Value < VS1.Max Then VS1.Value = VS1.Value + 1
     End If
  End Select
End Sub



Private Sub M_Font_FontChanged(ByVal PropertyName As String)
Set Pic.Font = M_Font
If Not mNodes Is Nothing Then Call DrawTree

End Sub



Public Property Get Font() As StdFont
Attribute Font.VB_UserMemId = -512
Set Font = M_Font
End Property

Public Property Set Font(ByVal vNewFont As StdFont)
With M_Font
      .Bold = vNewFont.Bold
      .Italic = vNewFont.Italic
      .Name = vNewFont.Name
      .Size = vNewFont.Size
      .Strikethrough = vNewFont.Strikethrough
      .Underline = vNewFont.Underline
      .Charset = vNewFont.Charset
   End With
   PropertyChanged "Font"
If Not mNodes Is Nothing Then Call DrawTree
End Property

Private Sub Pic_DblClick()
'Pic_MouseDown 1, 0, Xnow, Ynow
Dim rc As RECT, ItemNum As Long, T As Boolean
'Static OldNode As New ClsNode
'Set OldNode.mParent = Me

If mNodes Is Nothing Then Exit Sub

If BtNow = vbRightButton Then Exit Sub

For Each mNode In mNodes
  ItemNum = ItemNum + 1
'==========================Expand a node====================================
 If mNode.Visable = True Then
  rc.Left = IIf(m_Right, 0, mNode.rcLeft)
  rc.Right = IIf(m_Right, NodeRight, mRect.Right)
  rc.Top = mNode.rctop
  rc.Bottom = mNode.rcBottom
  If PtInRect(rc, Xnow, Ynow) Then
    
    If mNode.Expanded = True Then
      RaiseEvent Collapse(mNode)
    Else: RaiseEvent Expand(mNode)
    End If
    
    If VS1.Value = VS1.Max Then T = True
    ExpandItem mNode
  End If
End If
Next
End Sub

Private Sub Pic_GotFocus()
SelectedForeColor = OldeSelFClr
SelectedBackColor = OldeSelBClr
End Sub

Private Sub Pic_LostFocus()
OldeSelFClr = SelectedForeColor
OldeSelBClr = SelectedBackColor
SelectedBackColor = GetSysColor(COLOR_BTNFACE)
SelectedForeColor = vbBlack
Refresh

End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim rc As RECT, ItemNum As Long, T As Boolean
Static OldNode As New ClsNode
Set OldNode.mParent = Me


On Error Resume Next
BtNow = Button

If mNodes Is Nothing Then Exit Sub

If Button = vbRightButton Then Exit Sub

For Each mNode In mNodes
  ItemNum = ItemNum + 1
'==========================Expand a node====================================
 If mNode.Visable = True Then
  rc.Left = IIf(m_Right, NodeRight + 36 - (m_PIcon.Width / 21), mNode.rcLeft - 36)
  rc.Right = IIf(m_Right, rc.Left + (m_PIcon.Width / 21), rc.Left + (m_PIcon.Width / 21) + 2)
  rc.Top = mNode.rctop + (mNode.rcBottom - mNode.rctop) / 2 - m_PIcon.Width / 21
  rc.Bottom = rc.Top + (m_PIcon.Width / 21) + 5
  If PtInRect(rc, X, Y) Then
    
    If mNode.Expanded = True Then
      RaiseEvent Collapse(mNode)
    Else: RaiseEvent Expand(mNode)
    End If
    
    If VS1.Value = VS1.Max Then T = True
    ExpandItem mNode: GoTo Ext
  End If
'==========================Select a node====================================
  rc.Left = IIf(m_Right, 0, mNode.rcLeft)
  rc.Right = IIf(m_Right, NodeRight, mRect.Right)
  rc.Top = mNode.rctop
  rc.Bottom = mNode.rcBottom

    If PtInRect(rc, X, Y) Then
      If HoverButton Then
        RaiseEvent ButtonClick(mNode)
        DrawOneNode mNode: GoTo Ext
      Else
        RaiseEvent SelectedNode(mNode)
        Set OldNode = mNode: SelectItem mNode: GoTo Ext
      End If
    End If
    
 End If
Next

'If T = True And VS1.Value <> 0 Then
' VS1.Value = VS1.Max
'End If
If Button = vbLeftButton Then
  OldNode.Selected = False
  Refresh
End If
Ext:

End Sub


'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim i                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
           "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
           "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
           "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    i = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      i = i + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hWnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
End Function

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub
'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hWnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hWnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
  If Not bAdd Then
    Debug.Assert False                                                                  'hWnd not found, programmer error
  End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function


'Stop all subclassing
Private Sub Subclass_StopAll()
  Dim i As Long
  
  i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While i >= 0                                                                       'Iterate through each element
    With sc_aSubData(i)
      If .hWnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hWnd)                                                       'Subclass_Stop
      End If
    End With
    
    i = i - 1                                                                           'Next element
  Loop
End Sub
'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hWnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub
'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static OldNode As New ClsNode
Static Oldrc As RECT, OldNodeNum As Long
Set OldNode.mParent = Me
Dim rc As RECT, item As Long, Tick As Long
On Error Resume Next

If mNodes Is Nothing Then Exit Sub
Xnow = X
Ynow = Y
'#######################################################################
'check if the mouse still HoverButton on the same node. if true then do not
'execute rest of the code.
'OldNode:stored the old node properties
'#######################################################################
rc.Left = 0: rc.Right = IIf(m_Right, mRect.Right, mRect.Right)
rc.Top = OldNode.rctop: rc.Bottom = OldNode.rcBottom
If PtInRect(rc, X, Y) Then
 If Not OldNode.ButtonIcon Is Nothing Then
  If X >= IIf(m_Right, mRect.Left, rc.Right - 20) And X < IIf(m_Right, mRect.Left + 20, rc.Right) _
  And Y < rc.Top + 20 And Button = 0 Then
    HoverButton = True
    Pic.ToolTipText = OldNode.ButtonToolTip
    DrawRectangle Pic.hdc, IIf(m_Right, mRect.Left + 4, rc.Right - 20), rc.Top + 1, IIf(m_Right, mRect.Left + 20, mRect.Right - 4), rc.Top + 16, RGB(58, 122, 241), RGB(150, 189, 248)
    DrawButton OldNode, True
    Pic.Refresh
  ElseIf HoverButton = True Then
    HoverButton = False
    Pic.ToolTipText = ""
    DrawRectangle Pic.hdc, IIf(m_Right, mRect.Left + 4, rc.Right - 20), rc.Top + 1, IIf(m_Right, mRect.Left + 20, mRect.Right - 4), rc.Top + 16, IIf(OldNode.Selected, SelectedBackColor, OldNode.BackColor), IIf(OldNode.Selected, SelectedBackColor, OldNode.BackColor)
    DrawButton OldNode
    Pic.Refresh
  End If
 End If
Exit Sub
End If
'Tick = GetTickCount
Pic.ToolTipText = ""
HoverButton = False
'#######################################################################
'If the mouse moved from the old node recheck the new node.
'OldNode:store the old node properties
'item: not the nod number but the lines of the visable nodes.
'Note:there is a lot of [if..then] but not [if..And..then], this is intended
'for the speed improvement.
'#######################################################################
For Each mNode In mNodes
  If mNode.Visable = True Then
    item = item + mNode.Lines
      If item >= VS1.Value Then
        If item <= VS1.Value + MaxViewItems + mNode.Lines Then
          rc.Left = 0: rc.Right = IIf(m_Right, mRect.Right, mRect.Right)
          rc.Top = mNode.rctop: rc.Bottom = mNode.rcBottom
            If PtInRect(rc, X, Y) Then
              If OldSel <> item - 1 Then
                If OldSel <> 0 Then
                  mNodes(OldNodeNum, 1).HoverItem = False
                  If OldSel >= VS1.Value + 1 And OldSel < VS1.Value + MaxViewItems + mNodes(OldNodeNum, 1).Lines And OldNode.Visable = True Then DrawOneNode OldNode
                  End If
                  Set OldNode = mNode
                  
                  OldSel = item - 1: OldNodeNum = mNode.ItemNum
                  mNode.HoverItem = True
                  'Debug.Print mNode.HoverItem
                  DrawOneNode mNode
                  Exit For
                End If
              Exit For
            End If
        End If
      End If
  End If
Next
'MsgBox GetTickCount - Tick
End Sub

Private Sub Pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Pic_MouseMove 0, 0, Xnow, Ynow
RaiseEvent Click(Button)
BtNow = Button

End Sub


Private Sub UserControl_Initialize()
Set mNode = New ClsNode

Dim OS As OSVERSIONINFO
      
OS.dwOSVersionInfoSize = Len(OS)
Call GetVersionEx(OS)
    
mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)

End Sub

Private Sub UserControl_InitProperties()
BackColor = vbWhite
ForeColor = vbBlack
SelectedBackColor = vbBlue
SelectedForeColor = vbWhite
HoverBackColor = vbRed
CheckIntegrity = True
ShowTreeLines = True
LineStyle = Normal
TreeStyle = NormalColor

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Set m_PIcon = PropBag.ReadProperty("PlusIcon", Nothing)
Set m_MIcon = PropBag.ReadProperty("MinIcon", Nothing)
RightToLeft = PropBag.ReadProperty("RightToLeft", False)
BackColor = PropBag.ReadProperty("BackColor", vbWhite)
ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
SelectedBackColor = PropBag.ReadProperty("SelectedBackColor", vbBlue)
CheckIntegrity = PropBag.ReadProperty("CheckIntegrity", True)
ShowTreeLines = PropBag.ReadProperty("ShowTreeLines", True)
SelectedForeColor = PropBag.ReadProperty("SelectedForeColor", vbWhite)
LineStyle = PropBag.ReadProperty("LineStyle", 1)
TreeStyle = PropBag.ReadProperty("TreeStyle", 1)
HoverBackColor = PropBag.ReadProperty("HoverBackColor", vbRed)

Set M_Font = PropBag.ReadProperty("Font", Ambient.Font)
Set Pic.Font = M_Font
OldeSelBClr = SelectedBackColor
OldeSelFClr = SelectedForeColor
'========================================================================
With UserControl
'Start subclassing the UserControl
    Call Subclass_Start(.hWnd)
    Call Subclass_AddMsg(.hWnd, WM_MOUSEWHEEL, MSG_AFTER)
End With

End Sub

Private Sub UserControl_Resize()

With Pic1
 .Width = ScaleWidth - 2
 .Height = ScaleHeight - 2
 .Top = 1
 .Left = 1
End With

mRect.Right = Pic1.ScaleWidth
mRect.Bottom = Pic1.ScaleHeight
mWidth = ScaleWidth


With Pic
 .Width = ScaleWidth
 .Height = mRect.Bottom + (LineHeight) * 2
 .Left = 0
 .Top = -LineHeight
End With

VS1.Left = IIf(m_Right, 0, Pic1.ScaleWidth - VS1.Width + 1)
VS1.Height = ScaleHeight


If Not mNodes Is Nothing Then Call DrawTree

End Sub

Public Property Get Nodes() As ClsNodes
If mNodes Is Nothing Then Set mNodes = New ClsNodes: Set mNodes.ParentControl = Me
Set Nodes = mNodes
End Property

Friend Sub DrawTree(Optional Escape As Boolean = False, Optional StartV)
Dim rc As RECT, LineNum As Long, i As Long
Dim Tex As String
'#######################################################################
'Note:when we draw the tree, the VScroll determine the start node,
'MaxViewItems determine the end node to draw.
'The VScroll is a pointer to the number of the lines, so when we Loop throw all nodes
'to determine the target nodes to draw, there is a nother [For...Next] to
'loop throw each node lines.
'#######################################################################
'On Error Resume Next

'<<--begane from the top-->>
CurrBottom = 0
'<<-Check if the Scroll bars nedded befor draw the tree->>
If Escape = False Then
  Call CountAllLines
  Call CheckScroll
End If

If IsMissing(StartV) Then StartV = VS1.Value

CurrBottom = 0
LineNum = 0
Pic.Cls
ItemCounter = 0
'<<--calculate and draw all nodes-->>
For Each mNode In mNodes '<<1)Loop throw all nodes
  If mNode.Visable = True Then
   ItemCounter = ItemCounter + 1
    For i = 1 To mNode.Lines '<<2)Loop throw each node lines to found the target line
      LineNum = LineNum + 1
        If LineNum >= StartV + 1 Then
          If LineNum <= StartV + MaxViewItems + 2 Then
            If LineNum = StartV + 1 Then StartNode = mNode.ItemNum
            If i > 1 Then CurrBottom = CurrBottom - (i - 1) * LineHeight
            If i > 1 And i = mNode.Lines Then StartNode = mNode.ItemNum + 1
            i = mNode.Lines
            '<--Support Multiline, Draw Text-->>
            DrawTexts mNode.Caption, rc, mNode.Relative
            '<-- Draw the Tree Guide line-->>
            If mNode.Relative <> "" And ShowTreeLines And LineStyle = Group Then Call DrawTreeGuide
            '<-- if this node contain a children, then draw the Expanded/collapse icons-->>
            If mNode.Children > 0 Then
              DrawLine IIf(m_Right, NodeRight + 15, mNode.rcLeft - 15), mNode.rctop - 1 + CInt((mNode.rcBottom - mNode.rctop) / 2), IIf(m_Right, NodeRight + 30, mNode.rcLeft - 30), mNode.rctop - 1 + CInt((mNode.rcBottom - mNode.rctop) / 2), RGB(192, 192, 192)
                If mNode.Expanded = False Then
                  Pic.PaintPicture m_PIcon, IIf(m_Right, NodeRight - (m_MIcon.Width / 21) + 36, mNode.rcLeft - 36), mNode.rctop + (mNode.rcBottom - mNode.rctop) / 2 - ((m_MIcon.Height / 21) / 2)
                  Else
                  Pic.PaintPicture m_MIcon, IIf(m_Right, NodeRight - (m_MIcon.Width / 21) + 36, mNode.rcLeft - 36), mNode.rctop + (mNode.rcBottom - mNode.rctop) / 2 - ((m_MIcon.Height / 21) / 2)
                End If
            End If
            '<-- Draw the Tree Icon-->>
            If Not mNode.Icon Is Nothing Then
              Call DrawTreeIcons
            End If
            
          'ElseIf LineNum > StartV + MaxViewItems + 2 Then Exit For
          End If
        ElseIf LineNum < StartV + 1 Then mNode.rctop = -mNode.rcBottom: mNode.rcBottom = 1
        End If
    Next i '2)>>
     If LineNum > StartV + MaxViewItems + 2 Then
     mNode.rctop = LineHeight * (MaxViewItems + 2)
   End If
     If mNode.Relative <> "" And ShowTreeLines And LineStyle = Normal Then Call DrawNormalTreeGuide
  End If
Next mNode '1)>>

'<<--checkScroll again after draw-->>
If Escape = False Then
  Call CountAllLines
  Call CheckScroll
End If
'Pic.Refresh

DrawRectangle hdc, 0, 0, ScaleWidth, ScaleHeight, RGB(170, 203, 253)
End Sub

Friend Sub ExpandItem(wNode As ClsNode, Optional Expand)
Dim F As Integer, Counter As Integer
Dim ChNum As Long, KeyA() As Variant
ReDim KeyA(0)
Set mNode = wNode
On Error Resume Next

ChNum = mNode.Children
KeyA(0) = mNode.Key

If ChNum > 0 Then '<<--1)Just if the node is a parent
 If IsMissing(Expand) Then
   mNode.Expanded = Not mNode.Expanded
 Else
   If mNode.Expanded = Expand Then Exit Sub
   mNode.Expanded = Expand
 End If
 
 For Each mNode In mNodes
   For F = 0 To UBound(KeyA) '<<--2)Loop through parent nodes keys
     If mNode.Relative = KeyA(F) And KeyA(F) <> "" Then '<<--3)if a child found
       mNode.Visable = Not mNode.Visable
       'If mNode.Visable = True Then ExpandedLines = ExpandedLines + mNode.Lines
       
         If mNode.Children > 0 And mNode.Expanded = True Then '<<--4)if this child have Children
           ReDim Preserve KeyA(UBound(KeyA) + 1): KeyA(UBound(KeyA)) = mNode.Key
         End If '4)-->>
     End If '3)-->>
   Next F '2)-->>

  Counter = Counter + 1
 Next
 Call DrawTree
 Call CountAllLines
 Call CheckScroll
End If '1)-->>
End Sub

Private Sub SelectItem(wNode As ClsNode)
Dim F As Integer, Counter As Integer
Dim ChNum As Long, ItemNum As Long
Set mNode = wNode

ItemNum = mNode.ItemNum
mNode.Selected = True

 For Each mNode In mNodes
   If Counter + 1 <> ItemNum Then mNode.Selected = False

  Counter = Counter + 1
 Next
 Call DrawTree(True)
End Sub


Private Function DrawTexts(txt As String, ByRef rc As RECT, Relate As Variant, Optional NodeTop As Long)
Dim rctxt As RECT, Lng As Long

'If CurrBottom <> 0 Then CurrBottom = CurrBottom+5
rc.Right = IIf(m_Right, NodeRight, mRect.Right)
If Not mNode.ButtonIcon Is Nothing Then rc.Right = rc.Right - 20
rc.Top = IIf(NodeTop = 0, CurrBottom, NodeTop)
rc.Left = IIf(m_Right, 2, mNode.rcLeft)
If Not mNode.ButtonIcon Is Nothing And m_Right Then rc.Left = rc.Left + 20
rc.Bottom = mRect.Bottom
Pic.ForeColor = ForeColor

 'If NodeTop <> 0 Then Call DrawRectangle(Pic.hdc, rc.Left, rc.Top, rc.Right, mNode.rcBottom, vbBlue, vbBlue)
 '<<--Calculate the text height-->>
  If mWindowsNT Then
    DrawTextW Pic.hdc, StrPtr(txt), Len(txt), rc, DT_CALCRECT Or DT_WORDBREAK
  Else
    DrawText Pic.hdc, txt, Len(txt), rc, DT_CALCRECT Or DT_WORDBREAK
  End If
' If rc.Bottom - rc.Top < 20 Then rc.Bottom = rc.Top + 20
 If Not mNode.ButtonIcon Is Nothing And m_Right Then rc.Left = rc.Left - 20

 '<<--Draw Rectangle first-->>
 If mNode.BackColor <> -1 And TreeStyle = NormalColor Then
   DrawRectangle Pic.hdc, rc.Left - 2, rc.Top, IIf(m_Right, NodeRight, mRect.Right), rc.Bottom, Pic.BackColor, mNode.BackColor 'RGB(170, 203, 253)
 ElseIf mNode.BackColor <> -1 And TreeStyle <> NormalColor And mNode.Relative = "" Then
   DrawGradient rc, mNode.BackColor, TreeStyle, mNode
 ElseIf mNode.BackColor <> -1 And TreeStyle <> NormalColor And mNode.Relative <> "" Then
   DrawRectangle Pic.hdc, rc.Left - 2, rc.Top, IIf(m_Right, NodeRight, mRect.Right), rc.Bottom, Pic.BackColor, mNode.BackColor 'RGB(170, 203, 253)
 End If
 '<<--Draw HoverBackColor Rectangle-->>
 If mNode.HoverItem = True Then
   DrawRectangle Pic.hdc, rc.Left - 2, rc.Top, IIf(m_Right, NodeRight, mRect.Right), rc.Bottom, HoverBackColor, HoverBackColor 'RGB(102, 152, 244)
 End If
 
 If mNode.Selected = True Then
   Pic.ForeColor = SelectedForeColor
   DrawRectangle Pic.hdc, rc.Left - 2, rc.Top, IIf(m_Right, NodeRight, mRect.Right), rc.Bottom, SelectedBackColor, SelectedBackColor
 Else
   If mNode.ForeColor <> -1 Then Pic.ForeColor = mNode.ForeColor
 End If
 
 rc.Right = IIf(m_Right, NodeRight, mRect.Right)
 If Not mNode.ButtonIcon Is Nothing Then rc.Right = rc.Right - 20
 mNode.rctop = IIf(NodeTop = 0, CurrBottom, mNode.rctop)
 mNode.rcBottom = rc.Bottom
 mNode.Lines = (rc.Bottom - rc.Top) / LineHeight
 CurrBottom = CurrBottom + (rc.Bottom - rc.Top)
 Lng = IIf(m_Right, DT_RTLREADING Or DT_RIGHT Or DT_WORDBREAK, DT_LEFT Or DT_WORDBREAK)
 If Not mNode.ButtonIcon Is Nothing And m_Right Then rc.Left = rc.Left + 20
  '<<--Draw Text-->>
  If mWindowsNT Then
    DrawTextW Pic.hdc, StrPtr(txt), Len(txt), rc, Lng
  Else
    DrawText Pic.hdc, txt, Len(txt), rc, Lng
  End If

 Pic.ForeColor = ForeColor
End Function

Private Function DrawLine(x1 As Long, y1 As Long, x2 As Long, y2 As Long, LineColor As Long, Optional LWidth = 0)
Dim Old As Long
Dim RPen As Long

Call MoveToEx(Pic.hdc, x1, y1, 0)

RPen = CreatePen(0, LWidth, LineColor)
Old = SelectObject(Pic.hdc, RPen)

Call LineTo(Pic.hdc, x2, y2)
 
Old = SelectObject(Pic.hdc, Old)
DeleteObject RPen

End Function




Private Sub DrawTreeGuide()
Dim L As Long, T As Long, B As Long, Clr As Long
Dim PL As Long, PT As Long, PB As Long
Dim i As Integer

i = IIf(m_Right, -1, 1)

L = IIf(m_Right, NodeRight + 36, mNode.rcLeft - 36)
T = mNode.rctop + 10
B = mNode.rcBottom - 10

PT = mNodes(mNode.Relative, 1).rctop
PL = IIf(m_Right, mRect.Right - mNodes(mNode.Relative, 1).rcRight + 32, mNodes(mNode.Relative, 1).rcLeft - 32)
PB = mNodes(mNode.Relative, 1).rcBottom
Clr = RGB(192, 192, 192)

If mNodes(mNode.Relative, 1).Children > 1 Then T = PB + 10

If mNodes(mNode.Relative, 1).Children > 0 Then
  If ItemCounter = MaxViewItems + VS1.Value + 1 Or mNode.ChNum = mNodes(mNode.Relative, 1).Children Or mNode.ChNum = mNodes(mNode.Relative, 1).AllChildren Then

  DrawLine L + (2 * i), T, L - (5 * i), T, Clr
  DrawLine L + (2 * i), B, L - (5 * i), B, Clr
  DrawLine L + (-5 * i), T + 1, L + (-5 * i), B, Clr

  DrawLine L - (5 * i), T + CInt((B - T) / 2), PL, T + CInt((B - T) / 2), Clr
  
  DrawLine PL, PT + 4 + CInt((PB - PT) / 2), PL, T + CInt((B - T) / 2), Clr
  End If
End If

End Sub

Private Sub DrawNormalTreeGuide()
Dim L As Long, T As Long, B As Long, Clr As Long
Dim PL As Long, PT As Long, PB As Long
Dim i As Integer

i = IIf(m_Right, -1, 1)

L = IIf(m_Right, NodeRight + 36, mNode.rcLeft - 36)
T = mNode.rctop + 10
B = mNode.rcBottom - 10

PT = mNodes(mNode.Relative, 1).rctop
PL = IIf(m_Right, mRect.Right - mNodes(mNode.Relative, 1).rcRight + 32, mNodes(mNode.Relative, 1).rcLeft - 32)
PB = mNodes(mNode.Relative, 1).rcBottom
Clr = RGB(192, 192, 192)


If ItemCounter > (VS1.Value + MaxViewItems + 2) And mNodes(mNode.Relative, 1).rctop < LineHeight * MaxViewItems Then
  DrawLine L - (15 * i), PB, L - (15 * i), LineHeight * (MaxViewItems + 2), Clr
ElseIf ItemCounter >= VS1.Value + 1 And ItemCounter < (VS1.Value + MaxViewItems + 4) Then
  DrawLine L, T, L - (15 * i), T, Clr
  DrawLine L - (15 * i), T, L - (15 * i), PB - 5, Clr
End If

End Sub

Public Property Get PlusIcon() As StdPicture
Set PlusIcon = m_PIcon
End Property

Public Property Set PlusIcon(ByVal vNewValue As StdPicture)
If CInt(vNewValue.Height / 21) <> 9 Or CInt(vNewValue.Width / 21) <> 9 Then MsgBox "The icon must be 9X9", vbCritical: Exit Property

Set m_PIcon = vNewValue
 PropertyChanged "PlusIcon"

End Property

Public Property Get MinIcon() As StdPicture
Set MinIcon = m_MIcon
End Property

Public Property Set MinIcon(ByVal vNewValue As StdPicture)
If CInt(vNewValue.Height / 21) <> 9 Or CInt(vNewValue.Width / 21) <> 9 Then MsgBox "The icon must be 9X9", vbCritical: Exit Property
Set m_MIcon = vNewValue
 PropertyChanged "MinIcon"

End Property

Private Sub UserControl_Show()

If Ambient.UserMode = True And Not mNodes Is Nothing Then Call DrawTree
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "PlusIcon", m_PIcon, Nothing
    .WriteProperty "MinIcon", m_MIcon, Nothing
    .WriteProperty "RightToLeft", m_Right, False
    .WriteProperty "Backcolor", m_BackColor, vbWhite
    .WriteProperty "Font", M_Font, Ambient.Font
    .WriteProperty "ForeColor", m_ForeColor, vbBlack
    .WriteProperty "SelectedBackColor", m_SelectedBackColor, vbBlue
    .WriteProperty "CheckIntegrity", m_Checkintegrity, True
    .WriteProperty "ShowTreeLines", mShowTreeLines, True
    .WriteProperty "SelectedForeColor", m_SelectedForeColor, vbWhite
    .WriteProperty "LineStyle", mLineStyle, 1
    .WriteProperty "TreeStyle", mTreeStyle, 1
    .WriteProperty "HoverBackColor", m_HoverBackColor, vbRed

End With

End Sub

Private Sub DrawTreeIcons()
 
Pic.PaintPicture mNode.Icon, IIf(m_Right, NodeRight - 16 + 20, mNode.rcLeft - 20), mNode.rctop + 3, 16, 16

If Not mNode.ButtonIcon Is Nothing Then
 DrawButton
End If

End Sub

Private Sub DrawRectangle(mhdc As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, Clr As Long, Optional FillClr)
Dim BrushH As Long, Old As Long, RPen As Long, OldP As Long

 RPen = CreatePen(0, 1, Clr)
 OldP = SelectObject(mhdc, RPen)

If Not IsMissing(FillClr) Then
 BrushH = CreateSolidBrush(FillClr)
 Old = SelectObject(mhdc, BrushH)
 Rectangle mhdc, x1, y1, x2, y2
 Old = SelectObject(mhdc, Old)
 DeleteObject BrushH
Else
 Rectangle mhdc, x1, y1, x2, y2
End If
 Old = SelectObject(mhdc, OldP)
 DeleteObject RPen

End Sub

Private Sub VS1_Change()

If Scrolled = True Then LastValue = VS1.Value: Scrolled = False: Exit Sub

'<<--Slow Scroll-->>
If LastValue < VS1.Value Then
  Do Until Pic.Top <= -(LineHeight) * 2
    Pic.Top = Pic.Top - (VS1.Value - LastValue)
    Sleep 20
    Pic.Refresh
  Loop
Else
  Do Until Pic.Top >= 0
    Pic.Top = Pic.Top + (LastValue - VS1.Value)
    Sleep 20
    Pic.Refresh
  Loop
End If

'<<--Redraw the items-->>
Pic.Top = -LineHeight
LastValue = VS1.Value
Call DrawTree(True, VS1.Value)
Pic.Top = -LineHeight

End Sub


Private Sub VS1_Scroll()
DoEvents

Call DrawTree(True, VS1.Value)

If LastValue = VS1.Value Then
  Scrolled = False
Else
  Scrolled = True
End If
End Sub


Private Sub CheckScroll()


If AllLines > MaxViewItems Then
  VS1.Max = AllLines - MaxViewItems
  VS1.SmallChange = 1
  VS1.LargeChange = MaxViewItems
  VS1.Visible = True
  mRect.Right = Pic1.ScaleWidth - VS1.Width
  Pic.Width = mRect.Right
  Pic.Left = IIf(m_Right, VS1.Width, 0)
  VS1.Left = IIf(m_Right, 0, Pic1.ScaleWidth - VS1.Width + 1)
  
ElseIf AllLines <= MaxViewItems Then
  VS1.Max = 0: VS1.Visible = False
  mRect.Right = Pic1.ScaleWidth
  Pic.Width = mRect.Right
  Pic.Left = 0

End If

End Sub

Public Property Get RightToLeft() As Boolean
RightToLeft = m_Right
End Property

Public Property Let RightToLeft(ByVal vNewRightToLeft As Boolean)
m_Right = vNewRightToLeft

If Not mNodes Is Nothing Then Call DrawTree
 PropertyChanged "RightToLeft"

End Property

Private Property Get NodeRight() As Long
  NodeRight = mRect.Right - mNode.rcRight
End Property


Private Property Get MaxViewItems() As Long
MaxViewItems = CInt(mRect.Bottom / LineHeight)
End Property

Private Sub CountAllLines()
Dim rc As RECT
 AllLines = 0
'<<--Calculate the number of lines-->>

For Each mNode In mNodes
  If mNode.Visable = True Then
    rc.Right = IIf(m_Right, NodeRight, mRect.Right): rc.Top = CurrBottom
    If Not mNode.ButtonIcon Is Nothing Then rc.Right = rc.Right - 20
    rc.Left = IIf(m_Right, 2, mNode.rcLeft): rc.Bottom = mRect.Bottom
    If Not mNode.ButtonIcon Is Nothing And m_Right Then rc.Left = rc.Left + 20
      If mWindowsNT Then
        DrawTextW Pic.hdc, StrPtr(mNode.Caption), Len(mNode.Caption), rc, DT_CALCRECT Or DT_WORDBREAK
      Else
        DrawText Pic.hdc, mNode.Caption, Len(mNode.Caption), rc, DT_CALCRECT Or DT_WORDBREAK
     End If
     'If rc.Bottom - rc.Top < 20 Then rc.Bottom = rc.Top + 20
     
    mNode.Lines = CInt((rc.Bottom - rc.Top) / (LineHeight))
    AllLines = AllLines + mNode.Lines
    CurrBottom = CurrBottom + (rc.Bottom - rc.Top)
  End If
Next
AllLines = AllLines - 1
End Sub

Public Property Get LineHeight() As Single
LineHeight = Pic.TextHeight("A")
 'If LineHeight < 20 Then LineHeight = 20

End Property



Private Sub DrawOneNode(OldNode As ClsNode)
Dim rctxt As RECT, Lng As Long

rctxt.Right = IIf(m_Right, mRect.Right - OldNode.rcRight, mRect.Right)
If Not OldNode.ButtonIcon Is Nothing Then rctxt.Right = rctxt.Right - 20
rctxt.Top = OldNode.rctop
rctxt.Left = IIf(m_Right, 2, OldNode.rcLeft)
If Not OldNode.ButtonIcon Is Nothing And m_Right Then rctxt.Left = rctxt.Left + 20

rctxt.Bottom = OldNode.rcBottom
Pic.ForeColor = ForeColor

 '<<--Calculate the text height-->>
       If mWindowsNT Then
        DrawTextW Pic.hdc, StrPtr(OldNode.Caption), Len(OldNode.Caption), rctxt, DT_CALCRECT Or DT_WORDBREAK
      Else
        DrawText Pic.hdc, OldNode.Caption, Len(OldNode.Caption), rctxt, DT_CALCRECT Or DT_WORDBREAK
     End If

 'If rctxt.Bottom - rctxt.Top < 20 Then rctxt.Bottom = rctxt.Top + 20
 If Not OldNode.ButtonIcon Is Nothing And m_Right Then rctxt.Left = rctxt.Left - 20

 '<<--Draw main Rectangle-->>
 If OldNode.BackColor <> -1 And TreeStyle = NormalColor Then
   DrawRectangle Pic.hdc, rctxt.Left - 2, rctxt.Top, IIf(m_Right, mRect.Right - OldNode.rcRight, mRect.Right), rctxt.Bottom, Pic.BackColor, OldNode.BackColor 'RGB(170, 203, 253)
 ElseIf OldNode.BackColor <> -1 And TreeStyle <> NormalColor And OldNode.Relative = "" Then
   DrawGradient rctxt, OldNode.BackColor, TreeStyle, OldNode
 ElseIf OldNode.BackColor <> -1 And TreeStyle <> NormalColor And OldNode.Relative <> "" Then
   DrawRectangle Pic.hdc, rctxt.Left - 2, rctxt.Top, IIf(m_Right, mRect.Right - OldNode.rcRight, mRect.Right), rctxt.Bottom, Pic.BackColor, OldNode.BackColor 'RGB(170, 203, 253)
 Else
   DrawRectangle Pic.hdc, rctxt.Left - 2, rctxt.Top, IIf(m_Right, mRect.Right - OldNode.rcRight, mRect.Right), rctxt.Bottom, Pic.BackColor, Pic.BackColor  'RGB(170, 203, 253)
 End If
 '<<--Draw Hover Rectangle-->>
 If OldNode.HoverItem = True Then
   DrawRectangle Pic.hdc, rctxt.Left - 2, rctxt.Top, IIf(m_Right, mRect.Right - OldNode.rcRight, mRect.Right), rctxt.Bottom - 1, HoverBackColor, HoverBackColor
 End If
 '<<--Draw Selected Rectangle-->>
  If OldNode.Selected = True Then
   Pic.ForeColor = SelectedForeColor
   DrawRectangle Pic.hdc, rctxt.Left - 2, rctxt.Top, IIf(m_Right, mRect.Right - OldNode.rcRight, mRect.Right), rctxt.Bottom, SelectedBackColor, SelectedBackColor
 Else
   If OldNode.ForeColor <> -1 Then Pic.ForeColor = OldNode.ForeColor
 End If

If Not OldNode.ButtonIcon Is Nothing Then
  rctxt.Right = rctxt.Right - 20
  DrawButton OldNode
End If
 rctxt.Right = IIf(m_Right, mRect.Right - OldNode.rcRight, mRect.Right)
If Not OldNode.ButtonIcon Is Nothing And m_Right Then rctxt.Left = rctxt.Left + 20
 If Not OldNode.ButtonIcon Is Nothing Then rctxt.Right = rctxt.Right - 20
 Lng = IIf(m_Right, DT_RTLREADING Or DT_RIGHT Or DT_WORDBREAK, DT_LEFT Or DT_WORDBREAK)
       If mWindowsNT Then
        DrawTextW Pic.hdc, StrPtr(OldNode.Caption), Len(OldNode.Caption), rctxt, Lng
      Else
        DrawText Pic.hdc, OldNode.Caption, Len(OldNode.Caption), rctxt, Lng
     End If
Pic.Refresh
Pic.ForeColor = ForeColor

End Sub
Private Sub pvTransBlt( _
        ByVal hdcDest As Long, _
        ByVal xDest As Long, _
        ByVal yDest As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hdcSrc As Long, _
        Optional ByVal xSrc As Long = 0, _
        Optional ByVal ySrc As Long = 0, _
        Optional ByVal clrMask As OLE_COLOR, _
        Optional ByVal hPal As Long = 0)
    Dim hdcMask As Long                               ' hDC of the created mask image
    Dim hdcColor As Long                              ' hDC of the created color image
    Dim hbmMask As Long                               ' Bitmap handle to the mask image
    Dim hbmColor As Long                              ' Bitmap handle to the color image
    Dim hbmColorOld As Long
    Dim hbmMaskOld As Long
    Dim hpalOld As Long
    Dim hdcScreen As Long
    Dim hdcScnBuffer As Long                          ' Buffer to do all work on
    Dim hbmScnBuffer As Long
    Dim hbmScnBufferOld As Long
    Dim hPalBufferOld As Long
    Dim lMaskColor As Long
    Dim hpalHalftone As Long

    hdcScreen = GetDC(0&)
    ' Validate palette
    If hPal = 0 Then
        hpalHalftone = CreateHalftonePalette(hdcScreen)
        hPal = hpalHalftone
    End If
    OleTranslateColor clrMask, hPal, lMaskColor
    lMaskColor = lMaskColor And &HFFFFFF
    ' Create a color bitmap to server as a copy of the destination
    ' Do all work on this bitmap and then copy it back over the destination
    ' when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
    ' Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hdcScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    ' Copy the destination to the screen buffer
    ApiBitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcDest, xDest, yDest, vbSrcCopy
    ' Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    ' hdcSrc, because this will create a DIB section if the original bitmap
    ' is a DIB section)
    hbmColor = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
    ' Now create a monochrome bitmap for the mask
    hbmMask = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
    ' First, blt the source bitmap onto the cover.  We do this first
    ' and then use it instead of the source bitmap
    ' because the source bitmap may be
    ' a DIB section, which behaves differently than a bitmap.
    ' (Specifically, copying from a DIB section to a monochrome bitmap
    ' does a nearest-color selection rather than painting based on the
    ' backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hdcScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hpalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    ' In case hdcSrc contains a monochrome bitmap, we must set the destination
    ' foreground/background colors according to those currently set in hdcSrc
    ' (because Windows will associate these colors with the two monochrome colors)
    Call SetBkColor(hdcColor, GetBkColor(hdcSrc))
    Call SetTextColor(hdcColor, GetTextColor(hdcSrc))
    Call ApiBitBlt(hdcColor, 0, 0, nWidth, nHeight, hdcSrc, xSrc, ySrc, vbSrcCopy)
    ' Paint the mask.  What we want is white at the transparent color
    ' from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hdcScreen)
    hbmMaskOld = SelectObject(hdcMask, hbmMask)
    ' When ApiBitBlt'ing from color to monochrome, Windows sets to 1
    ' all pixels that match the background color of the source DC.  All
    ' other bits are set to 0.
    Call SetBkColor(hdcColor, lMaskColor)
    Call SetTextColor(hdcColor, vbWhite)
    Call ApiBitBlt(hdcMask, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcCopy)
    ' Paint the rest of the cover bitmap.
    '
    ' What we want here is black at the transparent color, and
    ' the original colors everywhere else.  To do this, we first
    ' paint the original onto the cover (which we already did), then we
    ' AND the inverse of the mask onto that using the DSna ternary raster
    ' operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    ' Operation Codes", "Ternary Raster Operations", or search in MSDN
    ' for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    ' When ApiBitBlt'ing from monochrome to color, Windows transforms all white
    ' bits (1) to the background color of the destination hDC.  All black (0)
    ' bits are transformed to the foreground color.
    Call SetTextColor(hdcColor, vbBlack)
    Call SetBkColor(hdcColor, vbWhite)
    Call ApiBitBlt(hdcColor, 0, 0, nWidth, nHeight, hdcMask, 0, 0, DSna)
    ' Paint the Mask to the Screen buffer
    Call ApiBitBlt(hdcScnBuffer, 0, 0, nWidth, nHeight, hdcMask, 0, 0, vbSrcAnd)
    ' Paint the Color to the Screen buffer
    Call ApiBitBlt(hdcScnBuffer, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcPaint)
    ' Copy the screen buffer to the screen
    Call ApiBitBlt(hdcDest, xDest, yDest, nWidth, nHeight, hdcScnBuffer, 0, 0, vbSrcCopy)
    ' All done!
    Call ApiDeleteObject(SelectObject(hdcColor, hbmColorOld))
    Call SelectPalette(hdcColor, hpalOld, True)
    Call RealizePalette(hdcColor)
    Call DeleteDC(hdcColor)
    Call ApiDeleteObject(SelectObject(hdcScnBuffer, hbmScnBufferOld))
    Call SelectPalette(hdcScnBuffer, hPalBufferOld, 0)
    Call RealizePalette(hdcScnBuffer)
    Call DeleteDC(hdcScnBuffer)
    Call ApiDeleteObject(SelectObject(hdcMask, hbmMaskOld))
    Call DeleteDC(hdcMask)
    Call ReleaseDC(0&, hdcScreen)
    If hpalHalftone <> 0 Then
        Call ApiDeleteObject(hpalHalftone)
    End If
End Sub


Private Function DrawButton(Optional wNode As ClsNode, Optional Hover As Boolean) As Long
Dim NewDC As Long, wPic As StdPicture

If Hover Then
 If wNode Is Nothing Then
  Set wPic = mNode.ButtonHoverIcon
 Else: Set wPic = wNode.ButtonHoverIcon
 End If
Else
 If wNode Is Nothing Then
  Set wPic = mNode.ButtonIcon
 Else: Set wPic = wNode.ButtonIcon
 End If
End If

NewDC = CreateCompatibleDC(0)
If wNode Is Nothing Then
    SelectObject NewDC, wPic.Handle
    
    If wPic.Type = 3 Then
      Pic.PaintPicture wPic, IIf(m_Right, 4, mRect.Right - 20), mNode.rctop, 16, 16
    Else
      pvTransBlt Pic.hdc, IIf(m_Right, 4, mRect.Right - 20), mNode.rctop, 16, 16, NewDC, , , mNode.ButtonMaskColor
    End If
Else
    SelectObject NewDC, wPic.Handle
    
    If wPic.Type = 3 Then
      Pic.PaintPicture wPic, IIf(m_Right, 4, mRect.Right - 20), wNode.rctop, 16, 16
    Else
      pvTransBlt Pic.hdc, IIf(m_Right, 4, mRect.Right - 20), wNode.rctop, 16, 16, NewDC, , , wNode.ButtonMaskColor
    End If
End If
DeleteDC NewDC

End Function

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
If vNewValue < 0 Then
  m_BackColor = GetSysColor(vNewValue And &HFFFFFF)
Else
 m_BackColor = vNewValue
End If
 Pic.BackColor = m_BackColor
 PropertyChanged "BackColor"
If Not mNodes Is Nothing Then Call DrawTree
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
If vNewValue < 0 Then
  m_ForeColor = GetSysColor(vNewValue And &HFFFFFF)
Else
 m_ForeColor = vNewValue
End If
 Pic.ForeColor = m_ForeColor
 PropertyChanged "ForeColor"
If Not mNodes Is Nothing Then Call DrawTree

End Property

Public Property Get SelectedBackColor() As OLE_COLOR
SelectedBackColor = m_SelectedBackColor
End Property

Public Property Let SelectedBackColor(ByVal vNewValue As OLE_COLOR)
If vNewValue < 0 Then
  m_SelectedBackColor = GetSysColor(vNewValue And &HFFFFFF)
Else
 m_SelectedBackColor = vNewValue
End If
 PropertyChanged "SelectedBackColor"
If Not mNodes Is Nothing Then Call DrawTree

End Property

Public Property Get SelectedForeColor() As OLE_COLOR
SelectedForeColor = m_SelectedForeColor
End Property

Public Property Let SelectedForeColor(ByVal vNewValue As OLE_COLOR)
If vNewValue < 0 Then
  m_SelectedForeColor = GetSysColor(vNewValue And &HFFFFFF)
Else
 m_SelectedForeColor = vNewValue
End If
 PropertyChanged "SelectedForeColor"
If Not mNodes Is Nothing Then Call DrawTree

End Property


Public Sub Refresh()
Call DrawTree
End Sub

Public Property Get CheckIntegrity() As Boolean
Attribute CheckIntegrity.VB_Description = "If true the Tree will check each node position every time you add a new node,So it will slow the loading time."
CheckIntegrity = m_Checkintegrity
End Property

Public Property Let CheckIntegrity(ByVal vNewValue As Boolean)
 m_Checkintegrity = vNewValue
 PropertyChanged "CheckIntegrity"

End Property

Public Sub Clear()
Set mNodes = Nothing
Pic.Cls
OldSel = 0
VS1.Visible = False
End Sub

Public Property Get OldSel() As Long
Attribute OldSel.VB_MemberFlags = "40"
OldSel = mOldSel
End Property

Public Property Let OldSel(ByVal vNewValue As Long)
mOldSel = vNewValue
End Property

Public Property Get ShowTreeLines() As Boolean
ShowTreeLines = mShowTreeLines
End Property

Public Property Let ShowTreeLines(ByVal vNewValue As Boolean)
mShowTreeLines = vNewValue
If Not mNodes Is Nothing Then Call DrawTree
 PropertyChanged "ShowTreeLines"

End Property

Public Property Get LineStyle() As LineStyleEnum
LineStyle = mLineStyle
End Property

Public Property Let LineStyle(ByVal vNewValue As LineStyleEnum)
mLineStyle = vNewValue
If Not mNodes Is Nothing Then Call DrawTree
 PropertyChanged "LineStyle"
End Property

Private Sub DrawGradient(rc As RECT, MyClr As Long, GradientType As Long, wNode As ClsNode)
Dim i As Long, W As Long
Dim Start As Long, Ends As Long, MyClr2 As Long
Dim R As Integer, G As Integer, B As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer
Dim LR As Long, LG As Long, LB As Long
Dim s As Integer

W = IIf(m_Right, mRect.Right - wNode.rcRight, mRect.Right)
MyClr2 = SetLumi(MyClr, -30)


Select Case GradientType
'<<--Vertical gradient-->>
  Case 2
    Start = rc.Top
    Ends = rc.Bottom - 1
    For i = Start To Ends
      MyClr = SetLumi(MyClr, -1)
      DrawLine rc.Left - 2, i, W, i, MyClr
    Next
'<<--Horizontal gradient-->>
  Case 3
    GetRGB MyClr, R, G, B
    GetRGB BackColor, R2, G2, B2
    LR = W / (R - R2)
    LG = W / (G - G2)
    LB = W / (B - B2)
    
    Start = IIf(m_Right, W, rc.Left - 2)
    Ends = IIf(m_Right, rc.Left - 2, W)
    s = IIf(m_Right, -1, 1)
    For i = Start To Ends Step s
      If i Mod Abs(LR) = 0 Then R = R - Sgn(LR): MyClr = RGB(R, G, B)
      If i Mod Abs(LG) = 0 Then G = G - Sgn(LG): MyClr = RGB(R, G, B)
      If i Mod Abs(LB) = 0 Then B = B - Sgn(LB): MyClr = RGB(R, G, B)
      DrawLine i, rc.Top, i, rc.Bottom - 1, MyClr, 1
    Next
'<<--Mac gradient-->>
  Case 4
    Start = rc.Top
    Ends = rc.Bottom - 1
    For i = Start To Ends
     If i <= CInt(CInt((Ends + Start)) / 2) Then
      MyClr = SetLumi(MyClr, -2)
     Else
      MyClr = SetLumi(MyClr2, 3)
      MyClr2 = MyClr
     End If
     DrawLine rc.Left - 2, i, W, i, MyClr
    Next i
End Select
DrawLine rc.Left - 2, rc.Top + 1, W, rc.Top + 1, SetLumi(MyClr, -20)

End Sub
Private Function GetRGB(ColorValue As Long, Red As Integer, Green As Integer, Blue As Integer) As Long
Dim TempClr As Long

TempClr = Abs(ColorValue)
  
    Red = TempClr Mod &H100
    TempClr = TempClr \ &H100
    Green = TempClr Mod &H100
    TempClr = TempClr \ &H100
    Blue = TempClr Mod &H100
    GetRGB = RGB(Red, Green, Blue)

End Function

Public Function SetLumi(RGBClr As Long, Lum As Single)
Dim Hu As Integer, Lu As Integer, Su As Integer

ColorRGBToHLS RGBClr, Hu, Lu, Su
 If Lu + Lum > 240 Then Lum = 0: Lu = 240
 If Lu + Lum < 40 Then Lum = 0: Lu = 40

SetLumi = ColorHLSToRGB(Hu, Lu + Lum, Su)

End Function


Public Property Get TreeStyle() As TreeStyleEnum
TreeStyle = mTreeStyle
End Property

Public Property Let TreeStyle(ByVal vNewValue As TreeStyleEnum)
mTreeStyle = vNewValue
If Not mNodes Is Nothing Then Call DrawTree
PropertyChanged "TreeStyle"

End Property

Public Property Get HoverBackColor() As OLE_COLOR
HoverBackColor = m_HoverBackColor
End Property

Public Property Let HoverBackColor(ByVal vNewValue As OLE_COLOR)
If vNewValue < 0 Then
  m_HoverBackColor = GetSysColor(vNewValue And &HFFFFFF)
Else
 m_HoverBackColor = vNewValue
End If
 PropertyChanged "HoverBackColor"
If Not mNodes Is Nothing Then Call DrawTree

End Property
