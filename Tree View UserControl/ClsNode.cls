VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_caption As String
Private m_Key As String
Private m_Relative As String
Private m_Top As Long
Private m_Left As Long
Private m_Bottom As Long
Private m_Right As Long
Private m_Children As Long

Private m_Visable As Boolean
Private m_Selected As Boolean
Private m_HoverItem As Boolean

Private m_Expanded As Boolean
Private m_Multiline As Boolean

Private mIcon As StdPicture

Private m_ButtonIcon As StdPicture
Private m_ButtonHoverIcon As StdPicture
Private m_ButtonToolTip As String
Private m_ButtonMaskColor As Long
Private m_BackColor As Long
Private m_ForeColor As Long

Public FromCls As Boolean
Attribute FromCls.VB_VarMemberFlags = "40"
Public ItemNum As Long
Attribute ItemNum.VB_VarMemberFlags = "40"
Public Lines As Single
Attribute Lines.VB_VarMemberFlags = "40"
Public AllChildren As Long
Public ChNum As Long
Attribute ChNum.VB_VarMemberFlags = "40"
Public mParent As PrjTreeBe.TreeBe
Attribute mParent.VB_VarMemberFlags = "40"


Public Property Get Caption() As String
Caption = m_caption
FromCls = False

End Property

Public Property Let Caption(ByVal vNewcaption As String)
m_caption = vNewcaption
If FromCls = True Then
 FromCls = False:  mParent.DrawTree
End If

End Property

Public Property Get Key() As Variant
Key = m_Key
FromCls = False

End Property

Public Property Let Key(ByVal vNewKey As Variant)
m_Key = vNewKey
FromCls = False

End Property


Public Property Get Relative() As Variant
Relative = m_Relative
FromCls = False

End Property

Public Property Let Relative(ByVal vNewRelative As Variant)
m_Relative = vNewRelative
If FromCls = True Then
 FromCls = False:  mParent.DrawTree
End If

End Property

Friend Property Get rcLeft() As Long
rcLeft = m_Left
FromCls = False

End Property

Friend Property Let rcLeft(ByVal vNewrcLeft As Long)
m_Left = vNewrcLeft
FromCls = False

End Property

Friend Property Get rctop() As Long
rctop = m_Top
FromCls = False

End Property

Friend Property Let rctop(ByVal vNewrctop As Long)
m_Top = vNewrctop
FromCls = False

End Property

Friend Property Get rcBottom() As Long
rcBottom = m_Bottom
FromCls = False

End Property

Friend Property Let rcRight(ByVal vNewrcRight As Long)
m_Right = vNewrcRight
FromCls = False

End Property

Friend Property Get rcRight() As Long
rcRight = m_Right
FromCls = False

End Property

Friend Property Let rcBottom(ByVal vNewrcBottom As Long)
m_Bottom = vNewrcBottom
FromCls = False

End Property
Public Property Get Children() As Long
Children = m_Children
FromCls = False

End Property

Public Property Let Children(ByVal vNewChilds As Long)
m_Children = vNewChilds
FromCls = False

End Property

Public Property Get Visable() As Boolean
Visable = m_Visable
FromCls = False

End Property

Public Property Let Visable(ByVal vNewVisable As Boolean)
m_Visable = vNewVisable
If FromCls = True Then
 FromCls = False:  mParent.DrawTree
End If

End Property

Public Property Get Selected() As Boolean
Selected = m_Selected
FromCls = False

End Property

Public Property Let Selected(ByVal vNewSelected As Boolean)
m_Selected = vNewSelected
If FromCls = True Then
 FromCls = False:  mParent.DrawTree
End If

End Property

Public Property Get HoverItem() As Boolean
HoverItem = m_HoverItem
FromCls = False

End Property

Public Property Let HoverItem(ByVal vNewHoverItem As Boolean)
m_HoverItem = vNewHoverItem
FromCls = False

End Property

Public Property Get Expanded() As Boolean
Expanded = m_Expanded
FromCls = False

End Property

Public Property Let Expanded(ByVal vNewExpanded As Boolean)
If FromCls = True Then
 FromCls = False:  mParent.ExpandItem Me, vNewExpanded
Else
 m_Expanded = vNewExpanded
End If

End Property



Public Property Get Icon() As StdPicture
Set Icon = mIcon
FromCls = False
End Property

Public Property Set Icon(ByVal vNewValue As StdPicture)
Set mIcon = vNewValue
If FromCls = True Then
 FromCls = False:  mParent.DrawTree
End If

End Property

Friend Property Get ButtonIcon() As StdPicture
Set ButtonIcon = m_ButtonIcon
FromCls = False

End Property

Friend Property Set ButtonIcon(ByVal vNewValue As StdPicture)
Set m_ButtonIcon = vNewValue
If FromCls = True Then
 FromCls = False:  mParent.DrawTree
End If

End Property

Friend Property Get ButtonHoverIcon() As StdPicture
Set ButtonHoverIcon = m_ButtonHoverIcon
FromCls = False

End Property

Friend Property Set ButtonHoverIcon(ByVal vNewValue As StdPicture)
Set m_ButtonHoverIcon = vNewValue
If FromCls = True Then
 FromCls = False:  mParent.DrawTree
End If

End Property

Public Property Get ButtonToolTip() As String
ButtonToolTip = m_ButtonToolTip
FromCls = False

End Property

Public Property Let ButtonToolTip(ByVal vNewValue As String)
m_ButtonToolTip = vNewValue
FromCls = False

End Property

Public Property Get ButtonMaskColor() As Long
ButtonMaskColor = m_ButtonMaskColor
FromCls = False

End Property

Public Property Let ButtonMaskColor(ByVal vNewValue As Long)
m_ButtonMaskColor = vNewValue
If FromCls = True Then
 FromCls = False:  mParent.DrawTree
End If

End Property

Public Property Get BackColor() As Long
BackColor = m_BackColor
FromCls = False
End Property

Public Property Let BackColor(ByVal vNewValue As Long)
m_BackColor = vNewValue
If FromCls = True Then
 FromCls = False:  mParent.DrawTree
End If

End Property


Public Property Get ForeColor() As Long
ForeColor = m_ForeColor
FromCls = False
End Property

Public Property Let ForeColor(ByVal vNewValue As Long)
m_ForeColor = vNewValue
If FromCls = True Then
 FromCls = False:  mParent.DrawTree
End If

End Property

