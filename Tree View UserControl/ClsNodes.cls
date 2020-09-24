VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private ItemsCol As Collection
Private mParent As PrjTreeBe.TreeBe

Public Sub AddItem(Key, Text As String, Optional Relative, Optional BackColor As Long = -1, Optional ForeColor As Long = -1, Optional Icon As StdPicture, Optional ButtonToolTip As String, Optional ButtonIcon As StdPicture, Optional ButtonHoverIcon As StdPicture, Optional ButtonMaskColor As Long = -1)
Dim NewNode As New ClsNode, tmpNode As ClsNode
Set tmpNode = NewNode
Set NewNode.mParent = Me.ParentControl
Dim tmpRelative As Variant
Dim i As Long

'<--Collect the node date-->>
NewNode.Caption = Text
NewNode.Key = Key
NewNode.rcLeft = 36
NewNode.rcRight = 36

NewNode.BackColor = BackColor
NewNode.ForeColor = ForeColor
If Not IsMissing(Icon) Then Set NewNode.Icon = Icon
If Not ButtonIcon Is Nothing Then Set NewNode.ButtonIcon = ButtonIcon
If ButtonToolTip <> "" Then NewNode.ButtonToolTip = ButtonToolTip
If Not ButtonHoverIcon Is Nothing Then Set NewNode.ButtonHoverIcon = ButtonHoverIcon
If ButtonMaskColor <> -1 Then NewNode.ButtonMaskColor = ButtonMaskColor

If Not IsMissing(Relative) Then
tmpRelative = Relative
 NewNode.Relative = Relative
 NewNode.Visable = False
 
 Set tmpNode = ItemsCol(Relative)
 ItemsCol(Relative).Children = ItemsCol(Relative).Children + 1

Do
 ItemsCol(tmpRelative).AllChildren = ItemsCol(tmpRelative).AllChildren + 1
 tmpRelative = ItemsCol(tmpRelative).Relative
Loop Until tmpRelative = ""

 NewNode.rcLeft = tmpNode.rcLeft + 36
 NewNode.rcRight = tmpNode.rcRight + 36
 NewNode.ChNum = ItemsCol(Relative).Children
'<--Assign the node date to the collection after its parent-->>
 ItemsCol.Add NewNode, Key, , CInt(tmpNode.ItemNum) + CInt(tmpNode.AllChildren - 1)
Else
  NewNode.Visable = True
 '<--Assign the node date to the collection-->>
 ItemsCol.Add NewNode, Key
End If

'mParent.DrawTree (True)

'------------------------------------------------------------------------
'If the CheckIntegrity is true the Tree will check each node position every
'time you add a new node. This have a bad effect on the speed.
'ie: set it to true in this case:-
 'TreeBe1.Nodes.AddItem "1", "Node1"
 'TreeBe1.Nodes.AddItem "2", "Node2"
 'TreeBe1.Nodes.AddItem "3", "Node1 Child", "1"
'set it to false in this case:-
 'TreeBe1.Nodes.AddItem "1", "Node1"
 'TreeBe1.Nodes.AddItem "2", "Node1 Child", "1"
 'TreeBe1.Nodes.AddItem "3", "Node2", "1"
'------------------------------------------------------------------------
If mParent.CheckIntegrity Then
  For Each NewNode In ItemsCol
   i = i + 1
   NewNode.ItemNum = i
  Next
Else
 NewNode.ItemNum = ItemsCol.Count
End If

End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = ItemsCol.[_NewEnum]
    
End Property

Private Sub Class_Initialize()
Dim NewNode As New ClsNode
Set NewNode.mParent = Me.ParentControl

NewNode.Caption = "First-Hidden-Node"
NewNode.Visable = True
NewNode.ItemNum = 1

Set ItemsCol = New Collection
ItemsCol.Add NewNode

End Sub

Public Property Get nItem(ItemNum As Variant, Optional Real As Boolean) As ClsNode
Attribute nItem.VB_UserMemId = 0
If Real = True Then
  Set nItem = ItemsCol(ItemNum)
  nItem.FromCls = True
Else
  If IsNumeric(ItemNum) Then ItemNum = ItemNum + 1
  Set nItem = ItemsCol(ItemNum)
  nItem.FromCls = True
End If
End Property


Private Sub Class_Terminate()
Set ItemsCol = Nothing

End Sub


Friend Property Get ParentControl() As PrjTreeBe.TreeBe
Set ParentControl = mParent
End Property

Friend Property Set ParentControl(ByVal vNewParentControl As PrjTreeBe.TreeBe)
Set mParent = vNewParentControl
End Property

Public Property Get Count() As Long
Count = ItemsCol.Count - 1
End Property

Public Sub Remove(item)
Dim NewNode As New ClsNode
Set NewNode.mParent = Me.ParentControl
Dim i As Long, ParentNode, Children, ItemNum As Long

If IsNumeric(item) Then item = item + 1
ParentNode = ItemsCol(item).Relative
Children = ItemsCol(item).AllChildren
ItemNum = ItemsCol(item).ItemNum

'<<--if the node is a child then remove it from it's parent list-->>
If ParentNode <> "" Then
 ItemsCol(ParentNode).AllChildren = ItemsCol(ParentNode).AllChildren - 1
 ItemsCol(ParentNode).Children = ItemsCol(ParentNode).Children - 1
End If

'<<--if the node is a parent then remove all it's children list-->>
If Children > 0 Then
  For i = ItemNum To ItemNum + Children
    Debug.Print ItemsCol(ItemNum).Caption
    ItemsCol.Remove ItemNum
  Next
Else
  ItemsCol.Remove item
End If
i = 0

'Reposition the nodes
For Each NewNode In ItemsCol
   i = i + 1
   NewNode.ItemNum = i
Next
mParent.OldSel = 0
mParent.DrawTree
End Sub
