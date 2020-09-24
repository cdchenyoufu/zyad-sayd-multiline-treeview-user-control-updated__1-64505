VERSION 5.00
Object = "*\APrjTree.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   135
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   427
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1440
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin PrjTreeBe.TreeBe TreeBe3 
      Height          =   2400
      Left            =   0
      TabIndex        =   6
      Top             =   3720
      Width           =   5892
      _ExtentX        =   10398
      _ExtentY        =   4233
      PlusIcon        =   "Form1.frx":0342
      MinIcon         =   "Form1.frx":0490
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectedBackColor=   14653050
      CheckIntegrity  =   0   'False
      LineStyle       =   2
      TreeStyle       =   4
   End
   Begin PrjTreeBe.TreeBe TreeBe2 
      Height          =   2895
      Left            =   3000
      TabIndex        =   5
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5106
      PlusIcon        =   "Form1.frx":05DE
      MinIcon         =   "Form1.frx":072C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      SelectedBackColor=   14653050
      CheckIntegrity  =   0   'False
      ShowTreeLines   =   0   'False
      TreeStyle       =   3
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   960
      Picture         =   "Form1.frx":087A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1200
      Picture         =   "Form1.frx":0BBC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   720
      Picture         =   "Form1.frx":0EFE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      Picture         =   "Form1.frx":1240
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin PrjTreeBe.TreeBe TreeBe1 
      Height          =   2892
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2652
      _ExtentX        =   4683
      _ExtentY        =   5106
      PlusIcon        =   "Form1.frx":1E82
      MinIcon         =   "Form1.frx":1FD0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      SelectedBackColor=   14653050
      CheckIntegrity  =   0   'False
      TreeStyle       =   4
      HoverBackColor  =   13348764
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TreeView With a multiline support (Resize the form)/Right Click For Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   0
      TabIndex        =   8
      Top             =   3360
      Width           =   6564
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TreeViw with a custem Colors"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   3180
   End
   Begin VB.Menu MainM 
      Caption         =   "ff"
      Visible         =   0   'False
      Begin VB.Menu RTL 
         Caption         =   "Right To Left"
      End
      Begin VB.Menu Spr1 
         Caption         =   "-"
      End
      Begin VB.Menu ExpAll 
         Caption         =   "Expand All"
      End
      Begin VB.Menu ColAll 
         Caption         =   "Collapse All"
      End
      Begin VB.Menu Spr2 
         Caption         =   "-"
      End
      Begin VB.Menu NC 
         Caption         =   "Normal Color"
      End
      Begin VB.Menu HG 
         Caption         =   "Horizontal Gradient"
      End
      Begin VB.Menu VG 
         Caption         =   "Vertical Gradient"
      End
      Begin VB.Menu MG 
         Caption         =   "Mac gradient"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ColAll_Click()
For i = 1 To TreeBe3.Nodes.Count
 TreeBe3.Nodes(i).Expanded = False
Next

End Sub

Private Sub ExpAll_Click()
For i = 1 To TreeBe3.Nodes.Count
 TreeBe3.Nodes(i).Expanded = True
Next

End Sub

Private Sub Form_Load()
For i = 1 To 100
  TreeBe1.Nodes.AddItem CStr(i), "Sample" & CStr(i), , RGB(156, 175, 203), vbWhite, Picture1
  TreeBe1.Nodes.AddItem "A" & CStr(i), "Child1", CStr(i), TreeBe1.BackColor, vbBlack, Picture5
  TreeBe1.Nodes.AddItem "B" & CStr(i), "Child2", CStr(i), TreeBe1.BackColor, vbBlack, Picture5
  TreeBe1.Nodes.AddItem "C" & CStr(i), "Child3", "B" & CStr(i), TreeBe1.BackColor, vbBlack, Picture5
Next i


For i = 1 To 100
  TreeBe2.Nodes.AddItem CStr(i), "Data" & CStr(i), , RGB(168, 200, 169), vbBlue, Picture1
  TreeBe2.Nodes.AddItem "A" & CStr(i), "Child1", CStr(i), TreeBe2.BackColor, vbBlack, Picture5
  TreeBe2.Nodes.AddItem "B" & CStr(i), "Child2", CStr(i), TreeBe2.BackColor, vbBlack, Picture5
  TreeBe2.Nodes.AddItem "C" & CStr(i), "Child3", "B" & CStr(i), TreeBe2.BackColor, vbBlack, Picture5
  TreeBe2.Nodes.AddItem "D" & CStr(i), "Child4", CStr(i), TreeBe2.BackColor, vbBlack, Picture5
Next i


'=========================================================================
  TreeBe3.Nodes.AddItem "1", "Welcome to Microsoft Visual Basic", , RGB(253, 236, 189), vbBlack, Picture1, "Add To Favorites", Picture3, Picture4, vbWhite
  TreeBe3.Nodes.AddItem "2", "Welcome to Microsoft Visual Basic, the fastest and easiest way to create applications for Microsoft Windows®.", "1", TreeBe3.BackColor, vbBlack, Picture2

  TreeBe3.Nodes.AddItem "3", "Installing Visual Basic", , RGB(253, 236, 189), vbBlack, Picture1, "Add To Favorites", Picture3, Picture4, vbWhite
  TreeBe3.Nodes.AddItem "4", "You install Visual Basic on your computer using the Setup program. The Setup program installs Visual Basic and other product components from the CD-ROM to your hard disc.", "3", TreeBe3.BackColor, vbBlack, Picture2

  TreeBe3.Nodes.AddItem "5", "Visual Basic Concepts", , RGB(253, 236, 189), vbBlack, Picture1, "Add To Favorites", Picture3, Picture4, vbWhite
  TreeBe3.Nodes.AddItem "6", "In order to understand the application development process, it is helpful to understand some of the key concepts upon which Visual Basic is built.", "5", TreeBe3.BackColor, vbBlack, Picture2

  TreeBe3.Nodes.AddItem "7", "Understanding Properties, Methods and Events", , RGB(253, 236, 189), vbBlack, Picture1, "Add To Favorites", Picture3, Picture4, vbWhite
  TreeBe3.Nodes.AddItem "8", "Visual Basic forms and controls are objects which expose their own properties, methods and events. Properties can be thought of as an object's attributes, methods as its actions, and events as its responses.", "7", TreeBe3.BackColor, vbBlack, Picture2

TreeBe1.Height = TreeBe1.LineHeight * 10
TreeBe1.HoverBackColor = RGB(231, 179, 150)
TreeBe2.Height = TreeBe2.LineHeight * 10
TreeBe2.HoverBackColor = RGB(215, 227, 196)
TreeBe3.HoverBackColor = RGB(255, 244, 215)

End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then
  TreeBe3.Width = Me.ScaleWidth - 5
  TreeBe3.Height = Me.ScaleHeight - TreeBe3.Top - 5
End If
End Sub

Private Sub HG_Click()
TreeBe3.TreeStyle = HGradient

End Sub

Private Sub MG_Click()
TreeBe3.TreeStyle = MacStyle

End Sub

Private Sub NC_Click()
TreeBe3.TreeStyle = NormalColor
End Sub

Private Sub RTL_Click()
TreeBe3.RightToLeft = Not TreeBe3.RightToLeft

End Sub

Private Sub TreeBe3_ButtonClick(ByVal Node As PrjTreeBe.ClsNode)
MsgBox "Button Clicked"
End Sub

Private Sub TreeBe3_Click(ByVal Button As Integer)
If Button = 2 Then
 Me.PopupMenu MainM
End If
End Sub

Private Sub VG_Click()
TreeBe3.TreeStyle = VGradient

End Sub
