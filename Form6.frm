VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选项卡"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "华文楷体"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Index           =   2
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   6555
      TabIndex        =   12
      Top             =   480
      Width           =   6615
      Begin VB.ComboBox Combo2 
         Height          =   330
         Index           =   0
         Left            =   3960
         TabIndex        =   18
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Height          =   330
         Index           =   1
         Left            =   3960
         TabIndex        =   17
         Top             =   1080
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Height          =   330
         Index           =   2
         Left            =   3960
         TabIndex        =   14
         Top             =   1800
         Width           =   2535
      End
      Begin VB.ComboBox Combo5 
         Height          =   330
         ItemData        =   "Form6.frx":1176A
         Left            =   1200
         List            =   "Form6.frx":11774
         TabIndex        =   13
         Top             =   2520
         Width           =   5295
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   2
         Left            =   3600
         TabIndex        =   15
         Top             =   1800
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Max             =   32767
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   16
         Top             =   1800
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   19
         Top             =   1080
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Max             =   32767
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   20
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   21
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Max             =   32767
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   22
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "缩   进："
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "右缩进："
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "悬挂缩进："
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "项目符号："
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取     消"
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确     定"
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   4320
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Index           =   1
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   480
      Width           =   6615
      Begin ComctlLib.Slider Slider1 
         Height          =   495
         Left            =   2520
         TabIndex        =   5
         Top             =   840
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         _Version        =   327682
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##磅"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   " 颜     色 ："
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "字体大小："
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "  字    体 ："
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4095
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7223
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "字体"
            Key             =   "Font"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "段落"
            Key             =   "Paragraph"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Created By 吴香礼 —— 2013年2月2日"
      BeginProperty Font 
         Name            =   "Adobe 楷体 Std R"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   2415
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long


Private Sub Combo2_Change(Index As Integer)
Static OldUnit(2) As String
Select Case Combo2(Index).Text
Case "磅"
UpDown1(Index).Increment = 100
If OldUnit(Index) = "厘米" Then
MaskEdBox2(Index).Text = Int(Val(MaskEdBox2(Index).Text) / 2.54 * 72)
End If
Case "厘米"
UpDown1(Index).Increment = 2
If OldUnit(Index) = "磅" Then
MaskEdBox2(Index).Text = CDbl(Int(Val(MaskEdBox2(Index).Text) / 72 * 2.54 * 100)) / 100
End If
End Select
OldUnit(Index) = Combo2(Index).Text
End Sub

Private Sub Command1_Click()
If Val(MaskEdBox1.ClipText) < 8 Or Val(MaskEdBox1.ClipText) > 80 Then
MsgBox "字体大小设置无效，它必须是介于8到80间的整数", , "错误"
Else
Slider1.Value = Val(MaskEdBox1.ClipText) / 4
Form6.Hide
End If
End Sub

Private Sub Command2_Click()
Form5.SetCurValue
Form6.Hide
End Sub

Private Sub Form_Load()
With TabStrip1
'把各选项卡所对应的pictureBox控件移到tabstrip1控件客户区
Picture1(1).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
Picture1(2).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
End With
'使用前所有的选项卡对应的picturebox控件位于其他控件的最上方
Picture1(TabStrip1.selectedItem.Index).ZOrder 0
For i = 0 To 2
Combo2(i).AddItem "磅"
Combo2(i).AddItem "厘米"
Combo2(i).ListIndex = 0
Next
End Sub

Private Sub Label10_Click()
On Error GoTo ErrHandler
Form5.CommonDialog1.DialogTitle = "字体颜色"
Form5.CommonDialog1.ShowColor
Label10.BackStyle = 1
Label10.BackColor = Form5.CommonDialog1.Color
Exit Sub
ErrHandler:
Exit Sub
End Sub

Private Sub MaskEdBox1_Change()
Slider1.Value = Val(MaskEdBox1.ClipText) / 4
End Sub

Private Sub Slider1_Click()
If Slider1.Value < 2 Then
MaskEdBox1.Text = Format(8, "00磅")
ElseIf Slider1.Value > 20 Then
MaskEdBox1.Text = Format(80, "00磅")
Else
MaskEdBox1.Text = Format(Slider1.Value * 4, "00磅")
End If
End Sub

Private Sub TabStrip1_Click()
Picture1(TabStrip1.selectedItem.Index).ZOrder 0
End Sub

Private Sub UpDown1_DownClick(Index As Integer)
MaskEdBox2(Index).Text = MaskEdBox2(Index).Text - UpDown1(Index).Increment / 100
End Sub

Private Sub UpDown1_UpClick(Index As Integer)
MaskEdBox2(Index).Text = MaskEdBox2(Index).Text + UpDown1(Index).Increment / 100
End Sub
