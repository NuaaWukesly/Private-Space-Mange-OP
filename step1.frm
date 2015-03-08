VERSION 5.00
Begin VB.Form step1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome "
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6120
   Icon            =   "step1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "安装设置"
      BeginProperty Font 
         Name            =   "Adobe 楷体 Std R"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2880
      TabIndex        =   5
      Top             =   1440
      Width           =   3135
      Begin VB.CheckBox Check2 
         Caption         =   "默认安装"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Adobe 楷体 Std R"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   7
         Top             =   960
         Value           =   2  'Grayed
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "创建桌面快捷方式"
         BeginProperty Font 
            Name            =   "Adobe 楷体 Std R"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "下一步（Next）"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取 消（Cancel）"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.Image Image1 
         Height          =   3135
         Left            =   120
         Picture         =   "step1.frx":4F32
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright （c）2013-2014 NUAA  161140225 吴香礼"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      X1              =   2760
      X2              =   2760
      Y1              =   3360
      Y2              =   4200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      X1              =   2760
      X2              =   2760
      Y1              =   0
      Y2              =   3360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Welcome Using The Private Space Manage OP1.0"
      BeginProperty Font 
         Name            =   "Adobe 楷体 Std R"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   0
      X2              =   6120
      Y1              =   3360
      Y2              =   3360
   End
End
Attribute VB_Name = "step1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancel_Click()
'点击取消时
If MsgBox("私人空间管理软件正在安装，退出安装将无法完成。您确定要退出软件的安装吗？", vbYesNo, "系统提示") = vbYes Then
Unload step1
Unload SetupMain
End If
End Sub

Private Sub CmdNext_Click()
Unload step1
Load step3
step3.Show 1, SetupMain
End Sub
