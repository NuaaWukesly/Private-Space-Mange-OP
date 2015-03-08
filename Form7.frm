VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form7 
   Caption         =   "进程"
   ClientHeight    =   1725
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   1725
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Created By 吴香礼 —— 于2013.02.02"
      BeginProperty Font 
         Name            =   "Adobe 楷体 Std R"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "正在初始化选项列表，请稍后..."
      BeginProperty Font 
         Name            =   "Adobe 黑体 Std R"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
