VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "产品说明"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5820
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5820
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   4215
      Begin VB.Label Label2 
         Caption         =   $"Form8.frx":4F32
         BeginProperty Font 
            Name            =   "华文楷体"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Created By 吴香礼㊣ Private Space Manage OP1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2760
      Width           =   4935
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   120
      Picture         =   "Form8.frx":4FDC
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
