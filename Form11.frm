VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5850
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5850
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label3 
      Caption         =   "24小时在线"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   1680
      X2              =   1680
      Y1              =   0
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   -120
      X2              =   5880
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      Caption         =   $"Form11.frx":0442
      BeginProperty Font 
         Name            =   "华文楷体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1920
      TabIndex        =   1
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Private Space Manage OP1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   240
      Picture         =   "Form11.frx":04A0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

