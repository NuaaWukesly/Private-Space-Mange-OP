VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyBrowse说明"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6030
   Icon            =   "Explain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6030
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "说明"
      Height          =   2055
      Left            =   1800
      TabIndex        =   2
      Top             =   2280
      Width           =   4215
      Begin VB.Label Label2 
         Caption         =   $"Explain.frx":4F32
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "声明"
      Height          =   1935
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.Label Label1 
         Caption         =   $"Explain.frx":50CE
         Height          =   1455
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Label Label3 
      Caption         =   $"Explain.frx":52BA
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   120
      Picture         =   "Explain.frx":534A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6000
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   1680
      Y1              =   0
      Y2              =   4440
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
