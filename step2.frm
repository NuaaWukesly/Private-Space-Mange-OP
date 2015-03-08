VERSION 5.00
Begin VB.Form step2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "安装提示"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4440
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "全程监督：15605188901"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   $"step2.frx":0000
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "step2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload step2
'Unload SetupMain
End Sub

Private Sub Form_Unload(Cancel As Integer)

Unload SetupMain
End Sub
