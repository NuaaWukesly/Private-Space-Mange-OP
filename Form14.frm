VERSION 5.00
Begin VB.Form Form14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择生日浏览器风格"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4590
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "取     消"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确      定"
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "风格二"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "风格一"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   2520
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   2520
      Y1              =   0
      Y2              =   2280
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Form14.Option1 = True Then
'风格一
Load Form13
Form13.Show
Else
If Form14.Option2 = True Then
Load MyBrowse
End If
End If
Unload Form14
End Sub

Private Sub Command2_Click()
Unload Form14
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form2.WindowsMediaPlayer1.Close
End Sub

