VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form setBirthday 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置生日"
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4245
   Icon            =   "SetBD.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   4245
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "确 定 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   167837696
      CurrentDate     =   41350
   End
End
Attribute VB_Name = "setBirthday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim birthday As String
Dim bd As Date
Dim awayBD As Integer
Dim fileFSO As New FileSystemObject
Private Sub Command1_Click()
Open App.Path & "\System\Birthday.bd" For Output As 1
'该方法直接将原内容覆盖
Print #1, setBirthday.DTPicker1.Value
Close #1
Open App.Path & "\System\Birthday.bd" For Input As 2
Line Input #2, birthday
Close #2
bd = DateValue(birthday)
MsgBox "你的生日为：" & (bd), , "Private Place Manage OP1.0"
'计算距生日的天数
awayBD = CInt(DateDiff("d", Date, CDate(Month(bd) & "-" & Day(bd))))
If awayBD = 0 Then
Form2.StatusBar1.Panels(3).Text = "今天是您的生日！请点击此处！！！"
ElseIf awayBD < 0 Then
awayBD = awayBD + 365
Form2.StatusBar1.Panels(3).Text = "你的生日是：" & Month(bd) & "月" & Day(bd) & "日,据您生日还有" & CStr(awayBD) & "天"
Else
Form2.StatusBar1.Panels(3).Text = "你的生日是：" & Month(bd) & "月" & Day(bd) & "日,据您生日还有" & CStr(awayBD) & "天"
End If
Unload setBirthday
End Sub

Private Sub Form_Load()
setBirthday.DTPicker1.Value = Date
End Sub
