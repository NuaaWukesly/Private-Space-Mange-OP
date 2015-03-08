VERSION 5.00
Begin VB.Form ModPW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "更改密码"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "ModPW.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "取 消"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   8.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确 定"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   8.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1560
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1560
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1560
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "确认："
      BeginProperty Font 
         Name            =   "华文楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "新密码："
      BeginProperty Font 
         Name            =   "华文楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "旧密码："
      BeginProperty Font 
         Name            =   "华文楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "ModPW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldpw As String
'加密函数
Private Function Code(OriStr As String) As String
Dim i As Integer, n As Integer
Dim tempStr As String
tempStr = OriStr
n = Len(OriStr)
For i = 1 To n
'加密过程，每个字符与整数3异或
Mid(tempStr, i, 1) = Chr((Asc((Mid(OriStr, i, 1))) Xor 90))
Next i
Code = tempStr
End Function

Private Sub Command1_Click()
'确定
If ModPW.Text1(0).Text = oldpw Then
'密码正确
If Not ModPW.Text1(1) = "" Then
'新密码非空
If ModPW.Text1(1).Text = ModPW.Text1(2).Text Then
'新密码与确认密码相同
'以输出形式打开文件
Open App.Path & "\System\key.pw" For Output As #2
'写入新密码，该方法直接将原内容覆盖
Print #2, Code(ModPW.Text1(1).Text)
'关闭文件
Close #2
MsgBox "密码更改成功！您的新密码为：" & ModPW.Text1(1).Text, , "Private Place Manage OP1.0"
Load LoadPage
LoadPage.Show
Unload ModPW
Else
choice = MsgBox("新密码与确认密码不相同！", 4, "Private Place Manage OP1.0")
If Not choice = vbYes Then
Load LoadPage
LoadPage.Show
Unload ModPW
End If
End If
Else
choice = MsgBox("新密码与确认密码不能为空！", 4, "Private Place Manage OP1.0")
If Not choice = vbYes Then
Load LoadPage
LoadPage.Show
Unload ModPW
End If
End If
Else
choice = MsgBox("密码错误！", 4, "Private Place Manage OP1.0")
If Not choice = vbYes Then
Load LoadPage
LoadPage.Show
Unload ModPW
End If
End If
End Sub

Private Sub Command2_Click()
Load LoadPage
LoadPage.Show
Unload ModPW
End Sub

Private Sub Form_Load()
'打开密码文件
Open App.Path & "\System\key.pw" For Input As #1
'读取密码
Line Input #1, oldpw
'解密
oldpw = Code(oldpw)
'关闭密码文件
Close #1
End Sub

