VERSION 5.00
Begin VB.Form step3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Destination Location"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6585
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "取消"
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "安装"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "上一步"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新建文件夹"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3840
      Width           =   3975
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   $"Step3.frx":0000
      Height          =   2055
      Left            =   2640
      TabIndex        =   10
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "版权声明"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   600
      Width           =   3615
   End
   Begin VB.Line Line5 
      X1              =   1320
      X2              =   1320
      Y1              =   3720
      Y2              =   4560
   End
   Begin VB.Line Line4 
      X1              =   2520
      X2              =   6600
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   6600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      X1              =   2520
      X2              =   2520
      Y1              =   480
      Y2              =   3720
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6600
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "安装路径："
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "请选择安装路径"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "step3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim installPath As String    '安装路径
Dim curPath As String        '当前路径
Dim newfold As String        '新文件夹名
Dim ExePath As String        '安装后的应用程序路径
Dim fold As New FileSystemObject

Public Sub mShellLnk(ByVal LnkName As String, IconFileIconIndex As String, ByVal FilePath As String, Optional ByVal FileName As String, Optional ByVal StrArg As String, Optional ByVal HookKey As String = "", Optional ByVal StrRemark As String = "", Optional ByVal strDesktop As String = "")
Dim WshShell As Object, WScript As Object, oShellLink As Object
Set WshShell = CreateObject("WScript.Shell")
If strDesktop = "" Then
strDesktop = WshShell.SpecialFolders("StartMenu")   '桌面路径StartupDesktop
End If
If UCase(Right(LnkName, 4)) = ".LNK" Then
Set oShellLink = WshShell.CreateShortcut(strDesktop & "\" & LnkName)
'创建快捷方式,参数为路径和名称
Else
Set oShellLink = WshShell.CreateShortcut(strDesktop & "\" & LnkName & ".lnk")
End If
oShellLink.TargetPath = FilePath & "\" & FileName
oShellLink.Arguments = StrArg
oShellLink.WindowStyle = 1 '风格,常规窗口
oShellLink.Hotkey = HookKey '热键
oShellLink.IconLocation = IconFileIconIndex '图标
oShellLink.Description = StrRemark '快捷方式备注内容
oShellLink.WorkingDirectory = FilePath '源文件所在目录
oShellLink.Save
'保存创建的快捷方式
Set WshShell = Nothing
Set oShellLink = Nothing
End Sub
Private Sub Command1_Click()
'新建文件夹
curPath = step3.Dir1.Path
'新建文件夹必须有一个名字
Do
newfold = InputBox("请输入新建文件夹名", "系统提示")
If newfold = "" Then
a = MsgBox("输入不能为空，您要请从新输入吗？", vbYesNo, 系统提示)
End If
Loop While a = vbYes
If Not newfold = "" Then
MkDir curPath & "\" & newfold
Dir1.Refresh
'定位到新建文件夹
step3.Dir1.Path = curPath & "\" & newfold
End If
End Sub

Private Sub Command2_Click()
Unload step3
Load step1
step1.Show 1, SetupMain
End Sub

Private Sub Command3_Click()
'点击安装时，显示全路径
installPath = step3.Dir1.Path
'将数据文件复制到安装路径下
fold.CopyFolder App.Path & "\SetupData\Private Space Manage OP1.0", installPath
step3.Dir1.Refresh
If Len(installPath) > 3 Then
mShellLnk "Private Space Manage", "notepad.exe", installPath & "\Private Space Manage OP1.0", "\Private Space Manage.exe" ', "c:\boot.ini"
Else
mShellLnk "Private Space Manage", "notepad.exe", installPath & "Private Space Manage OP1.0", "\Private Space Manage.exe"
End If
'安装完成
Unload step3
Load step2
step2.Show
End Sub

Private Sub Dir1_Change()
step3.Text1.Text = step3.Dir1.Path
installPath = step3.Dir1.Path
End Sub

Private Sub Drive1_Change()
'驱动器改变,驱动器同步
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
step3.Dir1 = App.Path
step3.Drive1.Drive = Mid(App.Path, 1, 2)
End Sub
