VERSION 5.00
Begin VB.Form Backup 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6060
   Icon            =   "Backup.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6060
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "是否打开输出文件夹"
      Height          =   1095
      Left            =   2520
      TabIndex        =   10
      Top             =   1440
      Width           =   3495
      Begin VB.CheckBox Check1 
         Caption         =   "打开输出文件夹"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Value           =   1  'Checked
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "将备份的文件夹路径"
      Height          =   735
      Left            =   2520
      TabIndex        =   8
      Top             =   600
      Width           =   3495
      Begin VB.Label Label3 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "备份保存路径"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   5895
      Begin VB.Label Label2 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "开始备份"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新建文件夹"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   0
      X2              =   6000
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "请您选择备份保存的路径"
      BeginProperty Font 
         Name            =   "Adobe 楷体 Std R"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim backupPath As String                  '备份保存的路径
Dim needPath As String                    '需要备份的文件夹路径
Dim ourFSO As New FileSystemObject        '用于文件操作
Dim backupSign As String                  '备份类型
Dim openStr As String                     '打开资源管理器时显示的路径
Dim sign As Boolean                       '标志是否为初始状态
'用资源管理器打开指定文件
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'ShellExecute hWnd, "open", "explorer.exe", "/e,/select,E:\焦点文件.MP3", "", 1
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'通过explorer.exe的开关实现.
'Explorer.exe的参数如下：
'命令格式Explorer [/n][/e][[,/root],[path]][[,/select],[path filename]]
'参数说明
'/n表示以“我的电脑”方式打开一个新的窗口，通常打开的是Windows安装分区的根目录。
'/e表示以“资源管理器”方式打开一个新的窗口，通常打开的也是Windows安装分区的根目录。
'/root,[path]表示打开指定的文件夹，
'/root表示只显示指定文件夹下面的文件（夹），
'不显示其它磁盘分区和文件夹；[path]表示指定的路径。
'如果不加/root参数，而只用[path]参数，
'则可以显示其它磁盘分区和文件夹中的内容。
'另外，[path]还可以指定网络共享文件夹。
'/select,[path filename]表示打开指定的文件夹并且选中指定的文件，
'[path filename]表示指定的路径和文件名。
'如果不加/select参数，则系统会用相应的关联程序打开该文件。
'如果[path filename]不跟文件名就会打开该文件夹的上级目录并选中该文件夹。
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'定义一个函数，判断备份类型，并将其返回
Function judge(ByVal judgestr As String) As String
If Not judgestr = "" Then
If judgestr = "备份我们的情书" Then
judge = "OurLetter"
End If
If judgestr = "备份我们的相片" Then
judge = "OurPicture"
End If
If judgestr = "备份我们的历程" Then
judge = "OurCourse"
End If
If judgestr = "备份我们的幻灯片" Then
judge = "PPS"
End If
If judgestr = "备份EXE" Then
judge = "EXE"
End If
End If
End Function
 
Private Sub Command1_Click()
Dim newfold As String                     '新建文件夹名
Dim curPath As String                     '记录当前路径
'新建文件夹
curPath = Backup.Dir1.Path
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
Backup.Dir1.Path = curPath & "\" & newfold
End If
End Sub

Private Sub Command2_Click()
'开始备份
'需要备份的文件夹路径，在form2中给出
needPath = Backup.Label3.Caption
backupSign = judge(Backup.Caption)
Backup.Label2.Caption = backupPath
If ourFSO.FolderExists(backupPath & "\" & backupSign) = False Then
'新建文件夹取名：我的备份
MkDir backupPath & "\" & backupSign
End If
If Len(backupPath) = 3 Then
backupPath = backupPath & backupSign
Else
backupPath = backupPath & "\" & backupSign
End If
'若文件已存在则覆盖
ourFSO.CopyFolder needPath, backupPath
MsgBox Backup.Caption & "成功！", , "系统信息"
If Backup.Check1.Value = 1 Then
Mid(backupPath, 1, 1) = UCase(Mid(backupPath, 1, 1))
'格式
openStr = "/e,/select," & backupPath
ShellExecute hWnd, "open", "explorer.exe", openStr, "", 1
End If
Unload Backup
End Sub

Private Sub Command3_Click()
a = MsgBox("你去定要退出" & Backup.Caption & "吗？", vbYesNo, "系统提示")
If a = vbYes Then
Unload Backup
End If
End Sub

Private Sub Dir1_Change()
If sign = False Then                      '如果不是初始
backupPath = Backup.Dir1.Path             '显示新路径
Backup.Label2.Caption = backupPath
Else
sign = False
End If
End Sub

Private Sub Drive1_Change()
'驱动器改变时,同步
If Not sign = True Then
Backup.Dir1.Path = Backup.Drive1.Drive
End If
End Sub

Private Sub Form_Load()
Dim WshShell As Object, WScript As Object, oShellLink As Object
Set WshShell = CreateObject("WScript.Shell")
sign = True                                       '初始
backupPath = WshShell.SpecialFolders("Desktop")   '桌面路径
Backup.Drive1 = Mid(backupPath, 1, 3)
Backup.Dir1.Path = backupPath
Backup.Label2.Caption = backupPath
Backup.Show
End Sub

