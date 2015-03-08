VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000008&
   Caption         =   "Our Place"
   ClientHeight    =   6030
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   9915
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6030
   ScaleWidth      =   9915
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   5655
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9975
      ExtentX         =   17595
      ExtentY         =   9975
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer Timer1 
      Left            =   1560
      Top             =   1200
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5655
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   855
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   999
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1931
      _cy             =   1508
   End
   Begin VB.Menu munFile 
      Caption         =   "系统（&F）"
      Begin VB.Menu munFileBackgroundMusic 
         Caption         =   "背景音乐"
         Begin VB.Menu SetBGM 
            Caption         =   "设置背景音乐"
            Begin VB.Menu AddBGM 
               Caption         =   "添加背景音乐"
            End
            Begin VB.Menu ChangeBGM 
               Caption         =   "改变背景音乐"
            End
         End
         Begin VB.Menu munFileBackgroundMusicOpen 
            Caption         =   "打开音乐"
         End
         Begin VB.Menu munFileBackgroundMusicClose 
            Caption         =   "关闭音乐"
         End
      End
      Begin VB.Menu munFileOpen 
         Caption         =   "打开（&O)"
         Begin VB.Menu munFileOpenPicture 
            Caption         =   "图片"
         End
         Begin VB.Menu munFileOpenPPS 
            Caption         =   "幻灯片"
         End
         Begin VB.Menu munFileOpenMusic 
            Caption         =   "音频"
         End
         Begin VB.Menu munFileOpenVedio 
            Caption         =   "视频"
         End
         Begin VB.Menu munFileOpenText 
            Caption         =   "文本"
         End
      End
      Begin VB.Menu setSys 
         Caption         =   "系统设置"
         Begin VB.Menu SetBGS 
            Caption         =   "设置背景风格"
            Begin VB.Menu AddBGS 
               Caption         =   "增加背景风格"
            End
            Begin VB.Menu ChangeBGS 
               Caption         =   "改变背景风格"
            End
         End
         Begin VB.Menu setBD 
            Caption         =   "设置生日"
         End
      End
      Begin VB.Menu munFileExit 
         Caption         =   "退出（&E）"
      End
   End
   Begin VB.Menu munAboutOur 
      Caption         =   "关于我们（&W）"
      Begin VB.Menu munAboutOurLetter 
         Caption         =   "我们的情书"
         Begin VB.Menu OurLetterScan 
            Caption         =   "浏览"
            Begin VB.Menu textWay 
               Caption         =   "文本方式"
            End
            Begin VB.Menu PPS 
               Caption         =   "幻灯片"
               Begin VB.Menu SWFStyle 
                  Caption         =   "SWF样式"
               End
               Begin VB.Menu PPSStyle 
                  Caption         =   "PPS样式"
               End
            End
         End
         Begin VB.Menu OurLetterEdit 
            Caption         =   "编写"
         End
         Begin VB.Menu OurLetterAdd 
            Caption         =   "添加"
         End
         Begin VB.Menu OurLetterBackup 
            Caption         =   "备份"
         End
      End
      Begin VB.Menu OurPPS 
         Caption         =   "我们的幻灯片"
         Begin VB.Menu AddOurPPS 
            Caption         =   "添加"
         End
         Begin VB.Menu ScanOurPPS 
            Caption         =   "浏览"
         End
         Begin VB.Menu BackupPPS 
            Caption         =   "备份"
         End
      End
      Begin VB.Menu munAboutOurPicture 
         Caption         =   "我们的相片"
         Begin VB.Menu OurPictureScan 
            Caption         =   "浏览"
         End
         Begin VB.Menu OurPictureAdd 
            Caption         =   "添加"
         End
         Begin VB.Menu OurPictureBackup 
            Caption         =   "备份"
         End
      End
      Begin VB.Menu munAboutOurMusic 
         Caption         =   "我们的音乐"
         Begin VB.Menu OurMusicScan 
            Caption         =   "浏览"
         End
         Begin VB.Menu OurMusicAdd 
            Caption         =   "添加"
         End
      End
      Begin VB.Menu munAboutOurVedio 
         Caption         =   "我们的视频"
         Begin VB.Menu OurVedioScan 
            Caption         =   "浏览"
         End
         Begin VB.Menu OurVedioAdd 
            Caption         =   "添加"
         End
      End
      Begin VB.Menu munAboutOurCourse 
         Caption         =   "我们的历程"
         Begin VB.Menu OurCourseScan 
            Caption         =   "浏览"
         End
         Begin VB.Menu OurCourseEdit 
            Caption         =   "编辑"
         End
         Begin VB.Menu OurCourseAdd 
            Caption         =   "添加"
         End
         Begin VB.Menu OurCourseBackup 
            Caption         =   "备份"
         End
      End
   End
   Begin VB.Menu EXEProgram 
      Caption         =   "EXE程序"
      Begin VB.Menu ScanEXE 
         Caption         =   "浏览"
      End
      Begin VB.Menu AddEXE 
         Caption         =   "添加"
      End
      Begin VB.Menu BackupEXE 
         Caption         =   "备份"
      End
   End
   Begin VB.Menu AboutThis 
      Caption         =   "关于产品（&A）"
      Begin VB.Menu AboutThisNew 
         Caption         =   "产品信息"
      End
      Begin VB.Menu AboutThisExplain 
         Caption         =   "产品说明"
      End
   End
   Begin VB.Menu help 
      Caption         =   "帮助（&H）"
      Begin VB.Menu helpText 
         Caption         =   "帮助文档"
      End
      Begin VB.Menu helpTechSupport 
         Caption         =   "技术支持"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strShow As String                   '显示字符串
Dim awayBD As Integer
Dim birthday As Date                        '生日
Dim BDstr As String                         '生日字符串
Dim BGSstr As String                        '背景路径
Dim BGMstr As String                        '背景音乐路径
Dim myFSO As New FileSystemObject           '文件对象

'用于幻灯片放映
Private Declare Function icePub_open Lib "icePubDll.dll" (ByVal strPath As String) As Integer

'定义一个函数获取桌面路径，用于初始化打开路径Auto
Function GetDesktopPath() As String
Dim WshShell As Object, WScript As Object, oShellLink As Object
Set WshShell = CreateObject("WScript.Shell")
GetDesktopPath = WshShell.SpecialFolders("Desktop") '桌面路径Startup
End Function


'定义一个函数,根据资源类型RecType，从相应记录文件读取记录数据，若无该文件，则创建该文件；文件无记录，则返回默认路径
Function GetRec(ByVal RecType As String) As String
Dim tempStr As String
Select Case RecType

'背景资源
Case "BGS"
If Not myFSO.FileExists(App.Path & "\System\BGStyle.rec") Then
'如果该文件不存在
'tempStr=Auto
tempStr = App.Path & "\OtherRes\BGStyle\BGStyle1\BGStyle1.html"
'将该记录写入记录文件
Open App.Path & "\System\BGStyle.rec" For Output As 1
Print #1, tempStr
Close #1
Else
'文件存在，读取数据
Open App.Path & "\System\BGStyle.rec" For Input As 1
Line Input #1, tempStr
Close #1
If tempStr = "" Then
'如果为空，tempStr=Auto
tempStr = App.Path & "\OtherRes\BGStyle\BGStyle1\BGStyle1.html"
End If
'不为空
End If

'背景音乐
Case "BGM"
If Not myFSO.FileExists(App.Path & "\System\BGMusic.rec") Then
'如果该文件不存在
'tempStr=Auto
tempStr = App.Path & "\Music\BGMusic\BGMusic.mp3"
'将该记录写入记录文件
Open App.Path & "\System\BGMusic.rec" For Output As 1
Print #1, tempStr
Close #1
Else
'，文件存在，读取数据
Open App.Path & "\System\BGMusic.rec" For Input As 1
Line Input #1, tempStr
Close #1
If tempStr = "" Then
'如果为空，tempStr=Auto
tempStr = App.Path & "\Music\BGMusic\BGMusic.mp3"
End If
'不为空
End If
End Select
GetRec = tempStr
End Function

'定义一个函数，根据资源类型RecType，将相应内容RecData写入相应的资源文件
Function WriteRec(ByVal RecType As String, ByVal RecData As String) As Integer
Select Case RecType
'背景风格资源
Case "BGS"
Open App.Path & "\System\BGStyle.rec" For Output As 1
'背景音乐
Case "BGM"
Open App.Path & "\System\BGMusic.rec" For Output As 1
End Select
'写入数据
Print #1, RecData
WriteRec = 1
Close #1
End Function

Private Sub AboutThisExplain_Click()
Form8.Show 1, Form2
End Sub

Private Sub AboutThisNew_Click()
Form9.Show 1, Form2
End Sub

Private Sub AddBGM_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "添加背景音乐"
CommonDialog1.Filter = "音频文件(*.mp3;*.wav)|*.mp3;*.wav"
'默认路径为桌面
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\Music\BGMusic\"
MsgBox "文件添加成功！", , "系统提示"
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

'定义一个函数，产生风格文件夹名,并新建该文件

Function CreateStyleFolder() As String
Dim i As Integer
i = 1
Do While (myFSO.FolderExists(App.Path & "\OtherRes\BGStyle\BGStyle" & CStr(i)))
i = i + 1
Loop
CreateStyleFolder = "BGStyle" & CStr(i) & "\"
MkDir (App.Path & "\OtherRes\BGStyle\BGStyle" & CStr(i))
End Function

Private Sub AddBGS_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "添加背景风格"
CommonDialog1.Filter = "HTML文件和Flash(*.html;*.swf)|*.html;*,swf"
'默认路径为桌面
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\OtherRes\BGStyle\" & CreateStyleFolder()
MsgBox "文件添加成功！注意相关资源的链接。", , "系统提示"
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub AddEXE_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "添加EXE程序"
CommonDialog1.Filter = "EXE文件(*.exe)|*.exe"
'默认路径为桌面
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\OtherRes\PPS\OurPPS\"
MsgBox "文件添加成功！注意相关资源的链接。", , "系统提示"
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub AddOurPPS_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "添加我们的幻灯片"
CommonDialog1.Filter = "幻灯片文件(*.ppt;*.pptx;*.pps;*.ppsx)|*.ppt;*.pptx;*.pps;*.ppsx"
'默认路径为桌面
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\OtherRes\PPS\OurPPS\"
MsgBox "文件添加成功！注意相关资源的链接。", , "系统提示"
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub BackupEXE_Click()
'选择备份我们的幻灯片
Backup.Caption = "备份EXE"
Backup.Label3.Caption = App.Path & "\OtherRes\EXE"
'加载备份窗口
Load Backup
End Sub

Private Sub BackupPPS_Click()
'选择备份我们的幻灯片
Backup.Caption = "备份我们的幻灯片"
Backup.Label3.Caption = App.Path & "\OtherRes\PPS"
'加载备份窗口
Load Backup
End Sub

Private Sub ChangeBGM_Click()
Dim i As Integer
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "选择背景的音乐"
CommonDialog1.Filter = "音频文件(*.mp3;*.wav)|*.mp3;*.wav"
'默认路径为桌面
Form2.CommonDialog1.InitDir = App.Path & "\Music\BGMusic"
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
Form2.WindowsMediaPlayer1.Close
Form2.WindowsMediaPlayer1.URL = Form2.CommonDialog1.FileName
Form2.WindowsMediaPlayer1.Controls.play
i = WriteRec("BGM", Form2.CommonDialog1.FileName)
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub ChangeBGS_Click()
Dim i As Integer
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "选择背景风格"
CommonDialog1.Filter = "HTML文件(*.html)|*.html"
'默认路径为桌面
Form2.CommonDialog1.InitDir = App.Path & "\OtherRes\BGStyle\"
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
Form2.WebBrowser2.Navigate Form2.CommonDialog1.FileName
i = WriteRec("BGS", Form2.CommonDialog1.FileName)
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub



Private Sub Form_Load()
Form2.StatusBar1.Panels(1).Text = App.Path              '记录安装目录
'获取背景、背景音乐路径
BGSstr = GetRec("BGS")
BGMstr = GetRec("BGM")
'加载窗口2
Form2.Show
'读取生日数据
Open App.Path & "\System\Birthday.bd" For Input As 1
Line Input #1, BDstr
Close #1
birthday = DateValue(BDstr)
'计算距生日的天数
awayBD = CInt(DateDiff("d", Date, CDate(Month(birthday) & "-" & Day(birthday))))
If awayBD = 0 Then
strShow = "今天是您的生日！请点击此处！！！"
ElseIf awayBD < 0 Then
awayBD = awayBD + 365
strShow = "你的生日是：" & Month(birthday) & "月" & Day(birthday) & "日,据您生日还有" & CStr(awayBD) & "天"
Else
strShow = "你的生日是：" & Month(birthday) & "月" & Day(birthday) & "日,据您生日还有" & CStr(awayBD) & "天"
End If
Form2.StatusBar1.Panels(3).Text = strShow
'背景风格
Form2.WebBrowser2.Navigate BGSstr
'播放背景音乐
Form2.WindowsMediaPlayer1.URL = BGMstr
Form2.WindowsMediaPlayer1.Controls.play
Form2.Timer1.Interval = 1000
End Sub

Private Sub Form_Resize()
'窗口大小改变时
If Form2.ScaleHeight <> 0 Then
Form2.WebBrowser2.Height = Form2.ScaleHeight - Form2.StatusBar1.Height
Else
Form2.WebBrowser2.Height = Form2.ScaleHeight
End If
Form2.WebBrowser2.Width = Form2.ScaleWidth
Form2.Refresh    '刷新
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Object
For Each frm In Forms
Unload frm
Next
End Sub


Private Sub helpTechSupport_Click()
Form11.Show 1, Form2
End Sub

Private Sub helpText_Click()
Form10.Icon = LoadPicture(App.Path & "\" & "Ico\Help.ico")
Form10.Caption = "帮助文档"
Form10.RichTextBox1.LoadFile App.Path & "\" & "RTF\HelpText\HelpText.rtf"
Load Form10
Form10.Show 1, Form2
End Sub

Private Sub munFileBackgroundMusicClose_Click()
'关闭背景音乐
Form2.WindowsMediaPlayer1.Controls.stop
End Sub

Private Sub munFileBackgroundMusicOpen_Click()
'打开背景音乐
Form2.WindowsMediaPlayer1.Controls.play
End Sub

Private Sub munFileExit_Click()
Unload Form2
End Sub

Private Sub munFileOpenMusic_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'打开选择音频时,文件必须存在
CommonDialog1.Flags = cdlOFNFileMustExist
'打开的文件类型
CommonDialog1.Filter = "音频文件(*.mp3;*.wav)|*.mp3;*.wav"
'默认路径为桌面
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.DialogTitle = "打开"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
'文件路径不能为空
If Not CommonDialog1.FileName = "" Then    '所选不为空
Form4.Icon = LoadPicture(App.Path & "\" & "Ico\Music.ico")
'URL赋值
Form4.WindowsMediaPlayer1.URL = CommonDialog1.FileName
'播放所选项
Load Form4
Form4.Show 1, Form2
Form4.WindowsMediaPlayer1.Controls.play
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub munFileOpenPicture_Click()
'选择打开图片时，
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'打开选择图片时,文件必须存在
CommonDialog1.Flags = cdlOFNFileMustExist
'打开的文件类型
'默认路径为桌面
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.DialogTitle = "打开"
CommonDialog1.Filter = "图片(*.jpg;*.bmp;*.gif;*.ico)|*.jpg;*.bmp;*.gif;*.ico"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
'文件路径不能为空
If Not CommonDialog1.FileName = "" Then
'读取浏览图片
Form3.Image1.Picture = LoadPicture(CommonDialog1.FileName)
'在窗口3显示图片
Form3.Show 1, Form2
End If
GoTo sign:    '结束该事件
ErrHandler:   '如果单击了取消键，那么将窗口2显示
'改变窗口2的背景
CommonDialog1.FileName = ""
sign:
End Sub

Private Sub munFileOpenPPS_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'打开选择视频时,文件必须存在
CommonDialog1.Flags = cdlOFNFileMustExist
'打开的文件类型
CommonDialog1.Filter = "幻灯片(*.ppt;*.pptx;*.pps;*.ppsx)|*.ppt;*.pptx;*.pps;*ppsx"
'默认路径为桌面
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.DialogTitle = "打开"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
'文件路径不能为空
If Not CommonDialog1.FileName = "" Then    '所选不为空
Dim a2 As Integer
a2 = icePub_open(Form2.CommonDialog1.FileName)
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示

End Sub

Private Sub munFileOpenText_Click()
Dim FileName As String
With Form5
On Error GoTo ErrHandler
.CommonDialog1.DialogTitle = "打开"
.CommonDialog1.Filter = "RTF文件|*.rtf"
'默认路径为桌面
.CommonDialog1.InitDir = GetDesktopPath()
.CommonDialog1.FileName = ""
.CommonDialog1.ShowOpen
On Error GoTo 0
FileName = .CommonDialog1.FileName
'使用RichTextbox1控件的LoadFile方法加载所以定的文件到控件中提供编辑
.RichTextBox1.LoadFile FileName
.StatusBar1.Panels(1).Text = "文件：" & FileName
End With
Form5.Show
ErrHandler:
End Sub

Private Sub munFileOpenVedio_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'打开选择视频时,文件必须存在
CommonDialog1.Flags = cdlOFNFileMustExist
'打开的文件类型
CommonDialog1.Filter = "视频文件(*.wmv;*.avi;*.mp4)|*.wmv;*.avi;*.mp4"
'默认路径为桌面
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.DialogTitle = "打开"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
'文件路径不能为空
If Not CommonDialog1.FileName = "" Then    '所选不为空
Form4.Icon = LoadPicture(App.Path & "\" & "Ico\VedioLook.ico")          '注意此方法
Form4.Show
'URL赋值
Form4.WindowsMediaPlayer1.URL = CommonDialog1.FileName
'播放所选项
Form4.WindowsMediaPlayer1.Controls.play
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub OurCourseAdd_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "添加我们的历程"
CommonDialog1.Filter = "RTF文件(*.rtf)|*.rtf"
'默认路径为桌面
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\RTF\OurCourse\"
MsgBox "文件添加成功！", , "系统提示"
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub OurCourseBackup_Click()
'选择备份我们的情书
Backup.Caption = "备份我们的历程"
Backup.Label3.Caption = App.Path & "\RTF\OurCourse"
'加载备份窗口
Load Backup
End Sub

Private Sub OurCourseEdit_Click()
'设置保存路径
Form5.CommonDialog1.InitDir = App.Path & "\RTF\OurCourse"
'默认文件保存名
Form5.CommonDialog1.FileName = "历程_n"
Load Form5
Form5.Show
End Sub

Private Sub OurCourseScan_Click()
'对话框主题
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
Form2.CommonDialog1.DialogTitle = "我们的历程"
'文件路径
Form2.CommonDialog1.InitDir = App.Path & "\" & "RTF\OurCourse"
CommonDialog1.Filter = "RTF文件(*.rtf)|*.rtf"
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
Load Form5
If Not Form2.CommonDialog1.FileName = "" Then
Form5.RichTextBox1.LoadFile Form2.CommonDialog1.FileName
Form5.StatusBar1.Panels(1).Text = Form2.CommonDialog1.FileName
End If
Form5.Show
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub OurLetterAdd_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "添加我们的情书"
CommonDialog1.Filter = "RTF文件(*.rtf)|*.rtf"
'默认路径为桌面
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\RTF\OurLetter\"
MsgBox "文件添加成功！", , "系统提示"
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub OurLetterBackup_Click()
'选择备份我们的情书
Backup.Caption = "备份我们的情书"
Backup.Label3.Caption = App.Path & "\RTF\OurLetter"
'加载备份窗口
Load Backup
End Sub

Private Sub OurLetterEdit_Click()
'设置保存路径
Form5.CommonDialog1.InitDir = App.Path & "\RTF\OurLetter"
'默认文件保存名
Form5.CommonDialog1.FileName = "情书_n"
Load Form5
Form5.Show
End Sub

Private Sub OurMusicAdd_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "添加我们的音乐"
CommonDialog1.Filter = "音频文件(*.mp3;*.wav)|*.mp3;*.wav"
'默认路径为桌面
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\Music\OurMusic\"
MsgBox "文件添加成功！", , "系统提示"
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub OurMusicScan_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "我们的音乐"
'文件路径
Form2.CommonDialog1.InitDir = App.Path & "\" & "Music\OurMusic"
CommonDialog1.Flags = cdlOFNFileMustExist
'打开的文件类型
CommonDialog1.Filter = "音频文件(*.mp3;*.wav)|*.mp3;*.wav"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
'文件路径不能为空
If Not CommonDialog1.FileName = "" Then    '所选不为空
Form4.Icon = LoadPicture(App.Path & "\" & "Ico\Music.ico")
Form4.WindowsMediaPlayer1.URL = Form2.CommonDialog1.FileName
Load Form4
Form4.Show 1, Form2
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub OurPictureAdd_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "添加我们的相片"
CommonDialog1.Filter = "图片(*.jpg;*.bmp;*.gif;*.ico)|*.jpg;*.bmp;*.gif;*.ico"
CommonDialog1.FileName = ""
'默认路径为桌面
Form2.CommonDialog1.InitDir = GetDesktopPath()
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\Image\OurPicture\"
MsgBox "文件添加成功！", , "系统提示"
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub OurPictureBackup_Click()
'选择备份我们的情书
Backup.Caption = "备份我们的相片"
Backup.Label3.Caption = App.Path & "\Image\OurPicture"
'加载备份窗口
Load Backup
End Sub

Private Sub OurPictureScan_Click()
'选择打开图片时，
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'打开选择图片时,文件必须存在
CommonDialog1.Flags = cdlOFNFileMustExist
'打开的文件类型
'对话框主题
Form2.CommonDialog1.DialogTitle = "我们的相片"
'文件路径
Form2.CommonDialog1.InitDir = App.Path & "\" & "Image\OurPicture"
CommonDialog1.Filter = "图片(*.jpg;*.bmp;*.gif;*.ico)|*.jpg;*.bmp;*.gif;*.ico"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
'文件路径不能为空
If Not CommonDialog1.FileName = "" Then
'读取浏览图片
Form3.Image1.Picture = LoadPicture(CommonDialog1.FileName)
'在窗口3显示图片
Form3.Show 1, Form2
End If
GoTo sign:    '结束该事件
ErrHandler:   '如果单击了取消键，那么将窗口2显示
'改变窗口2的背景
CommonDialog1.FileName = ""
sign:
End Sub

Private Sub OurVedioAdd_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "添加我们的视频"
CommonDialog1.Filter = "视频文件(*.wmv;*.avi;*.mp4)|*.wmv;*.avi;*.mp4"
CommonDialog1.FileName = ""
'默认路径为桌面
Form2.CommonDialog1.InitDir = GetDesktopPath()
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\Vedio\OurVideo\"
MsgBox "文件添加成功！", , "系统提示"
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub OurVedioScan_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "我们的视频"
'文件路径
Form2.CommonDialog1.InitDir = App.Path & "\" & "Vedio\OurVedio"
CommonDialog1.Flags = cdlOFNFileMustExist
'打开的文件类型
CommonDialog1.Filter = "视频文件(*.wmv;*.avi;*.mp4)|*.wmv;*.avi;*.mp4"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
'文件路径不能为空
If Not CommonDialog1.FileName = "" Then    '所选不为空
Form4.Icon = LoadPicture(App.Path & "\" & "Ico\VedioLook.ico")
Form4.WindowsMediaPlayer1.URL = Form2.CommonDialog1.FileName
Load Form4
Form4.Show 1, Form2
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub PPSStyle_Click()
'情书浏览方式之文本方式
Dim a2 As Integer
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "我们的情书PPS样式"
'文件路径
Form2.CommonDialog1.InitDir = App.Path & "\OtherRes\PPS\OurLetter\"
CommonDialog1.Filter = "PPS文件(*.ppsx;*pps)|*.ppsx;*pps"
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
a2 = icePub_open(Form2.CommonDialog1.FileName)
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub ScanEXE_Click()
Dim a2 As Integer
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "打开EXE程序"
'文件路径
Form2.CommonDialog1.InitDir = App.Path & "\OtherRes\EXE\"
CommonDialog1.Filter = "EXE文件(*.exe)|*.exe"
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
a2 = icePub_open(Form2.CommonDialog1.FileName)
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub ScanOurPPS_Click()
Dim a2 As Integer
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "我们的幻灯片"
'文件路径
Form2.CommonDialog1.InitDir = App.Path & "\OtherRes\PPS\OurPPS\"
CommonDialog1.Filter = "PPS文件(*.ppt;*pptx;*.ppsx;*pps)|*.ppt;*pptx;*.ppsx;*pps"
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
a2 = icePub_open(Form2.CommonDialog1.FileName)
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub setBD_Click()
Load setBirthday
setBirthday.Show 1, Form2
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)
'读取生日数据
Open App.Path & "\System\Birthday.bd" For Input As 1
Line Input #1, BDstr
Close #1
birthday = DateValue(BDstr)
awayBD = CInt(DateDiff("d", Date, CDate(Month(birthday) & "-" & Day(birthday))))
If Panel.Index = 3 Then
If awayBD = 0 Then
Load Form14
Form14.Show
End If
End If
If Panel.Index = 2 Then
Form12.Show 1, Form2
End If
End Sub

Private Sub SWFStyle_Click()
Form2.WebBrowser2.Navigate App.Path & "\OtherRes\SWF\OurLetter\Style1\Style1.html"
End Sub

Private Sub textWay_Click()
'情书浏览方式之文本方式
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form2.CommonDialog1.DialogTitle = "我们的情书"
'文件路径
Form2.CommonDialog1.InitDir = App.Path & "\" & "RTF\OurLetter"
CommonDialog1.Filter = "RTF文件(*.rtf)|*.rtf"
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
Load Form5
If Not Form2.CommonDialog1.FileName = "" Then
Form5.RichTextBox1.LoadFile Form2.CommonDialog1.FileName
Form5.StatusBar1.Panels(1).Text = Form2.CommonDialog1.FileName
End If
Form5.Show
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub Timer1_Timer()
Form2.StatusBar1.Panels(2).Text = "当前时间：" & Date & "  " & Time
End Sub

