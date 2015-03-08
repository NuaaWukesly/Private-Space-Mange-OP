VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form13 
   Caption         =   "Happy Birthday !"
   ClientHeight    =   5235
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   8430
   ForeColor       =   &H0000FFFF&
   Icon            =   "Form13.frx":0000
   LinkTopic       =   "Form13"
   ScaleHeight     =   5235
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   8415
      ExtentX         =   14843
      ExtentY         =   8493
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
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
      Location        =   "http:///"
   End
   Begin VB.Timer Timer2 
      Left            =   240
      Top             =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "生日快乐！Happy Birthday！"
      BeginProperty Font 
         Name            =   "华文楷体"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4800
      Width           =   8415
   End
   Begin VB.Menu Menu 
      Caption         =   "菜单"
      Begin VB.Menu ChangeTopic 
         Caption         =   "更换浏览器主题"
      End
      Begin VB.Menu ThemeExit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu returnMain 
      Caption         =   "GoHome(返回主页)"
   End
   Begin VB.Menu goback 
      Caption         =   "GoBack(后退)"
      Enabled         =   0   'False
   End
   Begin VB.Menu GoForward 
      Caption         =   "GoForward(前进)"
      Enabled         =   0   'False
   End
   Begin VB.Menu BirthDayPage 
      Caption         =   "生日主页"
      Begin VB.Menu BirthDayPageExplain 
         Caption         =   "说明"
      End
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private R As Integer, G As Integer, B As Integer, i As Integer, j As Integer, signal As Integer
Private change As Boolean

Private Sub BirthDayPageExplain_Click()
Load Explain
Explain.Show 1, Form13
End Sub

Private Sub ChangeTopic_Click()
Unload Form13
Load MyBrowse
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mblnRunning = False
    Unload Me
End Sub

Private Sub Form_Load()
'停止窗口2背景音乐
Form2.WindowsMediaPlayer1.Controls.stop
i = 0
j = 1
signal = 0
R = 3
G = 11
B = 23
change = True
Form13.WebBrowser1.Navigate App.Path & "\OtherRes\HTML\Birthday\HBDMain.html"
Form13.WebBrowser1.Width = Form13.ScaleWidth
Form13.Label1.Width = Form13.ScaleWidth
Form13.WebBrowser1.Height = Form13.ScaleHeight
Form13.Timer2.Interval = 100
Label1.Left = Form13.ScaleWidth
End Sub

Private Sub Form_Resize()
'窗口大小改变时
If Form13.ScaleHeight <> 0 Then
Form13.WebBrowser1.Height = Form13.ScaleHeight - Form13.Label1.Height
Else
Form13.WebBrowser1.Height = Form13.ScaleHeight
End If
Form13.WebBrowser1.Width = Form13.ScaleWidth
Form13.Label1.Width = Form13.ScaleWidth
Form13.Label1.Top = Form13.WebBrowser1.Height
Form13.Refresh    '刷新
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form13.WebBrowser1.stop
Form13.Timer2.Interval = 1000
End Sub

Private Sub GoBack_Click()
Form13.WebBrowser1.GoBack
End Sub

Private Sub GoForward_Click()
Form13.WebBrowser1.GoForward
End Sub

Private Sub returnMain_Click()
Form13.WebBrowser1.Navigate App.Path & "\OtherRes\HTML\Birthday\HBDMain.html"
End Sub

Private Sub ThemeExit_Click()
Form2.WindowsMediaPlayer1.Controls.play
Unload Form13
End Sub

Private Sub Timer2_Timer()
Label1.Left = Label1.Left - 60
If Label1.Left < -Label1.Width - 200 Then
Label1.Left = Form13.ScaleWidth + 130
End If
If R > 255 Then
R = R - 255
End If
R = R + 11
If G > 255 Then
G = G - 255
End If
G = G + 47
If B > 255 Then
B = B - 255
End If
B = B + 13
If R > 255 Then
R = R - 255
End If
If G > 255 Then
G = G - 255
End If
If B > 255 Then
B = B - 255
End If
Label1.ForeColor = RGB(R, G, B)
End Sub

Private Sub WebBrowser1_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
If (Command = CSC_NAVIGATEBACK) Then
Form13.GoBack.Enabled = Enable
End If
If (Command = CSC_NAVIGATEFORWARD) Then
Form13.GoForward.Enabled = Enable
End If
End Sub

