VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MyBrowse 
   Caption         =   "MyBrowse"
   ClientHeight    =   5655
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   8445
   Icon            =   "MyBrowse.frx":0000
   LinkTopic       =   "MyBrowse"
   ScaleHeight     =   5655
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4335
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   8415
      ExtentX         =   14843
      ExtentY         =   7646
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "《《转到（GO）"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "MyBrowse.frx":4F32
      Left            =   1800
      List            =   "MyBrowse.frx":4F34
      TabIndex        =   1
      Top             =   120
      Width           =   6615
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   873
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   8400
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8400
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Menu 
      Caption         =   "菜单"
      Begin VB.Menu ChangeTopic 
         Caption         =   "更换浏览器主题"
      End
      Begin VB.Menu Exit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu MainPage 
      Caption         =   "主页"
   End
   Begin VB.Menu GoBack 
      Caption         =   "上一页"
      Enabled         =   0   'False
   End
   Begin VB.Menu Refresh 
      Caption         =   "刷新"
   End
   Begin VB.Menu Navigation 
      Caption         =   "导航"
      Begin VB.Menu baidu 
         Caption         =   "百度"
      End
      Begin VB.Menu Google 
         Caption         =   "谷歌"
      End
      Begin VB.Menu Hao123 
         Caption         =   "hao123"
      End
      Begin VB.Menu youSchool 
         Caption         =   "鹿山学院"
      End
   End
   Begin VB.Menu GoForward 
      Caption         =   "下一页"
      Enabled         =   0   'False
   End
   Begin VB.Menu Menuxplain 
      Caption         =   "说明"
   End
End
Attribute VB_Name = "MyBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub baidu_Click()
WebBrowser1.Navigate "http://www.baidu.com/"
End Sub

Private Sub ChangeTopic_Click()
Unload MyBrowse
Load Form13
Form13.Show
End Sub

Private Sub Combo1_Change()
If Not MyBrowse.Combo1.Text = "" Then
MyBrowse.Command1.Enabled = True
Else
MyBrowse.Command1.Enabled = False
End If
'MyBrowse.WebBrowser1.Refresh
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long
    Dim existed As Boolean
    If KeyCode = 13 Then
    If Left(Combo1.Text, 7) <> "http://" Then   '/如果输入网址不是以“http://”开头则自动添加
    Combo1.Text = "http://" + Combo1.Text
    End If
    WebBrowser1.Navigate Combo1.Text '/ URL地址栏保存的网站地址
    For i = 0 To Combo1.ListCount - 1
    If Combo1.List(i) = Combo1.Text Then
    existed = True
    Exit For
    Else
    existed = False
    End If
    Next
    If Not existed Then
    Combo1.AddItem (Combo1.Text) '/ 如果输入新的网站则自动保存
    End If
    End If
End Sub

Private Sub Command1_Click()
MyBrowse.WebBrowser1.Navigate Combo1.Text
End Sub

Private Sub Exit_Click()
Form2.WindowsMediaPlayer1.Controls.play
Unload MyBrowse
End Sub

Private Sub Form_Load()
Form2.WindowsMediaPlayer1.Controls.stop
MyBrowse.WebBrowser1.Navigate App.Path & "\OtherRes\HTML\Birthday\HBDMain.html"
MyBrowse.Show
End Sub

Private Sub Form_Resize()
MyBrowse.Line1.X2 = MyBrowse.ScaleWidth
MyBrowse.Line2.X2 = MyBrowse.ScaleWidth
MyBrowse.WebBrowser1.Width = MyBrowse.ScaleWidth
If MyBrowse.ScaleWidth > 1800 Then
MyBrowse.Combo1.Width = MyBrowse.ScaleWidth - 1800
Else
MyBrowse.Combo1.Width = 0
End If
If MyBrowse.ScaleHeight > 1220 Then
MyBrowse.WebBrowser1.Height = MyBrowse.ScaleHeight - 1220
Else
MyBrowse.WebBrowser1.Height = 0
End If
End Sub

Private Sub GoBack_Click()
MyBrowse.WebBrowser1.GoBack
End Sub

Private Sub GoForward_Click()
MyBrowse.WebBrowser1.GoForward
End Sub

Private Sub Google_Click()
MyBrowse.WebBrowser1.Navigate "http://www.google.com.hk/"
End Sub

Private Sub Hao123_Click()
MyBrowse.WebBrowser1.Navigate "http://www.hao123.com/"
End Sub

Private Sub MainPage_Click()
MyBrowse.WebBrowser1.Navigate App.Path & "\OtherRes\HTML\Birthday\HBDMain.html"
End Sub

Private Sub Menuxplain_Click()
Load Explain
Explain.Show 1, MyBrowse
End Sub

Private Sub Refresh_Click()
WebBrowser1.Refresh
End Sub

Private Sub WebBrowser1_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
'判断前进后退是否可用
If (Command = CSC_NAVIGATEBACK) Then
MyBrowse.GoBack.Enabled = Enable
End If
If (Command = CSC_NAVIGATEFORWARD) Then
MyBrowse.GoForward.Enabled = Enable
End If
End Sub

Private Sub WebBrowser1_DownloadBegin()
MyBrowse.StatusBar1.SimpleText = "载入中..."
End Sub

Private Sub WebBrowser1_DownloadComplete()
MyBrowse.StatusBar1.SimpleText = "载入完成！"
End Sub

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
    Dim i As Long
    Dim existed As Boolean
    Combo1.Text = WebBrowser1.LocationURL
    MyBrowse.StatusBar1.SimpleText = MyBrowse.WebBrowser1.LocationName
    For i = 0 To Combo1.ListCount - 1
    If Combo1.List(i) = Combo1.Text Then
    existed = True
    Exit For
    Else
    existed = False
    End If
    Next
    If Not existed Then
    Combo1.AddItem (Combo1.Text) '/ 如果输入新的网站则自动保存
    End If
End Sub


Private Sub youSchool_Click()
MyBrowse.WebBrowser1.Navigate "http://www.lzls.gxut.edu.cn/"
End Sub
