VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "MyBrowse"
   ClientHeight    =   5655
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   8445
   Icon            =   "MyBrowse.frx":0000
   LinkTopic       =   "Form1"
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
      Caption         =   "����ת����GO��"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�˵�"
      Begin VB.Menu Exit 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu MainPage 
      Caption         =   "��ҳ"
   End
   Begin VB.Menu GoBack 
      Caption         =   "��һҳ"
      Enabled         =   0   'False
   End
   Begin VB.Menu Refresh 
      Caption         =   "ˢ��"
   End
   Begin VB.Menu Navigation 
      Caption         =   "����"
      Begin VB.Menu baidu 
         Caption         =   "�ٶ�"
      End
      Begin VB.Menu Google 
         Caption         =   "�ȸ�"
      End
      Begin VB.Menu Hao123 
         Caption         =   "hao123"
      End
      Begin VB.Menu youSchool 
         Caption         =   "¹ɽѧԺ"
      End
   End
   Begin VB.Menu GoForward 
      Caption         =   "��һҳ"
      Enabled         =   0   'False
   End
   Begin VB.Menu Explain 
      Caption         =   "˵��"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub baidu_Click()
Form1.WebBrowser1.Refresh
Form1.WebBrowser1.Navigate "http://www.baidu.com/"
End Sub

Private Sub Combo1_Change()
If Not Form1.Combo1.Text = "" Then
Form1.Command1.Enabled = True
Else
Form1.Command1.Enabled = False
End If
Form1.WebBrowser1.Refresh
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim I As Long
    Dim existed As Boolean
    If KeyCode = 13 Then
    If Left(Combo1.Text, 7) <> "http://" Then   '/���������ַ�����ԡ�http://����ͷ���Զ����
    Combo1.Text = "http://" + Combo1.Text
    End If
    WebBrowser1.Navigate Combo1.Text '/ URL��ַ���������վ��ַ
    For I = 0 To Combo1.ListCount - 1
    If Combo1.List(I) = Combo1.Text Then
    existed = True
    Exit For
    Else
    existed = False
    End If
    Next
    If Not existed Then
    Combo1.AddItem (Combo1.Text) '/ ��������µ���վ���Զ�����
    End If
    End If
End Sub

Private Sub Command1_Click()
Form1.WebBrowser1.Navigate Form1.Combo1.Text
End Sub

Private Sub Exit_Click()
Unload Form1
End Sub

Private Sub Explain_Click()
Load Form2
Form2.Show 1, Form1
End Sub

Private Sub Form_Load()
Form1.WebBrowser1.Navigate App.Path & "\OtherRes\HTML\Birthday\HBDMain.html"
End Sub

Private Sub Form_Resize()
Form1.Line1.X2 = Form1.ScaleWidth
Form1.Line2.X2 = Form1.ScaleWidth
Form1.WebBrowser1.Width = Form1.ScaleWidth
If Form1.ScaleWidth > 1800 Then
Form1.Combo1.Width = Form1.ScaleWidth - 1800
Else
Form1.Combo1.Width = 0
End If
If Form1.ScaleHeight > 1220 Then
Form1.WebBrowser1.Height = Form1.ScaleHeight - 1220
Else
Form1.WebBrowser1.Height = 0
End If
End Sub

Private Sub GoBack_Click()
Form1.WebBrowser1.GoBack
End Sub

Private Sub GoForward_Click()
Form1.WebBrowser1.GoForward
End Sub

Private Sub Google_Click()
Form1.WebBrowser1.Navigate "http://www.google.com.hk/"
End Sub

Private Sub Hao123_Click()
Form1.WebBrowser1.Navigate "http://www.hao123.com/"
End Sub

Private Sub MainPage_Click()
Form1.WebBrowser1.Navigate App.Path & "\OtherRes\HTML\Birthday\HBDMain.html"
End Sub

Private Sub Refresh_Click()
Form1.WebBrowser1.Refresh
End Sub

Private Sub WebBrowser1_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
'�ж�ǰ�������Ƿ����
If (Command = CSC_NAVIGATEBACK) Then
Form1.GoBack.Enabled = Enable
End If
If (Command = CSC_NAVIGATEFORWARD) Then
Form1.GoForward.Enabled = Enable
End If
End Sub

Private Sub WebBrowser1_DownloadBegin()
Form1.StatusBar1.SimpleText = "������..."
End Sub

Private Sub WebBrowser1_DownloadComplete()
Form1.StatusBar1.SimpleText = "������ɣ�"
End Sub

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
    Dim I As Long
    Dim existed As Boolean
    Combo1.Text = WebBrowser1.LocationURL
    Form1.StatusBar1.SimpleText = Form1.WebBrowser1.LocationName
    For I = 0 To Combo1.ListCount - 1
    If Combo1.List(I) = Combo1.Text Then
    existed = True
    Exit For
    Else
    existed = False
    End If
    Next
    If Not existed Then
    Combo1.AddItem (Combo1.Text) '/ ��������µ���վ���Զ�����
    End If
End Sub


Private Sub youSchool_Click()
Form1.WebBrowser1.Navigate "http://www.lzls.gxut.edu.cn/"
End Sub
