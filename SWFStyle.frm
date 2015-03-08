VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form14 
   Caption         =   "我的情书SWFStyle"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6630
   Icon            =   "SWFStyle.frx":0000
   LinkTopic       =   "Form14"
   ScaleHeight     =   5475
   ScaleWidth      =   6630
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      ExtentX         =   11668
      ExtentY         =   9763
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'情书浏览方式之SWF方式
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'对话框主题
Form14.CommonDialog1.DialogTitle = "我们的情书SWF"
'文件路径
Form14.CommonDialog1.InitDir = App.Path & "\OtherRes\SWF\OurLetter\"
CommonDialog1.Filter = "HTML文件(*.html)|*.html"
CommonDialog1.FileName = ""
Form14.CommonDialog1.ShowOpen
If Not Form14.CommonDialog1.FileName = "" Then
Form14.WebBrowser1.Navigate Form14.CommonDialog1.FileName
End If
ErrHandler:   '如果单击了取消键，那么将窗口2显示
End Sub

Private Sub Form_Resize()
Form14.WebBrowser1.Height = Form14.ScaleHeight
Form14.WebBrowser1.Width = Form14.ScaleWidth
End Sub
