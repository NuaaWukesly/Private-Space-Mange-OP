VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form5 
   Caption         =   "文本编辑"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6840
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   4830
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   926
      ButtonWidth     =   820
      ButtonHeight    =   767
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "新建"
            Object.Tag             =   ""
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "打开"
            Object.Tag             =   ""
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "保存"
            Object.Tag             =   ""
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "打印"
            Object.Tag             =   ""
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cut"
            Object.ToolTipText     =   "剪切"
            Object.Tag             =   ""
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Copy"
            Object.ToolTipText     =   "复制"
            Object.Tag             =   ""
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Paste"
            Object.ToolTipText     =   "粘贴"
            Object.Tag             =   ""
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Bold"
            Object.ToolTipText     =   "粗体字"
            Object.Tag             =   ""
            ImageKey        =   "Bold"
            Style           =   1
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Italic"
            Object.ToolTipText     =   "斜体字"
            Object.Tag             =   ""
            ImageKey        =   "Italic"
            Style           =   1
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "UnderLine"
            Object.ToolTipText     =   "下划线"
            Object.Tag             =   ""
            ImageKey        =   "UnderLine"
            Style           =   1
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   4
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Font"
            Object.ToolTipText     =   "字体"
            Object.Tag             =   ""
            ImageKey        =   "Font"
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Paragraph"
            Object.ToolTipText     =   "段落"
            Object.Tag             =   ""
            ImageKey        =   "Paragraph"
         EndProperty
      EndProperty
      Begin ComctlLib.ImageList ImageList1 
         Left            =   1440
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   23
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   12
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form5.frx":1176A
               Key             =   "New"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form5.frx":11DE4
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form5.frx":1245E
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form5.frx":12AD8
               Key             =   "Print"
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form5.frx":13152
               Key             =   "Cut"
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form5.frx":137CC
               Key             =   "Copy"
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form5.frx":13E46
               Key             =   "Paste"
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form5.frx":144C0
               Key             =   "Bold"
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form5.frx":1A97A
               Key             =   "Italic"
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form5.frx":1E2E4
               Key             =   "UnderLine"
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form5.frx":27906
               Key             =   "Font"
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Form5.frx":39058
               Key             =   "Paragraph"
            EndProperty
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4095
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7223
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form5.frx":4B40A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4455
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Adobe 仿宋 Std R"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'私有变量FileName保存当前打开文件的路径
Private FileName As String
'私有变量i用作循环计数
Private i As Long

Private Sub Form_Load()
Dim ctr As Object
Load Form6  '加载，但暂不显示
Load Form7  '加载，但暂不显示
On Error Resume Next
For Each ctr In Controls
ctr.Default = False '窗口5中所有控件的Default属性设置为False
Next
End Sub

Private Sub Form_Resize()
'Toolbar1.Width = Form5.ScaleWidth
'StatusBar1.Width = Form5.ScaleWidth
RichTextBox1.Width = Form5.ScaleWidth
If Not Form5.ScaleHeight = 0 Then
RichTextBox1.Height = Form5.ScaleHeight - Toolbar1.Height - StatusBar1.Height
End If
Form5.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form6
Unload Form7
End Sub

Private Sub RichTextBox1_GotFocus() '获得焦点时
On Error Resume Next
Dim ctr As Object
For Each ctr In Controls    '设置所有控件Tab键无效
ctr.TabStop = False
Next
End Sub

Private Sub RichTextBox1_LostFocus()    '失去焦点时
On Error Resume Next
Dim ctr As Object
For Each ctr In Controls
ctr.TabStop = True
Next
End Sub

Private Sub RichTextBox1_SelChange()
If RichTextBox1.SelBold Then
Toolbar1.Buttons("Bold").Value = tbrPressed     '按键处于按下状态
Else
Toolbar1.Buttons("Bold").Value = tbrUnpressed
End If
If RichTextBox1.SelItalic Then
Toolbar1.Buttons("Italic").Value = tbrPressed
Else
Toolbar1.Buttons("Italic").Value = tbrUnpressed
End If
If RichTextBox1.SelUnderline Then
Toolbar1.Buttons("UnderLine").Value = tbrPressed
Else
Toolbar1.Buttons("UnderLine").Value = tbrUnpressed
End If
End Sub

'当用户按下Toolbar1控件的某个按钮时，就会触发该Toolbar1控件所提供的ButtonClick事件
'ButtonClick事件所传递的Button参数代表用户所按下的按键，可以通过检查该控件的Key属性（关键字）
'可以区分用户按下去的是哪一个按键，从而执行相应操作
'也可以检查按钮的Index(索引)属性来判断用户按下的是哪一个按键，但是由于索引太多不好记
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
'这里使用一个很大的Select Case 结构来执行相应操作
Select Case Button.Key            '根据关键字判断
Case "New"                        '用户按下新建一个RTF文件
'由于前面已把CommonDialog1控件的CancelError属性设置为True，这样用户选择“取消”时就会发出一个错误
'从而据此可以判断用户是否选择了“取消”
On Error GoTo ErrHandler            '接收用户选择“取消”时发出的错误
CommonDialog1.DialogTitle = "新建"
CommonDialog1.Filter = "RTF文件|*.rtf"
'以保存文件的方式显示通用对话框
CommonDialog1.ShowSave
'取消对错误的捕捉，以使因为其他原因所引发的错误可以被Visual Basic 集成开发环境所捕获
On Error GoTo 0
'把用户输入的文件名存到变量FileName中
FileName = CommonDialog1.FileName
'如果所给的文件名已存在，弹出一个对话框询问用户是否覆盖
If Dir(FileName) <> "" Then
If MsgBox("是否覆盖已有文件" & FileName & "?", vbYesNo) = vbYes Then '选择覆盖
Kill FileName       '删除原有文件
Else
Exit Sub            '退出该事件
End If
End If
'清空RickTextBox1控件中的现有内容
RichTextBox1.TextRTF = ""
RichTextBox1.SaveFile FileName
'是以状态条反应当前打开的文件
StatusBar1.Panels(1).Text = "文件：" & FileName

Case "Open"         '用户按下一个打开文件按钮
On Error GoTo ErrHandler
CommonDialog1.DialogTitle = "打开"
CommonDialog1.Filter = "RTF文件|*.rtf"
CommonDialog1.ShowOpen
On Error GoTo 0
FileName = CommonDialog1.FileName
'使用RichTextbox1控件的LoadFile方法加载所以定的文件到控件中提供编辑
RichTextBox1.LoadFile FileName
StatusBar1.Panels(1).Text = "文件：" & FileName

Case "Save"         '用户按下了保存文件按钮
'如果当前还没有指定保存的文件名的话，就弹出一个保存文件的通用对话框
If FileName = "" Then
On Error GoTo ErrHandler
CommonDialog1.DialogTitle = "另存为"
CommonDialog1.Filter = "RTF文件|*.rtf"
CommonDialog1.ShowSave
On Error GoTo 0
FileName = CommonDialog1.FileName
If Dir(FileName) <> "" Then
If MsgBox("是否覆盖已有文件" & FileName & "?", vbYesNo, "Private Place Manage OP1.0") = vbYes Then
Kill FileName
Else
Exit Sub
End If
End If
End If
RichTextBox1.SaveFile FileName
StatusBar1.Panels(1).Text = "文件：" & FileName

Case "Print"       '用户选择打印
'弹出一个标准的打印对话框
On Error GoTo ErrHandler
CommonDialog1.DialogTitle = "打印"
CommonDialog1.Flags = cdlPDReturnDC + cdlPDNoPageNums
'如果没有选定文本，则打印RichTextBox1控件的所有内容，否则打印指定文本
If RichTextBox1.SelLength = 0 Then
CommonDialog1.Flags = CommonDialog1.Flags + cdlPDAllPages
Else
CommonDialog1.Flags = CommonDialog1.Filter + cdlPDSelection
End If
CommonDialog1.ShowPrinter
On Error GoTo 0
'使用RichTextBox1控件的SelPrint方法进行打印输出，所给的参数式打印机的设备句柄
RichTextBox1.SelPrint CommonDialog1.hDC

Case "Cut"          '用户选择剪切
'清空剪切板已有内容
Clipboard.Clear
'把当前选定文本复制到剪贴板上
Clipboard.SetText RichTextBox1.SelText
'删除选定文本
RichTextBox1.SelText = ""

Case "Copy"         '用户选择复制
'与剪切类似，只是不用删除选定文本
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText

Case "Paste"        '用户选择粘贴
RichTextBox1.SelText = Clipboard.GetText

Case "Bold"         '用户选择粗体
RichTextBox1.SelBold = Button.Value

Case "Italic"       '选择斜体
RichTextBox1.SelItalic = Button.Value

Case "UnderLine"    '下划线
RichTextBox1.SelUnderline = Button.Value

Case "Font"          '字体
'把Form6（即格式对话框）的标题改为“字体”
Form6.Caption = "字体"
'设置格式对话框中选择的选项卡是“字体”选项卡，这里是TabStrip控件的Tabs集合中的
'第一个选项卡（即第一个Tab对象）
Set Form6.TabStrip1.selectedItem = Form6.TabStrip1.Tabs(1)
'显示form7(初始化字体进度),并把Form5作为它的父窗口
Form7.Show , Form5
'把Progressbar1控件的Max值设置为比当前字体数少1，这样使得为进程的value属性赋值是无需换算
Form7.ProgressBar1.Max = Screen.FontCount - 1
'清除form6的字体组合框中已有的内容，然后检查系统的可用字体名称列表，并将它们加到字体组合框
Form6.Combo1.Clear          '字体组合框Combol1
For i = 0 To Screen.FontCount - 1
Form6.Combo1.AddItem Screen.Fonts(i)
Form7.ProgressBar1.Value = i
DoEvents
'避免系统死循环时系统失去响应
Next
'调用SetCurValue过程，使Form6中的控件能够反映当前值
SetCurValue
'隐藏Form7，以模式对话框的方式，显示form6，并把form1作为父窗口
Form7.Hide
Form6.Show 1, Form5
'调用SetTextValue，使form6中各相关控件的当前值来设置当前文本的格式
SetTextValue

Case "Paragraph"        '段落
'与上相似
Form6.Caption = "段落"
Set Form6.TabStrip1.selectedItem = Form6.TabStrip1.Tabs(2)
Form7.Show , Form5
Form7.ProgressBar1.Max = Screen.FontCount - 1
Form6.Combo1.Clear
For i = 0 To Screen.FontCount - 1
Form6.Combo1.AddItem Screen.Fonts(i)
Form7.ProgressBar1.Value = i
DoEvents
Next
SetCurValue
Form7.Hide
Form6.Show 1, Form5
SetTextValue
End Select
Exit Sub
ErrHandler:
Exit Sub
End Sub

Public Sub SetCurValue()
With Form6
If Not IsNull(RichTextBox1.SelIndent) Then
'selIndent返回当前选定段落的缩进量
.MaskEdBox2(0).Text = RichTextBox1.SelIndent
Else
.MaskEdBox2(0).Text = ""
End If
If Not IsNull(RichTextBox1.SelRightIndent) Then
'返回右缩进量
.MaskEdBox2(1).Text = RichTextBox1.SelRightIndent
Else
.MaskEdBox2(1).Text = ""
End If
If Not IsNull(RichTextBox1.SelHangingIndent) Then
'悬挂缩进量
.MaskEdBox2(2).Text = RichTextBox1.SelHangingIndent
Else
.MaskEdBox2(2).Text = ""
End If
If Not IsNull(RichTextBox1.SelBullet) Then
'当选定文本所在段落前带有项目符号的返回true
If RichTextBox1.SelBullet Then
.Combo5.ListIndex = 0
'项目符号
Else
.Combo5.ListIndex = 1
End If
Else
.Combo5.Text = ""
End If
If Not IsNull(RichTextBox1.SelFontName) Then
For i = 0 To .Combo1.ListCount - 1
If .Combo1.List(i) = RichTextBox1.SelFontName Then
.Combo1.ListIndex = i
End If
Next
Else
.Combo1.Text = ""
End If
If Not IsNull(RichTextBox1.SelFontSize) Then
'使用Format函数来生成格式化的字符串，并把他赋值给maskEdbox1的text属性
'这时如果直接复制可能不匹配
.MaskEdBox1.Text = Format(RichTextBox1.SelFontSize, "00磅")
Else
.MaskEdBox1.Text = "__磅"
End If
If Not IsNull(RichTextBox1.SelColor) Then
.Label10.BackStyle = 1
.Label10.BackColor = RichTextBox1.SelColor
Else
'如果有多种颜色，则设置为透明
.Label10.BackStyle = 0
End If
End With
End Sub

Public Sub SetTextValue()
'下面的代码将检测form6中各个控件的值与当前选定的文本的该属性是否一样
'若不一样则根据该控件的值修改当前选定文本或它所在段落的格式属性的值
With Form6
If Val(.MaskEdBox2(0).Text) > 0 Then
If Val(.MaskEdBox2(0).Text) <> RichTextBox1.SelIndent Then
Select Case .Combo2(0).Text
Case "磅"
RichTextBox1.SelIndent = Val(.MaskEdBox2(0).Text) * 20
Case "厘米"
RichTextBox1.SelIndent = Val(.MaskEdBox2(0)) / 2.54 * 72 * 20
End Select
End If
End If
If Val(.MaskEdBox2(1).Text) > 0 Then
If Val(.MaskEdBox2(1).Text) <> RichTextBox1.SelRightIndent Then
Select Case .Combo2(1).Text
Case "磅"
RichTextBox1.SelRightIndent = Val(.MaskEdBox2(1).Text) * 20
Case "厘米"
RichTextBox1.SelRightIndent = Val(.MaskEdBox2(1)) / 2.54 * 72 * 20
End Select
End If
End If
If Val(.MaskEdBox2(2).Text) > 0 Then
If Val(.MaskEdBox2(2).Text) <> RichTextBox1.SelHangingIndent Then
Select Case .Combo2(2).Text
Case "磅"
RichTextBox1.SelHangingIndent = Val(.MaskEdBox2(2).Text) * 207
Case "厘米"
RichTextBox1.SelHangingIndent = Val(.MaskEdBox2(2)) / 2.54 * 72 * 20
End Select
End If
End If
If Val(.MaskEdBox1.ClipText) > 0 Then
If Val(.MaskEdBox1.ClipText) <> RichTextBox1.SelFontSize Then
'cliptext返回控件中排除输入掩码之后的原义字符串
RichTextBox1.SelFontSize = Val(.MaskEdBox1.ClipText)
End If
End If
If Not .Label10.BackStyle = 0 Then
If .Label10.BackColor <> RichTextBox1.BackColor Then
RichTextBox1.SelColor = .Label10.BackColor
End If
End If
If .Combo1.Text <> "" Then
If .Combo1.Text <> RichTextBox1.SelFontName Then
RichTextBox1.SelFontName = .Combo1.Text
End If
End If
If .Combo5.Text <> "" Then
If .Combo5.Text = "True" Then
RichTextBox1.SelBullet = True
ElseIf .Combo5.Text = "False" Then
RichTextBox1.SelBullet = False
End If
End If
End With
End Sub












