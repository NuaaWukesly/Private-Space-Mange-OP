VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form5 
   Caption         =   "�ı��༭"
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
            Object.ToolTipText     =   "�½�"
            Object.Tag             =   ""
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "��"
            Object.Tag             =   ""
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "����"
            Object.Tag             =   ""
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "��ӡ"
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
            Object.ToolTipText     =   "����"
            Object.Tag             =   ""
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Copy"
            Object.ToolTipText     =   "����"
            Object.Tag             =   ""
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Paste"
            Object.ToolTipText     =   "ճ��"
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
            Object.ToolTipText     =   "������"
            Object.Tag             =   ""
            ImageKey        =   "Bold"
            Style           =   1
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Italic"
            Object.ToolTipText     =   "б����"
            Object.Tag             =   ""
            ImageKey        =   "Italic"
            Style           =   1
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "UnderLine"
            Object.ToolTipText     =   "�»���"
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
            Object.ToolTipText     =   "����"
            Object.Tag             =   ""
            ImageKey        =   "Font"
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Paragraph"
            Object.ToolTipText     =   "����"
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
         Name            =   "Adobe ���� Std R"
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
'˽�б���FileName���浱ǰ���ļ���·��
Private FileName As String
'˽�б���i����ѭ������
Private i As Long

Private Sub Form_Load()
Dim ctr As Object
Load Form6  '���أ����ݲ���ʾ
Load Form7  '���أ����ݲ���ʾ
On Error Resume Next
For Each ctr In Controls
ctr.Default = False '����5�����пؼ���Default��������ΪFalse
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

Private Sub RichTextBox1_GotFocus() '��ý���ʱ
On Error Resume Next
Dim ctr As Object
For Each ctr In Controls    '�������пؼ�Tab����Ч
ctr.TabStop = False
Next
End Sub

Private Sub RichTextBox1_LostFocus()    'ʧȥ����ʱ
On Error Resume Next
Dim ctr As Object
For Each ctr In Controls
ctr.TabStop = True
Next
End Sub

Private Sub RichTextBox1_SelChange()
If RichTextBox1.SelBold Then
Toolbar1.Buttons("Bold").Value = tbrPressed     '�������ڰ���״̬
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

'���û�����Toolbar1�ؼ���ĳ����ťʱ���ͻᴥ����Toolbar1�ؼ����ṩ��ButtonClick�¼�
'ButtonClick�¼������ݵ�Button���������û������µİ���������ͨ�����ÿؼ���Key���ԣ��ؼ��֣�
'���������û�����ȥ������һ���������Ӷ�ִ����Ӧ����
'Ҳ���Լ�鰴ť��Index(����)�������ж��û����µ�����һ��������������������̫�಻�ü�
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
'����ʹ��һ���ܴ��Select Case �ṹ��ִ����Ӧ����
Select Case Button.Key            '���ݹؼ����ж�
Case "New"                        '�û������½�һ��RTF�ļ�
'����ǰ���Ѱ�CommonDialog1�ؼ���CancelError��������ΪTrue�������û�ѡ��ȡ����ʱ�ͻᷢ��һ������
'�Ӷ��ݴ˿����ж��û��Ƿ�ѡ���ˡ�ȡ����
On Error GoTo ErrHandler            '�����û�ѡ��ȡ����ʱ�����Ĵ���
CommonDialog1.DialogTitle = "�½�"
CommonDialog1.Filter = "RTF�ļ�|*.rtf"
'�Ա����ļ��ķ�ʽ��ʾͨ�öԻ���
CommonDialog1.ShowSave
'ȡ���Դ���Ĳ�׽����ʹ��Ϊ����ԭ���������Ĵ�����Ա�Visual Basic ���ɿ�������������
On Error GoTo 0
'���û�������ļ����浽����FileName��
FileName = CommonDialog1.FileName
'����������ļ����Ѵ��ڣ�����һ���Ի���ѯ���û��Ƿ񸲸�
If Dir(FileName) <> "" Then
If MsgBox("�Ƿ񸲸������ļ�" & FileName & "?", vbYesNo) = vbYes Then 'ѡ�񸲸�
Kill FileName       'ɾ��ԭ���ļ�
Else
Exit Sub            '�˳����¼�
End If
End If
'���RickTextBox1�ؼ��е���������
RichTextBox1.TextRTF = ""
RichTextBox1.SaveFile FileName
'����״̬����Ӧ��ǰ�򿪵��ļ�
StatusBar1.Panels(1).Text = "�ļ���" & FileName

Case "Open"         '�û�����һ�����ļ���ť
On Error GoTo ErrHandler
CommonDialog1.DialogTitle = "��"
CommonDialog1.Filter = "RTF�ļ�|*.rtf"
CommonDialog1.ShowOpen
On Error GoTo 0
FileName = CommonDialog1.FileName
'ʹ��RichTextbox1�ؼ���LoadFile�����������Զ����ļ����ؼ����ṩ�༭
RichTextBox1.LoadFile FileName
StatusBar1.Panels(1).Text = "�ļ���" & FileName

Case "Save"         '�û������˱����ļ���ť
'�����ǰ��û��ָ��������ļ����Ļ����͵���һ�������ļ���ͨ�öԻ���
If FileName = "" Then
On Error GoTo ErrHandler
CommonDialog1.DialogTitle = "���Ϊ"
CommonDialog1.Filter = "RTF�ļ�|*.rtf"
CommonDialog1.ShowSave
On Error GoTo 0
FileName = CommonDialog1.FileName
If Dir(FileName) <> "" Then
If MsgBox("�Ƿ񸲸������ļ�" & FileName & "?", vbYesNo, "Private Place Manage OP1.0") = vbYes Then
Kill FileName
Else
Exit Sub
End If
End If
End If
RichTextBox1.SaveFile FileName
StatusBar1.Panels(1).Text = "�ļ���" & FileName

Case "Print"       '�û�ѡ���ӡ
'����һ����׼�Ĵ�ӡ�Ի���
On Error GoTo ErrHandler
CommonDialog1.DialogTitle = "��ӡ"
CommonDialog1.Flags = cdlPDReturnDC + cdlPDNoPageNums
'���û��ѡ���ı������ӡRichTextBox1�ؼ����������ݣ������ӡָ���ı�
If RichTextBox1.SelLength = 0 Then
CommonDialog1.Flags = CommonDialog1.Flags + cdlPDAllPages
Else
CommonDialog1.Flags = CommonDialog1.Filter + cdlPDSelection
End If
CommonDialog1.ShowPrinter
On Error GoTo 0
'ʹ��RichTextBox1�ؼ���SelPrint�������д�ӡ����������Ĳ���ʽ��ӡ�����豸���
RichTextBox1.SelPrint CommonDialog1.hDC

Case "Cut"          '�û�ѡ�����
'��ռ��а���������
Clipboard.Clear
'�ѵ�ǰѡ���ı����Ƶ���������
Clipboard.SetText RichTextBox1.SelText
'ɾ��ѡ���ı�
RichTextBox1.SelText = ""

Case "Copy"         '�û�ѡ����
'��������ƣ�ֻ�ǲ���ɾ��ѡ���ı�
Clipboard.Clear
Clipboard.SetText RichTextBox1.SelText

Case "Paste"        '�û�ѡ��ճ��
RichTextBox1.SelText = Clipboard.GetText

Case "Bold"         '�û�ѡ�����
RichTextBox1.SelBold = Button.Value

Case "Italic"       'ѡ��б��
RichTextBox1.SelItalic = Button.Value

Case "UnderLine"    '�»���
RichTextBox1.SelUnderline = Button.Value

Case "Font"          '����
'��Form6������ʽ�Ի��򣩵ı����Ϊ�����塱
Form6.Caption = "����"
'���ø�ʽ�Ի�����ѡ���ѡ��ǡ����塱ѡ���������TabStrip�ؼ���Tabs�����е�
'��һ��ѡ�������һ��Tab����
Set Form6.TabStrip1.selectedItem = Form6.TabStrip1.Tabs(1)
'��ʾform7(��ʼ���������),����Form5��Ϊ���ĸ�����
Form7.Show , Form5
'��Progressbar1�ؼ���Maxֵ����Ϊ�ȵ�ǰ��������1������ʹ��Ϊ���̵�value���Ը�ֵ�����軻��
Form7.ProgressBar1.Max = Screen.FontCount - 1
'���form6��������Ͽ������е����ݣ�Ȼ����ϵͳ�Ŀ������������б��������Ǽӵ�������Ͽ�
Form6.Combo1.Clear          '������Ͽ�Combol1
For i = 0 To Screen.FontCount - 1
Form6.Combo1.AddItem Screen.Fonts(i)
Form7.ProgressBar1.Value = i
DoEvents
'����ϵͳ��ѭ��ʱϵͳʧȥ��Ӧ
Next
'����SetCurValue���̣�ʹForm6�еĿؼ��ܹ���ӳ��ǰֵ
SetCurValue
'����Form7����ģʽ�Ի���ķ�ʽ����ʾform6������form1��Ϊ������
Form7.Hide
Form6.Show 1, Form5
'����SetTextValue��ʹform6�и���ؿؼ��ĵ�ǰֵ�����õ�ǰ�ı��ĸ�ʽ
SetTextValue

Case "Paragraph"        '����
'��������
Form6.Caption = "����"
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
'selIndent���ص�ǰѡ�������������
.MaskEdBox2(0).Text = RichTextBox1.SelIndent
Else
.MaskEdBox2(0).Text = ""
End If
If Not IsNull(RichTextBox1.SelRightIndent) Then
'������������
.MaskEdBox2(1).Text = RichTextBox1.SelRightIndent
Else
.MaskEdBox2(1).Text = ""
End If
If Not IsNull(RichTextBox1.SelHangingIndent) Then
'����������
.MaskEdBox2(2).Text = RichTextBox1.SelHangingIndent
Else
.MaskEdBox2(2).Text = ""
End If
If Not IsNull(RichTextBox1.SelBullet) Then
'��ѡ���ı����ڶ���ǰ������Ŀ���ŵķ���true
If RichTextBox1.SelBullet Then
.Combo5.ListIndex = 0
'��Ŀ����
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
'ʹ��Format���������ɸ�ʽ�����ַ�������������ֵ��maskEdbox1��text����
'��ʱ���ֱ�Ӹ��ƿ��ܲ�ƥ��
.MaskEdBox1.Text = Format(RichTextBox1.SelFontSize, "00��")
Else
.MaskEdBox1.Text = "__��"
End If
If Not IsNull(RichTextBox1.SelColor) Then
.Label10.BackStyle = 1
.Label10.BackColor = RichTextBox1.SelColor
Else
'����ж�����ɫ��������Ϊ͸��
.Label10.BackStyle = 0
End If
End With
End Sub

Public Sub SetTextValue()
'����Ĵ��뽫���form6�и����ؼ���ֵ�뵱ǰѡ�����ı��ĸ������Ƿ�һ��
'����һ������ݸÿؼ���ֵ�޸ĵ�ǰѡ���ı��������ڶ���ĸ�ʽ���Ե�ֵ
With Form6
If Val(.MaskEdBox2(0).Text) > 0 Then
If Val(.MaskEdBox2(0).Text) <> RichTextBox1.SelIndent Then
Select Case .Combo2(0).Text
Case "��"
RichTextBox1.SelIndent = Val(.MaskEdBox2(0).Text) * 20
Case "����"
RichTextBox1.SelIndent = Val(.MaskEdBox2(0)) / 2.54 * 72 * 20
End Select
End If
End If
If Val(.MaskEdBox2(1).Text) > 0 Then
If Val(.MaskEdBox2(1).Text) <> RichTextBox1.SelRightIndent Then
Select Case .Combo2(1).Text
Case "��"
RichTextBox1.SelRightIndent = Val(.MaskEdBox2(1).Text) * 20
Case "����"
RichTextBox1.SelRightIndent = Val(.MaskEdBox2(1)) / 2.54 * 72 * 20
End Select
End If
End If
If Val(.MaskEdBox2(2).Text) > 0 Then
If Val(.MaskEdBox2(2).Text) <> RichTextBox1.SelHangingIndent Then
Select Case .Combo2(2).Text
Case "��"
RichTextBox1.SelHangingIndent = Val(.MaskEdBox2(2).Text) * 207
Case "����"
RichTextBox1.SelHangingIndent = Val(.MaskEdBox2(2)) / 2.54 * 72 * 20
End Select
End If
End If
If Val(.MaskEdBox1.ClipText) > 0 Then
If Val(.MaskEdBox1.ClipText) <> RichTextBox1.SelFontSize Then
'cliptext���ؿؼ����ų���������֮���ԭ���ַ���
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












