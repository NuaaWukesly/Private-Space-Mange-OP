VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form2 
   BackColor       =   &H80000008&
   Caption         =   "Our Place"
   ClientHeight    =   6030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9915
   Icon            =   "frm2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6030
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   5655
      Left            =   0
      TabIndex        =   2
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
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      TabIndex        =   1
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
      Caption         =   "ϵͳ��&F��"
      Begin VB.Menu munFileBackgroundMusic 
         Caption         =   "��������"
         Begin VB.Menu SetBGM 
            Caption         =   "���ñ�������"
            Begin VB.Menu AddBGM 
               Caption         =   "��ӱ�������"
            End
            Begin VB.Menu ChangeBGM 
               Caption         =   "�ı䱳������"
            End
         End
         Begin VB.Menu munFileBackgroundMusicOpen 
            Caption         =   "������"
         End
         Begin VB.Menu munFileBackgroundMusicClose 
            Caption         =   "�ر�����"
         End
      End
      Begin VB.Menu munFileOpen 
         Caption         =   "�򿪣�&O)"
         Begin VB.Menu munFileOpenPicture 
            Caption         =   "ͼƬ"
         End
         Begin VB.Menu munFileOpenPPS 
            Caption         =   "�õ�Ƭ"
         End
         Begin VB.Menu munFileOpenMusic 
            Caption         =   "��Ƶ"
         End
         Begin VB.Menu munFileOpenVedio 
            Caption         =   "��Ƶ"
         End
         Begin VB.Menu munFileOpenText 
            Caption         =   "�ı�"
         End
      End
      Begin VB.Menu setSys 
         Caption         =   "ϵͳ����"
         Begin VB.Menu SetBGS 
            Caption         =   "���ñ������"
            Begin VB.Menu AddBGS 
               Caption         =   "���ӱ������"
            End
            Begin VB.Menu ChangeBGS 
               Caption         =   "�ı䱳�����"
            End
         End
         Begin VB.Menu setBD 
            Caption         =   "��������"
         End
      End
      Begin VB.Menu munFileExit 
         Caption         =   "�˳���&E��"
      End
   End
   Begin VB.Menu munAboutOur 
      Caption         =   "�������ǣ�&W��"
      Begin VB.Menu munAboutOurLetter 
         Caption         =   "���ǵ�����"
         Begin VB.Menu OurLetterScan 
            Caption         =   "���"
            Begin VB.Menu textWay 
               Caption         =   "�ı���ʽ"
            End
            Begin VB.Menu PPS 
               Caption         =   "�õ�Ƭ"
               Begin VB.Menu SWFStyle 
                  Caption         =   "SWF��ʽ"
               End
               Begin VB.Menu PPSStyle 
                  Caption         =   "PPS��ʽ"
               End
            End
         End
         Begin VB.Menu OurLetterEdit 
            Caption         =   "��д"
         End
         Begin VB.Menu OurLetterAdd 
            Caption         =   "���"
         End
         Begin VB.Menu OurLetterBackup 
            Caption         =   "����"
         End
      End
      Begin VB.Menu OurPPS 
         Caption         =   "���ǵĻõ�Ƭ"
         Begin VB.Menu AddOurPPS 
            Caption         =   "���"
         End
         Begin VB.Menu ScanOurPPS 
            Caption         =   "���"
         End
      End
      Begin VB.Menu munAboutOurPicture 
         Caption         =   "���ǵ���Ƭ"
         Begin VB.Menu OurPictureScan 
            Caption         =   "���"
         End
         Begin VB.Menu OurPictureAdd 
            Caption         =   "���"
         End
         Begin VB.Menu OurPictureBackup 
            Caption         =   "����"
         End
      End
      Begin VB.Menu munAboutOurMusic 
         Caption         =   "���ǵ�����"
         Begin VB.Menu OurMusicScan 
            Caption         =   "���"
         End
         Begin VB.Menu OurMusicAdd 
            Caption         =   "���"
         End
      End
      Begin VB.Menu munAboutOurVedio 
         Caption         =   "���ǵ���Ƶ"
         Begin VB.Menu OurVedioScan 
            Caption         =   "���"
         End
         Begin VB.Menu OurVedioAdd 
            Caption         =   "���"
         End
      End
      Begin VB.Menu munAboutOurCourse 
         Caption         =   "���ǵ�����"
         Begin VB.Menu OurCourseScan 
            Caption         =   "���"
         End
         Begin VB.Menu OurCourseEdit 
            Caption         =   "�༭"
         End
         Begin VB.Menu OurCourseAdd 
            Caption         =   "���"
         End
         Begin VB.Menu OurCourseBackup 
            Caption         =   "����"
         End
      End
   End
   Begin VB.Menu EXEProgram 
      Caption         =   "EXE����"
      Begin VB.Menu ScanEXE 
         Caption         =   "���"
      End
      Begin VB.Menu AddEXE 
         Caption         =   "���"
      End
   End
   Begin VB.Menu AboutThis 
      Caption         =   "���ڲ�Ʒ��&A��"
      Begin VB.Menu AboutThisNew 
         Caption         =   "��Ʒ��Ϣ"
      End
      Begin VB.Menu AboutThisExplain 
         Caption         =   "��Ʒ˵��"
      End
   End
   Begin VB.Menu help 
      Caption         =   "������&H��"
      Begin VB.Menu helpText 
         Caption         =   "�����ĵ�"
      End
      Begin VB.Menu helpTechSupport 
         Caption         =   "����֧��"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strShow As String                   '��ʾ�ַ���
Dim awayBD As Integer
Dim birthday As Date                        '����
Dim BDstr As String                         '�����ַ���
Dim BGSstr As String                        '����·��
Dim BGMstr As String                        '��������·��
Dim myFSO As New FileSystemObject           '�ļ�����

'���ڻõ�Ƭ��ӳ
Private Declare Function icePub_open Lib "icePubDll.dll" (ByVal strPath As String) As Integer

'����һ��������ȡ����·�������ڳ�ʼ����·��Auto
Function GetDesktopPath() As String
Dim WshShell As Object, WScript As Object, oShellLink As Object
Set WshShell = CreateObject("WScript.Shell")
GetDesktopPath = WshShell.SpecialFolders("Desktop") '����·��Startup
End Function


'����һ������,������Դ����RecType������Ӧ��¼�ļ���ȡ��¼���ݣ����޸��ļ����򴴽����ļ����ļ��޼�¼���򷵻�Ĭ��·��
Function GetRec(ByVal RecType As String) As String
Dim tempStr As String
Select Case RecType

'������Դ
Case "BGS"
If Not myFSO.FileExists(App.Path & "\System\BGStyle.rec") Then
'������ļ�������
'tempStr=Auto
tempStr = App.Path & "\OtherRec\BGStyle\BGStyle1.html"
'���ü�¼д���¼�ļ�
Open App.Path & "\System\BGStyle.rec" For Output As 1
Print #1, tempStr
Close #1
Else
'���ļ����ڣ���ȡ����
Open App.Path & "\System\BGStyle.rec" For Input As 1
Line Input #1, tempStr
Close #1
If tempStr = "" Then
'���Ϊ�գ�tempStr=Auto
tempStr = App.Path & "\OtherRec\BGStyle\BGStyle1.html"
End If
'��Ϊ��
End If

'��������
Case "BGM"
If Not myFSO.FileExists(App.Path & "\System\BGMusic.rec") Then
'������ļ�������
'tempStr=Auto
tempStr = App.Path & "\Music\BGMusic\BGMusic.mp3"
'���ü�¼д���¼�ļ�
Open App.Path & "\System\BGMusic.rec" For Output As 1
Print #1, tempStr
Close #1
Else
'���ļ����ڣ���ȡ����
Open App.Path & "\System\BGMusic.rec" For Input As 1
Line Input #1, tempStr
Close #1
If tempStr = "" Then
'���Ϊ�գ�tempStr=Auto
tempStr = App.Path & "\Music\BGMusic\BGMusic.mp3"
End If
'��Ϊ��
End If
End Select
GetRec = tempStr
End Function

'����һ��������������Դ����RecType������Ӧ����RecDataд����Ӧ����Դ�ļ�
Function WriteRec(ByVal RecType As String, ByVal RecData As String) As Integer
Select Case RecType
'���������Դ
Case "BGS"
Open App.Path & "\System\BGStyle.rec" For Output As 1
'��������
Case "BGM"
Open App.Path & "\System\BGMusic.rec" For Output As 1
End Select
'д������
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
'�Ի�������
Form2.CommonDialog1.DialogTitle = "��ӱ�������"
CommonDialog1.Filter = "��Ƶ�ļ�(*.mp3;*.wav)|*.mp3;*.wav"
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\Music\BGMusic\"
MsgBox "�ļ���ӳɹ���", , "ϵͳ��ʾ"
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

'����һ����������������ļ�����,���½����ļ�

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
'�Ի�������
Form2.CommonDialog1.DialogTitle = "��ӱ������"
CommonDialog1.Filter = "HTML�ļ���Flash(*.html;*.swf)|*.html;*,swf"
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\OtherRes\BGStyle\" & CreateStyleFolder()
MsgBox "�ļ���ӳɹ���ע�������Դ�����ӡ�", , "ϵͳ��ʾ"
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub AddEXE_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "���EXE����"
CommonDialog1.Filter = "EXE�ļ�(*.exe)|*.exe"
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\OtherRes\PPS\OurPPS\"
MsgBox "�ļ���ӳɹ���ע�������Դ�����ӡ�", , "ϵͳ��ʾ"
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub AddOurPPS_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "������ǵĻõ�Ƭ"
CommonDialog1.Filter = "�õ�Ƭ�ļ�(*.ppt;*.pptx;*.pps;*.ppsx)|*.ppt;*.pptx;*.pps;*.ppsx"
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\OtherRes\PPS\OurPPS\"
MsgBox "�ļ���ӳɹ���ע�������Դ�����ӡ�", , "ϵͳ��ʾ"
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub ChangeBGM_Click()
Dim i As Integer
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "ѡ�񱳾�������"
CommonDialog1.Filter = "��Ƶ�ļ�(*.mp3;*.wav)|*.mp3;*.wav"
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = App.Path & "\Music\BGMusic"
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
Form2.WindowsMediaPlayer1.Close
Form2.WindowsMediaPlayer1.URL = Form2.CommonDialog1.FileName
Form2.WindowsMediaPlayer1.Controls.Play
i = WriteRec("BGM", Form2.CommonDialog1.FileName)
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub ChangeBGS_Click()
Dim i As Integer
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "ѡ�񱳾����"
CommonDialog1.Filter = "HTML�ļ�(*.html)|*.html"
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = App.Path & "\OtherRes\BGStyle\"
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
Form2.WebBrowser2.Navigate Form2.CommonDialog1.FileName
i = WriteRec("BGS", Form2.CommonDialog1.FileName)
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub Form_Load()
Form2.StatusBar1.Panels(1).Text = App.Path              '��¼��װĿ¼
'��ȡ��������������·��
BGSstr = GetRec("BGS")
BGMstr = GetRec("BGM")
'���ش���2
Form2.Show
'��ȡ��������
Open App.Path & "\System\Birthday.bd" For Input As 1
Line Input #1, BDstr
Close #1
birthday = DateValue(BDstr)
'��������յ�����
awayBD = CInt(DateDiff("d", Date, CDate(Month(birthday) & "-" & Day(birthday))))
If awayBD = 0 Then
strShow = "�������������գ������˴�������"
ElseIf awayBD < 0 Then
awayBD = awayBD + 365
strShow = "��������ǣ�" & Month(birthday) & "��" & Day(birthday) & "��,�������ջ���" & CStr(awayBD) & "��"
Else
strShow = "��������ǣ�" & Month(birthday) & "��" & Day(birthday) & "��,�������ջ���" & CStr(awayBD) & "��"
End If
Form2.StatusBar1.Panels(3).Text = strShow
'�������
Form2.WebBrowser2.Navigate BGSstr
'���ű�������
Form2.WindowsMediaPlayer1.URL = BGMstr
Form2.WindowsMediaPlayer1.Controls.Play
Form2.Timer1.Interval = 1000
End Sub

Private Sub Form_Resize()
'���ڴ�С�ı�ʱ
If Form2.ScaleHeight <> 0 Then
Form2.WebBrowser2.Height = Form2.ScaleHeight - Form2.StatusBar1.Height
Else
Form2.WebBrowser2.Height = Form2.ScaleHeight
End If
Form2.WebBrowser2.Width = Form2.ScaleWidth
Form2.Refresh    'ˢ��
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
Form10.Caption = "�����ĵ�"
Form10.RichTextBox1.LoadFile App.Path & "\" & "RTF\HelpText\HelpText.rtf"
Load Form10
Form10.Show 1, Form2
End Sub

Private Sub munFileBackgroundMusicClose_Click()
'�رձ�������
Form2.WindowsMediaPlayer1.Controls.Stop
End Sub

Private Sub munFileBackgroundMusicOpen_Click()
'�򿪱�������
Form2.WindowsMediaPlayer1.Controls.Play
End Sub

Private Sub munFileExit_Click()
Unload Form2
End Sub

Private Sub munFileOpenMusic_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'��ѡ����Ƶʱ,�ļ��������
CommonDialog1.Flags = cdlOFNFileMustExist
'�򿪵��ļ�����
CommonDialog1.Filter = "��Ƶ�ļ�(*.mp3;*.wav)|*.mp3;*.wav"
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
'�ļ�·������Ϊ��
If Not CommonDialog1.FileName = "" Then    '��ѡ��Ϊ��
Form4.Icon = LoadPicture(App.Path & "\" & "Ico\Music.ico")
'URL��ֵ
Form4.WindowsMediaPlayer1.URL = CommonDialog1.FileName
'������ѡ��
Load Form4
Form4.Show 1, Form2
Form4.WindowsMediaPlayer1.Controls.Play
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub munFileOpenPicture_Click()
'ѡ���ͼƬʱ��
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'��ѡ��ͼƬʱ,�ļ��������
CommonDialog1.Flags = cdlOFNFileMustExist
'�򿪵��ļ�����
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.Filter = "ͼƬ(*.jpg;*.bmp;*.gif;*.ico)|*.jpg;*.bmp;*.gif;*.ico"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
'�ļ�·������Ϊ��
If Not CommonDialog1.FileName = "" Then
'��ȡ���ͼƬ
Form3.Image1.Picture = LoadPicture(CommonDialog1.FileName)
'�ڴ���3��ʾͼƬ
Form3.Show 1, Form2
End If
GoTo sign:    '�������¼�
ErrHandler:   '���������ȡ��������ô������2��ʾ
'�ı䴰��2�ı���
CommonDialog1.FileName = ""
sign:
End Sub

Private Sub munFileOpenPPS_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'��ѡ����Ƶʱ,�ļ��������
CommonDialog1.Flags = cdlOFNFileMustExist
'�򿪵��ļ�����
CommonDialog1.Filter = "�õ�Ƭ(*.ppt;*.pptx;*.pps;*.ppsx)|*.ppt;*.pptx;*.pps;*ppsx"
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
'�ļ�·������Ϊ��
If Not CommonDialog1.FileName = "" Then    '��ѡ��Ϊ��
Dim a2 As Integer
a2 = icePub_open(Form2.CommonDialog1.FileName)
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ

End Sub

Private Sub munFileOpenText_Click()
Dim FileName As String
With Form5
On Error GoTo ErrHandler
.CommonDialog1.DialogTitle = "��"
.CommonDialog1.Filter = "RTF�ļ�|*.rtf"
'Ĭ��·��Ϊ����
.CommonDialog1.InitDir = GetDesktopPath()
.CommonDialog1.FileName = ""
.CommonDialog1.ShowOpen
On Error GoTo 0
FileName = .CommonDialog1.FileName
'ʹ��RichTextbox1�ؼ���LoadFile�����������Զ����ļ����ؼ����ṩ�༭
.RichTextBox1.LoadFile FileName
.StatusBar1.Panels(1).Text = "�ļ���" & FileName
End With
Form5.Show
ErrHandler:
End Sub

Private Sub munFileOpenVedio_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'��ѡ����Ƶʱ,�ļ��������
CommonDialog1.Flags = cdlOFNFileMustExist
'�򿪵��ļ�����
CommonDialog1.Filter = "��Ƶ�ļ�(*.wmv;*.avi;*.mp4)|*.wmv;*.avi;*.mp4"
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
'�ļ�·������Ϊ��
If Not CommonDialog1.FileName = "" Then    '��ѡ��Ϊ��
Form4.Icon = LoadPicture(App.Path & "\" & "Ico\VedioLook.ico")          'ע��˷���
Form4.Show
'URL��ֵ
Form4.WindowsMediaPlayer1.URL = CommonDialog1.FileName
'������ѡ��
Form4.WindowsMediaPlayer1.Controls.Play
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub OurCourseAdd_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "������ǵ�����"
CommonDialog1.Filter = "RTF�ļ�(*.rtf)|*.rtf"
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\RTF\OurCourse\"
MsgBox "�ļ���ӳɹ���", , "ϵͳ��ʾ"
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub OurCourseBackup_Click()
'ѡ�񱸷����ǵ�����
Backup.Caption = "�������ǵ�����"
Backup.Label3.Caption = App.Path & "\RTF\OurCourse"
'���ر��ݴ���
Load Backup
End Sub

Private Sub OurCourseEdit_Click()
'���ñ���·��
Form5.CommonDialog1.InitDir = App.Path & "\RTF\OurCourse"
'Ĭ���ļ�������
Form5.CommonDialog1.FileName = "����_n"
Load Form5
Form5.Show
End Sub

Private Sub OurCourseScan_Click()
'�Ի�������
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
Form2.CommonDialog1.DialogTitle = "���ǵ�����"
'�ļ�·��
Form2.CommonDialog1.InitDir = App.Path & "\" & "RTF\OurCourse"
CommonDialog1.Filter = "RTF�ļ�(*.rtf)|*.rtf"
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
Load Form5
If Not Form2.CommonDialog1.FileName = "" Then
Form5.RichTextBox1.LoadFile Form2.CommonDialog1.FileName
Form5.StatusBar1.Panels(1).Text = Form2.CommonDialog1.FileName
End If
Form5.Show
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub OurLetterAdd_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "������ǵ�����"
CommonDialog1.Filter = "RTF�ļ�(*.rtf)|*.rtf"
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\RTF\OurLetter\"
MsgBox "�ļ���ӳɹ���", , "ϵͳ��ʾ"
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub OurLetterBackup_Click()
'ѡ�񱸷����ǵ�����
Backup.Caption = "�������ǵ�����"
Backup.Label3.Caption = App.Path & "\RTF\OurLetter"
'���ر��ݴ���
Load Backup
End Sub

Private Sub OurLetterEdit_Click()
'���ñ���·��
Form5.CommonDialog1.InitDir = App.Path & "\RTF\OurLetter"
'Ĭ���ļ�������
Form5.CommonDialog1.FileName = "����_n"
Load Form5
Form5.Show
End Sub

Private Sub OurMusicAdd_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "������ǵ�����"
CommonDialog1.Filter = "��Ƶ�ļ�(*.mp3;*.wav)|*.mp3;*.wav"
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = GetDesktopPath()
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\Music\OurMusic\"
MsgBox "�ļ���ӳɹ���", , "ϵͳ��ʾ"
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub OurMusicScan_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "���ǵ�����"
'�ļ�·��
Form2.CommonDialog1.InitDir = App.Path & "\" & "Music\OurMusic"
CommonDialog1.Flags = cdlOFNFileMustExist
'�򿪵��ļ�����
CommonDialog1.Filter = "��Ƶ�ļ�(*.mp3;*.wav)|*.mp3;*.wav"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
'�ļ�·������Ϊ��
If Not CommonDialog1.FileName = "" Then    '��ѡ��Ϊ��
Form4.Icon = LoadPicture(App.Path & "\" & "Ico\Music.ico")
Form4.WindowsMediaPlayer1.URL = Form2.CommonDialog1.FileName
Load Form4
Form4.Show 1, Form2
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub OurPictureAdd_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "������ǵ���Ƭ"
CommonDialog1.Filter = "ͼƬ(*.jpg;*.bmp;*.gif;*.ico)|*.jpg;*.bmp;*.gif;*.ico"
CommonDialog1.FileName = ""
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = GetDesktopPath()
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\Image\OurPicture\"
MsgBox "�ļ���ӳɹ���", , "ϵͳ��ʾ"
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub OurPictureBackup_Click()
'ѡ�񱸷����ǵ�����
Backup.Caption = "�������ǵ���Ƭ"
Backup.Label3.Caption = App.Path & "\Image\OurPicture"
'���ر��ݴ���
Load Backup
End Sub

Private Sub OurPictureScan_Click()
'ѡ���ͼƬʱ��
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'��ѡ��ͼƬʱ,�ļ��������
CommonDialog1.Flags = cdlOFNFileMustExist
'�򿪵��ļ�����
'�Ի�������
Form2.CommonDialog1.DialogTitle = "���ǵ���Ƭ"
'�ļ�·��
Form2.CommonDialog1.InitDir = App.Path & "\" & "Image\OurPicture"
CommonDialog1.Filter = "ͼƬ(*.jpg;*.bmp;*.gif;*.ico)|*.jpg;*.bmp;*.gif;*.ico"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
'�ļ�·������Ϊ��
If Not CommonDialog1.FileName = "" Then
'��ȡ���ͼƬ
Form3.Image1.Picture = LoadPicture(CommonDialog1.FileName)
'�ڴ���3��ʾͼƬ
Form3.Show 1, Form2
End If
GoTo sign:    '�������¼�
ErrHandler:   '���������ȡ��������ô������2��ʾ
'�ı䴰��2�ı���
CommonDialog1.FileName = ""
sign:
End Sub

Private Sub OurVedioAdd_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "������ǵ���Ƶ"
CommonDialog1.Filter = "��Ƶ�ļ�(*.wmv;*.avi;*.mp4)|*.wmv;*.avi;*.mp4"
CommonDialog1.FileName = ""
'Ĭ��·��Ϊ����
Form2.CommonDialog1.InitDir = GetDesktopPath()
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
myFSO.CopyFile Form2.CommonDialog1.FileName, App.Path & "\Vedio\OurVideo\"
MsgBox "�ļ���ӳɹ���", , "ϵͳ��ʾ"
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub OurVedioScan_Click()
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "���ǵ���Ƶ"
'�ļ�·��
Form2.CommonDialog1.InitDir = App.Path & "\" & "Vedio\OurVedio"
CommonDialog1.Flags = cdlOFNFileMustExist
'�򿪵��ļ�����
CommonDialog1.Filter = "��Ƶ�ļ�(*.wmv;*.avi;*.mp4)|*.wmv;*.avi;*.mp4"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
'�ļ�·������Ϊ��
If Not CommonDialog1.FileName = "" Then    '��ѡ��Ϊ��
Form4.Icon = LoadPicture(App.Path & "\" & "Ico\VedioLook.ico")
Form4.WindowsMediaPlayer1.URL = Form2.CommonDialog1.FileName
Load Form4
Form4.Show 1, Form2
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub PPSStyle_Click()
'���������ʽ֮�ı���ʽ
Dim a2 As Integer
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "���ǵ�����PPS��ʽ"
'�ļ�·��
Form2.CommonDialog1.InitDir = App.Path & "\OtherRes\PPS\OurLetter\"
CommonDialog1.Filter = "PPS�ļ�(*.ppsx;*pps)|*.ppsx;*pps"
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
a2 = icePub_open(Form2.CommonDialog1.FileName)
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub ScanEXE_Click()
Dim a2 As Integer
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "���ǵĻõ�Ƭ"
'�ļ�·��
Form2.CommonDialog1.InitDir = App.Path & "\OtherRes\EXE\"
CommonDialog1.Filter = "EXE�ļ�(*.exe)|*.exe"
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
a2 = icePub_open(Form2.CommonDialog1.FileName)
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub ScanOurPPS_Click()
Dim a2 As Integer
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "���ǵĻõ�Ƭ"
'�ļ�·��
Form2.CommonDialog1.InitDir = App.Path & "\OtherRes\PPS\OurPPS\"
CommonDialog1.Filter = "PPS�ļ�(*.ppt;*pptx;*.ppsx;*pps)|*.ppt;*pptx;*.ppsx;*pps"
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
If Not Form2.CommonDialog1.FileName = "" Then
a2 = icePub_open(Form2.CommonDialog1.FileName)
End If
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub setBD_Click()
Load setBirthday
setBirthday.Show 1, Form2
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)
Dim a As Double
'��ȡ��������
Open App.Path & "\System\Birthday.bd" For Input As 1
Line Input #1, BDstr
Close #1
birthday = DateValue(BDstr)
awayBD = CInt(DateDiff("d", Date, CDate(Month(birthday) & "-" & Day(birthday))))
If Panel.Index = 3 Then
If awayBD = 0 Then
'Load Form13
'Form13.Show
a = Shell(App.Path & "\MyBrowse.exe", vbMaximizedFocus)
Form2.WindowsMediaPlayer1.Controls.pause
End If
End If
If Panel.Index = 2 Then
Form12.Show 1, Form2
End If
End Sub

Private Sub SWFStyle_Click()
Load Form14
Form14.Show 1, Form2
End Sub

Private Sub textWay_Click()
'���������ʽ֮�ı���ʽ
CommonDialog1.CancelError = True
On Error GoTo ErrHandler:
'�Ի�������
Form2.CommonDialog1.DialogTitle = "���ǵ�����"
'�ļ�·��
Form2.CommonDialog1.InitDir = App.Path & "\" & "RTF\OurLetter"
CommonDialog1.Filter = "RTF�ļ�(*.rtf)|*.rtf"
CommonDialog1.FileName = ""
Form2.CommonDialog1.ShowOpen
Load Form5
If Not Form2.CommonDialog1.FileName = "" Then
Form5.RichTextBox1.LoadFile Form2.CommonDialog1.FileName
Form5.StatusBar1.Panels(1).Text = Form2.CommonDialog1.FileName
End If
Form5.Show
ErrHandler:   '���������ȡ��������ô������2��ʾ
End Sub

Private Sub Timer1_Timer()
Form2.StatusBar1.Panels(2).Text = "��ǰʱ�䣺" & Date & "  " & Time
End Sub
