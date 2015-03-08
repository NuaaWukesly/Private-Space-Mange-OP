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
      Caption         =   "�Ƿ������ļ���"
      Height          =   1095
      Left            =   2520
      TabIndex        =   10
      Top             =   1440
      Width           =   3495
      Begin VB.CheckBox Check1 
         Caption         =   "������ļ���"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Value           =   1  'Checked
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�����ݵ��ļ���·��"
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
      Caption         =   "���ݱ���·��"
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
      Caption         =   "ȡ��"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ʼ����"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�½��ļ���"
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
      Caption         =   "����ѡ�񱸷ݱ����·��"
      BeginProperty Font 
         Name            =   "Adobe ���� Std R"
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
Dim backupPath As String                  '���ݱ����·��
Dim needPath As String                    '��Ҫ���ݵ��ļ���·��
Dim ourFSO As New FileSystemObject        '�����ļ�����
Dim backupSign As String                  '��������
Dim openStr As String                     '����Դ������ʱ��ʾ��·��
Dim sign As Boolean                       '��־�Ƿ�Ϊ��ʼ״̬
'����Դ��������ָ���ļ�
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'ShellExecute hWnd, "open", "explorer.exe", "/e,/select,E:\�����ļ�.MP3", "", 1
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'ͨ��explorer.exe�Ŀ���ʵ��.
'Explorer.exe�Ĳ������£�
'�����ʽExplorer [/n][/e][[,/root],[path]][[,/select],[path filename]]
'����˵��
'/n��ʾ�ԡ��ҵĵ��ԡ���ʽ��һ���µĴ��ڣ�ͨ���򿪵���Windows��װ�����ĸ�Ŀ¼��
'/e��ʾ�ԡ���Դ����������ʽ��һ���µĴ��ڣ�ͨ���򿪵�Ҳ��Windows��װ�����ĸ�Ŀ¼��
'/root,[path]��ʾ��ָ�����ļ��У�
'/root��ʾֻ��ʾָ���ļ���������ļ����У���
'����ʾ�������̷������ļ��У�[path]��ʾָ����·����
'�������/root��������ֻ��[path]������
'�������ʾ�������̷������ļ����е����ݡ�
'���⣬[path]������ָ�����繲���ļ��С�
'/select,[path filename]��ʾ��ָ�����ļ��в���ѡ��ָ�����ļ���
'[path filename]��ʾָ����·�����ļ�����
'�������/select��������ϵͳ������Ӧ�Ĺ�������򿪸��ļ���
'���[path filename]�����ļ����ͻ�򿪸��ļ��е��ϼ�Ŀ¼��ѡ�и��ļ��С�
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'����һ���������жϱ������ͣ������䷵��
Function judge(ByVal judgestr As String) As String
If Not judgestr = "" Then
If judgestr = "�������ǵ�����" Then
judge = "OurLetter"
End If
If judgestr = "�������ǵ���Ƭ" Then
judge = "OurPicture"
End If
If judgestr = "�������ǵ�����" Then
judge = "OurCourse"
End If
If judgestr = "�������ǵĻõ�Ƭ" Then
judge = "PPS"
End If
If judgestr = "����EXE" Then
judge = "EXE"
End If
End If
End Function
 
Private Sub Command1_Click()
Dim newfold As String                     '�½��ļ�����
Dim curPath As String                     '��¼��ǰ·��
'�½��ļ���
curPath = Backup.Dir1.Path
'�½��ļ��б�����һ������
Do
newfold = InputBox("�������½��ļ�����", "ϵͳ��ʾ")
If newfold = "" Then
a = MsgBox("���벻��Ϊ�գ���Ҫ�����������", vbYesNo, ϵͳ��ʾ)
End If
Loop While a = vbYes
If Not newfold = "" Then
MkDir curPath & "\" & newfold
Dir1.Refresh
'��λ���½��ļ���
Backup.Dir1.Path = curPath & "\" & newfold
End If
End Sub

Private Sub Command2_Click()
'��ʼ����
'��Ҫ���ݵ��ļ���·������form2�и���
needPath = Backup.Label3.Caption
backupSign = judge(Backup.Caption)
Backup.Label2.Caption = backupPath
If ourFSO.FolderExists(backupPath & "\" & backupSign) = False Then
'�½��ļ���ȡ�����ҵı���
MkDir backupPath & "\" & backupSign
End If
If Len(backupPath) = 3 Then
backupPath = backupPath & backupSign
Else
backupPath = backupPath & "\" & backupSign
End If
'���ļ��Ѵ����򸲸�
ourFSO.CopyFolder needPath, backupPath
MsgBox Backup.Caption & "�ɹ���", , "ϵͳ��Ϣ"
If Backup.Check1.Value = 1 Then
Mid(backupPath, 1, 1) = UCase(Mid(backupPath, 1, 1))
'��ʽ
openStr = "/e,/select," & backupPath
ShellExecute hWnd, "open", "explorer.exe", openStr, "", 1
End If
Unload Backup
End Sub

Private Sub Command3_Click()
a = MsgBox("��ȥ��Ҫ�˳�" & Backup.Caption & "��", vbYesNo, "ϵͳ��ʾ")
If a = vbYes Then
Unload Backup
End If
End Sub

Private Sub Dir1_Change()
If sign = False Then                      '������ǳ�ʼ
backupPath = Backup.Dir1.Path             '��ʾ��·��
Backup.Label2.Caption = backupPath
Else
sign = False
End If
End Sub

Private Sub Drive1_Change()
'�������ı�ʱ,ͬ��
If Not sign = True Then
Backup.Dir1.Path = Backup.Drive1.Drive
End If
End Sub

Private Sub Form_Load()
Dim WshShell As Object, WScript As Object, oShellLink As Object
Set WshShell = CreateObject("WScript.Shell")
sign = True                                       '��ʼ
backupPath = WshShell.SpecialFolders("Desktop")   '����·��
Backup.Drive1 = Mid(backupPath, 1, 3)
Backup.Dir1.Path = backupPath
Backup.Label2.Caption = backupPath
Backup.Show
End Sub

