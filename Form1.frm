VERSION 5.00
Begin VB.Form LoadPage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ǵĿռ�-Loading Page"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5310
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton comCancel 
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton comLoad 
      Caption         =   "��½"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox inPutPassWord 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   0
      Picture         =   "Form1.frx":4F32
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "���Ŀ���"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   855
   End
End
Attribute VB_Name = "LoadPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim password As String
'���ܺ���
Private Function Code(OriStr As String) As String
Dim i As Integer, n As Integer
Dim tempStr As String
tempStr = OriStr
n = Len(OriStr)
For i = 1 To n
'���ܹ��̣�ÿ���ַ�������3���
Mid(tempStr, i, 1) = Chr((Asc((Mid(OriStr, i, 1))) Xor 90))
Next i
Code = tempStr
End Function

Private Sub comCancel_Click()
Unload LoadPage
Unload Form2
End Sub

Private Sub comLoad_Click()
If inPutPassWord.Text = password Then
Load Form2
Unload LoadPage
Else
MsgBox "���������������ȷ������½��", , "Private Place Manage OP1.0"
inPutPassWord.Text = ""
End If
End Sub

Private Sub Command1_Click()
If inPutPassWord.Text = password Then
Load ModPW
ModPW.Show
Unload LoadPage
Else
MsgBox "���������������ȷ������½��" & password, , "Private Place Manage OP1.0"
inPutPassWord.Text = ""
End If
End Sub

Private Sub Form_Load()
'ͨ��ʹ��line input# ���� input# ��������� print# ���ļ���д������ݣ�����input# �����ļ����� write#д�������
Open App.Path & "\System\key.pw" For Input As #1
'��ȡ����
Line Input #1, password
'����
password = Code(password)
'�ر��ļ�
Close #1
End Sub

