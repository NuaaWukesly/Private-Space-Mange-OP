VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Ʒ��Ϣ"
   ClientHeight    =   3630
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5580
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5580
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   360
      Picture         =   "Form9.frx":4F32
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   0
      Top             =   600
      Width           =   540
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5520
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      X1              =   1080
      X2              =   1080
      Y1              =   0
      Y2              =   2280
   End
   Begin VB.Label Label2 
      Caption         =   $"Form9.frx":9E64
      BeginProperty Font 
         Name            =   "���Ŀ���"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   1200
      TabIndex        =   4
      Top             =   2400
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   $"Form9.frx":9EBC
      BeginProperty Font 
         Name            =   "���Ŀ���"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label lblVersion 
      Caption         =   "Created By ������ ���� ��2013��2��2��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "   ˽�˿ռ������� ���� 1.0�汾"
      BeginProperty Font 
         Name            =   "���Ŀ���"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   3885
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

