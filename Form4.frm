VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ã½Ìåä¯ÀÀ´°¿Ú"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4620
   FillStyle       =   0  'Solid
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   254.226
   ScaleMode       =   0  'User
   ScaleWidth      =   311.899
   StartUpPosition =   1  'CenterOwner
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
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
      _cx             =   8281
      _cy             =   6588
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form2.WindowsMediaPlayer1.Controls.Stop
End Sub

Private Sub Form_Resize()
Form4.WindowsMediaPlayer1.Height = Form4.ScaleHeight + 0.5
Form4.WindowsMediaPlayer1.Width = Form4.ScaleWidth + 0.5
Form4.WindowsMediaPlayer1.stretchToFit = True
Form4.Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
Form4.WindowsMediaPlayer1.Controls.Stop
End Sub
