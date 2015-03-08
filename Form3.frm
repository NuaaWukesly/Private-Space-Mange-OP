VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õº∆¨‰Ø¿¿"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7350
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4845
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4815
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
Image1.Height = Form3.ScaleHeight
Image1.Width = Form3.ScaleWidth
Form3.Refresh
End Sub


