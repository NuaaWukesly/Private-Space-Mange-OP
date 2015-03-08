VERSION 5.00
Begin VB.Form Form15 
   Caption         =   "AddFile"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form15"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function icePub_open Lib "icePubDll.dll" (ByVal strPath As String) As Integer
Private Sub Form_Load()
Dim a2 As Integer
a2 = icePub_open(App.Path & "\OtherRes\PPS\ourLetter.ppsx")
End Sub
