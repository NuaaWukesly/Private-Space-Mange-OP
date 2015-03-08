VERSION 5.00
Begin VB.Form SetupMain 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Private Space Manage OP 1.0  Installation"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6465
   Icon            =   "SetupMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6465
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "SetupMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
SetupMain.Show
Load step1
step1.Show 1, SetupMain
End Sub
