VERSION 5.00
Begin VB.Form frmsplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   Picture         =   "frmsplash.frx":0000
   ScaleHeight     =   2235
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call StayOnTop(frmsplash)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmsplash.Hide
frmmain.Show
Unload frmsplash
End Sub

