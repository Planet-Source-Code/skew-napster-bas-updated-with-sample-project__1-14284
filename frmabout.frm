VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   Picture         =   "frmabout.frx":0000
   ScaleHeight     =   2235
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   4320
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
frmabout.Hide
frmmain.Show
Unload frmabout
End Sub

Private Sub Form_Load()
Call StayOnTop(frmabout)
Label1.Caption = "Napster TooLz" & vbCrLf & "by" & vbCrLf & "Skew && Jaze" & vbCrLf & "Version 1.0" & vbCrLf & "Ghetto Software Inc." & vbCrLf & "http://go.to/GSoftware/"
Label2.Caption = "This program was created for educaional purposes only." & vbCrLf & "Any use other then that is prohibited and we are not responsible for anything that happens." & vbCrLf & "Have Fun!"
End Sub

Private Sub Label1_Click()
frmabout.Hide
frmmain.Show
Unload frmabout
End Sub

Private Sub Label2_Click()
frmabout.Hide
frmmain.Show
Unload frmabout
End Sub
