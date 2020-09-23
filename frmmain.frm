VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   0  'None
   Caption         =   "Napster Toolz 1.0"
   ClientHeight    =   1800
   ClientLeft      =   2355
   ClientTop       =   2070
   ClientWidth     =   2985
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmmain.frx":030A
   ScaleHeight     =   1800
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2760
      Top             =   1200
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Toolz"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   960
      TabIndex        =   7
      Top             =   390
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2295
      TabIndex        =   6
      Top             =   395
      Width           =   375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Botz"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1665
      TabIndex        =   5
      Top             =   395
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   345
      TabIndex        =   4
      Top             =   395
      Width           =   315
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   1460
      Width           =   2745
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "__"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2460
      TabIndex        =   2
      Top             =   0
      Width           =   180
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   45
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Napster Toolz 1.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   25
      Width           =   1380
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Call RunMenuByString("&About Napster")
End Sub

Private Sub Form_Load()
Call StayOnTop(frmmain)
ChatSend ("(¯`·-> Napster Toolz 1.0 <-·´¯)")
ChatSend ("(¯`·-> By Skew & JaZe <-·´¯)")
ChatSend ("(¯`·-> Loaded By: " & UserSN() & " <-·´¯)")
ChatSend ("(¯`·-> http://go.to/GSoftware/ <-·´¯)")
End Sub

Private Sub Form_LostFocus()
Call RestoreColor(frmmain, vbWhite)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmmain)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmain, vbWhite)
End Sub

Private Sub Form_Unload(Cancel As Integer)
ChatSend ("(¯`·-> Napster Toolz 1.0 <-·´¯)")
ChatSend ("(¯`·-> By Skew & JaZe <-·´¯)")
ChatSend ("(¯`·-> Unloaded By: " & UserSN & " <-·´¯)")
ChatSend ("(¯`·-> http://go.to/GSoftware/  <-·´¯)")
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmmain)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmain, vbWhite)
Label1.ForeColor = vbBlue
End Sub

Private Sub Label2_Click()
Call UnloadAll
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmain, vbWhite)
Label2.ForeColor = vbBlue
End Sub

Private Sub Label3_Click()
frmmain.WindowState = vbMinimized
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmain, vbWhite)
Label3.ForeColor = vbBlue
End Sub

Private Sub Label4_Click()
frmmain.PopupMenu frmmenu.mnufile
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmain, vbWhite)
Label4.ForeColor = vbBlue
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmain, vbWhite)
Label5.ForeColor = vbBlue
End Sub

Private Sub Label6_Click()
frmmain.PopupMenu frmmenu.mnubotz
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmain, vbWhite)
Label6.ForeColor = vbBlue
End Sub

Private Sub Label7_Click()
frmmain.PopupMenu frmmenu.mnuhelp
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmain, vbWhite)
Label7.ForeColor = vbBlue
End Sub

Private Sub Label8_Click()
frmmain.PopupMenu frmmenu.mnuchat
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmain, vbWhite)
Label8.ForeColor = vbBlue
End Sub

Private Sub Timer1_Timer()
DoEvents
If IsUserOnline = False Then
   DoEvents
   Label5.Caption = "Disconnected"
Else
   DoEvents
   Label5.Caption = "Loaded By: " & UserSN()
End If
End Sub
