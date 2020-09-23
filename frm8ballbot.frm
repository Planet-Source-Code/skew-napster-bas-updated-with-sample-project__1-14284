VERSION 5.00
Begin VB.Form frm8ballbot 
   BorderStyle     =   0  'None
   Caption         =   "8-Ball Bot"
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "frm8ballbot.frx":0000
   ScaleHeight     =   1335
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2520
      Top             =   960
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
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
      Left            =   720
      TabIndex        =   4
      Top             =   600
      Width           =   600
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8-Ball Bot"
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
      TabIndex        =   2
      Top             =   25
      Width           =   795
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
      TabIndex        =   0
      Top             =   0
      Width           =   180
   End
End
Attribute VB_Name = "frm8ballbot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call StayOnTop(frm8ballbot)
End Sub

Private Sub Form_LostFocus()
Call RestoreColor(frm8ballbot, vbWhite)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frm8ballbot)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frm8ballbot, vbWhite)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frm8ballbot)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frm8ballbot, vbWhite)
Label1.ForeColor = vbBlue
End Sub

Private Sub Label2_Click()
If Label4.Caption = "Stop" Then Label4_Click
Unload frm8ballbot
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frm8ballbot, vbWhite)
Label2.ForeColor = vbBlue
End Sub

Private Sub Label3_Click()
frm8ballbot.WindowState = vbMinimized
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frm8ballbot, vbWhite)
Label3.ForeColor = vbBlue
End Sub

Private Sub Label4_Click()
If FindChatRoom <> 0& Then
If Label4.Caption = "Start" Then
   Label4.Caption = "Stop"
   Timer1.Enabled = True
   ChatSend ("(¯`·-> Napster Toolz 1.0 <-·´¯)")
   ChatSend ("(¯`·-> 8-Ball bot activated <-·´¯)")
   ChatSend ("(¯`·-> Type \8-Ball + 'question' <-·´¯)")
Else
   Label4.Caption = "Start"
   Timer1.Enabled = False
   ChatSend ("(¯`·-> Napster Toolz 1.0 <-·´¯)")
   ChatSend ("(¯`·-> 8-Ball Bot Deactivated <-·´¯)")
End If
Else
   ErrMsg ("You are not in a chatroom")
End If
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frm8ballbot, vbWhite)
Label4.ForeColor = vbBlue
End Sub

Private Sub Label5_Click()
Label2_Click
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frm8ballbot, vbWhite)
Label5.ForeColor = vbBlue
End Sub

Private Sub Timer1_Timer()
Static i As Integer
Static lastline As String
DoEvents
If FindChatRoom = 0& Then
   Timer1.Enabled = False
   Label4.Caption = "Start"
End If
lastchat$ = ValidChatLine
sn$ = ChatLineSN(lastchat$)
chat$ = ChatLineMsg(lastchat$)
If LCase(Left$(chat$, 7)) = "\8-ball" Then
   DoEvents
   If lastline <> lastchat Then
      DoEvents
      ChatSend ("(¯`·-> " & sn$ & " 8-Ball says " & BallAnswer() & " <-·´¯)")
    End If
End If
lastline = lastchat
i = i + 1
If i = 1200 Then
DoEvents
ChatSend ("(¯`·-> Napster Toolz 1.0 <-·´¯)")
ChatSend ("(¯`·-> 8-Ball bot activated <-·´¯)")
ChatSend ("(¯`·-> Type \8-Ball + 'question' <-·´¯)")
i = 0
End If
End Sub
