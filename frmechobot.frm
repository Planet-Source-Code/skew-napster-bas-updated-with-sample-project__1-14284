VERSION 5.00
Begin VB.Form frmechobot 
   BorderStyle     =   0  'None
   Caption         =   "Echo Bot"
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   Picture         =   "frmechobot.frx":0000
   ScaleHeight     =   1350
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "User..."
      Top             =   600
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   960
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
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   1080
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
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   1080
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
      Caption         =   "Echo Bot"
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
      Width           =   690
   End
End
Attribute VB_Name = "frmechobot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call StayOnTop(frmechobot)
End Sub

Private Sub Form_LostFocus()
Call RestoreColor(frmechobot, vbWhite)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmechobot)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmechobot, vbWhite)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmechobot)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmechobot, vbWhite)
Label1.ForeColor = vbBlue
End Sub

Private Sub Label2_Click()
If Label4.Caption = "Stop" Then Label4_Click
Unload frmechobot
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmechobot, vbWhite)
Label2.ForeColor = vbBlue
End Sub

Private Sub Label3_Click()
frmechobot.WindowState = vbMinimized
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmechobot, vbWhite)
Label3.ForeColor = vbBlue
End Sub

Private Sub Label4_Click()
If FindChatRoom <> 0& Then
If ReplaceString(Text1, " ", "") <> "" Then
   If Label4.Caption = "Start" Then
      Label4.Caption = "Stop"
      Timer1.Enabled = True
      ChatSend ("(¯`·-> Napster Toolz 1.0 <-·´¯)")
      ChatSend ("(¯`·-> Echo bot activated <-·´¯)")
      ChatSend ("(¯`·-> Now echoing " & Text1 & " <-·´¯)")
   Else
      Label4.Caption = "Start"
      Timer1.Enabled = False
      ChatSend ("(¯`·-> Napster Toolz 1.0 <-·´¯)")
      ChatSend ("(¯`·-> Echo bot deactivated <-·´¯)")
   End If
End If
Else
   ErrMsg ("you are not in a chatroom")
End If
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmechobot, vbWhite)
Label4.ForeColor = vbBlue
End Sub

Private Sub Label5_Click()
Timer1.Enabled = False
Unload frmechobot
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmechobot, vbWhite)
Label5.ForeColor = vbBlue
End Sub

Private Sub Timer1_Timer()
DoEvents
Dim lastchat As String
Dim chatlen As Integer
Static chatlen2 As Integer
If FindChatRoom = 0& Then
   Timer1.Enabled = False
   Label4.Caption = "Start"
End If
chatlen2 = Len(GetChatText())
DoEvents
lastchat = LastChatLineWithSN()
If lastchat <> "" Then
   DoEvents
   If LCase(User) = LCase(ChatLineSN(lastchat)) Then
      DoEvents
      chatlen = Len(GetChatText())
      If chatlen <> chatlen2 Then
         DoEvents
         ChatSend (ChatLineMsg(lastchat))
      End If
      chatlen2 = chatlen
   End If
End If
DoEvents
End Sub
