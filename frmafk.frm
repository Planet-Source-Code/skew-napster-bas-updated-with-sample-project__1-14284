VERSION 5.00
Begin VB.Form frmafk 
   BorderStyle     =   0  'None
   Caption         =   "AFK Bot"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "frmafk.frx":0000
   ScaleHeight     =   2985
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2400
      Top             =   480
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Text            =   "Reason..."
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   2640
      Width           =   600
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reason:"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   360
      Width           =   840
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
      Left            =   2040
      TabIndex        =   5
      Top             =   2640
      Width           =   600
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
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   600
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
      Caption         =   "AFK Bot"
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
      Top             =   30
      Width           =   675
   End
End
Attribute VB_Name = "frmafk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
tmp = "there you go"
Debug.Print (Left(tmp, 4))
End Sub

Private Sub Form_Load()
Call StayOnTop(frmafk)
End Sub

Private Sub Form_LostFocus()
Call RestoreColor(frmafk, vbWhite)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmafk)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmafk, vbWhite)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmafk)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frm8ballbot, vbWhite)
Label1.ForeColor = vbBlue
End Sub

Private Sub Label2_Click()
If Label4.Caption = "Stop" Then Label4_Click
Unload frmafk
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frm8ballbot, vbWhite)
Label2.ForeColor = vbBlue
End Sub

Private Sub Label3_Click()
frmafk.WindowState = vbMinimized
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
   ChatSend ("(¯`·-> Napster Toolz 1.0 AFK Bot <-·´¯)")
   ChatSend ("(¯`·-> AFK because " & Text1 & " <-·´¯)")
   ChatSend ("(¯`·-> Type \msg + 'msg' <-·´¯)")
Else
   Label4.Caption = "Start"
   Timer1.Enabled = False
   ChatSend ("(¯`·-> Napster Toolz 1.0 <-·´¯)")
   ChatSend ("(¯`·-> AFK Bot Deactivated <-·´¯)")
End If
Else
   ErrMsg ("You are not in a chat room")
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

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frm8ballbot, vbWhite)
Label6.ForeColor = vbBlue
End Sub

Private Sub Label7_Click()
List1.Clear
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frm8ballbot, vbWhite)
Label7.ForeColor = vbBlue
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
msg$ = ChatLineMsg(lastchat$)
If LCase(Left$(msg$, 4)) = "\msg" Then
   DoEvents
   If lastline <> lastchat Then
      DoEvents
      ChatSend ("(¯`·-> " & sn & " Your msg has been recorded <-·´¯)")
      msg$ = Mid(msg$, 5, Len(msg) - 4)
      List1.AddItem (sn$ & " : " & msg$)
   End If
End If
lastline = lastchat
i = i + 1
If i = 1200 Then
DoEvents
ChatSend ("(¯`·-> Napster Toolz 1.0 AFK Bot <-·´¯)")
ChatSend ("(¯`·-> AFK because " & Text1 & " <-·´¯)")
ChatSend ("(¯`·-> Type \msg + 'msg' <-·´¯)")
i = 0
End If
End Sub
