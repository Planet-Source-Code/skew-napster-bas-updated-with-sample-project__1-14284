VERSION 5.00
Begin VB.Form frmadvertise 
   BorderStyle     =   0  'None
   Caption         =   "Advertise bot"
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "frmadvertise.frx":0000
   ScaleHeight     =   1785
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Ad 2"
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Ad 1"
      Top             =   480
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   1560
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
      Left            =   480
      TabIndex        =   6
      Top             =   1440
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
      Left            =   1920
      TabIndex        =   5
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Advertise Bot"
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
      Top             =   30
      Width           =   1050
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
Attribute VB_Name = "frmadvertise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call StayOnTop(frmadvertise)
End Sub

Private Sub Form_LostFocus()
Call RestoreColor(frmadvertise, vbWhite)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmadvertise)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmadvertise, vbWhite)

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmadvertise)

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmadvertise, vbWhite)
Label1.ForeColor = vbBlue
End Sub

Private Sub Label2_Click()
If Label4.Caption = "Stop" Then Label4_Click
ChatSend ("(¯`·-> Napster Toolz 1.0 <-·´¯)")
ChatSend ("(¯`·-> Advertise Bot Deactivated<-·´¯)")
ChatSend ("(¯`·-> http://go.to/GSoftware/ <-·´¯)")
Unload frmadvertise
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmadvertise, vbWhite)
Label2.ForeColor = vbBlue
End Sub

Private Sub Label3_Click()
frmadvertise.WindowState = vbMinimized
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmadvertise, vbWhite)
Label3.ForeColor = vbBlue
End Sub

Private Sub Label4_Click()
If FindChatRoom <> 0& Then
If ReplaceString(Text1, " ", "") = "" And ReplaceString(Text2, " ", "") = "" Then
   ErrMsg ("Please enter something in the text boxes.")
Else
   If Label4.Caption = "Start" Then
      Label4.Caption = "Stop"
      Timer1.Enabled = True
   Else
      Label4.Caption = "Start"
      Timer1.Enabled = False
   End If
End If
Else
   ErrMsg ("You are not in a chatroom")
End If
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmadvertise, vbWhite)
Label4.ForeColor = vbBlue
End Sub

Private Sub Label5_Click()
Label2_Click
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmadvertise, vbWhite)
Label5.ForeColor = vbBlue
End Sub

Private Sub Timer1_Timer()
ChatSend ("(¯`·-> Napster Toolz 1.0 Advertise Bot<-·´¯)")
If ReplaceString(Text1, " ", "") <> "" Then
ChatSend ("(¯`·-> " & Text1 & " <-·´¯)")
End If
If ReplaceString(Text2, " ", "") <> "" Then
ChatSend ("(¯`·-> " & Text2 & " <-·´¯)")
End If
If FindChatRoom = 0& Then
   Timer1.Enabled = False
   Label4.Caption = "Start"
End If
End Sub
