VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmchat 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   165
      Width           =   7095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   3840
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3485
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Leave"
      Height          =   255
      Left            =   6240
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   5130
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   5280
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   5265
      _Version        =   393217
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmchat.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chat As String
Dim chat2 As String

Private Sub Command1_Click()
MsgBox RoomCount
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
tmp = NumChatPeople
End Sub

Private Sub Command4_Click()
rtb.TextRTF = rtb.Text & vbCrLf & "hey"
End Sub

Private Sub Form_Load()
Call StayOnTop(frmchat)
rtb.TextRTF = " You Have Entered: " & ChatRoomName
Text2.Text = NumChatPeople
Text3.Text = ChatRoomInfo
End Sub

Private Sub Timer1_Timer()
DoEvents
If FindChatRoom = 0& Then End
DoEvents
chat2 = LastChatLineWithSN
If chat2 <> chat Then
   DoEvents
   rtb.Text = rtb.Text & chat2
   rtb.SelStart = Len(rtb.Text)
End If
DoEvents
If GetAsyncKeyState(13) <> 0 Then
   DoEvents
   ChatSend (Text1)
   Text1.Text = ""
End If
DoEvents
Text2.Text = NumChatPeople
Text3.Text = ChatRoomInfo
DoEvents
chat = chat2
End Sub

Private Sub Timer2_Timer()
DoEvents
If GetAsyncKeyState(13) <> 0 Then
   DoEvents
   ChatSend (Text1)
   Text1.Text = ""
End If
DoEvents
End Sub
