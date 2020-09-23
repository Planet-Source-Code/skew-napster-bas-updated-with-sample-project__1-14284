VERSION 5.00
Begin VB.Form frmimanswer 
   BorderStyle     =   0  'None
   Caption         =   "IM Answer"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "frmimanswer.frx":0000
   ScaleHeight     =   2985
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   125
      Left            =   2400
      Top             =   2400
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmimanswer.frx":29FE
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label5 
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
      Left            =   2280
      TabIndex        =   6
      Top             =   765
      Width           =   645
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
      Left            =   2280
      TabIndex        =   5
      Top             =   435
      Width           =   645
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   0
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Im Answer"
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
      Width           =   840
   End
End
Attribute VB_Name = "frmimanswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call StayOnTop(frmimanswer)
End Sub

Private Sub Form_LostFocus()
Call RestoreColor(frmimanswer, vbWhite)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmimanswer)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmimanswer, vbWhite)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmimanswer)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmimanswer, vbWhite)
Label1.ForeColor = vbBlue
End Sub

Private Sub Label2_Click()
Timer1.Enabled = False
Unload frmimanswer
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmimanswer, vbWhite)
Label2.ForeColor = vbBlue
End Sub

Private Sub Label3_Click()
frmimanswer.WindowState = vbMinimized
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmimanswer, vbWhite)
Label3.ForeColor = vbBlue
End Sub

Private Sub Label4_Click()
If Label4.Caption = "Start" Then
   Label4.Caption = "Stop"
   Timer1.Enabled = True
Else
   Label4.Caption = "Start"
   Timer1.Enabled = False
End If
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmimanswer, vbWhite)
Label4.ForeColor = vbBlue
End Sub

Private Sub Label5_Click()
ans = MsgBox("Are you sure?", vbYesNo, frmmenu.Caption)
If ans = vbYes Then
   List1.Clear
End If
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmimanswer, vbWhite)
Label5.ForeColor = vbBlue
End Sub

Private Sub Timer1_Timer()
Dim line As String
DoEvents
imwin& = FindIM()
If imwin& <> 0& Then
   DoEvents
   line$ = LastIMLineWithSN(imwin&)
   List1.AddItem line$
   Call RespondIM(imwin&, Text1)
   Call sendmessagebynum(imwin&, WM_CLOSE, 0&, 0&)
   imwin& = 0&
End If
End Sub
