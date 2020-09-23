VERSION 5.00
Begin VB.Form frmmassim 
   BorderStyle     =   0  'None
   Caption         =   "Mass IM"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "frmmassim.frx":0000
   ScaleHeight     =   2985
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Text            =   "Message..."
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Name..."
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   1320
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
      Left            =   1440
      TabIndex        =   8
      Top             =   600
      Width           =   1440
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Name"
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
      Left            =   1440
      TabIndex        =   7
      Top             =   1560
      Width           =   1440
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Remove Name"
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
      Left            =   1440
      TabIndex        =   6
      Top             =   840
      Width           =   1440
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Remove All"
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
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   1440
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
      Caption         =   "Mass IM"
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
Attribute VB_Name = "frmmassim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call StayOnTop(frmmassim)
End Sub

Private Sub Form_LostFocus()
Call RestoreColor(frmmassim, vbWhite)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmmassim)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmassim, vbWhite)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmmassim)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmassim, vbWhite)
Label1.ForeColor = vbBlue
End Sub

Private Sub Label2_Click()
Timer1.Enabled = False
Unload frmmassim
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmassim, vbWhite)
Label2.ForeColor = vbBlue
End Sub

Private Sub Label3_Click()
frmmassim.WindowState = vbMinimized
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmassim, vbWhite)
Label3.ForeColor = vbBlue
End Sub

Private Sub Label4_Click()
If List2.ListCount = 0 Then
   ErrMsg ("No users in the list")
Else
   If Label4.Caption = "Start" Then
      Label4.Caption = "Stop"
      Timer1.Enabled = True
   Else
      Label4.Caption = "Start"
      Timer1.Enabled = False
   End If
End If
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmassim, vbWhite)
Label4.ForeColor = vbBlue
End Sub

Private Sub Label5_Click()
For i = 0 To List2.ListCount - 1
   If LCase(List2.List(i)) = LCase(Text1) Then
      ErrMsg ("Name already on the list")
      Exit Sub
    End If
Next i
If Text1 <> "" Then
   List2.AddItem Text1
End If
Text1 = ""
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmassim, vbWhite)
Label5.ForeColor = vbBlue
End Sub

Private Sub Label6_Click()
If List2.SelCount = 0 Then
   ErrMsg ("Please Select A user")
Else
   List2.RemoveItem (List2.ListIndex)
End If
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmassim, vbWhite)
Label6.ForeColor = vbBlue
End Sub

Private Sub Label7_Click()
ans = MsgBox("Are You Sure?", vbYesNo, frmmenu.Caption)
If ans = vbYes Then
   List2.Clear
End If
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmmassim, vbWhite)
Label7.ForeColor = vbBlue
End Sub

Private Sub Timer1_Timer()
For i = 0 To List2.ListCount
   DoEvents
   Call InstantMessage(List2.List(i), Text2)
   Pause (1)
   DoEvents
   win& = FindWindow("#32770", "(" & List2.List(i) & ") Instant Message")
   DoEvents
   Call PostMessage(b&, WM_CLOSE, 0&, 0&)
Next i
Label4_Click
End Sub
