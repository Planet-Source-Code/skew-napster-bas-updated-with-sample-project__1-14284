VERSION 5.00
Begin VB.Form frmonebox 
   BorderStyle     =   0  'None
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   Picture         =   "frmonebox.frx":0000
   ScaleHeight     =   1350
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "User or Msg..."
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
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
      Width           =   1200
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      TabIndex        =   3
      Top             =   840
      Width           =   1200
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
      Caption         =   "Caption"
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
      Width           =   630
   End
End
Attribute VB_Name = "frmonebox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox (LastPrivateLine)
MsgBox InStr(LastPrivateLine, "!") <> 0
End Sub

Private Sub Form_Load()
Call StayOnTop(frmonebox)
frmonebox.Caption = Label1.Caption
End Sub

Private Sub Form_LostFocus()
Call RestoreColor(frmonebox, vbWhite)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmonebox)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmonebox, vbWhite)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmonebox)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmonebox, vbWhite)
Label1.ForeColor = vbBlue
End Sub

Private Sub Label2_Click()
Unload frmonebox
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmonebox, vbWhite)
Label2.ForeColor = vbBlue
End Sub

Private Sub Label4_Click()
If ReplaceString(Text1, " ", "") <> "" Then
   If Label1.Caption = "Ping User" Then
      Call Command("Ping", Text1, "")
   End If
   If Label1.Caption = "Announce" Then
      Call Command("Announce", Text1, "")
   End If
   If Label1.Caption = "Opsay" Then
      Call Command("Opsay", Text1, "")
   End If
   If Label1.Caption = "Finger" Then
      Call Command("Finger", Text1, "")
   End If
   Call WaitForMsg
End If
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmonebox, vbWhite)
Label4.ForeColor = vbBlue
End Sub

Private Sub Label5_Click()
Unload frmonebox
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmonebox, vbWhite)
Label5.ForeColor = vbBlue
End Sub

