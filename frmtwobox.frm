VERSION 5.00
Begin VB.Form frmtwobox 
   BorderStyle     =   0  'None
   Caption         =   "MsgBox"
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "frmtwobox.frx":0000
   ScaleHeight     =   1350
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Reason or Topic..."
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "User or Channel..."
      Top             =   480
      Width           =   1695
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
      TabIndex        =   5
      Top             =   840
      Width           =   1200
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
      Caption         =   "Command"
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
      Width           =   810
   End
End
Attribute VB_Name = "frmtwobox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call StayOnTop(frmtwobox)
frmtwobox.Caption = Label1.Caption
End Sub

Private Sub Form_LostFocus()
Call RestoreColor(frmtwobox, vbWhite)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmtwobox)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmtwobox, vbWhite)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmtwobox)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmtwobox, vbWhite)
Label1.ForeColor = vbBlue
End Sub

Private Sub Label2_Click()
Unload frmtwobox
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmtwobox, vbWhite)
Label2.ForeColor = vbBlue
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmtwobox, vbWhite)
Label3.ForeColor = vbBlue
End Sub

Private Sub Label4_Click()
If ReplaceString(Text1, " ", "") <> "" And ReplaceString(Text2, " ", "") <> "" Then
   If Label1.Caption = "Kill User" Then
      Call Command("Kill", Text1, Text2)
   End If
   If Label1.Caption = "Change Topic" Then
      Call Command("Topic", Text1, Text2)
   End If
   If Label1.Caption = "Muzzle User" Then
      Call Command("Muzzle", Text1, Text2)
   End If
   If Label1.Caption = "Nuke User" Then
      Call Command("Nuke", Text1, Text2)
   End If
   If Label1.Caption = "Ban User" Then
      Call Command("Ban", Text1, Text2)
   End If
   Call WaitForMsg
End If
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmtwobox, vbWhite)
Label4.ForeColor = vbBlue
End Sub

Private Sub Label5_Click()
Unload frmtwobox
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmtwobox, vbWhite)
Label5.ForeColor = vbBlue
End Sub
