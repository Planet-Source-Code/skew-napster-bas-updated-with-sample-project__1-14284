VERSION 5.00
Begin VB.Form frmrollup 
   BorderStyle     =   0  'None
   Caption         =   "Napster Toolz 1.0"
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   Picture         =   "frmrollup.frx":0000
   ScaleHeight     =   300
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
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
      Left            =   360
      TabIndex        =   3
      Top             =   30
      Width           =   315
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
      TabIndex        =   2
      Top             =   30
      Width           =   360
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
      Left            =   2280
      TabIndex        =   1
      Top             =   30
      Width           =   375
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
      TabIndex        =   0
      Top             =   30
      Width           =   450
   End
End
Attribute VB_Name = "frmrollup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call StayOnTop(frmrollup)
frmrollup.Height = 0
frmrollup.Top = frmmain.Top
frmrollup.Left = frmmain.Left
End Sub

Private Sub Form_LostFocus()
Call RestoreColor(frmrollup, vbWhite)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormMove(frmrollup)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmrollup, vbWhite)
End Sub

Private Sub Label4_Click()
frmrollup.PopupMenu frmmenu.mnufile
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmrollup, vbWhite)
Label4.ForeColor = vbBlue
End Sub

Private Sub Label6_Click()
frmrollup.PopupMenu frmmenu.mnubotz
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmrollup, vbWhite)
Label6.ForeColor = vbBlue
End Sub

Private Sub Label7_Click()
frmrollup.PopupMenu frmmenu.mnuhelp
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmrollup, vbWhite)
Label7.ForeColor = vbBlue
End Sub

Private Sub Label8_Click()
frmrollup.PopupMenu frmmenu.mnuchat
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RestoreColor(frmrollup, vbWhite)
Label8.ForeColor = vbBlue
End Sub

