VERSION 5.00
Begin VB.Form frmmenu 
   Caption         =   "Napster Toolz 1.0 By Skew & Jaze"
   ClientHeight    =   285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   285
   ScaleWidth      =   4170
   Begin VB.Menu mnumain 
      Caption         =   "Main"
      Begin VB.Menu mnufile 
         Caption         =   "File"
         Begin VB.Menu mnudisclaimer 
            Caption         =   "Disclaimer"
         End
         Begin VB.Menu mnuspace3 
            Caption         =   "-"
         End
         Begin VB.Menu mnurollup 
            Caption         =   "Roll Up"
         End
         Begin VB.Menu mnuminimize 
            Caption         =   "Minimize"
         End
         Begin VB.Menu mnuspace1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuexit 
            Caption         =   "Exit"
         End
      End
      Begin VB.Menu mnuchat 
         Caption         =   "Toolz"
         Begin VB.Menu mnucommands 
            Caption         =   "Commands"
            Begin VB.Menu mnuping 
               Caption         =   "Ping"
            End
            Begin VB.Menu mnuannounce 
               Caption         =   "Announce"
            End
            Begin VB.Menu mnukill 
               Caption         =   "Kill"
            End
            Begin VB.Menu mnuopsay 
               Caption         =   "Opsay"
            End
            Begin VB.Menu mnufinger 
               Caption         =   "Finger"
            End
            Begin VB.Menu mnutopic 
               Caption         =   "Topic"
            End
            Begin VB.Menu mnumuzzle 
               Caption         =   "Muzzle"
            End
            Begin VB.Menu mnunuke 
               Caption         =   "Nuke"
            End
            Begin VB.Menu mnuban 
               Caption         =   "Ban"
            End
         End
         Begin VB.Menu mnuimanswer 
            Caption         =   "IM Answer"
         End
         Begin VB.Menu mnumassim 
            Caption         =   "Mass IM"
         End
         Begin VB.Menu mnuclearchat 
            Caption         =   "Chat Clear"
            Begin VB.Menu mnuclearprivate 
               Caption         =   "Private"
            End
            Begin VB.Menu mnuclearchat2 
               Caption         =   "Chat"
            End
         End
      End
      Begin VB.Menu mnubotz 
         Caption         =   "Botz"
         Begin VB.Menu mnu8ball 
            Caption         =   "8-Ball"
         End
         Begin VB.Menu mnuadvertise 
            Caption         =   "Advertise"
         End
         Begin VB.Menu mnuafk 
            Caption         =   "AFK"
         End
         Begin VB.Menu mnuecho 
            Caption         =   "Echo"
         End
         Begin VB.Menu mnufortune 
            Caption         =   "Fortune"
         End
         Begin VB.Menu mnuluckynum 
            Caption         =   "Lucky Number"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnunumguess 
            Caption         =   "Number Guess"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuscramble 
            Caption         =   "Scramble"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnustfu 
            Caption         =   "STFU"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnusup 
            Caption         =   "Sup"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuhelp 
         Caption         =   "Help"
         Begin VB.Menu mnuhelp2 
            Caption         =   "Help"
         End
         Begin VB.Menu mnuspace2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuabout 
            Caption         =   "About"
         End
         Begin VB.Menu mnucredits 
            Caption         =   "Credits"
         End
         Begin VB.Menu mnuwebpage 
            Caption         =   "Web Page"
         End
      End
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu8ball_Click()
If IsUserOnline = False Then
   ErrMsg ("Please connect to Napster first")
Else
   frm8ballbot.Show
End If
End Sub

Private Sub mnuabout_Click()
frmabout.Show
frmmain.Hide
End Sub

Private Sub mnuadvertise_Click()
If IsUserOnline = False Then
   ErrMsg ("Please connect to Napster first")
Else
   frmadvertise.Show
End If
End Sub

Private Sub mnuafk_Click()
If IsUserOnline = False Then
   ErrMsg ("Please connect to Napster first")
Else
   frmafk.Show
End If
End Sub

Private Sub mnuannounce_Click()
If IsUserOnline = True Then
   frmonebox.Show
   frmonebox.Label1.Caption = "Announce"
   frmonebox.Text1 = "Msg..."
Else
   ErrMsg ("Please Connect to napser first")
End If
End Sub

Private Sub mnuban_Click()
If IsUserOnline = True Then
   frmtwobox.Show
   frmtwobox.Label1.Caption = "Ban User"
   frmtwobox.Text1 = "User..."
   frmtwobox.Text2 = "Reason..."
Else
   ErrMsg ("Please Connect to napser first")
End If
End Sub

Private Sub mnuclearchat2_Click()
If IsUserOnline = False Then
   ErrMsg ("Please connect to Napster first")
Else
   Call ClearChat
End If
End Sub

Private Sub mnuclearprivate_Click()
If IsUserOnline = False Then
   ErrMsg ("Please connect to Napster first")
Else
   Call ClearPrivate
End If
End Sub

Private Sub mnucredits_Click()
frmcredits.Show
frmcredits.Left = frmmain.Left
frmcredits.Top = frmmain.Top
End Sub

Private Sub mnudisclaimer_Click()
MsgBox "This program was created for educaional purposes only. Any other use then that is prohibited and I am not responsible for anything that happens. Also You get kicked and banned from Napster for using this program so becareful. Try not to use it in public chatrooms where the moderators could bust you. Have Fun!", 64, "Napster Toolz 1.0 By Skew & Jaze"
End Sub

Private Sub mnuecho_Click()
If IsUserOnline = False Then
   ErrMsg ("Please connect to Napster first")
Else
   frmechobot.Show
End If
End Sub

Private Sub mnuexit_Click()
Call UnloadAll
End Sub

Private Sub mnumindesktop_Click()
If frmmain.Visible = True Then
   frmmain.WindowState = vbMinimized
Else
   frmrollup.WindowState = vbMinimized
End If
End Sub

Private Sub mnufinger_Click()
If IsUserOnline = True Then
   frmonebox.Show
   frmonebox.Label1.Caption = "Finger"
   frmonebox.Text1 = "User..."
Else
   ErrMsg ("Please Connect to napser first")
End If
End Sub

Private Sub mnufortune_Click()
If IsUserOnline = False Then
   ErrMsg ("Please connect to Napster first")
Else
   frmfortune.Show
End If
End Sub

Private Sub mnuhelp2_Click()
ErrMsg ("Basically all you have to do to be able to use this is to connect to napster and then activate this program. All the features are pretty much self explanitory and if you mess with them you should be able to figure it out.")
End Sub

Private Sub mnuimanswer_Click()
If IsUserOnline = False Then
   Call ErrMsg("Please connect to napster first.")
Else
   frmimanswer.Show
End If
End Sub

Private Sub mnuimignore_Click()
If IsUserOnline = False Then
   Call ErrMsg("You are not connected to Napster.")
Else
   frmimignore.Show
End If
End Sub

Private Sub mnukick_Click()

End Sub

Private Sub mnukill_Click()
If IsUserOnline = True Then
   frmtwobox.Show
   frmtwobox.Label1.Caption = "Kill User"
   frmtwobox.Text1 = "User..."
   frmtwobox.Text2 = "Reason..."
Else
   ErrMsg ("Please Connect to napser first")
End If
End Sub

Private Sub mnumassim_Click()
If IsUserOnline = False Then
   Call ErrMsg("Please connect to napster first.")
Else
   frmmassim.Show
End If
End Sub

Private Sub mnuminimize_Click()
If frmmain.Visible = True Then
   frmmain.WindowState = vbMinimized
Else
   frmrollup.WindowState = vbMinimized
End If
End Sub

Private Sub mnumuzzle_Click()
If IsUserOnline = True Then
   frmtwobox.Show
   frmtwobox.Label1.Caption = "Muzzle User"
   frmtwobox.Text1 = "User..."
   frmtwobox.Text2 = "Reason..."
Else
   ErrMsg ("Please Connect to napser first")
End If
End Sub

Private Sub mnunuke_Click()
If IsUserOnline = True Then
   frmtwobox.Show
   frmtwobox.Label1.Caption = "Nuke User"
   frmtwobox.Text1 = "User..."
   frmtwobox.Text2 = "Reason..."
Else
   ErrMsg ("Please Connect to napser first")
End If
End Sub

Private Sub mnuopsay_Click()
If IsUserOnline = True Then
   frmonebox.Show
   frmonebox.Label1.Caption = "Opsay"
   frmonebox.Text1 = "Msg..."
Else
   ErrMsg ("Please Connect to napser first")
End If
End Sub

Private Sub mnupassword_Click()

End Sub

Private Sub mnuping_Click()
If IsUserOnline = True Then
   frmonebox.Show
   frmonebox.Label1.Caption = "Ping User"
   frmonebox.Text1.Text = "User..."
Else
   ErrMsg ("Please Connect to napser first")
End If
End Sub

Private Sub mnurollup_Click()
If mnurollup.Caption = "Roll Up" Then
   mnurollup.Caption = "Roll Down"
   Call RollForm(frmmain, 0, 1800)
   frmrollup.Show
   frmmain.Hide
   Call RollForm(frmrollup, 0, 300)
Else
   mnurollup.Caption = "Roll Up"
   frmmain.Top = frmrollup.Top
   frmmain.Left = frmrollup.Left
   Call RollForm(frmrollup, 0, 300)
   frmmain.Show
   Unload frmrollup
   Call RollForm(frmmain, 0, 1800)
End If
End Sub

Private Sub mnutopic_Click()
If IsUserOnline = True Then
   frmtwobox.Show
   frmtwobox.Label1.Caption = "Change Topic"
   frmtwobox.Text1 = "Channel..."
   frmtwobox.Text2 = "Topic..."
Else
   ErrMsg ("Please Connect to napser first")
End If
End Sub

Private Sub mnuwebpage_Click()
tmp = ShellExecute(frmmenu.hWnd, vbNullString, "http://go.to/GSoftware/", vbNullString, "c:\", SW_HIDE)
End Sub
